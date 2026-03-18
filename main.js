import { currenciesData } from "./currencies.js";
import { FileStore } from "./file_store.js";
import { geoJsonService } from "./geo_json/geo_json_service.js";

const {
  xml,
  Component,
  whenReady,
  onWillStart,
  onMounted,
  useState,
  useExternalListener,
  onError,
  markRaw,
} = owl;

const { Spreadsheet, Model } = o_spreadsheet;
const { topbarMenuRegistry, ribbonRegistry } = o_spreadsheet.registries;
const { ActionButton } = o_spreadsheet.components;
const { useStoreProvider } = o_spreadsheet.stores;

const uuidGenerator = new o_spreadsheet.helpers.UuidGenerator();
let start;

// ── Import helpers ───────────────────────────────────────────────────────────

/** Parse a CSV string into a 2-D array of strings. */
function parseCSV(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"' && text[i + 1] === '"') {
        cell += '"';
        i++;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        cell += ch;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
      } else if (ch === ",") {
        row.push(cell);
        cell = "";
      } else if (ch === "\r") {
        /* skip */
      } else if (ch === "\n") {
        row.push(cell);
        rows.push(row);
        row = [];
        cell = "";
      } else {
        cell += ch;
      }
    }
  }
  if (cell !== "" || row.length > 0) {
    row.push(cell);
    rows.push(row);
  }
  return rows;
}

/** Convert DOCX word/document.xml text to notepad HTML. */
function docxXmlToHtml(xmlText) {
  const WN = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const doc = new DOMParser().parseFromString(xmlText, "application/xml");
  const body = doc.getElementsByTagNameNS(WN, "body")[0];
  if (!body) return "<p><br></p>";

  function processRun(run) {
    let text = "";
    for (const t of run.getElementsByTagNameNS(WN, "t")) text += t.textContent;
    if (!text) return run.getElementsByTagNameNS(WN, "br")[0] ? "<br>" : "";
    text = text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    const rPr = run.getElementsByTagNameNS(WN, "rPr")[0];
    if (rPr) {
      if (rPr.getElementsByTagNameNS(WN, "strike")[0]) text = `<s>${text}</s>`;
      if (rPr.getElementsByTagNameNS(WN, "u")[0]) text = `<u>${text}</u>`;
      if (rPr.getElementsByTagNameNS(WN, "i")[0]) text = `<i>${text}</i>`;
      if (rPr.getElementsByTagNameNS(WN, "b")[0]) text = `<b>${text}</b>`;
      const styles = [];
      // Font size: <w:sz w:val="N"/> where N is in half-points (e.g. 24 = 12pt)
      const szEl = rPr.getElementsByTagNameNS(WN, "sz")[0];
      if (szEl) {
        const halfPt = parseFloat(
          szEl.getAttributeNS(WN, "val") || szEl.getAttribute("w:val") || ""
        );
        if (!isNaN(halfPt) && halfPt > 0) {
          const px = Math.round((halfPt / 2) * 1.333);
          styles.push(`font-size:${px}px`);
        }
      }
      const colorEl = rPr.getElementsByTagNameNS(WN, "color")[0];
      if (colorEl) {
        const val = colorEl.getAttributeNS(WN, "val") || colorEl.getAttribute("w:val") || "";
        if (val && val !== "auto") styles.push(`color:#${val}`);
      }
      const shdEl = rPr.getElementsByTagNameNS(WN, "shd")[0];
      if (shdEl) {
        const fill = shdEl.getAttributeNS(WN, "fill") || shdEl.getAttribute("w:fill") || "";
        if (fill && fill !== "auto" && fill !== "FFFFFF" && fill !== "ffffff") {
          styles.push(`background-color:#${fill}`);
        }
      }
      if (styles.length) text = `<span style="${styles.join(";")}">${text}</span>`;
    }
    return text;
  }

  function processParaContent(para) {
    let content = "";
    for (const child of para.childNodes) {
      if (child.nodeType !== 1) continue;
      if (child.localName === "r") content += processRun(child);
      else if (child.localName === "hyperlink") {
        for (const rc of child.childNodes) {
          if (rc.nodeType === 1 && rc.localName === "r") content += processRun(rc);
        }
      }
    }
    return content;
  }

  let html = "";
  for (const node of body.childNodes) {
    if (node.nodeType !== 1 || node.localName !== "p") continue;
    const pPr = node.getElementsByTagNameNS(WN, "pPr")[0];
    const pStyle = pPr ? pPr.getElementsByTagNameNS(WN, "pStyle")[0] : null;
    const styleVal =
      (pStyle && (pStyle.getAttributeNS(WN, "val") || pStyle.getAttribute("w:val"))) || "";
    const content = processParaContent(node) || "<br>";
    const hm = styleVal.match(/^Heading(\d)$/);
    html += hm ? `<h${hm[1]}>${content}</h${hm[1]}>` : `<p>${content}</p>`;
  }
  return html || "<p><br></p>";
}

/** Very simple Markdown → HTML for notepad import. */
function markdownToHtml(md) {
  function inlineMd(text) {
    return text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\*\*\*(.*?)\*\*\*/g, "<b><i>$1</i></b>")
      .replace(/\*\*(.*?)\*\*/g, "<b>$1</b>")
      .replace(/__(.*?)__/g, "<b>$1</b>")
      .replace(/\*(.*?)\*/g, "<i>$1</i>")
      .replace(/_(.*?)_/g, "<i>$1</i>")
      .replace(/~~(.*?)~~/g, "<s>$1</s>")
      .replace(/`(.*?)`/g, "<code>$1</code>")
      .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>');
  }
  const lines = md.split("\n");
  let html = "";
  let listType = "";
  const closeList = () => {
    if (listType) {
      html += `</${listType}>`;
      listType = "";
    }
  };
  for (const line of lines) {
    const hm = line.match(/^(#{1,6})\s+(.*)/);
    if (hm) {
      closeList();
      html += `<h${hm[1].length}>${inlineMd(hm[2])}</h${hm[1].length}>`;
      continue;
    }
    const ulm = line.match(/^[-*+]\s+(.*)/);
    if (ulm) {
      if (listType !== "ul") {
        closeList();
        html += "<ul>";
        listType = "ul";
      }
      html += `<li>${inlineMd(ulm[1])}</li>`;
      continue;
    }
    const olm = line.match(/^\d+\.\s+(.*)/);
    if (olm) {
      if (listType !== "ol") {
        closeList();
        html += "<ol>";
        listType = "ol";
      }
      html += `<li>${inlineMd(olm[1])}</li>`;
      continue;
    }
    closeList();
    html += line.trim() ? `<p>${inlineMd(line)}</p>` : "<p><br></p>";
  }
  closeList();
  return html || "<p><br></p>";
}

/** Plain text → HTML for notepad import. */
function textToHtml(text) {
  return (
    text
      .split("\n")
      .map((line) => {
        const esc = line.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
        return esc ? `<p>${esc}</p>` : "<p><br></p>";
      })
      .join("") || "<p><br></p>"
  );
}
/**
 * Welcome screen shown when the app starts without a project loaded.
 * Allows the user to choose their initial activity.
 */
class WelcomeScreen extends Component {
  static props = { onChoice: Function };
  static template = xml/* xml */ `
    <div class="o-welcome-screen d-flex flex-column align-items-center justify-content-center w-100 h-100">
      <div class="o-welcome-container shadow-lg rounded-4 p-5 text-center">
        <h1 class="display-4 fw-bold mb-2">Welcome to Tabloid</h1>
        <p class="text-muted mb-5">Select an activity to get started</p>
        
        <div class="d-flex gap-4 justify-content-center">
          <div class="o-welcome-card p-4 rounded-3 shadow-sm d-flex flex-column align-items-center" t-on-click="() => props.onChoice('spreadsheet')">
            <div class="o-card-icon mb-3">📊</div>
            <h3 class="h5 fw-bold">Spreadsheet</h3>
            <p class="small text-muted text-center mt-2">Create a new grid with formulas and charts</p>
          </div>
          
          <div class="o-welcome-card p-4 rounded-3 shadow-sm d-flex flex-column align-items-center" t-on-click="() => props.onChoice('whiteboard')">
            <div class="o-card-icon mb-3">🎨</div>
            <h3 class="h5 fw-bold">Whiteboard</h3>
            <p class="small text-muted text-center mt-2">Start a free-form drawing and sketching area</p>
          </div>
          
          <div class="o-welcome-card p-4 rounded-3 shadow-sm d-flex flex-column align-items-center" t-on-click="() => props.onChoice('notepad')">
            <div class="o-card-icon mb-3">📝</div>
            <h3 class="h5 fw-bold">Notepad</h3>
            <p class="small text-muted text-center mt-2">Write documents with embedded live data</p>
          </div>

          <div class="o-welcome-card p-4 rounded-3 shadow-sm d-flex flex-column align-items-center" t-on-click="() => props.onChoice('code_editor')">
            <div class="o-card-icon mb-3">💻</div>
            <h3 class="h5 fw-bold">Code Editor</h3>
            <p class="small text-muted text-center mt-2">Write custom code and scripts</p>
          </div>
        </div>
        
        <div class="mt-5 text-muted small">
          Or <a href="#" class="text-primary text-decoration-none" t-on-click="() => props.onChoice('open_project')">open an existing project</a>
        </div>
      </div>
      
      <style>
        .o-welcome-screen {
          font-family: 'Inter', system-ui, -apple-system, sans-serif;
        }
        .o-welcome-container {
          background: rgba(255, 255, 255, 0.95);
          backdrop-filter: blur(20px);
          max-width: 900px;
          border: 1px solid rgba(255, 255, 255, 0.5);
          box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25);
        }
        .o-welcome-card {
          width: 240px;
          background: white;
          cursor: pointer;
          transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
          border: 1px solid #edf2f7;
          border-bottom: 4px solid #edf2f7;
        }
        .o-welcome-card:hover {
          transform: translateY(-12px);
          box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04) !important;
          border-color: #4c51bf;
          border-bottom-color: #4c51bf;
        }
        .o-welcome-card:hover .o-card-icon {
          transform: scale(1.2);
        }
        .o-card-icon {
          font-size: 56px;
          transition: transform 0.3s ease;
        }
        .o-welcome-card h3 {
          color: #2d3748;
        }
      </style>
    </div>
  `;
}

class Demo extends Component {
  setup() {
    this.state = useState({
      key: 0,
      colorScheme: "light",
      showWelcomeScreen: true,
      isTerminalVisible: false,
    });
    this.stateUpdateMessages = [];
    this.client = {
      id: uuidGenerator.uuidv4(),
      name: "Local",
    };
    this.fileStore = new FileStore();

    this._openProjectFile = () => {
      const input = document.createElement("input");
      input.setAttribute("type", "file");
      input.setAttribute("style", "display: none");
      input.setAttribute("accept", ".json");
      document.body.appendChild(input);
      input.addEventListener("change", async () => {
        if (input.files.length <= 0) {
          input.remove();
          return;
        }
        const file = input.files[0];
        try {
          const content = await file.text();
          const data = JSON.parse(content);
          stores.resetStores();
          await this.initiateConnection(data);
          this.state.showWelcomeScreen = false;
          this.state.key = this.state.key + 1;
        } catch (error) {
          console.error(error);
          this.notifyUser({ text: "An error occurred while opening the file.", type: "warning" });
        }
        input.remove();
      });
      input.click();
    };

    this.onChoice = async (choice) => {
      if (choice === "open_project") {
        this._openProjectFile();
        return;
      }

      let initialData = {};
      if (choice === "whiteboard") {
        initialData = {
          sheets: [{ id: "sheet1", name: "Whiteboard", isWhiteboard: true }],
        };
      } else if (choice === "notepad") {
        initialData = {
          sheets: [{ id: "sheet1", name: "Notepad", isNotepad: true }],
        };
      } else if (choice === "code_editor") {
        initialData = {
          sheets: [{ id: "sheet1", name: "Code Editor", isCodeEditor: true }],
        };
      } else {
        // Default spreadsheet
        initialData = {
          sheets: [{ id: "sheet1", name: "Sheet1" }],
        };
      }

      stores.resetStores();
      await this.initiateConnection(initialData);
      this.state.showWelcomeScreen = false;
      this.state.key = this.state.key + 1;
    };

    topbarMenuRegistry.addChild("new_project", ["file"], {
      name: "New project",
      sequence: 10,
      execute: async () => {
        this.state.showWelcomeScreen = true;
      },
      icon: "o-spreadsheet-Icon.NEW_PROJECT",
      isEnabledOnLockedSheet: true,
    });

    topbarMenuRegistry.addChild("dark_mode", ["view"], {
      name: "Toggle dark mode",
      sequence: 12.5,
      isReadonlyAllowed: true,
      execute: () =>
        (this.state.colorScheme = this.state.colorScheme === "dark" ? "light" : "dark"),
      icon: "o-spreadsheet-Icon.DARK_MODE",
      isEnabledOnLockedSheet: true,
    });

    topbarMenuRegistry.addChild("toggle_terminal", ["view"], {
      name: "Terminal",
      sequence: 15,
      execute: () => {
        this.state.isTerminalVisible = !this.state.isTerminalVisible;
        if (window.setTerminalVisible) {
          window.setTerminalVisible(this.state.isTerminalVisible);
        }
      },
      isActive: () => this.state.isTerminalVisible,
      icon: "o-spreadsheet-Icon.TERMINAL",
      isEnabledOnLockedSheet: true,
    });

    ribbonRegistry.addGroup("view", "terminal", "Terminal", 100);
    ribbonRegistry.addItem("view", "terminal", {
      id: "toggle_terminal_ribbon",
      component: ActionButton,
      props: {
        action: {
          name: "Terminal",
          icon: "o-spreadsheet-Icon.TERMINAL",
          execute: () => {
            this.state.isTerminalVisible = !this.state.isTerminalVisible;
            if (window.setTerminalVisible) {
              window.setTerminalVisible(this.state.isTerminalVisible);
            }
          },
          isActive: () => this.state.isTerminalVisible,
        },
        class: "o-hoverable-button o-toolbar-button",
      },
      sequence: 10,
    });

    topbarMenuRegistry.addChild("open_project", ["file"], {
      name: "Open project",
      sequence: 30,
      execute: async () => {
        this._openProjectFile();
      },
      icon: "o-spreadsheet-Icon.IMPORT_XLSX",
      isEnabledOnLockedSheet: true,
    });

    topbarMenuRegistry.addChild("import_file", ["file"], {
      name: "Import file",
      sequence: 40,
      isEnabledOnLockedSheet: true,
      separator: true,
      execute: async (env) => {
        // Utiliser le dialog déjà présent dans le DOM
        let dialog = document.getElementById("import-dialog");
        if (!dialog) {
          alert("Le dialogue d'import n'est pas présent dans la page.");
          return;
        }

        // Préparer l'input file
        const input = document.createElement("input");
        input.setAttribute("type", "file");
        input.setAttribute("style", "display: none");
        input.setAttribute(
          "accept",
          ".xlsx,.csv,.docx,.md,.markdown,.txt,.js,.ts,.py,.json,.xml,.html,.css,.yaml,.yml,.toml,.sh,.bash,.sql,.rb,.go,.rs,.java,.c,.cpp,.h,.hpp"
        );
        document.body.appendChild(input);

        // Ouvre le dialogue
        dialog.returnValue = "cancel";
        dialog.showModal();

        // Gestion de la validation
        dialog.querySelector("#import-confirm-btn").onclick = (e) => {
          e.preventDefault();
          dialog.close("ok");
        };

        dialog.querySelector("button[value='cancel']").onclick = (e) => {
          e.preventDefault();
          dialog.close("cancel");
        };

        dialog.onclose = async () => {
          if (dialog.returnValue !== "ok") {
            input.remove();
            return;
          }
          // Récupérer le choix
          const mode = dialog.querySelector("input[name='importMode']:checked").value;
          dialog.close();
          input.click();

          input.onchange = async () => {
            if (input.files.length <= 0) {
              input.remove();
              return;
            }
            const file = input.files[0];
            const ext = (file.name.match(/\.([^.]+)$/) || [])[1]?.toLowerCase() || "";
            try {
              if (mode === "new") {
                // Nouveau projet : reset + import
                let initialData = {};
                if (ext === "xlsx") {
                  const myjszip = new JSZip();
                  const zip = await myjszip.loadAsync(file);
                  const files = Object.keys(zip.files);
                  const images = [];
                  const contents = await Promise.all(
                    files.map((f) => {
                      if (f.includes("media/image")) {
                        images.push(f);
                        return zip.files[f].async("blob");
                      }
                      return zip.files[f].async("text");
                    })
                  );
                  const inputFiles = {};
                  for (let i = 0; i < contents.length; i++) inputFiles[files[i]] = contents[i];
                  for (let i = 0; i < images.length; i++) {
                    const imgFile = new File([inputFiles[images[i]]], images[i].split("/").at(-1));
                    inputFiles[images[i]] = { imageSrc: await this.fileStore.upload(imgFile) };
                  }
                  const tempModel = new Model(inputFiles, {
                    external: {
                      loadCurrencies: async () => currenciesData,
                      fileStore: this.fileStore,
                      geoJsonService: geoJsonService,
                    },
                    mode: "normal",
                  });
                  initialData = tempModel.exportData();
                } else if (ext === "csv") {
                  const text = await file.text();
                  const sheetName =
                    file.name.replace(/\.csv$/i, "").replace(/[^a-zA-Z0-9_\- ]/g, "_") || "Import";
                  const rows = parseCSV(text);
                  initialData = {
                    sheets: [{ id: "sheet1", name: sheetName }],
                  };
                  // Les cellules seront ajoutées après création du modèle
                } else if (["docx", "md", "markdown", "txt"].includes(ext)) {
                  let html = "";
                  let sheetName =
                    file.name.replace(/\.[^.]+$/, "").replace(/[^a-zA-Z0-9_\- ]/g, "_") || "Import";
                  if (ext === "docx") {
                    const myjszip = new JSZip();
                    const zip = await myjszip.loadAsync(file);
                    const xmlText = await zip.files["word/document.xml"].async("text");
                    html = docxXmlToHtml(xmlText);
                  } else if (["md", "markdown"].includes(ext)) {
                    const text = await file.text();
                    html = markdownToHtml(text);
                  } else if (ext === "txt") {
                    const text = await file.text();
                    html = textToHtml(text);
                  }
                  initialData = {
                    sheets: [{ id: "sheet1", name: sheetName, isNotepad: true }],
                  };
                  // Le contenu sera ajouté après création du modèle
                } else {
                  // Code editor ou autre
                  const text = await file.text();
                  const sheetName =
                    file.name.replace(/\.[^.]+$/, "").replace(/[^a-zA-Z0-9_\- ]/g, "_") || "Import";
                  initialData = {
                    sheets: [
                      { id: "sheet1", name: sheetName, isCodeEditor: true, codeExtension: ext },
                    ],
                  };
                  // Le contenu sera ajouté après création du modèle
                }
                stores.resetStores();
                await this.initiateConnection(initialData);
                this.state.showWelcomeScreen = false;
                this.state.key = this.state.key + 1;
                // Ajout du contenu pour CSV, Notepad, CodeEditor
                if (ext === "csv") {
                  const rows = parseCSV(await file.text());
                  const sheetId = this.model.getters.getSheetIds()[0];
                  for (let r = 0; r < rows.length; r++) {
                    for (let c = 0; c < rows[r].length; c++) {
                      const content = rows[r][c];
                      if (content !== undefined && content !== null && content !== "") {
                        this.model.dispatch("UPDATE_CELL", { sheetId, col: c, row: r, content });
                      }
                    }
                  }
                } else if (["docx", "md", "markdown", "txt"].includes(ext)) {
                  let html = "";
                  if (ext === "docx") {
                    const myjszip = new JSZip();
                    const zip = await myjszip.loadAsync(file);
                    const xmlText = await zip.files["word/document.xml"].async("text");
                    html = docxXmlToHtml(xmlText);
                  } else if (["md", "markdown"].includes(ext)) {
                    html = markdownToHtml(await file.text());
                  } else if (ext === "txt") {
                    html = textToHtml(await file.text());
                  }
                  const sheetId = this.model.getters.getSheetIds()[0];
                  this.model.dispatch("UPDATE_NOTEPAD_CONTENT", { sheetId, content: html });
                } else if (!["xlsx", "csv", "docx", "md", "markdown", "txt"].includes(ext)) {
                  const text = await file.text();
                  const sheetId = this.model.getters.getSheetIds()[0];
                  this.model.dispatch("UPDATE_CODE_EDITOR_CONTENT", { sheetId, content: text });
                }
                input.remove();
                return;
              }
              if (ext === "xlsx") {
                const myjszip = new JSZip();
                const zip = await myjszip.loadAsync(file);
                const files = Object.keys(zip.files);
                const images = [];
                const contents = await Promise.all(
                  files.map((f) => {
                    if (f.includes("media/image")) {
                      images.push(f);
                      return zip.files[f].async("blob");
                    }
                    return zip.files[f].async("text");
                  })
                );
                const inputFiles = {};
                for (let i = 0; i < contents.length; i++) inputFiles[files[i]] = contents[i];
                for (let i = 0; i < images.length; i++) {
                  const imgFile = new File([inputFiles[images[i]]], images[i].split("/").at(-1));
                  inputFiles[images[i]] = { imageSrc: await this.fileStore.upload(imgFile) };
                }

                // Parse XLSX into WorkbookData via a temporary Model
                const tempModel = new Model(inputFiles, {
                  external: {
                    loadCurrencies: async () => currenciesData,
                    fileStore: this.fileStore,
                    geoJsonService: geoJsonService,
                  },
                  mode: "normal",
                });
                const importedData = tempModel.exportData();

                // Export current project data and merge
                const currentData = env.model.exportData();

                // Compute offsets so imported ids don't collide with existing ones
                const maxId = (obj) => Math.max(0, ...Object.keys(obj || {}).map(Number));
                const styleOffset = maxId(currentData.styles) + 1;
                const formatOffset = maxId(currentData.formats) + 1;
                const borderOffset = maxId(currentData.borders) + 1;

                // Copy imported styles/formats/borders with shifted ids
                for (const [id, v] of Object.entries(importedData.styles || {})) {
                  currentData.styles[Number(id) + styleOffset] = v;
                }
                for (const [id, v] of Object.entries(importedData.formats || {})) {
                  currentData.formats[Number(id) + formatOffset] = v;
                }
                for (const [id, v] of Object.entries(importedData.borders || {})) {
                  currentData.borders[Number(id) + borderOffset] = v;
                }

                // Helper to shift all ids in a {zone: id} map
                const shiftIds = (map, offset) => {
                  if (!map) return map;
                  const out = {};
                  for (const [zone, id] of Object.entries(map)) out[zone] = Number(id) + offset;
                  return out;
                };

                // Process imported sheets: remap ids, deduplicate names
                const existingNames = currentData.sheets.map((s) => s.name);
                for (const sheet of importedData.sheets) {
                  sheet.id = env.model.uuidGenerator.smallUuid();
                  let name = sheet.name,
                    c = 1;
                  while (existingNames.includes(name)) name = `${sheet.name} (${c++})`;
                  sheet.name = name;
                  existingNames.push(name);

                  sheet.styles = shiftIds(sheet.styles, styleOffset);
                  sheet.formats = shiftIds(sheet.formats, formatOffset);
                  sheet.borders = shiftIds(sheet.borders, borderOffset);
                  if (sheet.figures) {
                    for (const fig of sheet.figures) fig.id = env.model.uuidGenerator.smallUuid();
                  }
                }

                currentData.sheets.push(...importedData.sheets);
                if (importedData.customTableStyles) {
                  currentData.customTableStyles = {
                    ...(currentData.customTableStyles || {}),
                    ...importedData.customTableStyles,
                  };
                }

                // Reload model with merged data
                stores.resetStores();
                await this.initiateConnection(currentData);
                this.state.key = this.state.key + 1;
                input.remove();
                return;
              }

              const activeSheetId = env.model.getters.getActiveSheetId();
              const position = env.model.getters.getSheetIds().length;
              const sheetId = env.model.uuidGenerator.smallUuid();

              // CSV files: open as a spreadsheet sheet
              if (ext === "csv") {
                const text = await file.text();
                const sheetName =
                  file.name.replace(/\.csv$/i, "").replace(/[^a-zA-Z0-9_\- ]/g, "_") || "Import";
                const rows = parseCSV(text);
                env.model.dispatch("CREATE_SHEET", { sheetId, position, name: sheetName });
                for (let r = 0; r < rows.length; r++) {
                  for (let c = 0; c < rows[r].length; c++) {
                    const content = rows[r][c];
                    if (content !== undefined && content !== null && content !== "") {
                      env.model.dispatch("UPDATE_CELL", { sheetId, col: c, row: r, content });
                    }
                  }
                }
                env.model.dispatch("ACTIVATE_SHEET", {
                  sheetIdFrom: activeSheetId,
                  sheetIdTo: sheetId,
                });
              }
              // Notepad files: .docx, .md, .txt
              else if (ext === "docx") {
                const sheetName =
                  file.name.replace(/\.docx$/i, "").replace(/[^a-zA-Z0-9_\- ]/g, "_") || "Import";
                const myjszip = new JSZip();
                const zip = await myjszip.loadAsync(file);
                const xmlText = await zip.files["word/document.xml"].async("text");
                const html = docxXmlToHtml(xmlText);
                env.model.dispatch("CREATE_SHEET", {
                  sheetId,
                  position,
                  name: sheetName,
                  isNotepad: true,
                });
                env.model.dispatch("UPDATE_NOTEPAD_CONTENT", { sheetId, content: html });
                env.model.dispatch("ACTIVATE_SHEET", {
                  sheetIdFrom: activeSheetId,
                  sheetIdTo: sheetId,
                });
              } else if (["md", "markdown"].includes(ext)) {
                const text = await file.text();
                const sheetName =
                  file.name.replace(/\.(md|markdown)$/i, "").replace(/[^a-zA-Z0-9_\- ]/g, "_") ||
                  "Import";
                const html = markdownToHtml(text);
                env.model.dispatch("CREATE_SHEET", {
                  sheetId,
                  position,
                  name: sheetName,
                  isNotepad: true,
                });
                env.model.dispatch("UPDATE_NOTEPAD_CONTENT", { sheetId, content: html });
                env.model.dispatch("ACTIVATE_SHEET", {
                  sheetIdFrom: activeSheetId,
                  sheetIdTo: sheetId,
                });
              } else if (ext === "txt") {
                const text = await file.text();
                const sheetName =
                  file.name.replace(/\.txt$/i, "").replace(/[^a-zA-Z0-9_\- ]/g, "_") || "Import";
                const html = textToHtml(text);
                env.model.dispatch("CREATE_SHEET", {
                  sheetId,
                  position,
                  name: sheetName,
                  isNotepad: true,
                });
                env.model.dispatch("UPDATE_NOTEPAD_CONTENT", { sheetId, content: html });
                env.model.dispatch("ACTIVATE_SHEET", {
                  sheetIdFrom: activeSheetId,
                  sheetIdTo: sheetId,
                });
              }
              // Everything else: open as code editor
              else {
                const text = await file.text();
                const sheetName =
                  file.name.replace(/\.[^.]+$/, "").replace(/[^a-zA-Z0-9_\- ]/g, "_") || "Import";
                env.model.dispatch("CREATE_SHEET", {
                  sheetId,
                  position,
                  name: sheetName,
                  isCodeEditor: true,
                  codeExtension: ext,
                });
                env.model.dispatch("UPDATE_CODE_EDITOR_CONTENT", { sheetId, content: text });
                env.model.dispatch("ACTIVATE_SHEET", {
                  sheetIdFrom: activeSheetId,
                  sheetIdTo: sheetId,
                });
              }
            } catch (error) {
              console.error(error);
              this.notifyUser({
                text: "An error occurred while importing the file.",
                type: "warning",
              });
            }
            input.remove();
          };
        };
      },
    });

    const stores = useStoreProvider();

    useExternalListener(window, "unhandledrejection", this.notifyError.bind(this));
    useExternalListener(window, "error", (ev) => {
      console.error("Global error caught: ", ev.error || ev.message);
      this.notifyError();
    });

    onMounted(() => console.log("Mounted: ", Date.now() - start));
    onError((error) => {
      console.error(error.cause || error);
      this.notifyError();
    });
  }

  notifyError() {
    this.notifyUser({
      text: "An unexpected error occurred. Open the developer console for details.",
      sticky: true,
      type: "warning",
    });
  }

  async initiateConnection(data = undefined) {
    this.stateUpdateMessages = [];
    this.createModel(data || {});
  }

  createModel(data) {
    this.model = new Model(
      data,
      {
        external: {
          loadCurrencies: async () => currenciesData,
          fileStore: this.fileStore,
          geoJsonService: geoJsonService,
        },
        custom: {},
        transportService: undefined,
        client: this.client,
        mode: "normal",
      },
      this.stateUpdateMessages
    );
    markRaw(this.model);
    window.spreadsheetModel = this.model;
    this.activateFirstSheet();
  }

  activateFirstSheet() {
    const sheetId = this.model.getters.getActiveSheetId();
    const firstSheetId = this.model.getters.getSheetIds()[0];
    if (firstSheetId !== sheetId) {
      this.model.dispatch("ACTIVATE_SHEET", { sheetIdFrom: sheetId, sheetIdTo: firstSheetId });
    }
  }

  notifyUser(notification) {
    const div = document.createElement("div");
    const text = document.createTextNode(notification.text);
    div.appendChild(text);
    div.classList.add(
      "o-test-notification",
      "bg-white",
      "p-3",
      "shadow",
      "rounded",
      notification.type
    );
    const element = document.querySelector(".o-spreadsheet") || document.body; // if we crash on launch, the spreadsheet is not mounted yet
    div.onclick = () => {
      element.removeChild(div);
    };
    element.appendChild(div);

    if (!notification.sticky) {
      setTimeout(() => {
        if (document.body.contains(div)) {
          element.removeChild(div);
        }
      }, 5000);
    }
  }
}

Demo.template = xml/* xml */ `
  <div class="w-100 h-100">
    <t-if t-if="state.showWelcomeScreen">
      <WelcomeScreen onChoice="onChoice"/>
    </t-if>
    <t-else>
      <t-if t-if="model">
        <Spreadsheet model="model" t-key="state.key" notifyUser="notifyUser" colorScheme="state.colorScheme"/>
      </t-if>
    </t-else>
  </div>
`;
Demo.components = { Spreadsheet, WelcomeScreen };
Demo.props = {};

// Setup code
async function setup() {
  const templates = await (await fetch("lib/o_spreadsheet.xml")).text();
  start = Date.now();

  const rootApp = new owl.App(Demo, { dev: true, warnIfNoStaticProps: true });
  rootApp.addTemplates(templates);
  rootApp.mount(document.body);
}

whenReady(setup);
