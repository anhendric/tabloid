export class FileStore {
  serverUrl = "http://localhost:9090/upload-image";

  /**
   * Upload a file to the server to be saved. Returns the path of the file.
   * If the server is unreachable or the origin is file://, falls back to Base64 (Data URL).
   */
  async upload(file) {
    if (window.location.protocol !== "file:") {
      try {
        const fd = new FormData();
        fd.append("image", file);
        const res = await fetch(this.serverUrl, {
          method: "POST",
          body: fd,
        });
        if (res.ok) {
          return await res.text();
        }
      } catch (e) {
        console.warn("Could not upload to server, using local data URL", e);
      }
    }
    // Fallback for standalone demo: convert to DataURL (Base64)
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  async delete(path) {
    console.warn("cannot delete file. Not implemented");
  }

  async getFile(path) {
    const response = await fetch(path);
    return await response.blob();
  }
}
