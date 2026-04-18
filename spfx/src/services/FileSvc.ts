export default class FileSvc {
  public static getFileBuffer(file:File):Promise<ArrayBuffer> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onloadend = (e) => {
        resolve(e.target.result as ArrayBuffer);
      };

      reader.onerror = (e) => {
        reject(null);
      };

      reader.readAsArrayBuffer(file);
    }).then((value) => {
      return value as ArrayBuffer;
    }).catch((reason) => {
      console.log(reason);

      return null;
    });
  }
}