export const ReadFileTextAsync = (file: File) => {
  return new Promise<string>((resolve, reject) => {
    const fileReader = new FileReader()
    fileReader.onloadend = (e: ProgressEvent<FileReader>) => {
      const data: string = e.target?.result as string
      resolve(data)
    }
    fileReader.onerror = (error) => {
      reject(error)
    }
    fileReader.readAsText(file)
  })
}

export const DownloadBlob = (fileName: string, blob: Blob) => {
  // Create blob link to download
  const url = window.URL.createObjectURL(
    new Blob([blob]),
  );
  const link = document.createElement('a');
  link.href = url;
  link.setAttribute(
    'download',
    `${fileName}.xlsx`,
  );

  // Append to html link element page
  document.body.appendChild(link);

  // Start download
  link.click();

  // Clean up and remove the link
  link.parentNode?.removeChild(link);
}