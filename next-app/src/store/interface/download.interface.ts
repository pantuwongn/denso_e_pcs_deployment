import { DownloadResponse } from "@/types/upload.type";

export interface IDownloadState {
  downloadFilePath?: string
  setDownload: (response: DownloadResponse) => void
}
