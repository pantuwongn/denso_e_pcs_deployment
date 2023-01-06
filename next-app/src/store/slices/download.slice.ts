import { StateCreator } from "zustand";
import { IDownloadState } from "../interface/download.interface";

export const DownloadSlice: StateCreator<IDownloadState> = (set, get) => ({
  downloadFilePath: undefined,

  setDownload(response) {
  },
})
