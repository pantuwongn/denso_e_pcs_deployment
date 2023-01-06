import create from "zustand";
import { DownloadSlice } from "./slices/download.slice";
import { IDownloadState } from "./interface/download.interface";

export const Store = create<IDownloadState>((...args) => ({
  ...DownloadSlice(...args),
}))
