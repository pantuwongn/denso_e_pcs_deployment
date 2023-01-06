import create from "zustand";
import { ILayoutState } from "./interface/layout.interface";
import { LayoutSlice } from "./slices/layout.slice";

export const LayoutStore = create<ILayoutState>((...args) => ({
  ...LayoutSlice(...args)
}))
