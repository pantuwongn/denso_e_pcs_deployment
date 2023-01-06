import { PCSUploadForm } from "@/types/upload.type";

export function UploadPCSRequestFormParser(pcsForm: PCSUploadForm) {
  return {
    ...pcsForm,
    processes: pcsForm.processes.map(process => ({
      ...process,
      items: process.items.map(item => ({
        ...item,
        control_method: {
          ...item.control_method,
          method_100: item.control_method["100_method"]
        }
      }))
    }))
  }
}