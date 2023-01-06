import { validateHaveElement, validateRequire } from ".";

export const UploadPCSFormValidator: Record<string, Array<(value: any) => string>> = {
  "pcs_no": [validateRequire],
  "date": [validateRequire],
  "processes": [validateHaveElement]
}
