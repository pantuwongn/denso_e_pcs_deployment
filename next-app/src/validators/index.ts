export function validateRequire(value: any): string {
  return ![null, undefined].includes(value) ? '': `Field is required`
}

export function validatePcsFileType(file: File): string {
  return ['application/json'].includes(file.type) ? '' : `Except JSON file only, got file type ${file.type}`
}

export function validateHaveElement(value: any): string {
  return value?.length > 0 ? '' : `Except array to have at least one element, got ${value}`
}