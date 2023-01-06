export function ValidateForm(form: any, config: Record<string, Array<(value: any) => string>>) {
  const errorList = []
  Object.entries(config).forEach(([key, value]) => {
    const nameNodes = key.split('.')
  })
}

function ValidateField(form: any, key: string, validators: Array<(value: any) => string>): string {
  // return validators.every(validator => validator(form?.[key]))
  return ''
}