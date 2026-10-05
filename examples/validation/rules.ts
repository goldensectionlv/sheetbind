import type { JsonValue, ValidationContext, ValidationOptions } from 'sheetbind'

function decimalArgs(args: readonly JsonValue[]): boolean {
  return args.length === 1 && Number.isSafeInteger(args[0]) && (args[0] as number) >= 0 && (args[0] as number) <= 15
}
function decimalPlaces(value: unknown, [places]: readonly JsonValue[]): boolean {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return false
  }
  const [coefficient, exponent = '0'] = value.toString().toLowerCase().split('e')
  return Math.max(0, (coefficient.split('.')[1]?.length ?? 0) - Number(exponent)) <= (places as number)
}
function fieldArgs(args: readonly JsonValue[]): boolean {
  return args.length === 1 && typeof args[0] === 'string' && /^(?:\$root\.)?[A-Za-z_]\w*$/.test(args[0])
}
function maxFromField(value: unknown, [field]: readonly JsonValue[], context: ValidationContext): boolean {
  const name = field as string
  const root = name.startsWith('$root.')
  const source = root ? context.root : context.current
  const key = root ? name.slice(6) : name
  const limit = Object.hasOwn(source, key) ? source[key] : undefined
  return typeof value === 'number' && typeof limit === 'number' && value <= limit
}

export const options: ValidationOptions = {
  validationRules: {
    decimalPlaces: { validate: decimalPlaces, validateArgs: decimalArgs, message: ({ args }) => `Use at most ${args[0]} decimal places` },
    maxFromField: { validate: maxFromField, validateArgs: fieldArgs, message: 'Limit exceeded' },
  },
  validationMessages: { required: 'Enter a value' },
}
