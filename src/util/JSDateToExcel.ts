export function JSDateToExcel(dt: Date) {
  return dt.valueOf() / 86400000 + 25569
}
