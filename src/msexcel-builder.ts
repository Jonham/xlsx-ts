import JSZip from 'jszip'
import xmlbuilder from 'xmlbuilder'
import { Workbook } from './workbook'

const excelbuilder = {
  createWorkbook: (fpath: string, fname: string) => new Workbook(fpath, fname),
}

export default excelbuilder
