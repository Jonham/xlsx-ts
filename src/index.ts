import { Workbook } from './workbook';

const xlsx = {
  createWorkbook: (fpath: string, fname: string) => new Workbook(fpath, fname),
};

export { xlsx, Workbook };
