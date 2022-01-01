import { writeFile } from 'fs';
import JSZip from 'jszip';
import { baseXl } from '../const/baseXl';
import { CalcChain } from '../lib/CalcChain';
import { ContentTypes } from '../lib/ContentTypes';
import { Image } from '../lib/Image';
import { SharedStrings } from '../lib/SharedStrings';
import { Style } from '../lib/Style';
import { errorHandler, Media, WorkBookSaveOption } from '../types';
import { DocPropsApp } from '../util/DocPropsApp';
import { XlDrawingRels } from '../util/XlDrawingRels';
import { XlWorkbook } from '../util/XlWorkbook';
import { XlWorkbookRels } from '../util/XlWorkbookRels';
import { XlWorksheetRels } from '../util/XlWorksheetRels';
import { Sheet } from './Sheet';

export class Workbook {
  id: string;
  filePath: string;
  fileName: string;

  sheets: Sheet[] = [];
  medias: Media[] = [];

  /** @ss = new SharedStrings */
  sharedStrings: SharedStrings;

  /** @cc: new ContentTypes(@) */
  contentType: ContentTypes;

  /** @da = new DocPropsApp(@) */
  docPropsApp: DocPropsApp;

  /** @wb = new XlWorkbook(@) */
  XlWorkbook: XlWorkbook;

  /** @wbre = new XlWorkbookRels(@) */
  XlWorkbookRels: XlWorkbookRels;

  /** @st = new Style(@) */
  style: Style;

  /** @cc = new CalcChain(@) */
  calcChain: CalcChain;

  /** @wsre = new XlSheetRels(@) */
  // XlSheetRels: XlWorksheetRels;

  // TODO Type
  /** @dw = new XlDrawing(@) */
  // XlDrawing: any;

  // TODO Type
  /** @dwre = new XlDrawingRels(@) */
  // XlDrawingRels: XlDrawingRels;

  constructor(filePath: string, fileName: string) {
    this.filePath = filePath;
    this.fileName = fileName;

    this.id = (Math.random() * 9999999).toFixed(0);
    // create temp folder & copy template data
    // init
    this.sheets = [];
    this.medias = [];

    this.sharedStrings = new SharedStrings();
    this.contentType = new ContentTypes(this);
    this.docPropsApp = new DocPropsApp(this);
    this.XlWorkbook = new XlWorkbook(this);
    this.XlWorkbookRels = new XlWorkbookRels(this);
    this.style = new Style(this);
    // @cc = new CalcChain(this)
    this.calcChain = new CalcChain(this);

    // @wsre = new XlSheetRels(this)
    // @dw = new XlDrawing(this)
    // @dwre = new XlDrawingRels(this)
  }

  createSheet(name: string, cols: number, rows: number) {
    const sheet = new Sheet(this, name, cols, rows);
    this.sheets.push(sheet);
    return sheet;
  }

  _addMediaFromImage(image: Image) {
    // converts image into proper media data structure
    this.medias.push({ image });
  }

  _removeMediaFromImage(image: Image) {
    // find image.
    const foundIndex = this.medias.findIndex(
      (media) => media.image.id === image.id,
    );
    // remove it if found.
    if (foundIndex !== -1) {
      this.medias.splice(foundIndex, 1);
    }
  }

  // declare
  save(cb: errorHandler): void;
  save(target: string, cb: errorHandler): void;
  save(target: string, opts: WorkBookSaveOption, cb: errorHandler): void;
  save(
    target: string | errorHandler,
    opts?: WorkBookSaveOption | errorHandler,
    cb?: errorHandler,
  ) {
    if (typeof target === 'function' && !opts && !cb) {
      cb = target;
      target = `${this.filePath}/${this.fileName}`;
      opts = {};
    }
    if (typeof opts === 'function' && !cb) {
      cb = opts;
      opts = {};
    }
    this._save(target as string, opts as WorkBookSaveOption, cb!);
  }

  private _save(
    target: string,
    opts: WorkBookSaveOption = {},
    cb: errorHandler,
  ) {
    this.generate((err: Error, zip: JSZip) => {
      // let buffer;
      let args = { type: 'nodebuffer' } as {
        type: 'nodebuffer';
      } & WorkBookSaveOption;
      if (opts.compressed) {
        args.compressed = 'DEFLATE';
      }

      // const buffer =
      zip.generateAsync(args).then((buffer: any) => {
        if (err) return cb(err);
        writeFile(target, buffer, (err) => cb(err));
      });
    });
  }

  // takes a callback function(err, zip) and returns a JSZip object on success
  generate(cb: Function) {
    const zip = new JSZip();

    for (const [key, value] of Object.entries(baseXl)) {
      zip.file(key, value);
    }

    // # 1 - build [Content_Types].xml
    zip.file('[Content_Types].xml', this.contentType.toxml());
    // # 2 - build docProps/app.xml
    zip.file('docProps/app.xml', this.docPropsApp.toxml());
    // # 3 - build xl/workbook.xml
    zip.file('xl/workbook.xml', this.XlWorkbook.toxml());
    // # 4 - build xl/sharedStrings.xml
    zip.file('xl/sharedStrings.xml', this.sharedStrings.toxml());
    // # 5 - build xl/_rels/workbook.xml.rels
    zip.file('xl/_rels/workbook.xml.rels', this.XlWorkbookRels.toxml());

    let wbMediaCounter = 1; // workbook media counter, per generation

    for (let i = 0; i < this.sheets.length; i++) {
      const sheet = this.sheets[i];

      sheet.wsRels = [];

      for (let j = 0; j < sheet.images.length; j++) {
        const image = sheet.images[j];

        const dwRels = []; // drawing media list, for this release, one per drawing
        const relId = 'rId' + (sheet.wsRels.length + 1); // same for both media to drawing OR sheet to drawing rels.

        // - build xl/media/image(1-N).xml
        const mediaFilename = [wbMediaCounter, '.', image.extension].join('');
        zip.file(`xl/media/image${mediaFilename}`, image.content, {
          base64: true,
        });
        dwRels.push({ id: relId, target: `../media/image${mediaFilename}` });

        // - build xl/drawings/drawing(1-N).xml
        const drawingFilename = `${wbMediaCounter}.xml`;
        sheet.wsRels.push({
          id: relId,
          target: `../drawings/drawing${drawingFilename}`,
        });
        zip.file(
          `xl/drawings/drawing${drawingFilename}`,
          image.toDrawingXml(relId, image),
        );

        // - build xl/drawings/_rels/drawing(1-N).xml.rels
        // TODO param type
        zip.file(
          `xl/drawings/_rels/drawing${wbMediaCounter}.xml.rels`,
          new XlDrawingRels(dwRels).toxml() as any,
        );

        wbMediaCounter++;
      }

      // - build xl/worksheets/_rels/sheet(1-N).xml.rels
      // TODO param type
      zip.file(
        `xl/worksheets/_rels/sheet${i + 1}.xml.rels`,
        new XlWorksheetRels(sheet.wsRels).toxml(),
      );

      // - build xl/worksheets/sheet(1-N).xml
      // TODO param type
      zip.file(`xl/worksheets/sheet${i + 1}.xml`, this.sheets[i].toxml());
    }

    // 7 - build xl/styles.xml
    // TODO param type
    zip.file('xl/styles.xml', this.style.toxml());

    // 8 - build xl/calcChain.xml
    if (Object.keys(this.calcChain.cache).length > 0) {
      zip.file('xl/calcChain.xml', this.calcChain.toxml());
    }
    // 9 - build xl/worksheets/sheet(1-N).xml

    cb(null, zip);
  }

  cancel() {
    // delete temp folder
    console.error('workbook.cancel() is deprecated');
  }
}
