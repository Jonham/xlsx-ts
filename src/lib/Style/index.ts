import xmlbuilder from 'xmlbuilder';
import { _Fill, _Font, _FONT_ID, _Border, _Style } from '../../types';
import { Workbook } from '../../workbook';
import { numberFormats } from './numberFormats';

export class Style {
  numberFormats: Record<number, string> = { ...numberFormats };

  // TODO
  book: Workbook;
  // font.toString cache
  cache: Record<string, _FONT_ID>;
  mfonts: _Font[];
  mfills: _Fill[];
  mbders: _Border[];
  mstyle: _Style[];
  numFmtNextId: number;

  def_font_id: number;
  def_fill_id: number;
  def_bder_id: number;
  def_align: string;
  def_valign: string;
  def_rotate: string;
  def_wrap: string;
  def_numfmt_id: number;
  def_style_id: number;

  constructor(book: Workbook) {
    this.book = book;

    this.cache = {};
    this.mfonts = []; // font style
    this.mfills = []; // fill style
    this.mbders = []; // border style
    this.mstyle = []; // cell style<ref-font,ref-fill,ref-border,align>
    this.numFmtNextId = 164;

    // this.with_default()

    this.def_font_id = this.font2id();
    this.def_fill_id = this.fill2id();
    this.def_bder_id = this.bder2id();
    this.def_align = '-';
    this.def_valign = '-';
    this.def_rotate = '-';
    this.def_wrap = '-';
    this.def_numfmt_id = 0;
    this.def_style_id = this.style2id({
      font_id: this.def_font_id,
      fill_id: this.def_fill_id,
      bder_id: this.def_bder_id,
      align: this.def_align,
      valign: this.def_valign,
      rotate: this.def_rotate,
    });
  }

  private with_default() {
    this.def_font_id = this.font2id();
    this.def_fill_id = this.fill2id();
    this.def_bder_id = this.bder2id();
    this.def_align = '-';
    this.def_valign = '-';
    this.def_rotate = '-';
    this.def_wrap = '-';
    this.def_numfmt_id = 0;
    this.def_style_id = this.style2id({
      font_id: this.def_font_id,
      fill_id: this.def_fill_id,
      bder_id: this.def_bder_id,
      align: this.def_align,
      valign: this.def_valign,
      rotate: this.def_rotate,
    });
  }

  font2id(font: Partial<_Font> = {}) {
    // Default
    const defaultFont: Partial<_Font> = {
      bold: '-',
      iter: '-',
      sz: '11',
      color: '-',
      name: 'Calibri',
      scheme: 'minor',
      family: '2',
      underline: '-',
      strike: '-',
      outline: '-',
      shadow: '-',
    };
    font = {
      ...defaultFont,
      ...((font as any) || {}),
    };

    const str =
      'font_' +
      font.bold +
      font.iter +
      font.sz +
      font.color +
      font.name +
      font.scheme +
      font.family +
      font.underline +
      font.strike +
      font.outline +
      font.shadow;
    const id = this.cache[str];
    if (id) {
      return id;
    } else {
      this.mfonts.push(font as _Font);
      this.cache[str] = this.mfonts.length;
      return this.mfonts.length;
    }
  }

  fill2id(fill: Partial<_Fill> = {}) {
    const defaultFill: Partial<_Fill> = {
      type: 'none',
      bgColor: '-',
      fgColor: '-',
    };
    fill = {
      ...defaultFill,
      ...fill,
    };

    const str = 'fill_' + fill.type + fill.bgColor + fill.fgColor;
    const id = this.cache[str];
    if (id) {
      return id;
    } else {
      this.mfills.push(fill as _Fill);
      this.cache[str] = this.mfills.length;
      return this.mfills.length;
    }
  }

  bder2id(border: Partial<_Border> = {}) {
    const defaultBorder: _Border = {
      left: '-',
      right: '-',
      top: '-',
      bottom: '-',
    };

    border = {
      ...defaultBorder,
      ...border,
    };

    const { left, right, top, bottom } = border;

    const str = JSON.stringify(['bder_', left, right, top, bottom]);
    const id = this.cache[str];
    if (id) {
      return id;
    } else {
      this.mbders.push(border as _Border);
      this.cache[str] = this.mbders.length;
      return this.mbders.length;
    }
  }

  numfmt2id(numfmt: number | string) {
    if (typeof numfmt == 'number') {
      return numfmt;
    } else if (typeof numfmt == 'string') {
      if (!numfmt) {
        throw 'Invalid format specification';
      }

      for (const [key, value] of Object.entries(this.numberFormats)) {
        if (value === numfmt) {
          return parseInt(key);
        }
      }

      // if it's not in numberFormats, we parse the string and add it the end of numberFormats
      // numfmt = numfmt
      //   .replace(/&/g, '&amp')
      //   .replace(/</g, '&lt;')
      //   .replace(/>/g, '&gt;')
      //   .replace(/"/g, '&quot;')
      this.numberFormats[++this.numFmtNextId] = numfmt;
      return this.numFmtNextId;
    }
  }

  private parseStyleKey(style: _Style) {
    const joint = [
      style.font_id,
      style.fill_id,
      style.bder_id,
      style.align,
      style.valign,
      style.wrap,
      style.rotate,
      style.numfmt_id,
    ].join('_');
    const key = `s_${joint}`;
    return key;
  }

  style2id(styleOpt: Partial<_Style> = {}) {
    const style: _Style = {
      align: this.def_align,
      valign: this.def_valign,
      rotate: this.def_rotate,
      wrap: this.def_wrap,
      font_id: this.def_font_id,
      fill_id: this.def_fill_id,
      bder_id: this.def_bder_id,
      numfmt_id: this.def_numfmt_id,

      ...styleOpt,
    };

    const key = this.parseStyleKey(style);
    const id = this.cache[key];
    if (id) {
      return id;
    } else {
      this.mstyle.push(style);
      this.cache[key] = this.mstyle.length;
      return this.mstyle.length;
    }
  }

  toxml(): string {
    const ss = xmlbuilder.create('styleSheet', {
      version: '1.0',
      encoding: 'UTF-8',
      standalone: true,
    });
    ss.att(
      'xmlns',
      'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    );
    // add all numFmts >= 164 as <numFmt numFmtId="${o.num_fmt_id}" formatCode="numFmt"/>
    const customNumFmts = [];
    for (const [key, fmt] of Object.entries(this.numberFormats)) {
      if (parseInt(key) >= 164) {
        customNumFmts.push({ numFmtId: key, formatCode: fmt });
      }
    }
    if (customNumFmts.length > 0) {
      const numFmts = ss.ele('numFmts', {
        count: customNumFmts.length,
      });
      for (const o in customNumFmts) {
        numFmts.ele('numFmt', o);
      }
    }

    const fonts = ss.ele('fonts', { count: this.mfonts.length });
    for (const o of this.mfonts) {
      const e = fonts.ele('font');
      if (o.bold !== '-') e.ele('b');
      if (o.iter !== '-') e.ele('i');
      if (o.underline !== '-') e.ele('u');
      if (o.strike !== '-') e.ele('strike');
      if (o.outline !== '-') e.ele('outline');
      if (o.shadow !== '-') e.ele('shadow');

      e.ele('sz', { val: o.sz });
      if (o.color !== '-') e.ele('color', { rgb: o.color });
      e.ele('name', { val: o.name });
      e.ele('family', { val: o.family });
      e.ele('charset', { val: '134' });
      if (o.scheme !== '-') e.ele('scheme', { val: 'minor' });
    }
    const fills = ss.ele('fills', { count: this.mfills.length + 2 });
    fills.ele('fill').ele('patternFill', { patternType: 'none' });
    fills.ele('fill').ele('patternFill', { patternType: 'gray125' });
    //<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
    for (const o of this.mfills) {
      const e = fills.ele('fill');
      const es = e.ele('patternFill', { patternType: o.type });
      if (o.fgColor !== '-') es.ele('fgColor', { rgb: o.fgColor });
      if (o.bgColor !== '-') es.ele('bgColor', { indexed: o.bgColor });
    }

    // borders
    const borders = ss.ele('borders', { count: this.mbders.length });
    for (const o of this.mbders) {
      const e = borders.ele('border');
      const dirs: (keyof _Border)[] = ['left', 'right', 'top', 'bottom'];
      dirs.forEach((borderDir) => {
        if (o[borderDir] !== '-') {
          if (typeof o.left === 'string') {
            e.ele('left', { style: o.left }).ele('color', { auto: '1' });
          } else {
            e.ele('left', { style: o.left.style }).ele('color', o.left.color);
          }
        } else {
          e.ele('borderDir');
        }
      });

      e.ele('diagonal');
    }

    // cellStyleXfs
    ss.ele('cellStyleXfs', { count: '1' })
      .ele('xf', {
        numFmtId: '0',
        fontId: '0',
        fillId: '0',
        borderId: '0',
      })
      .ele('alignment', { vertical: 'center' });

    const cs = ss.ele('cellXfs', { count: this.mstyle.length });
    for (const o of this.mstyle) {
      const e = cs.ele('xf', {
        numFmtId: o.numfmt_id || '0',
        fontId: o.font_id - 1,
        fillId: o.fill_id + 1,
        borderId: o.bder_id - 1,
        xfId: '0',
      });
      if (o.font_id !== 1) e.att('applyFont', '1');
      if (o.fill_id !== 1) e.att('applyFill', '1');
      if (o.numfmt_id !== undefined) e.att('applyNumberFormat', '1');
      if (o.bder_id !== 1) e.att('applyBorder', '1');

      if (o.align !== '-' || o.valign !== '-' || o.wrap !== '-') {
        e.att('applyAlignment', '1');
        const ea = e.ele('alignment', {
          textRotation: o.rotate === '-' ? '0' : o.rotate,
          horizontal: o.align === '-' ? 'left' : o.align,
          vertical: o.valign === '-' ? 'bottom' : o.valign,
        });
        if (o.wrap !== '-') ea.att('wrapText', '1');
      }
    }
    ss.ele('cellStyles', { count: '1' }).ele('cellStyle', {
      name: 'Normal',
      xfId: '0',
      builtinId: '0',
    });
    ss.ele('dxfs', { count: '0' });
    ss.ele('tableStyles', {
      count: '0',
      defaultTableStyle: 'TableStyleMedium9',
      defaultPivotStyle: 'PivotStyleLight16',
    });
    return ss.end({ pretty: false });
  }
}
