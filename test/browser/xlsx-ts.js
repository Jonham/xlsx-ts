(function (global, factory) {
  typeof exports === 'object' && typeof module !== 'undefined' ? factory(exports, require('fs'), require('jszip'), require('xmlbuilder')) :
  typeof define === 'function' && define.amd ? define(['exports', 'fs', 'jszip', 'xmlbuilder'], factory) :
  (global = typeof globalThis !== 'undefined' ? globalThis : global || self, factory(global.XLSXts = {}, global.fs, global.JSZip, global.xmlbuilder));
})(this, (function (exports, fs, JSZip, xmlbuilder) { 'use strict';

  function _interopDefaultLegacy (e) { return e && typeof e === 'object' && 'default' in e ? e : { 'default': e }; }

  var JSZip__default = /*#__PURE__*/_interopDefaultLegacy(JSZip);
  var xmlbuilder__default = /*#__PURE__*/_interopDefaultLegacy(xmlbuilder);

  var baseXl = {
      '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
      'docProps/core.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Administrator</dc:creator><cp:lastModifiedBy></cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2006-09-13T11:21:51Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2006-09-13T11:21:55Z</dcterms:modified></cp:coreProperties>',
      'xl/theme/theme1.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office 主题"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>',
      'xl/styles.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><color theme="1"/><name val="宋体"/><family val="2"/><charset val="134"/><scheme val="minor"/></font><font><sz val="9"/><name val="宋体"/><family val="2"/><charset val="134"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>',
  };

  function i2a(column) {
      var temp;
      var letter = '';
      while (column > 0) {
          temp = (column - 1) % 26;
          letter = String.fromCharCode(temp + 65) + letter;
          column = (column - temp - 1) / 26;
      }
      return letter;
  }

  var CalcChain = (function () {
      function CalcChain(book) {
          this.cache = {};
          this.book = book;
      }
      CalcChain.prototype.add_ref = function (idx, col, row) {
          var num = idx + 1;
          if (!this.cache.hasOwnProperty(num))
              this.cache[num] = [];
          this.cache[num].push(i2a(col) + row);
      };
      CalcChain.prototype.toxml = function () {
          var cc = xmlbuilder__default["default"].create('calcChain', {
              version: '1.0',
              encoding: 'UTF-8',
              standalone: true,
          });
          cc.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
          for (var _i = 0, _a = Object.entries(this.cache); _i < _a.length; _i++) {
              var _b = _a[_i], key = _b[0], val = _b[1];
              for (var _c = 0, val_1 = val; _c < val_1.length; _c++) {
                  var el = val_1[_c];
                  cc.ele('c', { r: '' + el, i: '' + key });
              }
          }
          return cc.end({ pretty: false });
      };
      return CalcChain;
  }());

  var ContentTypes = (function () {
      function ContentTypes(book) {
          this.book = book;
      }
      ContentTypes.prototype._getKnowImageTypes = function () {
          var imagesToAdd = [];
          var imagesToAddDistinct = {};
          for (var _i = 0, _a = this.book.sheets; _i < _a.length; _i++) {
              var sheet = _a[_i];
              for (var _b = 0, _c = sheet.images; _b < _c.length; _b++) {
                  var image = _c[_b];
                  if (!imagesToAddDistinct[image.extension]) {
                      imagesToAdd.push({
                          Extension: image.extension,
                          ContentType: image.contentType,
                      });
                  }
              }
          }
          return imagesToAdd;
      };
      ContentTypes.prototype.toxml = function () {
          var types = xmlbuilder__default["default"].create('Types', {
              version: '1.0',
              encoding: 'UTF-8',
              standalone: true,
          });
          types.att('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');
          types.ele('Default', { Extension: 'png', ContentType: 'image/png' });
          types.ele('Default', { Extension: 'svg', ContentType: 'image/svg+xml' });
          types.ele('Default', {
              Extension: 'rels',
              ContentType: 'application/vnd.openxmlformats-package.relationships+xml',
          });
          types.ele('Default', { Extension: 'xml', ContentType: 'application/xml' });
          types.ele('Override', {
              PartName: '/xl/workbook.xml',
              ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
          });
          for (var i = 1; i <= this.book.sheets.length; i++) {
              types.ele('Override', {
                  PartName: '/xl/worksheets/sheet' + i + '.xml',
                  ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
              });
          }
          types.ele('Override', {
              PartName: '/xl/theme/theme1.xml',
              ContentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
          });
          types.ele('Override', {
              PartName: '/xl/calcChain.xml',
              ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml',
          });
          types.ele('Override', {
              PartName: '/xl/styles.xml',
              ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
          });
          for (var _i = 0, _a = this.book.sheets; _i < _a.length; _i++) {
              var sheet = _a[_i];
              for (var _b = 0, _c = sheet.images; _b < _c.length; _b++) {
                  var image = _c[_b];
                  types.ele('Override', {
                      PartName: '/xl/drawings/drawing' + image.id + '.xml',
                      ContentType: 'application/vnd.openxmlformats-officedocument.drawing+xml',
                  });
              }
          }
          types.ele('Override', {
              PartName: '/docProps/core.xml',
              ContentType: 'application/vnd.openxmlformats-package.core-properties+xml',
          });
          types.ele('Override', {
              PartName: '/docProps/app.xml',
              ContentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
          });
          types.ele('Override', {
              PartName: '/xl/sharedStrings.xml',
              ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
          });
          return types.end({ pretty: false });
      };
      return ContentTypes;
  }());

  var SharedStrings = (function () {
      function SharedStrings() {
          this.arr = [];
          this.cache = {};
          this.arr = [];
      }
      SharedStrings.prototype.str2id = function (s) {
          var id = this.cache[s];
          if (id) {
              return id;
          }
          else {
              this.arr.push(s);
              this.cache[s] = this.arr.length;
              return this.arr.length;
          }
      };
      SharedStrings.prototype.toxml = function () {
          var sst = xmlbuilder__default["default"].create('sst', {
              version: '1.0',
              encoding: 'UTF-8',
              standalone: true,
          });
          sst.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
          sst.att('count', '' + this.arr.length);
          sst.att('uniqueCount', '' + this.arr.length);
          for (var i = 0; i <= this.arr.length; i++) {
              var si = sst.ele('si');
              si.ele('t', this.arr[i]);
              si.ele('phoneticPr', { fontId: 1, type: 'noConversion' });
          }
          return sst.end({ pretty: false });
      };
      return SharedStrings;
  }());

  /*! *****************************************************************************
  Copyright (c) Microsoft Corporation.

  Permission to use, copy, modify, and/or distribute this software for any
  purpose with or without fee is hereby granted.

  THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES WITH
  REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF MERCHANTABILITY
  AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY SPECIAL, DIRECT,
  INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES WHATSOEVER RESULTING FROM
  LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR
  OTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR
  PERFORMANCE OF THIS SOFTWARE.
  ***************************************************************************** */

  var __assign = function() {
      __assign = Object.assign || function __assign(t) {
          for (var s, i = 1, n = arguments.length; i < n; i++) {
              s = arguments[i];
              for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p)) t[p] = s[p];
          }
          return t;
      };
      return __assign.apply(this, arguments);
  };

  var numberFormats = {
      0: 'General',
      1: '0',
      2: '0.00',
      3: '#,##0',
      4: '#,##0.00',
      9: '0%',
      10: '0.00%',
      11: '0.00E+00',
      12: '# ?/?',
      13: '# ??/??',
      14: 'm/d/yy',
      15: 'd-mmm-yy',
      16: 'd-mmm',
      17: 'mmm-yy',
      18: 'h:mm AM/PM',
      19: 'h:mm:ss AM/PM',
      20: 'h:mm',
      21: 'h:mm:ss',
      22: 'm/d/yy h:mm',
      37: '#,##0 ;(#,##0)',
      38: '#,##0 ;[Red](#,##0)',
      39: '#,##0.00;(#,##0.00)',
      40: '#,##0.00;[Red](#,##0.00)',
      45: 'mm:ss',
      46: '[h]:mm:ss',
      47: 'mmss.0',
      48: '##0.0E+0',
      49: '@',
      56: '"上午/下午 "hh"時"mm"分"ss"秒 "',
  };

  var Style = (function () {
      function Style(book) {
          this.numberFormats = __assign({}, numberFormats);
          this.book = book;
          this.cache = {};
          this.mfonts = [];
          this.mfills = [];
          this.mbders = [];
          this.mstyle = [];
          this.numFmtNextId = 164;
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
      Style.prototype.font2id = function (_font) {
          if (_font === void 0) { _font = {}; }
          var font = __assign({ bold: '-', iter: '-', sz: '11', color: '-', name: 'Calibri', scheme: 'minor', family: '2', underline: '-', strike: '-', outline: '-', shadow: '-' }, _font);
          var strKeyOrder = [
              'bold',
              'iter',
              'sz',
              'color',
              'name',
              'scheme',
              'family',
              'underline',
              'strike',
              'outline',
              'shadow',
          ];
          var str = 'font_' + strKeyOrder.map(function (k) { return font[k]; }).join('');
          var id = this.cache[str];
          if (id) {
              return id;
          }
          else {
              this.mfonts.push(font);
              this.cache[str] = this.mfonts.length;
              return this.mfonts.length;
          }
      };
      Style.prototype.fill2id = function (_fill) {
          if (_fill === void 0) { _fill = {}; }
          var fill = __assign({ type: 'none', bgColor: '-', fgColor: '-' }, _fill);
          var str = 'fill_' + fill.type + fill.bgColor + fill.fgColor;
          var id = this.cache[str];
          if (id) {
              return id;
          }
          else {
              this.mfills.push(fill);
              this.cache[str] = this.mfills.length;
              return this.mfills.length;
          }
      };
      Style.prototype.bder2id = function (_border) {
          if (_border === void 0) { _border = {}; }
          var border = __assign({ left: '-', right: '-', top: '-', bottom: '-' }, _border);
          var left = border.left, right = border.right, top = border.top, bottom = border.bottom;
          var str = JSON.stringify(['bder_', left, right, top, bottom]);
          var id = this.cache[str];
          if (id) {
              return id;
          }
          else {
              this.mbders.push(border);
              this.cache[str] = this.mbders.length;
              return this.mbders.length;
          }
      };
      Style.prototype.numfmt2id = function (numfmt) {
          if (typeof numfmt == 'number') {
              return numfmt;
          }
          if (typeof numfmt == 'string') {
              if (!numfmt) {
                  throw 'Invalid format specification';
              }
              for (var _i = 0, _a = Object.entries(this.numberFormats); _i < _a.length; _i++) {
                  var _b = _a[_i], key = _b[0], value = _b[1];
                  if (value === numfmt) {
                      return parseInt(key);
                  }
              }
              this.numberFormats[++this.numFmtNextId] = numfmt;
              return this.numFmtNextId;
          }
          return -1;
      };
      Style.prototype.parseStyleKey = function (style) {
          var joint = [
              style.font_id,
              style.fill_id,
              style.bder_id,
              style.align,
              style.valign,
              style.wrap,
              style.rotate,
              style.numfmt_id,
          ].join('_');
          var key = "s_".concat(joint);
          return key;
      };
      Style.prototype.style2id = function (styleOpt) {
          if (styleOpt === void 0) { styleOpt = {}; }
          var style = __assign({ align: this.def_align, valign: this.def_valign, rotate: this.def_rotate, wrap: this.def_wrap, font_id: this.def_font_id, fill_id: this.def_fill_id, bder_id: this.def_bder_id, numfmt_id: this.def_numfmt_id }, styleOpt);
          var key = this.parseStyleKey(style);
          var id = this.cache[key];
          if (id) {
              return id;
          }
          else {
              this.mstyle.push(style);
              this.cache[key] = this.mstyle.length;
              return this.mstyle.length;
          }
      };
      Style.prototype.toxml = function () {
          var ss = xmlbuilder__default["default"].create('styleSheet', {
              version: '1.0',
              encoding: 'UTF-8',
              standalone: true,
          });
          ss.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
          var customNumFmts = [];
          for (var _i = 0, _a = Object.entries(this.numberFormats); _i < _a.length; _i++) {
              var _b = _a[_i], key = _b[0], fmt = _b[1];
              if (parseInt(key) >= 164) {
                  customNumFmts.push({ numFmtId: key, formatCode: fmt });
              }
          }
          if (customNumFmts.length > 0) {
              var numFmts = ss.ele('numFmts', {
                  count: customNumFmts.length,
              });
              for (var o in customNumFmts) {
                  numFmts.ele('numFmt', o);
              }
          }
          var fonts = ss.ele('fonts', { count: this.mfonts.length });
          for (var _c = 0, _d = this.mfonts; _c < _d.length; _c++) {
              var o = _d[_c];
              var e = fonts.ele('font');
              if (o.bold !== '-')
                  e.ele('b');
              if (o.iter !== '-')
                  e.ele('i');
              if (o.underline !== '-')
                  e.ele('u');
              if (o.strike !== '-')
                  e.ele('strike');
              if (o.outline !== '-')
                  e.ele('outline');
              if (o.shadow !== '-')
                  e.ele('shadow');
              e.ele('sz', { val: o.sz });
              if (o.color !== '-')
                  e.ele('color', { rgb: o.color });
              e.ele('name', { val: o.name });
              e.ele('family', { val: o.family });
              e.ele('charset', { val: '134' });
              if (o.scheme !== '-')
                  e.ele('scheme', { val: 'minor' });
          }
          var fills = ss.ele('fills', { count: this.mfills.length + 2 });
          fills.ele('fill').ele('patternFill', { patternType: 'none' });
          fills.ele('fill').ele('patternFill', { patternType: 'gray125' });
          for (var _e = 0, _f = this.mfills; _e < _f.length; _e++) {
              var o = _f[_e];
              var e = fills.ele('fill');
              var es = e.ele('patternFill', { patternType: o.type });
              if (o.fgColor !== '-')
                  es.ele('fgColor', { rgb: o.fgColor });
              if (o.bgColor !== '-')
                  es.ele('bgColor', { indexed: o.bgColor });
          }
          var borders = ss.ele('borders', { count: this.mbders.length });
          var _loop_1 = function (o) {
              var e = borders.ele('border');
              var dirs = ['left', 'right', 'top', 'bottom'];
              dirs.forEach(function (borderDir) {
                  if (o[borderDir] !== '-') {
                      if (typeof o.left === 'string') {
                          e.ele('left', { style: o.left }).ele('color', { auto: '1' });
                      }
                      else {
                          e.ele('left', { style: o.left.style }).ele('color', o.left.color);
                      }
                  }
                  else {
                      e.ele('borderDir');
                  }
              });
              e.ele('diagonal');
          };
          for (var _g = 0, _h = this.mbders; _g < _h.length; _g++) {
              var o = _h[_g];
              _loop_1(o);
          }
          ss.ele('cellStyleXfs', { count: '1' })
              .ele('xf', {
              numFmtId: '0',
              fontId: '0',
              fillId: '0',
              borderId: '0',
          })
              .ele('alignment', { vertical: 'center' });
          var cs = ss.ele('cellXfs', { count: this.mstyle.length });
          for (var _j = 0, _k = this.mstyle; _j < _k.length; _j++) {
              var o = _k[_j];
              var e = cs.ele('xf', {
                  numFmtId: o.numfmt_id || '0',
                  fontId: o.font_id - 1,
                  fillId: o.fill_id + 1,
                  borderId: o.bder_id - 1,
                  xfId: '0',
              });
              if (o.font_id !== 1)
                  e.att('applyFont', '1');
              if (o.fill_id !== 1)
                  e.att('applyFill', '1');
              if (o.numfmt_id !== undefined)
                  e.att('applyNumberFormat', '1');
              if (o.bder_id !== 1)
                  e.att('applyBorder', '1');
              if (o.align !== '-' || o.valign !== '-' || o.wrap !== '-') {
                  e.att('applyAlignment', '1');
                  var ea = e.ele('alignment', {
                      textRotation: o.rotate === '-' ? '0' : o.rotate,
                      horizontal: o.align === '-' ? 'left' : o.align,
                      vertical: o.valign === '-' ? 'bottom' : o.valign,
                  });
                  if (o.wrap !== '-')
                      ea.att('wrapText', '1');
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
      };
      return Style;
  }());

  var DocPropsApp = (function () {
      function DocPropsApp(book) {
          this.book = book;
      }
      DocPropsApp.prototype.toxml = function () {
          var props = xmlbuilder__default["default"].create('Properties', {
              version: '1.0',
              encoding: 'UTF-8',
              standalone: true,
          });
          props.att('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties');
          props.att('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes');
          props.ele('Application', 'Microsoft Excel');
          props.ele('DocSecurity', '0');
          props.ele('ScaleCrop', 'false');
          var tmp = props
              .ele('HeadingPairs')
              .ele('vt:vector', { size: 2, baseType: 'variant' });
          tmp.ele('vt:variant').ele('vt:lpstr', 'Worksheets');
          tmp.ele('vt:variant').ele('vt:i4', '' + this.book.sheets.length);
          tmp = props
              .ele('TitlesOfParts')
              .ele('vt:vector', { size: this.book.sheets.length, baseType: 'lpstr' });
          for (var i = 1; i <= this.book.sheets.length; i++) {
              tmp.ele('vt:lpstr', this.book.sheets[i - 1].name);
          }
          props.ele('Company');
          props.ele('LinksUpToDate', 'false');
          props.ele('SharedDoc', 'false');
          props.ele('HyperlinksChanged', 'false');
          props.ele('AppVersion', '12.0000');
          return props.end({ pretty: false });
      };
      return DocPropsApp;
  }());

  var XlDrawingRels = (function () {
      function XlDrawingRels(dwRels) {
          this.dwRels = dwRels;
      }
      XlDrawingRels.prototype.generate = function () { };
      XlDrawingRels.prototype.toxml = function () {
          var rs = xmlbuilder__default["default"].create('Relationships', {
              version: '1.0',
              encoding: 'UTF-8',
              standalone: true,
          });
          rs.att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
          for (var _i = 0, _a = this.dwRels; _i < _a.length; _i++) {
              var dwRel = _a[_i];
              rs.ele('Relationship', {
                  Id: dwRel.id,
                  Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                  Target: dwRel.target,
              });
          }
          return rs.end({ pretty: false });
      };
      return XlDrawingRels;
  }());

  var XlWorkbook = (function () {
      function XlWorkbook(book) {
          this.book = book;
      }
      XlWorkbook.prototype.toxml = function () {
          var wb = xmlbuilder__default["default"].create('workbook', {
              version: '1.0',
              encoding: 'UTF-8',
              standalone: true,
          });
          wb.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
          wb.att('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
          wb.att('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
          wb.att('mc:Ignorable', 'x15 xr xr6 xr10 xr2');
          wb.att('xmlns:x15', 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main');
          wb.att('xmlns:xr', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision');
          wb.att('xmlns:xr6', 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision6');
          wb.att('xmlns:xr10', 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision10');
          wb.att('xmlns:xr2', 'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2');
          wb.ele('fileVersion', {
              appName: 'xl',
              lastEdited: '4',
              lowestEdited: '4',
              rupBuild: '4505',
          });
          wb.ele('workbookPr', { filterPrivacy: '1', defaultThemeVersion: '124226' });
          wb.ele('bookViews').ele('workbookView', {
              xWindow: '0',
              yWindow: '90',
              windowWidth: '19200',
              windowHeight: '11640',
          });
          var tmp = wb.ele('sheets');
          for (var i = 1; i <= this.book.sheets.length; i++) {
              tmp.ele('sheet', {
                  name: this.book.sheets[i - 1].name,
                  sheetId: '' + i,
                  'r:id': 'rId' + i,
              });
          }
          var definedNames = wb.ele('definedNames');
          this.book.sheets.forEach(function (sheet, idx) {
              if (sheet.autofilter) {
                  definedNames
                      .ele('definedName', {
                      name: '_xlnm._FilterDatabase',
                      hidden: '1',
                      localSheetId: idx,
                  })
                      .raw("'" + sheet.name + "'!" + sheet.getRange());
              }
              if (sheet._repeatRows || sheet._repeatCols) {
                  var range = '';
                  if (sheet._repeatCols) {
                      range +=
                          "'" +
                              sheet.name +
                              "'!$" +
                              i2a(sheet._repeatCols.start) +
                              ':$' +
                              i2a(sheet._repeatCols.end);
                  }
                  if (sheet._repeatRows) {
                      range +=
                          ",'" +
                              sheet.name +
                              "'!$" +
                              sheet._repeatRows.start +
                              ':$' +
                              sheet._repeatRows.end;
                  }
                  definedNames
                      .ele('definedName', {
                      name: '_xlnm.Print_Titles',
                      localSheetId: idx,
                  })
                      .raw(range);
              }
          });
          wb.ele('calcPr', { calcId: '124519' });
          return wb.end({ pretty: false });
      };
      return XlWorkbook;
  }());

  var XlWorkbookRels = (function () {
      function XlWorkbookRels(book) {
          this.book = book;
      }
      XlWorkbookRels.prototype.toxml = function () {
          var rs = xmlbuilder__default["default"].create('Relationships', {
              version: '1.0',
              encoding: 'UTF-8',
              standalone: true,
          });
          rs.att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
          for (var i = 1; i <= this.book.sheets.length; i++) {
              rs.ele('Relationship', {
                  Id: 'rId' + i,
                  Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                  Target: 'worksheets/sheet' + i + '.xml',
              });
          }
          rs.ele('Relationship', {
              Id: 'rId' + (this.book.sheets.length + 1),
              Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
              Target: 'theme/theme1.xml',
          });
          rs.ele('Relationship', {
              Id: 'rId' + (this.book.sheets.length + 2),
              Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
              Target: 'styles.xml',
          });
          rs.ele('Relationship', {
              Id: 'rId' + (this.book.sheets.length + 3),
              Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
              Target: 'sharedStrings.xml',
          });
          rs.ele('Relationship', {
              Id: 'rId' + (this.book.sheets.length + 4),
              Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain',
              Target: 'calcChain.xml',
          });
          return rs.end();
      };
      return XlWorkbookRels;
  }());

  var XlWorksheetRels = (function () {
      function XlWorksheetRels(wsRels) {
          this.wsRels = [];
          this.wsRels = wsRels;
      }
      XlWorksheetRels.prototype.generate = function () { };
      XlWorksheetRels.prototype.toxml = function () {
          var rs = xmlbuilder__default["default"].create('Relationships', {
              version: '1.0',
              encoding: 'UTF-8',
              standalone: true,
          });
          rs.att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
          for (var _i = 0, _a = this.wsRels; _i < _a.length; _i++) {
              var wsRel = _a[_i];
              rs.ele('Relationship', {
                  Id: wsRel.id,
                  Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
                  Target: wsRel.target,
              });
          }
          return rs.end({ pretty: false });
      };
      return XlWorksheetRels;
  }());

  var addressRegex = /^[A-Z]+\d+$/;

  var AtoZ = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  var colCache = {
      _dictionary: AtoZ.split(''),
      _l2nFill: 0,
      _l2n: {},
      _n2l: [],
      _level: function (n) {
          if (n <= 26)
              return 1;
          if (n <= 26 * 26)
              return 2;
          return 3;
      },
      _fill: function (level) {
          var c = undefined;
          var v = undefined;
          var l1 = undefined;
          var l2 = undefined;
          var l3 = undefined;
          var n = 1;
          if (level >= 4)
              throw new Error('Out of bounds. Excel supports columns from 1 to 16384');
          if (this._l2nFill < 1 && level >= 1) {
              while (n <= 26) {
                  var c_1 = this._dictionary[n - 1];
                  this._n2l[n] = c_1;
                  this._l2n[c_1] = n;
                  n++;
              }
              this._l2nFill = 1;
          }
          if (this._l2nFill < 2 && level >= 2) {
              n = 27;
              var max = 26 + 26 * 26;
              while (n <= max) {
                  v = n - (26 + 1);
                  l1 = v % 26;
                  l2 = Math.floor(v / 26);
                  c = this._dictionary[l2] + this._dictionary[l1];
                  this._n2l[n] = c;
                  this._l2n[c] = n;
                  n++;
              }
              this._l2nFill = 2;
          }
          if (this._l2nFill < 3 && level >= 3) {
              n = 26 + 26 * 26 + 1;
              while (n <= 16384) {
                  v = n - (26 * 26 + 26 + 1);
                  l1 = v % 26;
                  l2 = Math.floor(v / 26) % 26;
                  l3 = Math.floor(v / (26 * 26));
                  c = this._dictionary[l3] + this._dictionary[l2] + this._dictionary[l1];
                  this._n2l[n] = c;
                  this._l2n[c] = n;
                  n++;
              }
              this._l2nFill = 3;
          }
          return;
      },
      l2n: function (l) {
          if (!this._l2n[l]) {
              this._fill(l.length);
          }
          if (!this._l2n[l])
              throw new Error('Out of bounds. Invalid column letter: ' + l);
          return this._l2n[l];
      },
      n2l: function (n) {
          if (n < 1 || n > 16384)
              throw new Error(n + ' is out of bounds. Excel supports columns from 1 to 16384');
          if (!this._n2l[n])
              this._fill(this._level(n));
          return this._n2l[n];
      },
      _hash: {},
      validateAddress: function (value) {
          if (!addressRegex.test(value))
              throw new Error('Invalid Address: ' + value);
          return true;
      },
      decodeAddress: function (value) {
          var addr = value.length < 5 && this._hash[value];
          if (addr)
              return addr;
          var hasCol = false;
          var col = '';
          var colNumber = 0;
          var hasRow = false;
          var row = '';
          var rowNumber = 0;
          var i = 0;
          while (i < value.length) {
              var char_1 = value.charCodeAt(i);
              if (!hasRow && char_1 >= 65 && char_1 <= 90) {
                  hasCol = true;
                  col += value[i];
                  colNumber = colNumber * 26 + char_1 - 64;
              }
              else if (char_1 >= 48 && char_1 <= 57) {
                  hasRow = true;
                  row += value[i];
                  rowNumber = rowNumber * 10 + char_1 - 48;
              }
              else if (hasRow && hasCol && char_1 != 36) {
                  break;
              }
              i++;
          }
          if (!hasCol) {
              colNumber = undefined;
          }
          else if (colNumber > 16384) {
              throw new Error('Out of bounds. Invalid column letter: ' + col);
          }
          if (!hasRow) {
              rowNumber = undefined;
          }
          value = col + row;
          var address = {
              address: value,
              col: colNumber,
              row: rowNumber,
              $col$row: '$' + col + '$' + row,
          };
          if (colNumber <= 100 && rowNumber <= 100) {
              this._hash[value] = address;
              this._hash[address.$col$row] = address;
          }
          return address;
      },
      getAddress: function (r, c) {
          if (c) {
              var address = this.n2l(c) + r;
              return this.decodeAddress(address);
          }
          return this.decodeAddress(r);
      },
      decode: function (value) {
          var parts = value.split(':');
          if (parts.length == 2) {
              var tl = this.decodeAddress(parts[0]);
              var br = this.decodeAddress(parts[1]);
              var result = {
                  top: Math.min(tl.row, br.row),
                  left: Math.min(tl.col, br.col),
                  bottom: Math.max(tl.row, br.row),
                  right: Math.max(tl.col, br.col),
                  tl: '',
                  br: '',
                  dimensions: '',
              };
              result.tl = this.n2l(result.left) + result.top;
              result.br = this.n2l(result.right) + result.bottom;
              result.dimensions = result.tl + ':' + result.br;
              return result;
          }
          return this.decodeAddress(value);
      },
      decodeEx: function (value) {
          var groups = value.match(/(?:(?:(?:'((?:[^']|'')*)')|([^'^ !]*))!)?(.*)/);
          var sheetName = groups[1] || groups[2];
          var reference = groups[3];
          var parts = reference.split(':');
          if (parts.length > 1) {
              var tl = this.decodeAddress(parts[0]);
              var br = this.decodeAddress(parts[1]);
              var top_1 = Math.min(tl.row, br.row);
              var left = Math.min(tl.col, br.col);
              var bottom = Math.max(tl.row, br.row);
              var right = Math.max(tl.col, br.col);
              var tl1 = this.n2l(left) + top_1;
              var br1 = this.n2l(right) + bottom;
              return {
                  top: top_1,
                  left: left,
                  bottom: bottom,
                  right: right,
                  sheetName: sheetName,
                  tl: {
                      address: tl1,
                      col: left,
                      row: top_1,
                      $col$row: '$' + this.n2l(left) + '$' + top_1,
                      sheetName: sheetName,
                  },
                  br: {
                      address: br1,
                      col: right,
                      row: bottom,
                      $col$row: '$' + this.n2l(right) + '$' + bottom,
                      sheetName: sheetName,
                  },
                  dimensions: tl + ':' + br,
              };
          }
          if (reference.startsWith('#')) {
              if (sheetName) {
                  return {
                      sheetName: sheetName,
                      error: reference,
                  };
              }
              else {
                  return {
                      error: reference,
                  };
              }
          }
          var address = this.decodeAddress(reference);
          if (sheetName) {
              return {
                  sheetName: sheetName,
                  address: address.address,
                  col: address.col,
                  row: address.row,
                  $col$row: '$' + address.col + '$' + address.row,
              };
          }
          else {
              return address;
          }
      },
      encodeAddress: function (row, col) {
          return colCache.n2l(col) + row;
      },
      encode: function () {
          var args = [];
          for (var _i = 0; _i < arguments.length; _i++) {
              args[_i] = arguments[_i];
          }
          switch (args.length) {
              case 2:
                  return colCache.encodeAddress(args[0], args[1]);
              case 4:
                  return (colCache.encodeAddress(args[0], args[1]) +
                      ':' +
                      colCache.encodeAddress(args[2], args[3]));
              default:
                  throw new Error('Can only encode with 2 or 4 arguments');
          }
      },
      inRange: function (range, address) {
          var left = range[0], top = range[1], right = range[2], bottom = range[3];
          var col = address[0], row = address[1];
          return col >= left && col <= right && row >= top && row <= bottom;
      },
  };

  var Anchor = (function () {
      function Anchor(worksheet, address, offset) {
          if (offset === void 0) { offset = 0; }
          this.nativeCol = 0;
          this.nativeColOff = 0;
          this.nativeRow = 0;
          this.nativeRowOff = 0;
          this.worksheet = worksheet;
          if (!address) {
              this.nativeCol = 0;
              this.nativeColOff = 0;
              this.nativeRow = 0;
              this.nativeRowOff = 0;
          }
          else if (typeof address === 'string') {
              var decoded = colCache.decodeAddress(address);
              this.nativeCol = decoded.col + offset;
              this.nativeColOff = 0;
              this.nativeRow = decoded.row + offset;
              this.nativeRowOff = 0;
          }
          else if (address.nativeCol !== undefined) {
              this.nativeCol = address.nativeCol || 0;
              this.nativeColOff = address.nativeColOff || 0;
              this.nativeRow = address.nativeRow || 0;
              this.nativeRowOff = address.nativeRowOff || 0;
          }
          else if (address.col !== undefined) {
              this.col = address.col + offset;
              this.row = address.row + offset;
          }
          else {
              this.nativeCol = 0;
              this.nativeColOff = 0;
              this.nativeRow = 0;
              this.nativeRowOff = 0;
          }
          return;
      }
      Object.defineProperty(Anchor.prototype, "col", {
          get: function () {
              return (this.nativeCol +
                  Math.min(this.colWidth - 1, this.nativeColOff) / this.colWidth);
          },
          set: function (v) {
              this.nativeCol = Math.floor(v);
              this.nativeColOff = Math.floor((v - this.nativeCol) * this.colWidth);
              return;
          },
          enumerable: false,
          configurable: true
      });
      Object.defineProperty(Anchor.prototype, "row", {
          get: function () {
              return (this.nativeRow +
                  Math.min(this.rowHeight - 1, this.nativeRowOff) / this.rowHeight);
          },
          set: function (v) {
              this.nativeRow = Math.floor(v);
              this.nativeRowOff = Math.floor((v - this.nativeRow) * this.rowHeight);
              return;
          },
          enumerable: false,
          configurable: true
      });
      Object.defineProperty(Anchor.prototype, "colWidth", {
          get: function () {
              return 0;
          },
          enumerable: false,
          configurable: true
      });
      Object.defineProperty(Anchor.prototype, "rowHeight", {
          get: function () {
              if (this.worksheet &&
                  this.worksheet.getRow(this.nativeRow + 1) &&
                  this.worksheet.getRow(this.nativeRow + 1).height) {
                  return Math.floor(this.worksheet.getRow(this.nativeRow + 1).height * 10000);
              }
              return 180000;
          },
          enumerable: false,
          configurable: true
      });
      Object.defineProperty(Anchor.prototype, "model", {
          get: function () {
              return {
                  nativeCol: this.nativeCol,
                  nativeColOff: this.nativeColOff,
                  nativeRow: this.nativeRow,
                  nativeRowOff: this.nativeRowOff,
              };
          },
          set: function (value) {
              this.nativeCol = value.nativeCol;
              this.nativeColOff = value.nativeColOff;
              this.nativeRow = value.nativeRow;
              this.nativeRowOff = value.nativeRowOff;
              return;
          },
          enumerable: false,
          configurable: true
      });
      Anchor.prototype.asInstance = function (model) {
          return model instanceof Anchor || model == null
              ? model
              : new Anchor(this.worksheet, model);
      };
      Anchor.prototype.toxml = function () {
          var wb = xmlbuilder__default["default"].create('workbook', {
              version: '1.0',
              encoding: 'UTF-8',
              standalone: true,
          });
          wb.ele('from').ele('workbookView', {
              xWindow: '0',
              yWindow: '90',
              windowWidth: '19200',
              windowHeight: '11640',
          });
          wb.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
          wb.att('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
          wb.ele('fileVersion', {
              appName: 'xl',
              lastEdited: '4',
              lowestEdited: '4',
              rupBuild: '4505',
          });
          wb.ele('workbookPr', { filterPrivacy: '1', defaultThemeVersion: '124226' });
          wb.ele('bookViews').ele('workbookView', {
              xWindow: '0',
              yWindow: '90',
              windowWidth: '19200',
              windowHeight: '11640',
          });
      };
      return Anchor;
  }());

  var Image = (function () {
      function Image(id, extension, content, range, options) {
          this.base64 = false;
          this.id = id;
          this.extension = extension;
          this.content = content;
          this.range = range;
          this.options = options;
          this.editAs = 'oneCell';
      }
      Image.prototype.publish = function (sheet, zip, media) {
      };
      Image.prototype.toDrawingXml = function (relId, spec) {
          var dr = xmlbuilder__default["default"].create('xdr:wsDr', {
              version: '1.0',
              encoding: 'UTF-8',
              standalone: true,
          });
          dr.att('xmlns:xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing');
          dr.att('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
          var twoCellAnchor = dr.ele('xdr:twoCellAnchor', { editAs: this.editAs });
          var _from = twoCellAnchor.ele('xdr:from');
          _from.ele('xdr:col', this.range.from.model.nativeCol);
          _from.ele('xdr:colOff', this.range.from.model.nativeColOff);
          _from.ele('xdr:row', this.range.from.model.nativeRow);
          _from.ele('xdr:rowOff', this.range.from.model.nativeRowOff);
          var _to = twoCellAnchor.ele('xdr:to');
          _to.ele('xdr:col', this.range.to.model.nativeCol);
          _to.ele('xdr:colOff', this.range.to.model.nativeColOff);
          _to.ele('xdr:row', this.range.to.model.nativeRow);
          _to.ele('xdr:rowOff', this.range.to.model.nativeRowOff);
          var pic = twoCellAnchor.ele('xdr:pic');
          var nvPicPr = pic.ele('xdr:nvPicPr');
          var graphic_index = 30924 ;
          var cNvPr = nvPicPr.ele('xdr:cNvPr', {
              id: 3,
              name: 'Graphic ' + graphic_index,
          });
          cNvPr
              .ele('a:extLst')
              .ele('a:ext', { uri: '{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}' })
              .ele('a16:creationId', {
              'xmlns:a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
              id: '{9D66B5F7-2553-484C-A5BE-4D0B8D57E08B}',
          });
          nvPicPr.ele('xdr:cNvPicPr').ele('a:picLocks', { noChangeAspect: 1 });
          var blipFill = pic.ele('xdr:blipFill');
          var blipModel = {
              'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
              'r:embed': '',
              cstate: '',
          };
          if (relId) {
              blipModel['r:embed'] = relId;
          }
          blipModel['cstate'] = 'print';
          var blip = blipFill.ele('a:blip', blipModel);
          var extLst = blip.ele('a:extLst');
          var ext = extLst.ele('a:ext', {
              uri: '{28A0092B-C50C-407E-A947-70E740481C1C}',
          });
          ext.ele('a14:useLocalDpi', {
              'xmlns:a14': 'http://schemas.microsoft.com/office/drawing/2010/main',
              val: 0,
          });
          if (this.extension == 'svg') {
              var ext_1 = extLst.ele('a:ext', {
                  uri: '{96DAC541-7B7A-43D3-8B79-37D633B846F1}',
              });
              ext_1.ele('asvg:svgBlip', {
                  'xmlns:asvg': 'http://schemas.microsoft.com/office/drawing/2016/SVG/main',
                  'r:embed': relId,
              });
          }
          if (this.options.stretch) {
              blipFill.ele('a:stretch');
          }
          blipFill.ele('srcRect');
          var spPr = pic.ele('xdr:spPr');
          var xfrm = spPr.ele('a:xfrm');
          xfrm.ele('a:off', { x: 609600, y: 190500 });
          xfrm.ele('a:ext', { cx: 2857500, cy: 2857500 });
          spPr.ele('a:prstGeom', { prst: 'rect' }).ele('a:avLst');
          twoCellAnchor.ele('xdr:clientData');
          return dr.end({ pretty: false });
      };
      return Image;
  }());

  function JSDateToExcel(dt) {
      return dt.valueOf() / 86400000 + 25569;
  }

  function getDefaultPageMargin() {
      return {
          left: '0.7',
          right: '0.7',
          top: '0.75',
          bottom: '0.75',
          header: '0.3',
          footer: '0.3',
      };
  }

  var Sheet = (function () {
      function Sheet(book, name, colCount, rowCount) {
          this.wsRels = [];
          this.colCount = 0;
          this.rowCount = 0;
          this.col_wd = [];
          this.row_ht = [];
          this._sheetViews = {
              workbookViewId: '0',
          };
          this._sheetViewsPane = {};
          this._pageSetup = {
              paperSize: '9',
              orientation: 'portrait',
              horizontalDpi: '200',
              verticalDpi: '200',
          };
          this._rowBreaks = [];
          this._colBreaks = [];
          this.book = book;
          this.name = name;
          this.colCount = colCount;
          this.rowCount = rowCount;
          this.data = {};
          for (var i = 1; i <= rowCount; i++) {
              this.data[i] = [];
              for (var j = 1; j <= colCount; j++) {
                  this.data[i][j] = { v: 0 };
              }
          }
          this.merges = [];
          this.colWidths = [];
          this.rowHeights = {};
          this.styles = {};
          this.formulas = [];
          this._pageMargins = getDefaultPageMargin();
          this.images = [];
      }
      Sheet.prototype.addImage = function (image) {
          if (!image || !image.range || !image.base64 || !image.extension)
              throw Error('please verify your image format');
          if (typeof image.range != 'string' || !/\w+\d+:\w+\d/i.test(image.range))
              throw Error('Please provide range parameter like `B2:F6`.');
          var decoded = colCache.decode(image.range);
          this.range = {
              from: new Anchor(this.worksheet, decoded.tl, -1),
              to: new Anchor(this.worksheet, decoded.br, 0),
              editAs: 'oneCell',
          };
          var id = this.book.medias.length + 1;
          var imageToAdd = new Image(id, image.extension, image.base64, this.range, image.options || {});
          this.book._addMediaFromImage(imageToAdd);
          console.log('imageToAdd', imageToAdd);
          this.images.push(imageToAdd);
          return id;
      };
      Sheet.prototype.getImage = function (id) {
          return this.images[id];
      };
      Sheet.prototype.getImages = function () {
          return this.images;
      };
      Sheet.prototype.removeImage = function (id) {
          this.images = this.images.filter(function (i) { return i.id !== id; });
      };
      Sheet.prototype.set = function () {
          var args = [];
          for (var _i = 0; _i < arguments.length; _i++) {
              args[_i] = arguments[_i];
          }
          var col = args[0], row = args[1], str = args[2];
          if (args.length === 1 && col && typeof col == 'object') {
              var cells = col;
              for (var _a = 0, _b = Object.entries(cells); _a < _b.length; _a++) {
                  var _c = _b[_a], c = _c[0], col_1 = _c[1];
                  for (var _d = 0, _e = Object.entries(col_1); _d < _e.length; _d++) {
                      var _f = _e[_d], r = _f[0], cell = _f[1];
                      this.set(parseInt(c), parseInt(r), cell);
                  }
              }
          }
          else if (str instanceof Date) {
              this.set(col, row, JSDateToExcel(str));
              this.fill(col, row, {
                  type: 'solid',
                  fgColor: 'FFFFFF',
              });
              this.numberFormat(col, row, 'd-mmm');
          }
          else if (typeof str === 'object') {
              for (var _g = 0, _h = Object.entries(str); _g < _h.length; _g++) {
                  var _j = _h[_g], key = _j[0], value = _j[1];
                  this[key](col, row, value);
              }
          }
          else if (typeof str === 'string') {
              if (str !== '') {
                  this.data[row][col].v = this.book.sharedStrings.str2id('' + str);
              }
              this.data[row][col].dataType = 'string';
              return;
          }
          else if (typeof str === 'number') {
              this.data[row][col].v = str;
              this.data[row][col].dataType = 'number';
              return;
          }
          else {
              this.data[row][col].v = str;
          }
          return;
      };
      Sheet.prototype.formula = function (col, row, str) {
          var _this = this;
          if (typeof str == 'string') {
              this.formulas = this.formulas || [];
              this.formulas[row] = this.formulas[row] || [];
              var sheet_idx = this.book.sheets.findIndex(function (sheet) { return sheet.name === _this.name; });
              this.book.calcChain.add_ref(sheet_idx, col, row);
              this.formulas[row][col] = str;
          }
      };
      Sheet.prototype.merge = function (from_cell, to_cell) {
          this.merges.push({ from: from_cell, to: to_cell });
      };
      Sheet.prototype.width = function (col, wd) {
          return this.col_wd.push({ c: col, cw: wd });
      };
      Sheet.prototype.getColWidth = function (col) {
          for (var _i = 0, _a = this.col_wd; _i < _a.length; _i++) {
              var _col = _a[_i];
              if (_col.c == col)
                  return Math.floor(_col.cw * 10000);
          }
          return 640000;
      };
      Sheet.prototype.height = function (row, ht) {
          this.row_ht[row] = ht;
      };
      Sheet.prototype.font = function (col, row, font_s) {
          return (this.styles['font_' + col + '_' + row] =
              this.book.style.font2id(font_s));
      };
      Sheet.prototype.fill = function (col, row, fill_style) {
          var key = 'fill_' + col + '_' + row;
          return (this.styles[key] = this.book.style.fill2id(fill_style));
      };
      Sheet.prototype.border = function (col, row, bder_s) {
          return (this.styles['bder_' + col + '_' + row] =
              this.book.style.bder2id(bder_s));
      };
      Sheet.prototype.numberFormat = function (col, row, numfmt_s) {
          this.styles['numfmt_' + col + '_' + row] =
              this.book.style.numfmt2id(numfmt_s);
      };
      Sheet.prototype.align = function (col, row, alignValue) {
          return (this.styles['algn_' + col + '_' + row] = alignValue);
      };
      Sheet.prototype.valign = function (col, row, valignValue) {
          return (this.styles['valgn_' + col + '_' + row] = valignValue);
      };
      Sheet.prototype.rotate = function (col, row, textRotation) {
          return (this.styles['rotate_' + col + '_' + row] = textRotation);
      };
      Sheet.prototype.wrap = function (col, row, wrap_s) {
          return (this.styles['wrap_' + col + '_' + row] = wrap_s);
      };
      Sheet.prototype.autoFilter = function (filter_s) {
          return (this.autofilter =
              typeof filter_s === 'string' ? filter_s : this.getRange());
      };
      Sheet.prototype.sheetViews = function (obj) {
          for (var _i = 0, _a = Object.entries(obj); _i < _a.length; _i++) {
              var _b = _a[_i], key = _b[0], val = _b[1];
              var k = key;
              var fn = this[k];
              if (typeof fn === 'function') {
                  this[k](val);
              }
              else {
                  this._sheetViews[k] = val;
              }
          }
      };
      Sheet.prototype.split = function (ncols, nrows, state, activePane, _topLeftCell) {
          if (state === void 0) { state = 'frozen'; }
          if (activePane === void 0) { activePane = 'bottomRight'; }
          var topLeftCell = _topLeftCell || i2a((ncols || 0) + 1) + ((nrows || 0) + 1);
          if (ncols)
              this._sheetViewsPane.xSplit = '' + ncols;
          if (nrows)
              this._sheetViewsPane.ySplit = '' + nrows;
          if (state)
              this._sheetViewsPane.state = state;
          if (activePane)
              this._sheetViewsPane.activePane = activePane;
          if (topLeftCell)
              this._sheetViewsPane.topLeftCell = topLeftCell;
      };
      Sheet.prototype.printBreakRows = function (arr) {
          this._rowBreaks = arr;
      };
      Sheet.prototype.printBreakColumns = function (arr) {
          this._colBreaks = arr;
      };
      Sheet.prototype.printRepeatRows = function (start, end) {
          if (Array.isArray(start)) {
              this._repeatRows = { start: start[0], end: start[1] };
          }
          else {
              this._repeatRows = { start: start, end: end };
          }
      };
      Sheet.prototype.printRepeatColumns = function (start, end) {
          if (Array.isArray(start)) {
              this._repeatCols = { start: start[0], end: start[1] };
          }
          else {
              this._repeatCols = { start: start, end: end };
          }
      };
      Sheet.prototype.pageSetup = function (obj) {
          for (var _i = 0, _a = Object.entries(obj); _i < _a.length; _i++) {
              var _b = _a[_i], key = _b[0], val = _b[1];
              this._pageSetup[key] = val;
          }
      };
      Sheet.prototype.pageMargins = function (obj) {
          for (var _i = 0, _a = Object.entries(obj); _i < _a.length; _i++) {
              var _b = _a[_i], key = _b[0], val = _b[1];
              this._pageMargins[key] = val;
          }
      };
      Sheet.prototype.style_id = function (col, row) {
          var inx = '_' + col + '_' + row;
          var style = {
              numfmt_id: this.styles['numfmt' + inx],
              font_id: this.styles['font' + inx],
              fill_id: this.styles['fill' + inx],
              bder_id: this.styles['bder' + inx],
              align: this.styles['algn' + inx],
              valign: this.styles['valgn' + inx],
              rotate: this.styles['rotate' + inx],
              wrap: this.styles['wrap' + inx],
          };
          var id = this.book.style.style2id(style);
          return id;
      };
      Sheet.prototype.getRange = function () {
          return '$A$1:$' + i2a(this.colCount) + '$' + this.rowCount;
      };
      Sheet.prototype.toxml = function () {
          var ws = xmlbuilder__default["default"].create('worksheet', {
              version: '1.0',
              encoding: 'UTF-8',
              standalone: true,
          });
          ws.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
          ws.att('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
          ws.ele('dimension', { ref: 'A1' });
          ws.ele('sheetViews')
              .ele('sheetView', this._sheetViews)
              .ele('pane', this._sheetViewsPane);
          ws.ele('sheetFormatPr', { defaultRowHeight: '13.5' });
          if (this.col_wd.length > 0) {
              var cols = ws.ele('cols');
              for (var _i = 0, _a = this.col_wd; _i < _a.length; _i++) {
                  var cw = _a[_i];
                  cols.ele('col', {
                      min: '' + cw.c,
                      max: '' + cw.c,
                      width: cw.cw,
                      customWidth: '1',
                  });
              }
          }
          var sd = ws.ele('sheetData');
          for (var i = 1; i <= this.rowCount; i++) {
              var r = sd.ele('row', { r: '' + i, spans: '1:' + this.colCount });
              var ht = this.row_ht[i];
              if (ht) {
                  r.att('ht', ht);
                  r.att('customHeight', '1');
              }
              for (var j = 1; j <= this.colCount; j++) {
                  var ix = this.data[i][j];
                  var sid = this.style_id(j, i);
                  if ((ix.v !== null && ix.v !== undefined) || sid !== 1) {
                      var c = r.ele('c', { r: '' + i2a(j) + i });
                      if (sid !== 1)
                          c.att('s', '' + (sid - 1));
                      if (this.formulas[i] && this.formulas[i][j]) {
                          c.ele('f', '' + this.formulas[i][j]);
                          c.ele('v');
                      }
                      else if (ix.dataType == 'string') {
                          c.att('t', 's');
                          c.ele('v', '' + (ix.v - 1));
                      }
                      else if (ix.dataType == 'number') {
                          c.ele('v', '' + ix.v);
                      }
                  }
              }
          }
          if (this.merges.length > 0) {
              var mc = ws.ele('mergeCells', { count: this.merges.length });
              for (var _b = 0, _c = this.merges; _b < _c.length; _b++) {
                  var m = _c[_b];
                  mc.ele('mergeCell', {
                      ref: '' + i2a(m.from.col) + m.from.row + ':' + i2a(m.to.col) + m.to.row,
                  });
              }
          }
          if (typeof this.autofilter == 'string') {
              ws.ele('autoFilter', { ref: this.autofilter });
          }
          ws.ele('phoneticPr', { fontId: '1', type: 'noConversion' });
          ws.ele('pageMargins', this._pageMargins);
          ws.ele('pageSetup', this._pageSetup);
          if (this._rowBreaks && this._rowBreaks.length) {
              var cb = ws.ele('rowBreaks', {
                  count: this._rowBreaks.length,
                  manualBreakCount: this._rowBreaks.length,
              });
              for (var _d = 0, _e = this._rowBreaks; _d < _e.length; _d++) {
                  var i = _e[_d];
                  cb.ele('brk', { id: i, man: '1' });
              }
          }
          if (this._colBreaks && this._colBreaks.length) {
              var cb = ws.ele('colBreaks', {
                  count: this._colBreaks.length,
                  manualBreakCount: this._colBreaks.length,
              });
              for (var _f = 0, _g = this._colBreaks; _f < _g.length; _f++) {
                  var i = _g[_f];
                  cb.ele('brk', { id: i, man: '1' });
              }
          }
          for (var _h = 0, _j = this.wsRels; _h < _j.length; _h++) {
              var wsRel = _j[_h];
              ws.ele('drawing', { 'r:id': wsRel.id });
          }
          return ws.end({ pretty: false });
      };
      Sheet.prototype.getRow = function (rn) {
          console.log(rn);
          console.error('Sheet.getRow() is NOT implement');
          return {
              height: 0,
          };
      };
      return Sheet;
  }());

  var Workbook = (function () {
      function Workbook(filePath, fileName) {
          this.sheets = [];
          this.medias = [];
          this.filePath = filePath;
          this.fileName = fileName;
          this.id = (Math.random() * 9999999).toFixed(0);
          this.sheets = [];
          this.medias = [];
          this.sharedStrings = new SharedStrings();
          this.contentType = new ContentTypes(this);
          this.docPropsApp = new DocPropsApp(this);
          this.XlWorkbook = new XlWorkbook(this);
          this.XlWorkbookRels = new XlWorkbookRels(this);
          this.style = new Style(this);
          this.calcChain = new CalcChain(this);
      }
      Workbook.prototype.createSheet = function (name, cols, rows) {
          var sheet = new Sheet(this, name, cols, rows);
          this.sheets.push(sheet);
          return sheet;
      };
      Workbook.prototype._addMediaFromImage = function (image) {
          this.medias.push({ image: image });
      };
      Workbook.prototype._removeMediaFromImage = function (image) {
          var foundIndex = this.medias.findIndex(function (media) { return media.image.id === image.id; });
          if (foundIndex !== -1) {
              this.medias.splice(foundIndex, 1);
          }
      };
      Workbook.prototype.save = function (target, opts, cb) {
          if (typeof target === 'function' && !opts && !cb) {
              cb = target;
              target = "".concat(this.filePath, "/").concat(this.fileName);
              opts = {};
          }
          if (typeof opts === 'function' && !cb) {
              cb = opts;
              opts = {};
          }
          this._save(target, opts, cb);
      };
      Workbook.prototype._save = function (target, opts, cb) {
          if (opts === void 0) { opts = {}; }
          this.generate(function (err, zip) {
              var args = { type: 'nodebuffer' };
              if (opts.compressed) {
                  args.compressed = 'DEFLATE';
              }
              zip.generateAsync(args).then(function (buffer) {
                  if (err)
                      return cb(err);
                  fs.writeFile(target, buffer, function (err) { return cb(err); });
              });
          });
      };
      Workbook.prototype.generate = function (cb) {
          var zip = new JSZip__default["default"]();
          for (var _i = 0, _a = Object.entries(baseXl); _i < _a.length; _i++) {
              var _b = _a[_i], key = _b[0], value = _b[1];
              zip.file(key, value);
          }
          zip.file('[Content_Types].xml', this.contentType.toxml());
          zip.file('docProps/app.xml', this.docPropsApp.toxml());
          zip.file('xl/workbook.xml', this.XlWorkbook.toxml());
          zip.file('xl/sharedStrings.xml', this.sharedStrings.toxml());
          zip.file('xl/_rels/workbook.xml.rels', this.XlWorkbookRels.toxml());
          var wbMediaCounter = 1;
          for (var i = 0; i < this.sheets.length; i++) {
              var sheet = this.sheets[i];
              sheet.wsRels = [];
              for (var j = 0; j < sheet.images.length; j++) {
                  var image = sheet.images[j];
                  var dwRels = [];
                  var relId = 'rId' + (sheet.wsRels.length + 1);
                  var mediaFilename = [wbMediaCounter, '.', image.extension].join('');
                  zip.file("xl/media/image".concat(mediaFilename), image.content, {
                      base64: true,
                  });
                  dwRels.push({ id: relId, target: "../media/image".concat(mediaFilename) });
                  var drawingFilename = "".concat(wbMediaCounter, ".xml");
                  sheet.wsRels.push({
                      id: relId,
                      target: "../drawings/drawing".concat(drawingFilename),
                  });
                  zip.file("xl/drawings/drawing".concat(drawingFilename), image.toDrawingXml(relId, image));
                  zip.file("xl/drawings/_rels/drawing".concat(wbMediaCounter, ".xml.rels"), new XlDrawingRels(dwRels).toxml());
                  wbMediaCounter++;
              }
              zip.file("xl/worksheets/_rels/sheet".concat(i + 1, ".xml.rels"), new XlWorksheetRels(sheet.wsRels).toxml());
              zip.file("xl/worksheets/sheet".concat(i + 1, ".xml"), this.sheets[i].toxml());
          }
          zip.file('xl/styles.xml', this.style.toxml());
          if (Object.keys(this.calcChain.cache).length > 0) {
              zip.file('xl/calcChain.xml', this.calcChain.toxml());
          }
          cb(null, zip);
      };
      Workbook.prototype.cancel = function () {
          console.error('workbook.cancel() is deprecated');
      };
      return Workbook;
  }());

  var xlsx = {
      createWorkbook: function (fpath, fname) { return new Workbook(fpath, fname); },
  };

  exports.Workbook = Workbook;
  exports.xlsx = xlsx;

  Object.defineProperty(exports, '__esModule', { value: true });

}));
