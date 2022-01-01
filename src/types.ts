import { Image } from './lib/Image';

export type _Image = MediaImage & {
  id: string;
  extension: string;
  content: any;
  toDrawingXml: (relId: string, image: _Image) => any;
};
export type Media = {
  image: Image;
};

export type _dwRel = {
  id: string;
  target: string;
};

export type _wsRel = {
  id: string;
  target: string;
};

export type TODO = unknown;

export type FontDef = {
  bold: string;
  iter: string;
  sz: string;
  color: string;
  name: string;
  scheme: string;
  family: string;
  underline: string;
  strike: string;
  outline: string;
  shadow: string;
};

export type FontID = number;

export type Fill = {
  /** 'none' */
  type: string;
  /** '-' */
  bgColor: string;
  /** '-' */
  fgColor: string;
};

export type BorderStyle =
  | {
      style?: string;
      color?: string;
    }
  | '-';
export type Border = {
  left: BorderStyle;
  right: BorderStyle;
  top: BorderStyle;
  bottom: BorderStyle;
};

export type StyleDef = {
  align: string;
  valign: string;
  rotate: string;
  wrap: string;
  font_id: number;
  fill_id: number;
  bder_id: number;
  numfmt_id: number;
};
