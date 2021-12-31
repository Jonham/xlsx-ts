export type _Image = MediaImage & {
  id: string;
  extension: string;
  content: any;
  toDrawingXml: (relId: string, image: _Image) => any;
};
export type _Media = {
  image: _Image;
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

export type _Font = {
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

export type _FONT_ID = number;

export type _Fill = {
  /** 'none' */
  type: string;
  /** '-' */
  bgColor: string;
  /** '-' */
  fgColor: string;
};

export type _Border = {
  left: string;
  right: string;
  top: string;
  bottom: string;
};
