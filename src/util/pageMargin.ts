export type PageMargin = {
  /** 默认值: '0.7' */
  left: string;
  /** 默认值: '0.7' */
  right: string;
  /** 默认值: '0.75' */
  top: string;
  /** 默认值: '0.75' */
  bottom: string;
  /** 默认值: '0.3' */
  header: string;
  /** 默认值: '0.3' */
  footer: string;
};

export function getDefaultPageMargin(): PageMargin {
  return {
    left: '0.7',
    right: '0.7',
    top: '0.75',
    bottom: '0.75',
    header: '0.3',
    footer: '0.3',
  };
}
