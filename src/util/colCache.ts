import { addressRegex } from '../const/addressRegex';

const AtoZ = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
/** letter A-Z */
export type Letter = string;
export type _Address = {
  address: string;
  col: number;
  row: number;
  $col$row: string;
};

/** Column Letter to Number conversion */
export const colCache = {
  _dictionary: AtoZ.split(''),
  _l2nFill: 0,
  /** letter map number */
  _l2n: {} as Record<Letter, number>,
  // private num2letter: string[] = [],
  _n2l: [] as Letter[],

  _level(n: number) {
    if (n <= 26) return 1;
    if (n <= 26 * 26) return 2;
    return 3;
  },
  _fill(level: number) {
    let c = undefined;
    let v = undefined;
    let l1 = undefined;
    let l2 = undefined;
    let l3 = undefined;
    let n = 1;
    if (level >= 4)
      throw new Error('Out of bounds. Excel supports columns from 1 to 16384');
    if (this._l2nFill < 1 && level >= 1) {
      while (n <= 26) {
        const c = this._dictionary[n - 1];
        this._n2l[n] = c;
        this._l2n[c] = n;
        n++;
      }
      this._l2nFill = 1;
    }
    if (this._l2nFill < 2 && level >= 2) {
      n = 27;
      const max = 26 + 26 * 26;
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
  l2n(l: Letter) {
    if (!this._l2n[l]) {
      this._fill(l.length);
    }
    if (!this._l2n[l])
      throw new Error('Out of bounds. Invalid column letter: ' + l);
    return this._l2n[l];
  },
  n2l(n: number) {
    if (n < 1 || n > 16384)
      throw new Error(
        n + ' is out of bounds. Excel supports columns from 1 to 16384',
      );
    if (!this._n2l[n]) this._fill(this._level(n));
    return this._n2l[n];
  },
  _hash: {} as Record<string, _Address>,
  validateAddress(value: string) {
    if (!addressRegex.test(value)) throw new Error('Invalid Address: ' + value);
    return true;
  },
  decodeAddress(value: string) {
    const addr = value.length < 5 && this._hash[value];
    if (addr) return addr;
    let hasCol = false;
    let col = '';
    let colNumber: number | undefined = 0;
    let hasRow = false;
    let row = '';
    let rowNumber: number | undefined = 0;
    let char = undefined;

    let i = 0;
    while (i < value.length) {
      const char = value.charCodeAt(i);
      // col should before row
      if (!hasRow && char >= 65 && char <= 90) {
        // 65 = 'A'.charCodeAt(0)
        // 90 = 'Z'.charCodeAt(0)
        hasCol = true;
        col += value[i];
        // colNumber starts from 1
        colNumber = colNumber * 26 + char - 64;
      } else if (char >= 48 && char <= 57) {
        // 48 = '0'.charCodeAt(0)
        // 57 = '9'.charCodeAt(0)
        hasRow = true;
        row += value[i];
        // rowNumber starts from 0
        rowNumber = rowNumber * 10 + char - 48;
      } else if (hasRow && hasCol && char != 36) {
        // 36 = '$'.charCodeAt(0)
        break;
      }
      i++;
    }
    if (!hasCol) {
      colNumber = undefined;
    } else if (colNumber > 16384) {
      throw new Error('Out of bounds. Invalid column letter: ' + col);
    }
    if (!hasRow) {
      rowNumber = undefined;
    }
    //   // in case $row$col
    value = col + row;
    const address: _Address = {
      address: value,
      col: colNumber as number, // TODO undefined
      row: rowNumber as number, // TODO undefined
      $col$row: '$' + col + '$' + row,
    };

    // TODO mem fix - cache only the tl 100x100 square
    if ((colNumber as number) <= 100 && (rowNumber as number) <= 100) {
      this._hash[value] = address;
      this._hash[address.$col$row] = address;
    }
    return address;
  },
  getAddress(r: string, c: number) {
    if (c) {
      const address = this.n2l(c) + r;
      return this.decodeAddress(address);
    }
    return this.decodeAddress(r);
  },
  decode(value: string) {
    const parts = value.split(':');
    if (parts.length == 2) {
      const tl = this.decodeAddress(parts[0]);
      const br = this.decodeAddress(parts[1]);
      const result = {
        top: Math.min(tl.row, br.row),
        left: Math.min(tl.col, br.col),
        bottom: Math.max(tl.row, br.row),
        right: Math.max(tl.col, br.col),
        tl: '',
        br: '',
        dimensions: '',
      };
      // reconstruct tl, br and dimensions
      result.tl = this.n2l(result.left) + result.top;
      result.br = this.n2l(result.right) + result.bottom;
      result.dimensions = result.tl + ':' + result.br;
      return result;
    }
    return this.decodeAddress(value);
  },
  // decodeEx: (value) ->
  //   groups = value.match(/(?:(?:(?:'((?:[^']|'')*)')|([^'^ !]*))!)?(.*)/)
  //   sheetName = groups[1] or groups[2]
  //   # Qouted and unqouted groups
  //   reference = groups[3]
  //   # Remaining address
  //   parts = reference.split(':')
  //   if parts.length > 1
  //     tl = @decodeAddress(parts[0])
  //     br = @decodeAddress(parts[1])
  //     top = Math.min(tl.row, br.row)
  //     left = Math.min(tl.col, br.col)
  //     bottom = Math.max(tl.row, br.row)
  //     right = Math.max(tl.col, br.col)
  //     tl = @n2l(left) + top
  //     br = @n2l(right) + bottom
  //     return {
  //       top: top
  //       left: left
  //       bottom: bottom
  //       right: right
  //       sheetName: sheetName
  //       tl:
  //         address: tl
  //         col: left
  //         row: top
  //         $col$row: '$' + @n2l(left) + '$' + top
  //         sheetName: sheetName
  //       br:
  //         address: br
  //         col: right
  //         row: bottom
  //         $col$row: '$' + @n2l(right) + '$' + bottom
  //         sheetName: sheetName
  //       dimensions: tl + ':' + br
  //     }
  //   if reference.startsWith('#')
  //     return if sheetName then {
  //       sheetName: sheetName,
  //       error: reference
  //     } else {error: reference}
  //   address = @decodeAddress(reference)
  //   if sheetName then {
  //     sheetName: sheetName,
  //     address: address.address,
  //     col: address.col,
  //     row: address.row,
  //     $col$row: '$' + col + '$' + row
  //   } else address
  encodeAddress(row: number, col: number) {
    return colCache.n2l(col) + row;
  },
  encode(...args: number[]) {
    switch (args.length) {
      case 2:
        return colCache.encodeAddress(args[0], args[1]);
      case 4:
        return (
          colCache.encodeAddress(args[0], args[1]) +
          ':' +
          colCache.encodeAddress(args[2], args[3])
        );
      default:
        throw new Error('Can only encode with 2 or 4 arguments');
    }
    // return
  },
  inRange(
    range: [number, number, number, number],
    address: [number, number],
  ): boolean {
    const [left, top, right, bottom] = range;
    // const left = range[0]
    // const top = range[1]
    // const right = range[range.length - 2]
    // const bottom = range[range.length - 1]
    // const [left, top, , right, bottom] = range;
    const [col, row] = address;
    // row = address[1]
    return col >= left && col <= right && row >= top && row <= bottom;
  },
};
