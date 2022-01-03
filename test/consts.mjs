import { dirname } from 'path';
import { fileURLToPath } from 'url';

export const __dirname = dirname(fileURLToPath(import.meta.url));

// ✅
export const ICON_CHECK = '\u001B[32m✓\u001B[39m';
export const CHAR_CHECK = '✅';
