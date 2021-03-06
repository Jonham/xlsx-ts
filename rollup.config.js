import resolve from 'rollup-plugin-node-resolve';
import commonjs from 'rollup-plugin-commonjs';
import typescript from 'rollup-plugin-typescript2';
import { terser } from 'rollup-plugin-terser';
import pkg from './package.json';

// CommonJS (for Node) and ES module (for bundlers) build.
// (We could have three entries in the configuration array
// instead of two, but it's quicker to generate multiple
// builds from a single configuration where possible, using
// an array for the `output` option, where we can specify
// `file` and `format` for each target)
export default [
  // browser-friendly UMD build
  {
    input: 'src/index.ts',
    output: [
      {
        name: 'XLSXts',
        file: pkg.browser + '.js',
        format: 'umd',
      },
      {
        name: 'XLSXts',
        file: pkg.browser + '.min.js',
        format: 'umd',
        plugins: [terser()],
      },
    ],
    external: ['xmlbuilder', 'fs', 'jszip'],
    plugins: [
      resolve({
        browser: true,
      }), // so Rollup can find external
      commonjs({}), // so Rollup can convert external to an ES module
      typescript({
        rollupCommonJSResolveHack: false,
        clean: true,
      }),
    ],
  },
  {
    input: 'src/index.ts',
    external: ['xmlbuilder', 'fs', 'jszip'],
    plugins: [
      typescript({
        rollupCommonJSResolveHack: false,
        clean: true,
      }), // so Rollup can convert TypeScript to JavaScript
      // handlebars(),
      // babel(),
    ],
    output: [
      { file: pkg.cjs, format: 'cjs' },
      { file: pkg.esModule, format: 'es' },
    ],
  },
];
