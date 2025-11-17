import typescript from '@rollup/plugin-typescript';
import resolve from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';

export default {
  input: 'src/index.ts',
  output: [
    {
      file: 'dist/index.cjs.js',
      format: 'cjs',
      sourcemap: false,
    },
  ],
  external: [],
  plugins: [
    resolve(),
    commonjs({
      preferBuiltins: false,
    }),
    typescript({
      tsconfig: './tsconfig.json',
    }),
  ],
}
