import js from '@eslint/js';
import { defineConfig, globalIgnores } from 'eslint/config';
import ts from 'typescript-eslint';

export default defineConfig([
  globalIgnores(['apm_modules', 'node_modules']),
  {
    files: ['**/*.mjs'],
    plugins: {
      js,
    },
    extends: [js.configs.recommended],
    rules: {
      'consistent-return': 'error',
      eqeqeq: 'error',
      'no-else-return': 'error',
      'prefer-const': 'error',
      'prefer-template': 'error',
      semi: 'error',
    },
  },
  {
    files: ['**/*.ts'],
    plugins: {
      ts,
    },
    extends: [js.configs.recommended, ts.configs.strict, ts.configs.stylistic],
    rules: {
      '@typescript-eslint/consistent-type-imports': ['error'],
      'consistent-return': 'error',
      eqeqeq: 'error',
      'no-else-return': 'error',
      'prefer-const': 'error',
      'prefer-template': 'error',
    },
  },
]);
