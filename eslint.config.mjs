import pluginJs from '@eslint/js';
import tseslint from 'typescript-eslint';
export default [
  ...tseslint.configs.recommended,
  pluginJs.configs.recommended,
  {
    ignores: ['node_modules'],
  },
  {
    files: ['**/*.js', '**/*.mjs'],
    rules: {
      'consistent-return': 'error',
      eqeqeq: 'error',
      'no-else-return': 'error',
      quotes: ['error', 'single'],
      semi: 'error',
    },
  },
  {
    files: ['**/*.ts'],
    languageOptions: {
      parser: tseslint.parser,
      parserOptions: {
        projectService: true,
        tsconfigRootDir: import.meta.dirname,
      },
    },
    rules: {
      '@typescript-eslint/consistent-type-imports': [
        'error',
        { prefer: 'type-imports' },
      ],
      'consistent-return': 'error',
      eqeqeq: 'error',
      'no-else-return': 'error',
      'prefer-const': 'error',
      'prefer-template': 'error',
    },
  },
];
