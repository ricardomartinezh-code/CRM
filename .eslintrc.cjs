const { Legacy } = require('@eslint/eslintrc');
const appsScriptPlugin = require('@google/eslint-plugin-apps-script');

const appsScriptEnv = appsScriptPlugin.environments.googleappsscript;
Legacy.environments.set('googleappsscript', appsScriptEnv);

module.exports = {
  root: true,
  parser: 'espree',
  parserOptions: {
    ecmaVersion: 2022,
    sourceType: 'module',
  },
  plugins: ['html', 'import', '@google/apps-script'],
  extends: ['eslint:recommended', 'plugin:import/recommended', 'prettier'],
  env: {
    browser: true,
    es2021: true,
  },
  overrides: [
    {
      files: ['**/*.html'],
      plugins: ['html'],
      processor: 'html/html',
      globals: {
        XLSX: 'readonly',
      },
      rules: {
        'no-unused-vars': 'off',
        'no-undef': 'off',
        'no-useless-escape': 'off',
        'no-useless-catch': 'off',
        'no-empty': 'off',
        'no-irregular-whitespace': 'off',
      },
    },
    {
      files: ['apps-script.gs'],
      plugins: ['@google/apps-script'],
      env: {
        googleappsscript: true,
      },
      parserOptions: {
        sourceType: 'script',
      },
      globals: appsScriptEnv.globals,
      rules: {
        'no-unused-vars': 'off',
      },
    },
    {
      files: ['packages/**/*.js', '.eslintrc.cjs'],
      env: {
        node: true,
      },
    },
    {
      files: ['config/**/*.js', 'js/**/*.js', 'tests/**/*.js'],
      env: {
        browser: true,
        node: true,
      },
    },
  ],
  settings: {
    'html/html-extensions': ['.html'],
  },
};
