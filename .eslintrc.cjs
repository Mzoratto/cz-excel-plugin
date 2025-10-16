module.exports = {
  parser: "@typescript-eslint/parser",
  parserOptions: {
    ecmaVersion: "latest",
    sourceType: "module"
  },
  env: {
    browser: true,
    es2021: true
  },
  plugins: ["@typescript-eslint"],
  extends: [
    "eslint:recommended",
    "plugin:@typescript-eslint/recommended",
    "plugin:import/recommended",
    "plugin:promise/recommended",
    "prettier"
  ],
  globals: {
    Office: "readonly",
    Excel: "readonly"
  },
  rules: {
    "import/no-unresolved": "off",
    "import/named": "off"
  }
};
