import js from "@eslint/js";
import { defineConfig } from "eslint/config";
import stylistic from "@stylistic/eslint-plugin";
import stylisticRules from "../stylistic.rules.mjs";
import googleAppsScript from "eslint-plugin-googleappsscript";
import globals from "globals";
import htmlParser from "@html-eslint/parser";
import html from "@html-eslint/eslint-plugin";
import jsInHtml from "eslint-plugin-html";

export default defineConfig([
  {
    ignores: ['dist/**']
  },
  {
    files: ["**/*.{js,mjs,cjs}"],
    ...js.configs.recommended
  },
  {
    files: ["src/**/*.js"],
    extends: [stylistic.configs.recommended],
    languageOptions: {
      sourceType: "module",
      globals: googleAppsScript.environments.googleappsscript.globals
    },
    rules: {
      ...stylisticRules
    },
  },
  {
    files: ["src/**/nexus.js"],
    languageOptions: {
      sourceType: "script",
      parserOptions: {
        ecmaFeatures: {
          globalReturn: true
        }
      }
    }
  },
  {
    files: ["src/**/*.html"],
    plugins: {
      js,
      html,
      jsInHtml
    },
    extends: [
      "js/recommended",
      "html/recommended",
      stylistic.configs.recommended
    ],
    languageOptions: {
      parser: htmlParser,
      parserOptions: {
        templateEngineSyntax: {
          "<?": "?>"
        }
      },
      globals: {
        ...globals.browser,
        ...googleAppsScript.environments.googleappsscript.globals,
        google: "readonly"
      }
    },
    rules: {
      ...stylisticRules,
      "html/indent": ["error", 2, { "tagChildrenIndent": { "html": 0 } }]
    }
  }
]);
