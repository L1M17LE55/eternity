import js from "@eslint/js";
import { defineConfig } from "eslint/config";
import stylistic from "@stylistic/eslint-plugin";
import stylisticRules from "../stylistic.rules.mjs";
import googleAppsScript from "eslint-plugin-googleappsscript";
import globals from "globals";
import html from "eslint-plugin-html"

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
    rules: {
      ...stylisticRules
    },
    languageOptions: {
      sourceType: "module",
      globals: googleAppsScript.environments.googleappsscript.globals
    }
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
    extends: [
      js.configs.recommended,
      stylistic.configs.recommended
    ],
    rules: {
      ...stylisticRules
    },
    plugins: {
      html: html,
    },
    languageOptions: {
      globals: {
        ...globals.browser,
        ...googleAppsScript.environments.googleappsscript.globals,
        google: "readonly"
      }
    }
  }
]);
