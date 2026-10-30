import js from "@eslint/js";
import { defineConfig } from "eslint/config";
import stylistic from "@stylistic/eslint-plugin";
import stylisticRules from "../stylistic.rules.mjs";
import googleAppsScript from "eslint-plugin-googleappsscript";

export default defineConfig([
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
      sourceType: "script",
      globals: googleAppsScript.environments.googleappsscript.globals,
      parserOptions: {
        ecmaFeatures: {
          globalReturn: true
        }
      }
    }
  }
]);
