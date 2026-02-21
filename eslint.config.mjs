import js from '@eslint/js';
import tseslint from 'typescript-eslint';
import vitest from 'eslint-plugin-vitest';
import prettierConfig from 'eslint-config-prettier';
import globals from 'globals';

export default [
  // Global ignores
  {
    ignores: [
      '**/node_modules/**',
      '**/dist/**',
      '**/build/**',
      '**/coverage/**',
      '**/*.d.ts',
      '**/package-lock.json',
    ],
  },

  // Base configuration for all JS/TS files
  {
    files: ['**/*.{js,mjs,cjs,ts,tsx}'],
    languageOptions: {
      ecmaVersion: 2023,
      sourceType: 'module',
      globals: {
        ...globals.browser,
        ...globals.es2023,
      },
    },
  },

  // Apply recommended configurations
  js.configs.recommended,

  // TypeScript specific configuration (type-checked rules)
  {
    files: ['**/*.{ts,tsx}'],
    plugins: {
      '@typescript-eslint': tseslint.plugin,
    },
    languageOptions: {
      parser: tseslint.parser,
      parserOptions: {
        projectService: true,
        tsconfigRootDir: import.meta.dirname,
      },
    },
    rules: {
      // Add TypeScript recommended rules
      ...tseslint.configs.recommended.reduce((acc, config) => ({ ...acc, ...config.rules }), {}),
      ...tseslint.configs.recommendedTypeChecked.reduce(
        (acc, config) => ({ ...acc, ...config.rules }),
        {}
      ),
      ...tseslint.configs.stylisticTypeChecked.reduce(
        (acc, config) => ({ ...acc, ...config.rules }),
        {}
      ),

      // TypeScript specific rules
      '@typescript-eslint/no-unused-vars': [
        'error',
        {
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
        },
      ],
      '@typescript-eslint/no-explicit-any': 'warn',
      '@typescript-eslint/explicit-function-return-type': 'off',
      '@typescript-eslint/explicit-module-boundary-types': 'off',
      '@typescript-eslint/no-non-null-assertion': 'warn',
      '@typescript-eslint/prefer-nullish-coalescing': 'error',
      '@typescript-eslint/prefer-optional-chain': 'error',
      '@typescript-eslint/no-unnecessary-type-assertion': 'error',
      '@typescript-eslint/no-floating-promises': 'error',
      '@typescript-eslint/await-thenable': 'error',

      // General code quality rules
      'no-console': 'warn',
      'no-debugger': 'error',
      'prefer-const': 'error',
      'no-var': 'error',
      'object-shorthand': 'error',
      'prefer-arrow-callback': 'error',
    },
  },

  // JavaScript files configuration (no TypeScript type checking)
  {
    files: ['**/*.{js,mjs,cjs}'],
    plugins: {
      '@typescript-eslint': tseslint.plugin,
    },
    languageOptions: {
      parser: tseslint.parser,
    },
    rules: {
      ...tseslint.configs.recommended.reduce((acc, config) => {
        if (config.rules) {
          const filteredRules = Object.entries(config.rules).reduce((ruleAcc, [key, value]) => {
            if (
              !key.includes('await-thenable') &&
              !key.includes('no-floating-promises') &&
              !key.includes('no-unnecessary-type-assertion') &&
              !key.includes('prefer-nullish-coalescing') &&
              !key.includes('prefer-optional-chain')
            ) {
              ruleAcc[key] = value;
            }
            return ruleAcc;
          }, {});
          return { ...acc, ...filteredRules };
        }
        return acc;
      }, {}),

      'no-console': 'warn',
      'no-debugger': 'error',
      'prefer-const': 'error',
      'no-var': 'error',
      'object-shorthand': 'error',
      'prefer-arrow-callback': 'error',
    },
  },

  // Office Add-in specific overrides
  {
    files: ['src/**/*.{ts,tsx,mjs}'],
    languageOptions: {
      globals: {
        Excel: 'readonly',
        Office: 'readonly',
        console: 'readonly',
        navigator: 'readonly',
      },
    },
    rules: {
      '@typescript-eslint/no-unused-vars': 'warn',
      '@typescript-eslint/no-explicit-any': 'warn',
      'no-undef': 'off',
      'no-console': 'off',
    },
  },

  // Test files configuration
  {
    files: ['**/*.test.{ts,tsx}', '**/*.spec.{ts,tsx}', '**/tests/**/*.{ts,tsx}'],
    plugins: {
      vitest,
    },
    languageOptions: {
      globals: {
        ...vitest.environments.env.globals,
      },
    },
    rules: {
      ...vitest.configs.recommended.rules,
      'vitest/no-disabled-tests': 'warn',
      'vitest/no-focused-tests': 'error',
      'vitest/no-identical-title': 'error',
      'vitest/prefer-to-have-length': 'warn',
      'vitest/valid-expect': 'error',
      'vitest/expect-expect': 'error',
      'vitest/no-standalone-expect': 'error',
      '@typescript-eslint/no-explicit-any': 'off',
      '@typescript-eslint/no-unsafe-assignment': 'off',
      '@typescript-eslint/no-unsafe-member-access': 'off',
      '@typescript-eslint/no-unsafe-call': 'off',
      'no-console': 'off',
    },
  },

  // Configuration files
  {
    files: ['**/*.config.{js,mjs,ts}', '**/eslint.config.mjs'],
    languageOptions: {
      globals: {
        ...globals.node,
      },
      parserOptions: {
        project: false,
      },
    },
    rules: {
      '@typescript-eslint/no-var-requires': 'off',
      '@typescript-eslint/await-thenable': 'off',
      '@typescript-eslint/no-floating-promises': 'off',
      '@typescript-eslint/no-unnecessary-type-assertion': 'off',
      '@typescript-eslint/prefer-nullish-coalescing': 'off',
      '@typescript-eslint/prefer-optional-chain': 'off',
    },
  },

  // Prettier integration (must be last)
  prettierConfig,
];
