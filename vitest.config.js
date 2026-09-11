import { defineConfig } from 'vitest/config';
import path from 'path';

export default defineConfig({
  test: {
    globals: true,
    environment: 'jsdom',
    setupFiles: './tests/setup.js',
    coverage: {
      provider: 'v8',
      reporter: ['text', 'json', 'html'],
      exclude: [
        'node_modules/**',
        'tests/**',
        'dist/**',
        'webpack.*.js',
        'vitest.config.js',
        'scripts/**',
      ],
    },
  },
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './js'),
    },
  },
});
