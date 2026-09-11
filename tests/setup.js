/**
 * Vitest 测试环境设置
 */
import { vi } from 'vitest';

// jsdom already provides localStorage, just clear it before each test
beforeEach(() => {
  localStorage.clear();
});

// Suppress console output in tests to reduce noise
global.console = {
  ...console,
  log: vi.fn(),
  debug: vi.fn(),
  info: vi.fn(),
};
