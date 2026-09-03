import { defineConfig } from 'vite';

export default defineConfig({
  // Use relative asset paths so the built site works under GitHub Pages
  // project subpaths (for example /My_Test/) and remains portable.
  base: './',
  build: {
    outDir: 'dist',
  },
});
