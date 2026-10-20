import {defineConfig} from 'vite';
import {fileURLToPath} from 'node:url';

// Two pages: the dashboard, and the small page Microsoft returns to after sign-in.
export default defineConfig({
  build: {
    rolldownOptions: {
      input: {
        main: fileURLToPath(new URL('./index.html', import.meta.url)),
        auth: fileURLToPath(new URL('./auth/index.html', import.meta.url)),
      },
    },
  },
});
