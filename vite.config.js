import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// ✅ Single defineConfig call that includes both `plugins` and `build`
export default defineConfig({
  plugins: [react()],
  build: {
    rollupOptions: {
      external: ['xlsx']
    }
  }
});
