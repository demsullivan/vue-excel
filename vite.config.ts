import { fileURLToPath, URL } from 'node:url'
import { resolve } from 'path'
import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [vue()],
  build: {
    lib: {
      entry: resolve(__dirname, 'src/index.ts'),
      name: 'vue-excel',
      fileName: 'vue-excel'
    },
    rollupOptions: {
      external: ['vue', 'lodash', 'moment-msdate'],
      output: {
        globals: {
          vue: 'Vue',
          'moment-msdate': 'moment'
        }
      }
    }
  },
  resolve: {
    alias: {
      '@vue-excel': fileURLToPath(new URL('./src', import.meta.url))
    }
  }
})
