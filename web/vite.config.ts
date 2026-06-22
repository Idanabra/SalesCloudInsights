import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import vuetify from 'vite-plugin-vuetify'
import basicSsl from '@vitejs/plugin-basic-ssl'

export default defineConfig({
  plugins: [
    basicSsl(),   // self-signed cert — Outlook add-ins require HTTPS even in dev
    vue(),
    vuetify({ autoImport: true }),
  ],
  server: {
    https: true,
    port:  5173,
  },
  build: {
    outDir: 'dist',
  },
})
