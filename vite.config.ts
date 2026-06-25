import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import basicSsl from '@vitejs/plugin-basic-ssl';

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [
    react(),
    basicSsl()
  ],
  server: {
    port: 3000,
    https: true,
    // Izinkan CORS agar Excel online atau desktop bisa mengakses resource dengan aman jika diperlukan
    headers: {
      "Access-Control-Allow-Origin": "*",
    }
  }
});
