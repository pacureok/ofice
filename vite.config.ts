import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  // Configuración opcional para base path en Netlify si usas rutas absolutas
  base: '/', 
});
