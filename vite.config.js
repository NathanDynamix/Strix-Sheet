import{defineConfig} from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins : [react()],
  server : {
    host : "0.0.0.0",
    port : 5002,
    proxy : {
      "/api" : {
        target : "http://localhost:2000",
        changeOrigin : true,
        secure : false
      }
    }
  }
});
