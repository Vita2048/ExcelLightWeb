import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { deploymentBase } from './config/base-path.ts';

export default defineConfig({
  plugins: [react()],
  base: deploymentBase(process.env.VITE_BASE_PATH),
  build: { target: 'es2022' },
});
