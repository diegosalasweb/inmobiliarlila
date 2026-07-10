import { defineConfig } from 'astro/config';
import sitemap from '@astrojs/sitemap';

export default defineConfig({
  site: 'https://inmobiliarlila.cl',
  output: 'static',
  integrations: [
    sitemap(),
  ],
});
