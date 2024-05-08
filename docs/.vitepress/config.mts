import { defineConfig } from 'vitepress'

// https://vitepress.dev/reference/site-config
export default defineConfig({
  title: 'Vue Excel',
  description: 'A vue plugin for building modern, reactive Excel Add-ins',
  base: '/vue-excel/',
  themeConfig: {
    // https://vitepress.dev/reference/default-theme-config
    nav: [
      { text: 'Home', link: '/' },
      { text: 'Guide', link: '/guide/' }
    ],

    sidebar: [
      {
        text: 'Setup',
        items: [
          { text: 'Introduction', link: '/introduction' },
          { text: 'Installation', link: '/installation' }
        ]
      },
      {
        text: 'Basics',
        items: [{ text: 'Getting Started', link: '/guide/' }]
      },
      {
        text: 'Routing',
        items: [{ text: 'Simple', link: '/routing/simple' }]
      }
    ],

    socialLinks: [{ icon: 'github', link: 'https://github.com/demsullivan/vue-excel' }]
  }
})
