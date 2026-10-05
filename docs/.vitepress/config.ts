import { defineConfig } from 'vitepress'

const english = [
  {
    text: 'Guide',
    items: [
      { text: 'Overview', link: '/' },
      { text: 'First report', link: '/getting-started' },
      { text: 'Templates', link: '/templates' },
      { text: 'Validation and lists', link: '/fields' },
      { text: 'Fill in a form', link: '/forms' },
    ],
  },
  {
    text: 'Reference',
    items: [
      { text: 'API', link: '/api' },
      { text: 'Excel and limitations', link: '/xlsx' },
    ],
  },
  {
    text: 'Contributing',
    collapsed: true,
    items: [
      { text: 'Architecture', link: '/architecture' },
      { text: 'Development', link: '/development' },
    ],
  },
]
const russian = [
  {
    text: 'Руководство',
    items: [
      { text: 'Обзор', link: '/ru/' },
      { text: 'Первый отчёт', link: '/ru/getting-started' },
      { text: 'Разметка шаблона', link: '/ru/templates' },
      { text: 'Проверки и списки', link: '/ru/fields' },
      { text: 'Заполнение формы', link: '/ru/forms' },
    ],
  },
  {
    text: 'Справочник',
    items: [
      { text: 'API', link: '/ru/api' },
      { text: 'Excel и ограничения', link: '/ru/xlsx' },
    ],
  },
  {
    text: 'Разработка библиотеки',
    collapsed: true,
    items: [
      { text: 'Архитектура', link: '/ru/architecture' },
      { text: 'Разработка', link: '/ru/development' },
    ],
  },
]

export default defineConfig({
  cleanUrls: true,
  markdown: {
    config(md) {
      const renderLink = md.renderer.rules.link_open
      md.renderer.rules.link_open = (tokens, index, options, env, self) => {
        // Download examples directly; VitePress otherwise treats .xlsx as a page.
        if (/^\/examples\/.*\.(xlsx|ts|json)$/.test(tokens[index].attrGet('href') ?? '')) {
          tokens[index].attrSet('download', '')
        }
        return renderLink?.(tokens, index, options, env, self) ?? self.renderToken(tokens, index, options)
      }
    },
  },
  head: [['link', { href: '/logo.svg', rel: 'icon', type: 'image/svg+xml' }]],
  lastUpdated: true,
  locales: {
    root: {
      title: 'Sheetbind',
      description: 'Workbook templates for XLSX reports and editable forms.',
      label: 'English',
      lang: 'en',
      themeConfig: { nav: [{ text: 'Guide', link: '/getting-started' }, { text: 'API', link: '/api' }], sidebar: english, outline: { level: [2, 3] } },
    },
    ru: {
      title: 'Sheetbind',
      description: 'Шаблоны книг для XLSX-отчётов и редактируемых форм.',
      label: 'Русский',
      lang: 'ru-RU',
      link: '/ru/',
      themeConfig: {
        nav: [{ text: 'Руководство', link: '/ru/getting-started' }, { text: 'API', link: '/ru/api' }],
        sidebar: russian,
        outline: { level: [2, 3], label: 'На этой странице' },
        docFooter: { prev: 'Назад', next: 'Далее' },
        lastUpdatedText: 'Обновлено',
        sidebarMenuLabel: 'Содержание',
        returnToTopLabel: 'Наверх',
      },
    },
  },
  themeConfig: { logo: '/logo.svg', search: { provider: 'local' } },
})
