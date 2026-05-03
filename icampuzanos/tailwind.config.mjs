/** @type {import('tailwindcss').Config} */
export default {
  content: ['./src/**/*.{astro,html,js,jsx,md,mdx,ts,tsx}'],
  theme: {
    extend: {
      colors: {
        // Teal — color primario de UI (hero, footer, nav, headers)
        brand: {
          50:  '#EEF5F6',
          100: '#C8DDE0',
          200: '#90BAC0',
          300: '#5C97A0',
          400: '#4A7A80',   // Slate teal — secundario
          500: '#2D6068',
          600: '#1A4850',
          700: '#123840',   // Teal profundo
          800: '#0D2E38',
          900: '#08202A',
          950: '#041218',   // Negro teal ← hero / footer
        },
        // Dorado — acentos, CTAs, links, iconos
        gold: {
          50:  '#FBF5E0',
          100: '#F5E9C0',
          200: '#EDD380',
          300: '#E8C855',
          400: '#D4B030',
          500: '#C9A030',   // Dorado principal ← CTA
          600: '#A67C18',   // Hover
          700: '#8A6010',   // Texto sobre fondo claro
          800: '#6B4808',
          900: '#4A3005',
        },
      },
      fontFamily: {
        sans: ['Inter', 'system-ui', 'sans-serif'],
        serif: ['Playfair Display', 'Georgia', 'serif'],
      },
    },
  },
  plugins: [],
};
