/** @type {import('tailwindcss').Config} */
export default {
  darkMode: 'class',
  content: ['./index.html', './src/**/*.{js,jsx}'],
  theme: {
    extend: {
      boxShadow: {
        soft: '0 24px 70px rgba(15,23,42,.12)',
        paper: '0 18px 50px rgba(15,23,42,.10)'
      },
      fontFamily: {
        doc: ['Inter','Helvetica Neue','Helvetica','Arial','PingFang TC','Noto Sans TC','sans-serif']
      }
    }
  },
  plugins: []
}
