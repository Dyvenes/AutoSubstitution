/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    "./templates/**/*.html",  // все HTML файлы в папке templates
    "./static/**/*.js",       // если есть JavaScript файлы
  ],
  theme: {
    extend: {},
  },
  plugins: [],
}

