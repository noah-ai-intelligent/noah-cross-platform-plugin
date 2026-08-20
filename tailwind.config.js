/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./src/**/*.{html,js,ts,jsx,tsx}",
    "./*.html"
  ],
  theme: {
    extend: {
      colors: {
        canvas: "#ffffff",
        surface: "#fafafa",
        "surface-hover": "#f4f4f5",
        border: "#e4e4e7",
        "border-strong": "#d4d4d8",
        ink: "#18181b",
        "ink-secondary": "#52525b",
        "ink-muted": "#71717a",
        "ink-placeholder": "#a1a1aa",
        accent: "#2eb369",
        "accent-hover": "#24955a",
        "accent-soft": "#e6f6ed",
        danger: "#dc2626",
        "danger-soft": "#fef2f2",
        warn: "#b45309",
        "disabled-bg": "#e4e4e7",
        "disabled-ink": "#a1a1aa"
      }
    },
  },
  plugins: [],
}
