const tintColorLight = "#134e4a";
const tintColorDark = "#ccfbf1";
const slate50 = "#f8fafc";
const slate100 = "#f1f5f9";
const slate200 = "#e2e8f0";
const slate300 = "#cbd5e1";
const slate400 = "#94a3b8";
const slate500 = "#64748b";
const slate600 = "#475569";
const slate700 = "#334155";
const slate800 = "#1e293b";
const slate900 = "#0f172a";
const slate950 = "#020617";

// from tailwind color https://tailwindcss.com/docs/customizing-colors
const red50 = "#fef2f2";
const red100 = "#fee2e2";
const red200 = "#fecaca";
const red300 = "#fca5a5";
const red400 = "#f87171";
const red500 = "#ef4444";
const red600 = "#dc2626";
const red700 = "#b91c1c";
const red800 = "#991b1b";
const red900 = "#7f1d1d";
const red950 = "#450a0a";
const green50 = "#f0fdf4";
const green100 = "#dcfce7";
const green200 = "#bbf7d0";
const green300 = "#86efac";
const green400 = "#4ade80";
const green500 = "#22c55e";
const green600 = "#16a34a";
const green700 = "#15803d";
const green800 = "#166534";
const green900 = "#14532d";
const green950 = "#052e16";
const sky50 = "#f0f9ff";
const sky100 = "#e0f2fe";
const sky200 = "#bae6fd";
const sky300 = "#7dd3fc";
const sky400 = "#38bdf8";
const sky500 = "#0ea5e9";
const sky600 = "#0284c7";
const sky700 = "#0369a1";
const sky800 = "#075985";
const sky900 = "#0c4a6e";
const sky950 = "#082f49";

const teal50 = "#f0fdfa";
const teal100 = "#ccfbf1";
const teal200 = "#99f6e4";
const teal300 = "#5eead4";
const teal400 = "#2dd4bf";
const teal500 = "#14b8a6";
const teal600 = "#0d9488";
const teal700 = "#0f766e";
const teal800 = "#115e59";
const teal900 = "#134e4a";
const teal950 = "#042f2e";

export default {
  light: {
    text: "#000",
    slateText: slate600,
    tealText: teal900,
    background: "#fff",
    slateBackground: slate100,
    tealBackground: teal50,
    border: slate300,
    tint: tintColorLight,
    tabIconDefault: "#ccc",
    tabIconSelected: tintColorLight,
    homeTabDefault: "#fff",
    homeTabSelected: "rgba(226,232,240,1)",
    shadowColor: "#000",
    icon: slate600,
    button: "#000022",
    buttonText: "#fff",
    dangerButton: "red",
    dangerButtonText: "white",
    signInMainColor: teal900,
    signInTextColor: slate100,
  },
  dark: {
    text: "#fff",
    slateText: slate400,
    tealText: teal100,
    background: "#000022",
    slateBackground: slate900,
    tealBackground: teal950,
    border: slate700,
    tint: tintColorDark,
    tabIconDefault: "#ccc",
    tabIconSelected: tintColorDark,
    homeTabDefault: "#000022",
    homeTabSelected: "rgba(51,65,85,1)",
    shadowColor: "#fff",
    icon: slate400,
    button: "#fff",
    buttonText: "#000",
    dangerButton: "red",
    dangerButtonText: "white",
    signInMainColor: teal100,
    signInTextColor: slate900,
  },
  slate50,
  slate100,
  slate200,
  slate300,
  slate400,
  slate500,
  slate600,
  slate700,
  slate800,
  slate900,
  slate950,
  red50,
  red100,
  red200,
  red300,
  red400,
  red500,
  red600,
  red700,
  red800,
  red900,
  red950,
  green50,
  green100,
  green200,
  green300,
  green400,
  green500,
  green600,
  green700,
  green800,
  green900,
  green950,
  sky50,
  sky100,
  sky200,
  sky300,
  sky400,
  sky500,
  sky600,
  sky700,
  sky800,
  sky900,
  sky950,
  teal50,
  teal100,
  teal200,
  teal300,
  teal400,
  teal500,
  teal600,
  teal700,
  teal800,
  teal900,
  teal950,
};

export const hexToRgb = (hex: string) => {
  const bigint = parseInt(hex.slice(1), 16);
  const r = (bigint >> 16) & 255;
  const g = (bigint >> 8) & 255;
  const b = bigint & 255;

  return [r, g, b];
};

export const addOpacityToHex = (hex: string, opacity: number) => {
  const [r, g, b] = hexToRgb(hex);
  return `rgba(${r},${g},${b},${opacity})`;
};
