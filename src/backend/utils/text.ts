export function normalizeCzechText(text: string): string {
  return text
    .toLowerCase()
    .replace(/[_-]+/g, " ")
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .replace(/\s+/g, " ")
    .trim();
}
