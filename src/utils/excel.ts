export function columnLetterFromIndex(index: number): string {
  let dividend = index + 1;
  let columnName = "";

  while (dividend > 0) {
    const modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }

  return columnName;
}

export function columnIndexFromLetter(letter: string): number | null {
  const normalized = letter.trim().toUpperCase();
  if (!/^[A-Z]+$/.test(normalized)) {
    return null;
  }

  let index = 0;
  for (let i = 0; i < normalized.length; i += 1) {
    index *= 26;
    index += normalized.charCodeAt(i) - 64;
  }

  return index - 1;
}

export function buildPlanList(items: string[]): string {
  return items.map((item, idx) => `${idx + 1}. ${item}`).join("\n");
}
