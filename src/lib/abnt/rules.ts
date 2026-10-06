// Conversão de centímetros para twips (unidade usada pela lib docx)
export function cmToTwip(cm: number): number {
  return Math.round((cm / 2.54) * 1440);
}

export const ABNT_RULES = {
  margin: {
    topTwip: cmToTwip(3),
    leftTwip: cmToTwip(3),
    bottomTwip: cmToTwip(2),
    rightTwip: cmToTwip(2),
  },
  font: {
    family: "Arial",
    sizeBodyHalfPt: 24, // 12pt (a lib docx usa "meio-ponto": 12 * 2)
    sizeSmallHalfPt: 20, // 10pt (citações longas, notas)
  },
  spacing: {
    bodyLine: 360, // 1,5 entrelinhas (240 = simples)
    singleLine: 240, // espaço simples
  },
  indent: {
    firstLineTwip: cmToTwip(1.25),
    quoteLeftTwip: cmToTwip(4),
  },
};