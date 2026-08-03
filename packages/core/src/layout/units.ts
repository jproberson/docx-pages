export const EMU_PER_POINT = 12700;
export const TWIPS_PER_POINT = 20;

export const emuToPoints = (emu: number): number => emu / EMU_PER_POINT;
export const twipsToPoints = (twips: number): number => twips / TWIPS_PER_POINT;
