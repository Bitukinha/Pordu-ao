export const CATEGORIAS = ["Extrusão", "Flotação", "Exportação", "Germen", "Mercado interno", "Milho"];

// Dados históricos foram gravados com grafias diferentes (ex.: "Mercado Interno" vs "Mercado interno").
// Normaliza para a grafia canônica de CATEGORIAS, ignorando maiúsculas/minúsculas e espaços nas pontas.
export function normalizeCategoria(input: string): string {
  const trimmed = input.trim();
  const found = CATEGORIAS.find((c) => c.toLowerCase() === trimmed.toLowerCase());
  return found ?? trimmed;
}

export const PRODUTOS = [
  "Germen", "Fubá", "Fubá Mimoso", "Fubá 0,3", "Fubá (Fouani)", "Fubá (Heleng)", "Fubá (master)",
  "Fubá Pré-Cozido", "Pré Cozido", "Flocão", "Creme de Milho",
  "N-Form-D25", "N-Form-D48", "N-Form-F28", "N-Form-F35", "N-Form-F48", "N-Form-NT48",
  "Nutrigel", "Nutrigel Pro", "Grits Fino", "Grits Remoido",
  "Canjica Amarela", "Canjiquinha Fina", "Farinha Média", "Milho",
];

// Dados históricos foram gravados com grafias diferentes (ex.: "N-Form D25" / "N-Form-D25" / "N-form D25").
// Normaliza para a grafia canônica de PRODUTOS, ignorando maiúsculas/minúsculas e diferenças de espaço/hífen.
function produtoKey(s: string): string {
  return s.trim().toLowerCase().replace(/[-\s]+/g, "");
}

export function normalizeProduto(input: string): string {
  const trimmed = input.trim();
  const key = produtoKey(trimmed);
  const found = PRODUTOS.find((p) => produtoKey(p) === key);
  return found ?? trimmed;
}
