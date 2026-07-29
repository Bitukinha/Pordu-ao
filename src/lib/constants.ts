export const CATEGORIAS = ["Extrusão", "Flotação", "Exportação", "Germen", "Mercado interno", "Milho"];

// Dados históricos foram gravados com grafias diferentes (ex.: "Mercado Interno" vs "Mercado interno").
// Normaliza para a grafia canônica de CATEGORIAS, ignorando maiúsculas/minúsculas e espaços nas pontas.
export function normalizeCategoria(input: string): string {
  const trimmed = input.trim();
  const found = CATEGORIAS.find((c) => c.toLowerCase() === trimmed.toLowerCase());
  return found ?? trimmed;
}
