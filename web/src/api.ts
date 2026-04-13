/** Базовый URL API: в dev Vite проксирует /api на uvicorn */

const base =
  import.meta.env.VITE_API_URL?.replace(/\/$/, "") ?? "";

export type CatalogItem = {
  id?: number | null;
  excel_row: number;
  name: string;
  size: string;
  kind: string;
  file: string;
  price: string;
  price_new: string;
  qty_per_sheet: string;
  qty_per_year: string;
  gmp_code: string;
  updated_at: string;
};

export type CatalogResponse = {
  items: CatalogItem[];
  total: number;
};

export async function fetchCatalog(): Promise<CatalogResponse> {
  const path = `${base}/api/v1/items`;
  const res = await fetch(path);
  if (!res.ok) {
    throw new Error(`API ${res.status}: ${await res.text()}`);
  }
  return res.json() as Promise<CatalogResponse>;
}

export type HealthResponse = { status: string };

/** Живость API (прокси `/health` в dev). */
export async function fetchHealth(): Promise<HealthResponse> {
  const path = `${base}/health`;
  const res = await fetch(path);
  if (!res.ok) {
    throw new Error(`health ${res.status}`);
  }
  return res.json() as Promise<HealthResponse>;
}
