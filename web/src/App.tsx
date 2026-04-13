import type { CSSProperties } from "react";
import { useEffect, useMemo, useState } from "react";
import { fetchCatalog, fetchHealth, type CatalogItem } from "./api";

function useCatalog() {
  const [items, setItems] = useState<CatalogItem[]>([]);
  const [total, setTotal] = useState(0);
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const data = await fetchCatalog();
        if (!cancelled) {
          setItems(data.items);
          setTotal(data.total);
        }
      } catch (e) {
        if (!cancelled) {
          setErr(e instanceof Error ? e.message : String(e));
        }
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  return { items, total, loading, err };
}

function useApiHealth() {
  const [status, setStatus] = useState<"pending" | "ok" | "err">("pending");
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const h = await fetchHealth();
        if (!cancelled && h.status === "ok") setStatus("ok");
        else if (!cancelled) setStatus("err");
      } catch {
        if (!cancelled) setStatus("err");
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);
  return status;
}

export default function App() {
  const { items, total, loading, err } = useCatalog();
  const health = useApiHealth();
  const [q, setQ] = useState("");

  const filtered = useMemo(() => {
    const s = q.trim().toLowerCase();
    if (!s) return items;
    return items.filter(
      (r) =>
        String(r.excel_row).includes(s) ||
        r.name.toLowerCase().includes(s) ||
        r.gmp_code.toLowerCase().includes(s) ||
        r.file.toLowerCase().includes(s)
    );
  }, [items, q]);

  return (
    <div style={{ padding: "1.25rem 1.5rem", maxWidth: 1400, margin: "0 auto" }}>
      <header style={{ marginBottom: "1rem" }}>
        <h1 style={{ margin: "0 0 0.25rem", fontSize: "1.35rem" }}>
          Каталог макетов
        </h1>
        <p style={{ margin: 0, color: "#555", fontSize: "0.9rem" }}>
          Данные с read-only API · всего строк: {loading ? "…" : total}
          {" · "}
          <span
            title="GET /health"
            style={{
              color:
                health === "ok" ? "#2e7d32" : health === "err" ? "#c62828" : "#757575",
            }}
          >
            API:{" "}
            {health === "pending"
              ? "проверка…"
              : health === "ok"
                ? "доступен"
                : "недоступен"}
          </span>
        </p>
      </header>

      <label style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 12 }}>
        <span style={{ fontSize: "0.875rem" }}>Поиск</span>
        <input
          type="search"
          value={q}
          onChange={(e) => setQ(e.target.value)}
          placeholder="строка Excel, имя, GMP, файл…"
          style={{
            flex: 1,
            maxWidth: 420,
            padding: "0.4rem 0.6rem",
            borderRadius: 6,
            border: "1px solid #ccc",
          }}
        />
      </label>

      {loading && <p>Загрузка…</p>}
      {err && (
        <p style={{ color: "#b71c1c" }}>
          {err} — запустите API:{" "}
          <code style={{ fontSize: "0.85rem" }}>
            uvicorn api.main:app --reload
          </code>
        </p>
      )}

      {!loading && !err && (
        <div
          style={{
            overflow: "auto",
            borderRadius: 8,
            border: "1px solid #ddd",
            background: "#fff",
            boxShadow: "0 1px 3px rgba(0,0,0,0.06)",
          }}
        >
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.8125rem" }}>
            <thead>
              <tr style={{ background: "#eceff1", textAlign: "left" }}>
                <th style={th}># Excel</th>
                <th style={th}>GMP</th>
                <th style={th}>Наименование</th>
                <th style={th}>Вид</th>
                <th style={th}>Размер</th>
                <th style={th}>PDF</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((r) => (
                <tr key={r.excel_row} style={{ borderTop: "1px solid #eee" }}>
                  <td style={td}>{r.excel_row}</td>
                  <td style={td}>{r.gmp_code}</td>
                  <td style={td}>{r.name}</td>
                  <td style={td}>{r.kind}</td>
                  <td style={td}>{r.size}</td>
                  <td style={{ ...td, maxWidth: 220, wordBreak: "break-all" }}>{r.file}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {filtered.length === 0 && (
            <p style={{ padding: "1rem", margin: 0, color: "#666" }}>Нет строк.</p>
          )}
        </div>
      )}
    </div>
  );
}

const th: CSSProperties = {
  padding: "0.5rem 0.65rem",
  fontWeight: 600,
  whiteSpace: "nowrap",
};
const td: CSSProperties = { padding: "0.4rem 0.65rem", verticalAlign: "top" };
