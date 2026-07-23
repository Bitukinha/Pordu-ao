import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useRef, useState } from "react";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, LabelList,
} from "recharts";
import { addEntries, addEntry, clearEntries, deleteEntry, listEntries } from "@/lib/entries";
import logo from "@/assets/nutrimilho-logo.png";

const MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];

function excelSerialToISO(n: number): string {
  const dt = new Date(Math.round((n - 25569) * 86400 * 1000));
  return dt.toISOString().slice(0, 10);
}

export const Route = createFileRoute("/")({
  component: App,
});

type Entry = {
  id: string;
  data: string; // YYYY-MM-DD
  categoria: string;
  produto: string;
  qteTon: number;
};

type NewEntry = Omit<Entry, "id">;

const CATEGORIAS = ["Extrusão", "Flotação", "Exportação", "Germen", "Mercado interno", "Milho"];
const PRODUTOS = [
  "Germen", "Fubá", "Fubá Mimoso", "Pré Cozido", "Flocão",
  "N-Form-D25", "N-Form-D48", "N-Form-F28", "N-Form-F35", "N-Form-F48",
  "Nutrigel", "Nutrigel Pro", "Grits Fino", "Grits Remoido",
  "Canjica Amarela", "Canjiquinha Fina", "Farinha Média", "Milho",
];

// palette derived from Nutrimilho logo (greens/yellows/warm)
const CAT_COLORS: Record<string, string> = {
  "Extrusão": "#E8833A",
  "Flotação": "#8FA6A2",
  "Exportação": "#3E7CB1",
  "Germen": "#F2B807",
  "Mercado interno": "#7CB342",
  "Milho": "#C0392B",
};
const DATE_PALETTE = ["#3E7CB1", "#E8833A", "#8FA6A2", "#F2B807", "#5DADE2", "#7CB342", "#C0392B", "#9B59B6"];

function fmt(n: number) {
  return n.toLocaleString("pt-BR", { maximumFractionDigits: 2 });
}

function App() {
  const [tab, setTab] = useState<"dashboard" | "dados">("dashboard");
  const queryClient = useQueryClient();

  const { data: entries = [], isLoading } = useQuery({
    queryKey: ["entries"],
    queryFn: () => listEntries(),
  });

  function invalidate() {
    return queryClient.invalidateQueries({ queryKey: ["entries"] });
  }

  const addMutation = useMutation({
    mutationFn: (entry: NewEntry) => addEntry({ data: entry }),
    onSuccess: invalidate,
  });
  const addManyMutation = useMutation({
    mutationFn: (entries: NewEntry[]) => addEntries({ data: entries }),
    onSuccess: invalidate,
  });
  const removeMutation = useMutation({
    mutationFn: (id: string) => deleteEntry({ data: { id } }),
    onSuccess: invalidate,
  });
  const clearMutation = useMutation({
    mutationFn: () => clearEntries(),
    onSuccess: invalidate,
  });

  return (
    <div className="flex min-h-screen flex-col bg-background">
      <header className="border-b bg-white">
        <div className="mx-auto flex max-w-7xl flex-wrap items-center justify-between gap-3 px-4 py-3 sm:gap-4 sm:px-6 sm:py-4">
          <div className="flex items-center gap-2 sm:gap-3">
            <img src={logo} alt="Nutrimilho" className="h-8 w-auto sm:h-10" />
            <span className="hidden text-xs text-muted-foreground sm:inline">Controle de Produção Diária</span>
          </div>
          <nav className="flex gap-1 rounded-lg bg-muted p-1">
            {(["dashboard", "dados"] as const).map((t) => (
              <button
                key={t}
                onClick={() => setTab(t)}
                className={`rounded-md px-2.5 py-1.5 text-xs font-medium transition sm:px-4 sm:text-sm ${
                  tab === t ? "bg-primary text-primary-foreground shadow-sm" : "text-muted-foreground hover:text-foreground"
                }`}
              >
                {t === "dashboard" ? "Indicadores" : "Lançar Dados"}
              </button>
            ))}
          </nav>
        </div>
      </header>

      <main className="mx-auto w-full max-w-7xl flex-1 px-4 py-6 sm:px-6 sm:py-8">
        {isLoading ? (
          <p className="py-12 text-center text-sm text-muted-foreground">Carregando...</p>
        ) : tab === "dashboard" ? (
          <Dashboard entries={entries} />
        ) : (
          <DataEntry
            entries={entries}
            onAdd={(entry) => addMutation.mutateAsync(entry)}
            onImport={(imported) => addManyMutation.mutateAsync(imported)}
            onRemove={(id) => removeMutation.mutateAsync(id)}
            onClear={() => clearMutation.mutateAsync()}
          />
        )}
      </main>

      <footer className="border-t bg-white py-4">
        <p className="px-4 text-center text-xs text-muted-foreground">
          © 2026 Nutrimilho - (Novaes Tech) | Todos os direitos reservados
        </p>
      </footer>
    </div>
  );
}

function exportXlsx(entries: Entry[]) {
  const rows = entries
    .slice()
    .sort((a, b) => a.data.localeCompare(b.data))
    .map((e) => ({
      Data: new Date(e.data + "T00:00"),
      Categoria: e.categoria,
      Produto: e.produto,
      "Qte Ton": e.qteTon,
    }));
  const ws = XLSX.utils.json_to_sheet(rows, { cellDates: true });
  ws["!cols"] = [{ wch: 12 }, { wch: 18 }, { wch: 22 }, { wch: 12 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Dados");
  XLSX.writeFile(wb, `nutrimilho-producao-${new Date().toISOString().slice(0, 10)}.xlsx`);
}

/* -------------------- DATA ENTRY -------------------- */
function DataEntry({
  entries,
  onAdd,
  onImport,
  onRemove,
  onClear,
}: {
  entries: Entry[];
  onAdd: (entry: NewEntry) => Promise<unknown>;
  onImport: (entries: NewEntry[]) => Promise<unknown>;
  onRemove: (id: string) => Promise<unknown>;
  onClear: () => Promise<unknown>;
}) {
  const today = new Date().toISOString().slice(0, 10);
  const [data, setData] = useState(today);
  const [categoria, setCategoria] = useState(CATEGORIAS[0]);
  const [produto, setProduto] = useState(PRODUTOS[0]);
  const [qteTon, setQteTon] = useState("");
  const fileRef = useRef<HTMLInputElement>(null);
  const [importMsg, setImportMsg] = useState<string | null>(null);

  async function importXlsx(file: File) {
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { cellDates: false });
      const wsName =
        wb.SheetNames.find((n) => n.trim().toLowerCase() === "dados") ||
        wb.SheetNames.find((n) => n.trim().toLowerCase().startsWith("dados")) ||
        wb.SheetNames.find((n) => n.trim().toLowerCase().includes("dado"));
      if (!wsName) {
        setImportMsg("Aba 'Dados' não encontrada no arquivo.");
        return;
      }
      const ws = wb.Sheets[wsName];
      const rawRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { raw: true, defval: null });
      // normalize keys: trim + lowercase, so " Qte Ton " matches "qte ton"
      const rows = rawRows.map((r) => {
        const o: Record<string, unknown> = {};
        for (const k of Object.keys(r)) o[k.trim().toLowerCase()] = r[k];
        return o;
      });
      const pick = (r: Record<string, unknown>, keys: string[]) => {
        for (const k of keys) if (r[k] != null && r[k] !== "") return r[k];
        return null;
      };
      const imported: NewEntry[] = [];
      let skipped = 0;
      for (const r of rows) {
        const d = pick(r, ["data", "dt", "dia"]);
        const cat = pick(r, ["categoria", "cat"]);
        const prod = pick(r, ["produto"]);
        const q = pick(r, ["qte ton", "qtd ton", "quantidade", "qte_ton", "qte", "qtd"]);
        if (d == null || !cat || !prod || q == null || isNaN(Number(q))) { skipped++; continue; }
        let iso: string;
        if (d instanceof Date) iso = new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
        else if (typeof d === "number") iso = excelSerialToISO(d);
        else {
          const s = String(d);
          const m = s.match(/^(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})/);
          if (m) {
            const yy = m[3].length === 2 ? "20" + m[3] : m[3];
            iso = `${yy}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
          } else iso = s.slice(0, 10);
        }
        imported.push({
          data: iso,
          categoria: String(cat).trim(),
          produto: String(prod).trim(),
          qteTon: Number(q),
        });
      }
      await onImport(imported);
      setImportMsg(`${imported.length} lançamentos importados de "${wsName}"${skipped ? ` (${skipped} linhas ignoradas)` : ""}.`);
    } catch (err) {
      setImportMsg("Erro ao ler o arquivo: " + (err as Error).message);
    }
    if (fileRef.current) fileRef.current.value = "";
  }

  async function submit(e: React.FormEvent) {
    e.preventDefault();
    const n = parseFloat(qteTon.replace(",", "."));
    if (!data || !categoria || !produto || isNaN(n)) return;
    await onAdd({ data, categoria, produto, qteTon: n });
    setQteTon("");
  }

  async function clearAll() {
    if (confirm("Excluir todos os lançamentos?")) await onClear();
  }


  async function remove(id: string) {
    await onRemove(id);
  }

  return (
    <div className="grid gap-6 lg:grid-cols-[380px_1fr]">
      <section className="rounded-xl border bg-card p-6 shadow-sm">
        <h2 className="mb-4 text-base font-semibold">Novo Lançamento</h2>
        <form onSubmit={submit} className="space-y-4">
          <Field label="Data">
            <input type="date" value={data} onChange={(e) => setData(e.target.value)} className={inputCls} required />
          </Field>
          <Field label="Categoria">
            <select value={categoria} onChange={(e) => setCategoria(e.target.value)} className={inputCls}>
              {CATEGORIAS.map((c) => <option key={c}>{c}</option>)}
            </select>
          </Field>
          <Field label="Produto">
            <select value={produto} onChange={(e) => setProduto(e.target.value)} className={inputCls}>
              {PRODUTOS.map((p) => <option key={p}>{p}</option>)}
            </select>
          </Field>
          <Field label="Quantidade (Ton)">
            <input
              type="text" inputMode="decimal" value={qteTon}
              onChange={(e) => setQteTon(e.target.value)}
              placeholder="0,00" className={inputCls} required
            />
          </Field>
          <button type="submit" className="w-full rounded-md bg-primary px-4 py-2.5 text-sm font-semibold text-primary-foreground transition hover:opacity-90">
            Adicionar Lançamento
          </button>
        </form>

        <div className="mt-6 border-t pt-4">
          <p className="mb-2 text-xs font-medium uppercase text-muted-foreground">Importar Excel</p>
          <input
            ref={fileRef}
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => e.target.files?.[0] && importXlsx(e.target.files[0])}
            className="block w-full text-xs file:mr-3 file:rounded-md file:border-0 file:bg-brand-leaf file:px-3 file:py-2 file:text-xs file:font-semibold file:text-white hover:file:opacity-90"
          />
          <p className="mt-2 text-[11px] text-muted-foreground">Lê a aba <strong>Dados</strong> (colunas: Data, Categoria, Produto, Qte Ton).</p>
          {importMsg && <p className="mt-2 text-xs text-foreground">{importMsg}</p>}
          {entries.length > 0 && (
            <div className="mt-4 flex flex-col gap-2 border-t pt-4">
              <p className="text-xs font-medium uppercase text-muted-foreground">Exportar Excel</p>
              <button
                onClick={() => exportXlsx(entries)}
                className="w-full rounded-md bg-brand-yellow px-3 py-2 text-xs font-semibold text-foreground transition hover:opacity-90"
              >
                Baixar planilha (.xlsx)
              </button>
              <button onClick={clearAll} className="mt-1 text-xs text-destructive hover:underline">Excluir todos os lançamentos</button>
            </div>
          )}
        </div>
      </section>


      <section className="rounded-xl border bg-card p-6 shadow-sm">
        <h2 className="mb-4 text-base font-semibold">Lançamentos ({entries.length})</h2>
        {entries.length === 0 ? (
          <p className="py-12 text-center text-sm text-muted-foreground">Nenhum lançamento ainda.</p>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b text-left text-xs uppercase text-muted-foreground">
                  <th className="py-2 pr-4">Data</th>
                  <th className="py-2 pr-4">Categoria</th>
                  <th className="py-2 pr-4">Produto</th>
                  <th className="py-2 pr-4 text-right">Qte (Ton)</th>
                  <th className="py-2"></th>
                </tr>
              </thead>
              <tbody>
                {entries.map((e) => (
                  <tr key={e.id} className="border-b last:border-0 hover:bg-muted/50">
                    <td className="py-2 pr-4">{new Date(e.data + "T00:00").toLocaleDateString("pt-BR")}</td>
                    <td className="py-2 pr-4">
                      <span className="inline-block h-2 w-2 rounded-full" style={{ background: CAT_COLORS[e.categoria] || "#999" }} /> {e.categoria}
                    </td>
                    <td className="py-2 pr-4">{e.produto}</td>
                    <td className="py-2 pr-4 text-right font-medium tabular-nums">{fmt(e.qteTon)}</td>
                    <td className="py-2 text-right">
                      <button onClick={() => remove(e.id)} className="text-xs text-destructive hover:underline">Excluir</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </section>
    </div>
  );
}

const inputCls =
  "w-full rounded-md border border-input bg-white px-3 py-2 text-sm outline-none focus:border-primary focus:ring-2 focus:ring-primary/20";

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <label className="block">
      <span className="mb-1 block text-xs font-medium text-muted-foreground">{label}</span>
      {children}
    </label>
  );
}

/* -------------------- DASHBOARD -------------------- */
function Dashboard({ entries }: { entries: Entry[] }) {
  // Available years/months from data
  const years = useMemo(
    () => Array.from(new Set(entries.map((e) => e.data.slice(0, 4)))).sort(),
    [entries]
  );
  const [ano, setAno] = useState<string>("");
  const [mes, setMes] = useState<string>(""); // "" = todos, "01".."12"
  const [dtIni, setDtIni] = useState<string>("");
  const [dtFim, setDtFim] = useState<string>("");
  const [categoriasSel, setCategoriasSel] = useState<string[]>([]); // [] = todas

  useEffect(() => {
    if (years.length && !years.includes(ano)) setAno(years[years.length - 1]);
  }, [years, ano]);

  const monthsForYear = useMemo(() => {
    const s = new Set(entries.filter((e) => e.data.startsWith(ano)).map((e) => e.data.slice(5, 7)));
    return Array.from(s).sort();
  }, [entries, ano]);

  const usaPeriodo = Boolean(dtIni || dtFim);

  const hoje = new Date().toISOString().slice(0, 10);
  const seteDiasAtras = new Date(Date.now() - 6 * 86400000).toISOString().slice(0, 10);
  const anoAtual = String(new Date().getFullYear());
  const mesAtual = String(new Date().getMonth() + 1).padStart(2, "0");
  const presetSeteDias = usaPeriodo && dtIni === seteDiasAtras && dtFim === hoje;
  const presetMensal = !usaPeriodo && ano === anoAtual && mes === mesAtual;
  const presetAnual = !usaPeriodo && ano === anoAtual && mes === "";

  function aplicarUltimos7Dias() {
    setDtIni(seteDiasAtras);
    setDtFim(hoje);
  }
  function aplicarMensal() {
    setDtIni("");
    setDtFim("");
    setAno(years.includes(anoAtual) ? anoAtual : years[years.length - 1] ?? anoAtual);
    setMes(mesAtual);
  }
  function aplicarAnual() {
    setDtIni("");
    setDtFim("");
    setAno(years.includes(anoAtual) ? anoAtual : years[years.length - 1] ?? anoAtual);
    setMes("");
  }
  function toggleCategoria(c: string) {
    setCategoriasSel((prev) => (prev.includes(c) ? prev.filter((x) => x !== c) : [...prev, c]));
  }

  const filtered = useMemo(
    () =>
      entries.filter((e) => {
        if (categoriasSel.length && !categoriasSel.includes(e.categoria)) return false;
        if (usaPeriodo) {
          if (dtIni && e.data < dtIni) return false;
          if (dtFim && e.data > dtFim) return false;
          return true;
        }
        if (ano && !e.data.startsWith(ano)) return false;
        if (mes && e.data.slice(5, 7) !== mes) return false;
        return true;
      }),
    [entries, ano, mes, dtIni, dtFim, usaPeriodo, categoriasSel]
  );

  const sortedDates = useMemo(
    () => Array.from(new Set(filtered.map((e) => e.data))).sort(),
    [filtered]
  );


  // Chart 1: Produção diária por produto — grouped bars per produto, series = data
  const chart1 = useMemo(() => {
    const produtos = Array.from(new Set(filtered.map((e) => `${e.categoria}||${e.produto}`))).sort();
    return produtos.map((key) => {
      const [categoria, produto] = key.split("||");
      const row: Record<string, string | number> = { categoria, produto, label: `${categoria}\n${produto}` };
      for (const d of sortedDates) {
        const total = filtered
          .filter((e) => e.data === d && e.categoria === categoria && e.produto === produto)
          .reduce((s, e) => s + e.qteTon, 0);
        if (total) row[d] = total;
      }
      return row;
    });
  }, [filtered, sortedDates]);

  // Chart 2: Total diário — stacked bar per data, series = categoria
  const chart2 = useMemo(() => {
    return sortedDates.map((d) => {
      const row: Record<string, string | number> = { data: d };
      let total = 0;
      for (const c of CATEGORIAS) {
        const v = filtered.filter((e) => e.data === d && e.categoria === c).reduce((s, e) => s + e.qteTon, 0);
        if (v) row[c] = v;
        total += v;
      }
      row.__total = total;
      return row;
    });
  }, [filtered, sortedDates]);


  const grandTotal = chart2.reduce((s, r) => s + (r.__total as number), 0);
  const activeCats = CATEGORIAS.filter((c) => chart2.some((r) => r[c]));

  if (entries.length === 0) {
    return (
      <div className="rounded-xl border-2 border-dashed bg-card p-16 text-center">
        <p className="text-lg font-medium text-foreground">Nenhum dado ainda</p>
        <p className="mt-1 text-sm text-muted-foreground">Vá em <strong>Lançar Dados</strong> para começar a inserir a produção.</p>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      {/* Filters */}
      <div className="space-y-3 rounded-xl border bg-card p-4 shadow-sm">
        <div className="flex flex-wrap items-center gap-2">
          <span className="text-xs font-semibold uppercase text-muted-foreground">Período rápido</span>
          <button
            onClick={aplicarUltimos7Dias}
            className={`rounded-md px-3 py-1 text-xs font-medium transition ${
              presetSeteDias ? "bg-primary text-primary-foreground" : "bg-muted text-muted-foreground hover:text-foreground"
            }`}
          >
            Últimos 7 dias
          </button>
          <button
            onClick={aplicarMensal}
            className={`rounded-md px-3 py-1 text-xs font-medium transition ${
              presetMensal ? "bg-primary text-primary-foreground" : "bg-muted text-muted-foreground hover:text-foreground"
            }`}
          >
            Mensal
          </button>
          <button
            onClick={aplicarAnual}
            className={`rounded-md px-3 py-1 text-xs font-medium transition ${
              presetAnual ? "bg-primary text-primary-foreground" : "bg-muted text-muted-foreground hover:text-foreground"
            }`}
          >
            Anual
          </button>
        </div>

        <div className="flex flex-wrap items-end gap-3">
          <span className="text-xs font-semibold uppercase text-muted-foreground">Filtro</span>
          <label className="flex items-center gap-2 text-sm">
            <span className="text-xs text-muted-foreground">Ano</span>
            <select
              value={ano}
              onChange={(e) => { setAno(e.target.value); setMes(""); setDtIni(""); setDtFim(""); }}
              className={inputCls + " w-28"}
              disabled={usaPeriodo}
            >
              {years.map((y) => <option key={y}>{y}</option>)}
            </select>
          </label>
          <label className="flex items-center gap-2 text-sm">
            <span className="text-xs text-muted-foreground">Mês</span>
            <select
              value={mes}
              onChange={(e) => setMes(e.target.value)}
              className={inputCls + " w-36"}
              disabled={usaPeriodo}
            >
              <option value="">Todos</option>
              {monthsForYear.map((m) => <option key={m} value={m}>{MESES[parseInt(m, 10) - 1]}</option>)}
            </select>
          </label>
          <div className="mx-2 hidden h-8 w-px bg-border sm:block" />
          <label className="flex items-center gap-2 text-sm">
            <span className="text-xs text-muted-foreground">De</span>
            <input type="date" value={dtIni} onChange={(e) => setDtIni(e.target.value)} className={inputCls + " w-40"} />
          </label>
          <label className="flex items-center gap-2 text-sm">
            <span className="text-xs text-muted-foreground">Até</span>
            <input type="date" value={dtFim} onChange={(e) => setDtFim(e.target.value)} className={inputCls + " w-40"} />
          </label>
          <button
            onClick={() => { setMes(""); setDtIni(""); setDtFim(""); setCategoriasSel([]); if (years.length) setAno(years[years.length - 1]); }}
            className="ml-auto text-xs text-muted-foreground hover:text-foreground hover:underline"
          >
            Limpar filtros
          </button>
        </div>

        <div className="flex flex-wrap items-center gap-2 border-t pt-3">
          <span className="text-xs font-semibold uppercase text-muted-foreground">Categoria</span>
          <button
            onClick={() => setCategoriasSel([])}
            className={`rounded-full px-3 py-1 text-xs font-medium transition ${
              categoriasSel.length === 0 ? "bg-primary text-primary-foreground" : "bg-muted text-muted-foreground hover:text-foreground"
            }`}
          >
            Todas
          </button>
          {CATEGORIAS.map((c) => (
            <button
              key={c}
              onClick={() => toggleCategoria(c)}
              className={`rounded-full px-3 py-1 text-xs font-medium transition ${
                categoriasSel.includes(c) ? "bg-primary text-primary-foreground" : "bg-muted text-muted-foreground hover:text-foreground"
              }`}
            >
              {c}
            </button>
          ))}
        </div>
      </div>

      <div className="grid gap-4 sm:grid-cols-3">
        <Stat label="Total (Ton)" value={fmt(grandTotal)} tint="green" />
        <Stat label="Dias com produção" value={String(sortedDates.length)} tint="yellow" />
        <Stat label="Lançamentos" value={String(filtered.length)} tint="leaf" />
      </div>


      {/* Chart 1 */}
      <section className="rounded-xl border bg-card p-6 shadow-sm">
        <img src={logo} alt="Nutrimilho" className="mx-auto mb-3 h-8 w-auto" />
        <h2 className="mb-1 text-center text-lg font-bold uppercase tracking-wide text-foreground">Produção Diária por Produto</h2>
        <p className="mb-4 text-center text-xs text-muted-foreground">Agrupado por produto — cada cor representa uma data</p>
        <div style={{ width: "100%", height: 420 }}>
          <ResponsiveContainer>
            <BarChart data={chart1} margin={{ top: 24, right: 16, left: 0, bottom: 40 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
              <XAxis
                dataKey="produto"
                tick={(props) => <TwoLineTick {...props} row={chart1[props.index]} />}
                interval={0}
                height={60}
              />
              <YAxis tick={{ fontSize: 11 }} />
              <Tooltip
                formatter={(v: number) => fmt(v) + " Ton"}
                labelFormatter={(_, p) => {
                  const r = p?.[0]?.payload as any;
                  return r ? `${r.categoria} — ${r.produto}` : "";
                }}
              />
              <Legend wrapperStyle={{ fontSize: 12 }} formatter={(v) => new Date(v + "T00:00").toLocaleDateString("pt-BR")} />
              {sortedDates.map((d, i) => (
                <Bar key={d} dataKey={d} fill={DATE_PALETTE[i % DATE_PALETTE.length]}>
                  <LabelList dataKey={d} position="top" style={{ fontSize: 10 }} formatter={(v: number) => (v ? fmt(v) : "")} />
                </Bar>
              ))}
            </BarChart>
          </ResponsiveContainer>
        </div>
      </section>

      {/* Chart 2 */}
      <section className="rounded-xl border bg-card p-6 shadow-sm">
        <img src={logo} alt="Nutrimilho" className="mx-auto mb-3 h-8 w-auto" />
        <h2 className="mb-4 text-center text-lg font-bold uppercase tracking-wide text-foreground">Total Diário</h2>
        <div style={{ width: "100%", height: 420 }}>
          <ResponsiveContainer>
            <BarChart data={chart2} margin={{ top: 32, right: 16, left: 0, bottom: 20 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
              <XAxis
                dataKey="data"
                tick={{ fontSize: 11 }}
                tickFormatter={(v) => new Date(v + "T00:00").toLocaleDateString("pt-BR")}
              />
              <YAxis tick={{ fontSize: 11 }} />
              <Tooltip
                formatter={(v: number) => fmt(v) + " Ton"}
                labelFormatter={(v) => new Date(v + "T00:00").toLocaleDateString("pt-BR")}
              />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              {activeCats.map((c) => (
                <Bar key={c} dataKey={c} stackId="a" fill={CAT_COLORS[c]}>
                  <LabelList dataKey={c} position="center" style={{ fontSize: 10, fill: "#fff", fontWeight: 600 }} formatter={(v: number) => (v ? fmt(v) : "")} />
                </Bar>
              ))}
              <Bar dataKey="__total" fill="transparent" legendType="none">
                <LabelList dataKey="__total" position="top" style={{ fontSize: 11, fontWeight: 700 }} formatter={(v: number) => (v ? fmt(v) : "")} />
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </section>
    </div>
  );
}

function TwoLineTick({ x, y, row }: any) {
  if (!row) return null;
  return (
    <g transform={`translate(${x},${y + 12})`}>
      <text textAnchor="middle" fontSize={11} fontWeight={600} fill="#333">{row.categoria}</text>
      <text textAnchor="middle" fontSize={11} fill="#666" dy={14}>{row.produto}</text>
    </g>
  );
}

function Stat({ label, value, tint }: { label: string; value: string; tint: "green" | "yellow" | "leaf" }) {
  const bg = { green: "bg-brand-green", yellow: "bg-brand-yellow", leaf: "bg-brand-leaf" }[tint];
  return (
    <div className="overflow-hidden rounded-xl border bg-card shadow-sm">
      <div className={`h-1 ${bg}`} />
      <div className="p-5">
        <p className="text-xs font-medium uppercase tracking-wide text-muted-foreground">{label}</p>
        <p className="mt-1 text-4xl font-bold tabular-nums text-foreground">{value}</p>
      </div>
    </div>
  );
}
