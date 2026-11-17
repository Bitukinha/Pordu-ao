import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useRef, useState } from "react";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, LabelList, Cell, ReferenceLine,
} from "recharts";
import { ChevronDown } from "lucide-react";
import { addEntries, addEntry, clearEntries, deleteEntry, listEntries, updateEntry } from "@/lib/entries";
import {
  CATEGORIAS, normalizeCategoria, PRODUTOS, normalizeProduto, TURNOS, TURNO_OBRIGATORIO_DESDE, normalizeTurno,
} from "@/lib/constants";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { Command, CommandEmpty, CommandGroup, CommandInput, CommandItem, CommandList } from "@/components/ui/command";
import { Checkbox } from "@/components/ui/checkbox";
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
  turno: string | null; // "A" | "B" | "C" | null (lançamentos antigos)
};

type NewEntry = Omit<Entry, "id">;

// palette derived from Nutrimilho logo (greens/yellows/warm)
const CAT_COLORS: Record<string, string> = {
  "Extrusão": "#E8833A",
  "Flotação": "#8FA6A2",
  "Exportação": "#3E7CB1",
  "Germen": "#F2B807",
  "Mercado interno": "#7CB342",
  "Milho": "#C0392B",
};
// cores por produto (índice na lista de produtos existentes, estável independente do filtro)
const PRODUTO_PALETTE = [
  "#F2B807", "#3E7CB1", "#E8833A", "#7CB342", "#8FA6A2", "#C0392B", "#9B59B6", "#5DADE2",
  "#16A085", "#D35400", "#F48FB1", "#795548", "#AED581", "#607D8B", "#FFD54F", "#4DB6AC",
];
const TURNO_COLORS: Record<string, string> = { A: "#3E7CB1", B: "#F2B807", C: "#7CB342", "Sem turno": "#B0B0B0" };
const SEM_TURNO = "Sem turno";
const turnoLabel = (t: string | null) => t ?? SEM_TURNO;

const DATE_PALETTE = ["#3E7CB1", "#E8833A", "#8FA6A2", "#F2B807", "#5DADE2", "#7CB342", "#C0392B", "#9B59B6"];

function fmt(n: number) {
  return n.toLocaleString("pt-BR", { maximumFractionDigits: 2 });
}

function App() {
  const [tab, setTab] = useState<"dashboard" | "turnos" | "dados">("dashboard");
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
  const updateMutation = useMutation({
    mutationFn: (entry: Entry) => updateEntry({ data: entry }),
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
            {(["dashboard", "turnos", "dados"] as const).map((t) => (
              <button
                key={t}
                onClick={() => setTab(t)}
                className={`rounded-md px-2.5 py-1.5 text-xs font-medium transition sm:px-4 sm:text-sm ${
                  tab === t ? "bg-primary text-primary-foreground shadow-sm" : "text-muted-foreground hover:text-foreground"
                }`}
              >
                {t === "dashboard" ? "Indicadores" : t === "turnos" ? "Por Turno" : "Lançar Dados"}
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
        ) : tab === "turnos" ? (
          <TurnoDashboard entries={entries} />
        ) : (
          <DataEntry
            entries={entries}
            onAdd={(entry) => addMutation.mutateAsync(entry)}
            onImport={(imported) => addManyMutation.mutateAsync(imported)}
            onUpdate={(entry) => updateMutation.mutateAsync(entry)}
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
      Turno: e.turno ?? "",
      "Qte Ton": e.qteTon,
    }));
  const ws = XLSX.utils.json_to_sheet(rows, { cellDates: true });
  ws["!cols"] = [{ wch: 12 }, { wch: 18 }, { wch: 22 }, { wch: 8 }, { wch: 12 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Dados");
  XLSX.writeFile(wb, `nutrimilho-producao-${new Date().toISOString().slice(0, 10)}.xlsx`);
}

/* -------------------- DATA ENTRY -------------------- */
function DataEntry({
  entries,
  onAdd,
  onImport,
  onUpdate,
  onRemove,
  onClear,
}: {
  entries: Entry[];
  onAdd: (entry: NewEntry) => Promise<unknown>;
  onImport: (entries: NewEntry[]) => Promise<unknown>;
  onUpdate: (entry: Entry) => Promise<unknown>;
  onRemove: (id: string) => Promise<unknown>;
  onClear: () => Promise<unknown>;
}) {
  const today = new Date().toISOString().slice(0, 10);
  const [data, setData] = useState(today);
  const [categoria, setCategoria] = useState(CATEGORIAS[0]);
  const [produto, setProduto] = useState(PRODUTOS[0]);
  const [qteTon, setQteTon] = useState("");
  const [turno, setTurno] = useState<string>("");
  const turnoObrigatorio = data >= TURNO_OBRIGATORIO_DESDE;
  const [editingId, setEditingId] = useState<string | null>(null);
  const [semTurnoOnly, setSemTurnoOnly] = useState(false);
  const semTurnoCount = entries.filter((e) => !e.turno).length;
  const visiveis = semTurnoOnly ? entries.filter((e) => !e.turno) : entries;
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
        const t = pick(r, ["turno"]);
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
          categoria: normalizeCategoria(String(cat)),
          produto: normalizeProduto(String(prod)),
          qteTon: Number(q),
          turno: normalizeTurno(t),
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
    if (turnoObrigatorio && !turno) return;
    await onAdd({ data, categoria, produto, qteTon: n, turno: turno || null });
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
          <Field label={turnoObrigatorio ? "Turno" : "Turno (opcional para datas anteriores a nov/2026)"}>
            <div className="flex gap-2">
              {TURNOS.map((t) => (
                <button
                  key={t}
                  type="button"
                  onClick={() => setTurno(turno === t ? "" : t)}
                  className={`flex-1 rounded-md border px-3 py-2 text-sm font-semibold transition ${
                    turno === t ? "border-primary bg-primary text-primary-foreground" : "border-input bg-white hover:border-primary"
                  }`}
                >
                  {t}
                </button>
              ))}
            </div>
            {turnoObrigatorio && !turno && <span className="mt-1 block text-[11px] text-destructive">Selecione o turno.</span>}
          </Field>
          <Field label="Quantidade (Ton)">
            <input
              type="text" inputMode="decimal" value={qteTon}
              onChange={(e) => setQteTon(e.target.value)}
              placeholder="0,00" className={inputCls} required
            />
          </Field>
          <button
            type="submit"
            disabled={turnoObrigatorio && !turno}
            className="w-full rounded-md bg-primary px-4 py-2.5 text-sm font-semibold text-primary-foreground transition hover:opacity-90 disabled:cursor-not-allowed disabled:opacity-50"
          >
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
          <p className="mt-2 text-[11px] text-muted-foreground">Lê a aba <strong>Dados</strong> (colunas: Data, Categoria, Produto, Turno, Qte Ton).</p>
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
        <div className="mb-4 flex flex-wrap items-center justify-between gap-2">
          <h2 className="text-base font-semibold">Lançamentos ({visiveis.length})</h2>
          {semTurnoCount > 0 && (
            <label className="flex cursor-pointer items-center gap-2 text-xs text-muted-foreground">
              <Checkbox checked={semTurnoOnly} onCheckedChange={(v) => setSemTurnoOnly(v === true)} />
              Mostrar só sem turno ({semTurnoCount})
            </label>
          )}
        </div>
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
                  <th className="py-2 pr-4">Turno</th>
                  <th className="py-2 pr-4 text-right">Qte (Ton)</th>
                  <th className="py-2"></th>
                </tr>
              </thead>
              <tbody>
                {visiveis.map((e) =>
                  editingId === e.id ? (
                    <EditRow
                      key={e.id}
                      entry={e}
                      onCancel={() => setEditingId(null)}
                      onSave={async (upd) => {
                        await onUpdate(upd);
                        setEditingId(null);
                      }}
                    />
                  ) : (
                    <tr key={e.id} className="border-b last:border-0 hover:bg-muted/50">
                      <td className="py-2 pr-4">{new Date(e.data + "T00:00").toLocaleDateString("pt-BR")}</td>
                      <td className="py-2 pr-4">
                        <span className="inline-block h-2 w-2 rounded-full" style={{ background: CAT_COLORS[e.categoria] || "#999" }} /> {e.categoria}
                      </td>
                      <td className="py-2 pr-4">{e.produto}</td>
                      <td className="py-2 pr-4">
                        {e.turno ? (
                          <span className="inline-block rounded px-2 py-0.5 text-xs font-semibold text-white" style={{ background: TURNO_COLORS[e.turno] }}>
                            {e.turno}
                          </span>
                        ) : (
                          <span className="text-xs text-muted-foreground">—</span>
                        )}
                      </td>
                      <td className="py-2 pr-4 text-right font-medium tabular-nums">{fmt(e.qteTon)}</td>
                      <td className="whitespace-nowrap py-2 text-right">
                        <button onClick={() => setEditingId(e.id)} className="mr-3 text-xs text-primary hover:underline">Editar</button>
                        <button onClick={() => remove(e.id)} className="text-xs text-destructive hover:underline">Excluir</button>
                      </td>
                    </tr>
                  )
                )}
              </tbody>
            </table>
          </div>
        )}
      </section>
    </div>
  );
}

function EditRow({
  entry,
  onSave,
  onCancel,
}: {
  entry: Entry;
  onSave: (entry: Entry) => Promise<unknown>;
  onCancel: () => void;
}) {
  const [data, setData] = useState(entry.data);
  const [categoria, setCategoria] = useState(entry.categoria);
  const [produto, setProduto] = useState(entry.produto);
  const [turno, setTurno] = useState(entry.turno ?? "");
  const [qte, setQte] = useState(String(entry.qteTon).replace(".", ","));
  const [saving, setSaving] = useState(false);
  const n = parseFloat(qte.replace(",", "."));
  const faltaTurno = data >= TURNO_OBRIGATORIO_DESDE && !turno;
  const cls = inputCls + " px-2 py-1";
  const categorias = CATEGORIAS.includes(categoria) ? CATEGORIAS : [categoria, ...CATEGORIAS];
  const produtos = PRODUTOS.includes(produto) ? PRODUTOS : [produto, ...PRODUTOS];

  async function save() {
    if (isNaN(n) || faltaTurno) return;
    setSaving(true);
    try {
      await onSave({ id: entry.id, data, categoria, produto, qteTon: n, turno: turno || null });
    } finally {
      setSaving(false);
    }
  }

  return (
    <tr className="border-b bg-muted/40 last:border-0">
      <td className="py-2 pr-2">
        <input type="date" value={data} onChange={(e) => setData(e.target.value)} className={cls + " w-36"} />
      </td>
      <td className="py-2 pr-2">
        <select value={categoria} onChange={(e) => setCategoria(e.target.value)} className={cls}>
          {categorias.map((c) => <option key={c}>{c}</option>)}
        </select>
      </td>
      <td className="py-2 pr-2">
        <select value={produto} onChange={(e) => setProduto(e.target.value)} className={cls}>
          {produtos.map((p) => <option key={p}>{p}</option>)}
        </select>
      </td>
      <td className="py-2 pr-2">
        <select
          value={turno}
          onChange={(e) => setTurno(e.target.value)}
          className={cls + " w-20" + (faltaTurno ? " border-destructive" : "")}
          autoFocus
        >
          <option value="">—</option>
          {TURNOS.map((t) => <option key={t} value={t}>{t}</option>)}
        </select>
      </td>
      <td className="py-2 pr-2">
        <input type="text" inputMode="decimal" value={qte} onChange={(e) => setQte(e.target.value)} className={cls + " w-24 text-right"} />
      </td>
      <td className="whitespace-nowrap py-2 text-right">
        <button
          onClick={save}
          disabled={saving || isNaN(n) || faltaTurno}
          className="mr-3 text-xs font-semibold text-primary hover:underline disabled:opacity-50"
        >
          {saving ? "Salvando..." : "Salvar"}
        </button>
        <button onClick={onCancel} className="text-xs text-muted-foreground hover:underline">Cancelar</button>
      </td>
    </tr>
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

function ProdutoFilter({
  produtos,
  selecionados,
  onToggle,
  onLimpar,
}: {
  produtos: string[];
  selecionados: string[];
  onToggle: (p: string) => void;
  onLimpar: () => void;
}) {
  const [open, setOpen] = useState(false);
  const ativo = selecionados.length > 0;

  return (
    <Popover open={open} onOpenChange={setOpen}>
      <PopoverTrigger asChild>
        <button
          className={`flex items-center gap-1 rounded-full px-3 py-1 text-xs font-medium transition ${
            ativo ? "bg-primary text-primary-foreground" : "bg-muted text-muted-foreground hover:text-foreground"
          }`}
        >
          Produto{ativo ? ` (${selecionados.length})` : ""}
          <ChevronDown className="h-3 w-3" />
        </button>
      </PopoverTrigger>
      <PopoverContent className="w-64 p-0" align="start">
        <Command>
          <CommandInput placeholder="Buscar produto..." className="text-sm" />
          <CommandList>
            <CommandEmpty>Nenhum produto encontrado.</CommandEmpty>
            <CommandGroup>
              {produtos.map((p) => {
                const checked = selecionados.includes(p);
                return (
                  <CommandItem key={p} value={p} onSelect={() => onToggle(p)} className="cursor-pointer">
                    <Checkbox checked={checked} className="mr-2" />
                    <span className="text-sm">{p}</span>
                  </CommandItem>
                );
              })}
            </CommandGroup>
          </CommandList>
          {ativo && (
            <div className="border-t p-2">
              <button
                onClick={() => onLimpar()}
                className="w-full rounded-md px-2 py-1 text-xs text-muted-foreground hover:text-foreground hover:underline"
              >
                Limpar seleção de produtos
              </button>
            </div>
          )}
        </Command>
      </PopoverContent>
    </Popover>
  );
}

/* -------------------- DASHBOARD -------------------- */
function Dashboard({ entries }: { entries: Entry[] }) {
  // Available years/months from data
  const years = useMemo(
    () => Array.from(new Set(entries.map((e) => e.data.slice(0, 4)))).sort(),
    [entries]
  );

  // Ao entrar no app o dashboard abre sempre nos últimos 7 dias, com todas as categorias e produtos
  const [ano, setAno] = useState<string>("");
  const [mes, setMes] = useState<string>(""); // "" = todos, "01".."12"
  const [dtIni, setDtIni] = useState<string>(() => new Date(Date.now() - 6 * 86400000).toISOString().slice(0, 10));
  const [dtFim, setDtFim] = useState<string>(() => new Date().toISOString().slice(0, 10));
  const [categoriasSel, setCategoriasSel] = useState<string[]>([]); // [] = todas
  const [produtosSel, setProdutosSel] = useState<string[]>([]); // [] = todos

  const produtosDisponiveis = useMemo(
    () => Array.from(new Set(entries.map((e) => e.produto))).sort((a, b) => a.localeCompare(b, "pt-BR")),
    [entries]
  );

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
  const presetDiaria = usaPeriodo && dtIni === hoje && dtFim === hoje;
  const presetSeteDias = usaPeriodo && dtIni === seteDiasAtras && dtFim === hoje;
  const presetMensal = !usaPeriodo && ano === anoAtual && mes === mesAtual;
  const presetAnual = !usaPeriodo && ano === anoAtual && mes === "";

  function aplicarDiaria() {
    setDtIni(hoje);
    setDtFim(hoje);
  }
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
  function toggleProduto(p: string) {
    setProdutosSel((prev) => (prev.includes(p) ? prev.filter((x) => x !== p) : [...prev, p]));
  }

  const filtered = useMemo(
    () =>
      entries.filter((e) => {
        if (categoriasSel.length && !categoriasSel.includes(e.categoria)) return false;
        if (produtosSel.length && !produtosSel.includes(e.produto)) return false;
        if (usaPeriodo) {
          if (dtIni && e.data < dtIni) return false;
          if (dtFim && e.data > dtFim) return false;
          return true;
        }
        if (ano && !e.data.startsWith(ano)) return false;
        if (mes && e.data.slice(5, 7) !== mes) return false;
        return true;
      }),
    [entries, ano, mes, dtIni, dtFim, usaPeriodo, categoriasSel, produtosSel]
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


  // Chart 5: Total diário por produto — barra empilhada por data, série = produto
  const chart5 = useMemo(() => {
    return sortedDates.map((d) => {
      const row: Record<string, string | number> = { data: d };
      let total = 0;
      for (const e of filtered) {
        if (e.data !== d) continue;
        row[e.produto] = ((row[e.produto] as number) ?? 0) + e.qteTon;
        total += e.qteTon;
      }
      row.__total = total;
      return row;
    });
  }, [filtered, sortedDates]);
  const activeProdutos = produtosDisponiveis.filter((p) => chart5.some((r) => r[p]));

  const grandTotal = chart2.reduce((s, r) => s + (r.__total as number), 0);
  const activeCats = CATEGORIAS.filter((c) => chart2.some((r) => r[c]));

  const META_GERMEN_PCT = 30;
  const SEGMENTOS_PRODUCAO_GERMEN = ["Mercado interno", "Flotação", "Extrusão", "Exportação", "Germen"];
  // Só Germen selecionado (na categoria e/ou no produto), sem nenhum outro filtro junto
  const soGermen = (sel: string[]) => sel.length === 1 && sel[0] === "Germen";
  const filtrandoGermen =
    (soGermen(categoriasSel) || soGermen(produtosSel)) &&
    (categoriasSel.length === 0 || soGermen(categoriasSel)) &&
    (produtosSel.length === 0 || soGermen(produtosSel));

  // Entradas do período selecionado (ignora filtro de categoria/produto, para poder comparar Germen com as demais categorias)
  const entriesPeriodo = useMemo(
    () =>
      entries.filter((e) => {
        if (usaPeriodo) {
          if (dtIni && e.data < dtIni) return false;
          if (dtFim && e.data > dtFim) return false;
          return true;
        }
        if (ano && !e.data.startsWith(ano)) return false;
        if (mes && e.data.slice(5, 7) !== mes) return false;
        return true;
      }),
    [entries, ano, mes, dtIni, dtFim, usaPeriodo]
  );

  const germenAnalise = useMemo(() => {
    const dias = Array.from(new Set(entriesPeriodo.map((e) => e.data))).sort();
    let germenTotal = 0;
    let totalProcessado = 0;
    const percPorDia: number[] = [];
    const porDia: { data: string; percentual: number }[] = [];
    for (const d of dias) {
      const doDia = entriesPeriodo.filter((e) => e.data === d);
      // Germen = lançamentos na categoria Germen ou com produto Germen (em qualquer categoria)
      const isGermen = (e: (typeof doDia)[number]) => e.categoria === "Germen" || e.produto === "Germen";
      const gDia = doDia.filter(isGermen).reduce((s, e) => s + e.qteTon, 0);
      // Produção total = Mercado interno + Flotação + Extrusão + Exportação + Germen
      const pDia = doDia
        .filter((e) => SEGMENTOS_PRODUCAO_GERMEN.includes(e.categoria) || isGermen(e))
        .reduce((s, e) => s + e.qteTon, 0);
      germenTotal += gDia;
      totalProcessado += pDia;
      if (pDia > 0) {
        percPorDia.push((gDia / pDia) * 100);
        porDia.push({ data: d, percentual: (gDia / pDia) * 100 });
      }
    }
    const percentual = totalProcessado > 0 ? (germenTotal / totalProcessado) * 100 : 0;
    const media = percPorDia.length ? percPorDia.reduce((s, v) => s + v, 0) / percPorDia.length : 0;
    return {
      germenTotal,
      totalProcessado,
      percentual,
      media,
      // Dia mais recente no topo do gráfico horizontal
      porDia: porDia.reverse(),
      status: percentual > META_GERMEN_PCT ? "alerta" : "boa",
    } as const;
  }, [entriesPeriodo]);

  // Entradas filtradas apenas por categoria/produto (ignora período/ano/mês) para os comparativos gerais
  const porCategoria = useMemo(
    () =>
      entries.filter(
        (e) =>
          (!categoriasSel.length || categoriasSel.includes(e.categoria)) &&
          (!produtosSel.length || produtosSel.includes(e.produto))
      ),
    [entries, categoriasSel, produtosSel]
  );

  // Total histórico do(s) produto(s) selecionado(s), independente do período filtrado
  const totalGeralProduto = useMemo(
    () => porCategoria.reduce((s, e) => s + e.qteTon, 0),
    [porCategoria]
  );

  // Chart 3: Comparativo mensal — total por mês do ano selecionado, série = categoria
  const chart3 = useMemo(() => {
    const doAno = porCategoria.filter((e) => e.data.startsWith(ano));
    return MESES.map((label, idx) => {
      const mKey = String(idx + 1).padStart(2, "0");
      const row: Record<string, string | number> = { mes: label };
      let total = 0;
      for (const c of CATEGORIAS) {
        const v = doAno
          .filter((e) => e.data.slice(5, 7) === mKey && e.categoria === c)
          .reduce((s, e) => s + e.qteTon, 0);
        if (v) row[c] = v;
        total += v;
      }
      row.__total = total;
      return row;
    });
  }, [porCategoria, ano]);
  const activeCatsChart3 = CATEGORIAS.filter((c) => chart3.some((r) => r[c]));

  // Chart 4: Recordes de produção diária — melhor dia de cada mês do ano selecionado
  const chart4 = useMemo(() => {
    const doAno = porCategoria.filter((e) => e.data.startsWith(ano));
    return MESES.map((label, idx) => {
      const mKey = String(idx + 1).padStart(2, "0");
      const doMes = doAno.filter((e) => e.data.slice(5, 7) === mKey);
      const porDia = new Map<string, number>();
      for (const e of doMes) porDia.set(e.data, (porDia.get(e.data) ?? 0) + e.qteTon);
      let recorde = 0;
      let dataRecorde = "";
      for (const [d, v] of porDia) {
        if (v > recorde) {
          recorde = v;
          dataRecorde = d;
        }
      }
      return { mes: label, recorde, dataRecorde };
    });
  }, [porCategoria, ano]);

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
            onClick={aplicarDiaria}
            className={`rounded-md px-3 py-1 text-xs font-medium transition ${
              presetDiaria ? "bg-primary text-primary-foreground" : "bg-muted text-muted-foreground hover:text-foreground"
            }`}
          >
            Diária
          </button>
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
            onClick={() => { setMes(""); setDtIni(""); setDtFim(""); setCategoriasSel([]); setProdutosSel([]); if (years.length) setAno(years[years.length - 1]); }}
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
          <div className="mx-1 hidden h-5 w-px bg-border sm:block" />
          <ProdutoFilter
            produtos={produtosDisponiveis}
            selecionados={produtosSel}
            onToggle={toggleProduto}
            onLimpar={() => setProdutosSel([])}
          />
        </div>
      </div>

      <div className={`grid gap-4 ${produtosSel.length ? "sm:grid-cols-4" : "sm:grid-cols-3"}`}>
        <Stat label="Total (Ton)" value={fmt(grandTotal)} tint="green" />
        <Stat label="Dias com produção" value={String(sortedDates.length)} tint="yellow" />
        <Stat label="Lançamentos" value={String(filtered.length)} tint="leaf" />
        {produtosSel.length > 0 && (
          <Stat label="Total Geral do Produto (Ton)" value={fmt(totalGeralProduto)} tint="blue" />
        )}
      </div>

      {filtrandoGermen && (
        <section
          className={`rounded-xl border p-5 shadow-sm ${
            germenAnalise.status === "alerta" ? "border-red-300 bg-red-50" : "border-green-300 bg-green-50"
          }`}
        >
          <div className="flex flex-wrap items-center justify-between gap-4">
            <div>
              <p className="text-xs font-semibold uppercase text-muted-foreground">
                Percentual de Germen sobre a produção total
              </p>
              <p
                className={`mt-1 text-4xl font-bold tabular-nums ${
                  germenAnalise.status === "alerta" ? "text-red-600" : "text-green-700"
                }`}
              >
                {fmt(germenAnalise.percentual)}%
              </p>
              <p className="mt-1 text-xs text-muted-foreground">
                Meta: até {META_GERMEN_PCT}% · Média do período: {fmt(germenAnalise.media)}%
              </p>
            </div>
            <span
              className={`rounded-full px-4 py-1.5 text-sm font-semibold ${
                germenAnalise.status === "alerta" ? "bg-red-600 text-white" : "bg-green-600 text-white"
              }`}
            >
              {germenAnalise.status === "alerta" ? "⚠ Alerta" : "✓ Boa"}
            </span>
          </div>
          <p className="mt-3 text-[11px] text-muted-foreground">
            Germen: {fmt(germenAnalise.germenTotal)} Ton · Produção total (Mercado interno + Flotação + Extrusão + Exportação + Germen): {fmt(germenAnalise.totalProcessado)} Ton
          </p>
        </section>
      )}

      {filtrandoGermen && germenAnalise.porDia.length > 0 && (
        <section className="rounded-xl border bg-card p-6 shadow-sm">
          <img src={logo} alt="Nutrimilho" className="mx-auto mb-3 h-8 w-auto" />
          <h2 className="mb-1 text-center text-lg font-bold uppercase tracking-wide text-foreground">% Germen por Produção Diária</h2>
          <p className="mb-4 text-center text-xs text-muted-foreground">
            Germen ÷ (Mercado interno + Flotação + Extrusão + Exportação + Germen) · Meta: até {META_GERMEN_PCT}%
          </p>
          <div style={{ height: Math.max(240, germenAnalise.porDia.length * 28 + 60) }}>
            <ResponsiveContainer>
              <BarChart data={germenAnalise.porDia} layout="vertical" margin={{ top: 8, right: 64, left: 8, bottom: 8 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#eee" horizontal={false} />
                <XAxis type="number" domain={[0, "auto"]} tick={{ fontSize: 12, fill: "#000" }} tickFormatter={(v) => `${v}%`} />
                <YAxis
                  type="category"
                  dataKey="data"
                  width={90}
                  interval={0}
                  tick={{ fontSize: 12, fill: "#000" }}
                  tickFormatter={(v) => new Date(v + "T00:00").toLocaleDateString("pt-BR")}
                />
                <Tooltip
                  formatter={(v: number) => [fmt(v) + "%", "% Germen"]}
                  labelFormatter={(v) => new Date(v + "T00:00").toLocaleDateString("pt-BR")}
                />
                <ReferenceLine x={META_GERMEN_PCT} stroke="#dc2626" strokeDasharray="4 4" />
                <Bar dataKey="percentual" barSize={16}>
                  {germenAnalise.porDia.map((r) => (
                    <Cell key={r.data} fill={r.percentual > META_GERMEN_PCT ? "#dc2626" : "#1F3A6E"} />
                  ))}
                  <LabelList
                    dataKey="percentual"
                    position="right"
                    style={{ fontSize: 12, fontWeight: 600, fill: "#000" }}
                    formatter={(v: number) => fmt(v) + "%"}
                  />
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </section>
      )}

      {/* Chart 1 */}
      <section className="rounded-xl border bg-card p-6 shadow-sm">
        <img src={logo} alt="Nutrimilho" className="mx-auto mb-3 h-8 w-auto" />
        <h2 className="mb-1 text-center text-lg font-bold uppercase tracking-wide text-foreground">Produção Diária por Produto</h2>
        <p className="mb-4 text-center text-xs text-muted-foreground">Agrupado por produto — cada cor representa uma data</p>
        <div className="w-full overflow-x-auto">
          <div style={{ minWidth: Math.max(760, chart1.length * Math.max(sortedDates.length, 1) * 30), height: 440 }}>
            <ResponsiveContainer>
              <BarChart data={chart1} margin={{ top: 28, right: 16, left: 0, bottom: 40 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
                <XAxis
                  dataKey="produto"
                  tick={(props) => <TwoLineTick {...props} row={chart1[props.index]} />}
                  interval={0}
                  height={60}
                />
                <YAxis tick={{ fontSize: 12, fill: "#000" }} />
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
                    <LabelList
                      dataKey={d}
                      position="top"
                      style={{ fontSize: 11, fontWeight: 600, fill: "#000" }}
                      formatter={(v: number) => (v ? fmt(v) : "")}
                    />
                  </Bar>
                ))}
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </section>

      {/* Chart 2 */}
      <section className="rounded-xl border bg-card p-6 shadow-sm">
        <img src={logo} alt="Nutrimilho" className="mx-auto mb-3 h-8 w-auto" />
        <h2 className="mb-4 text-center text-lg font-bold uppercase tracking-wide text-foreground">Total Diário</h2>
        <div className="w-full overflow-x-auto">
          <div style={{ minWidth: Math.max(760, sortedDates.length * 70), height: 440 }}>
            <ResponsiveContainer>
              <BarChart data={chart2} margin={{ top: 34, right: 16, left: 0, bottom: 20 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
                <XAxis
                  dataKey="data"
                  tick={{ fontSize: 12, fill: "#000" }}
                  tickFormatter={(v) => new Date(v + "T00:00").toLocaleDateString("pt-BR")}
                />
                <YAxis tick={{ fontSize: 12, fill: "#000" }} />
                <Tooltip
                  formatter={(v: number) => fmt(v) + " Ton"}
                  labelFormatter={(v) => new Date(v + "T00:00").toLocaleDateString("pt-BR")}
                />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                {activeCats.map((c) => (
                  <Bar key={c} dataKey={c} stackId="a" fill={CAT_COLORS[c]}>
                    <LabelList dataKey={c} position="center" style={{ fontSize: 12, fill: "#000" }} formatter={(v: number) => (v ? fmt(v) : "")} />
                  </Bar>
                ))}
                <Bar dataKey="__total" fill="transparent" legendType="none">
                  <LabelList dataKey="__total" position="top" style={{ fontSize: 13, fontWeight: 700, fill: "#000" }} formatter={(v: number) => (v ? fmt(v) : "")} />
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </section>

      {/* Chart 5: Total Diário por Produto */}
      <section className="rounded-xl border bg-card p-6 shadow-sm">
        <img src={logo} alt="Nutrimilho" className="mx-auto mb-3 h-8 w-auto" />
        <h2 className="mb-1 text-center text-lg font-bold uppercase tracking-wide text-foreground">Total Diário por Produto</h2>
        <p className="mb-4 text-center text-xs text-muted-foreground">Cada cor representa um produto — acompanha os filtros aplicados</p>
        <div className="w-full overflow-x-auto">
          <div style={{ minWidth: Math.max(760, sortedDates.length * 70), height: 440 }}>
            <ResponsiveContainer>
              <BarChart data={chart5} margin={{ top: 34, right: 16, left: 0, bottom: 20 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
                <XAxis
                  dataKey="data"
                  tick={{ fontSize: 12, fill: "#000" }}
                  tickFormatter={(v) => new Date(v + "T00:00").toLocaleDateString("pt-BR")}
                />
                <YAxis tick={{ fontSize: 12, fill: "#000" }} />
                <Tooltip
                  formatter={(v: number) => fmt(v) + " Ton"}
                  labelFormatter={(v) => new Date(v + "T00:00").toLocaleDateString("pt-BR")}
                />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                {activeProdutos.map((p) => (
                  <Bar key={p} dataKey={p} stackId="a" fill={PRODUTO_PALETTE[produtosDisponiveis.indexOf(p) % PRODUTO_PALETTE.length]}>
                    <LabelList dataKey={p} position="center" style={{ fontSize: 12, fill: "#000" }} formatter={(v: number) => (v ? fmt(v) : "")} />
                  </Bar>
                ))}
                <Bar dataKey="__total" fill="transparent" legendType="none">
                  <LabelList dataKey="__total" position="top" style={{ fontSize: 13, fontWeight: 700, fill: "#000" }} formatter={(v: number) => (v ? fmt(v) : "")} />
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </section>

      {/* Chart 3: Comparativo Mensal */}
      <section className="rounded-xl border bg-card p-6 shadow-sm">
        <img src={logo} alt="Nutrimilho" className="mx-auto mb-3 h-8 w-auto" />
        <h2 className="mb-1 text-center text-lg font-bold uppercase tracking-wide text-foreground">Comparativo Mensal — {ano}</h2>
        <p className="mb-4 text-center text-xs text-muted-foreground">Total de produção por mês do ano selecionado (independe do filtro de período)</p>
        <div style={{ width: "100%", height: 400 }}>
          <ResponsiveContainer>
            <BarChart data={chart3} margin={{ top: 34, right: 16, left: 0, bottom: 20 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
              <XAxis dataKey="mes" tick={{ fontSize: 12, fill: "#000" }} />
              <YAxis tick={{ fontSize: 12, fill: "#000" }} />
              <Tooltip formatter={(v: number) => fmt(v) + " Ton"} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              {activeCatsChart3.map((c) => (
                <Bar key={c} dataKey={c} stackId="a" fill={CAT_COLORS[c]}>
                  <LabelList dataKey={c} position="center" style={{ fontSize: 12, fill: "#000" }} formatter={(v: number) => (v ? fmt(v) : "")} />
                </Bar>
              ))}
              <Bar dataKey="__total" fill="transparent" legendType="none">
                <LabelList dataKey="__total" position="top" style={{ fontSize: 13, fontWeight: 700, fill: "#000" }} formatter={(v: number) => (v ? fmt(v) : "")} />
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </section>

      {/* Chart 4: Recorde de Produção Diária */}
      <section className="rounded-xl border bg-card p-6 shadow-sm">
        <img src={logo} alt="Nutrimilho" className="mx-auto mb-3 h-8 w-auto" />
        <h2 className="mb-1 text-center text-lg font-bold uppercase tracking-wide text-foreground">Recorde de Produção Diária — {ano}</h2>
        <p className="mb-4 text-center text-xs text-muted-foreground">Maior produção em um único dia por mês do ano selecionado — respeita o filtro de categoria</p>
        <div style={{ width: "100%", height: 380 }}>
          <ResponsiveContainer>
            <BarChart data={chart4} margin={{ top: 32, right: 16, left: 0, bottom: 20 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
              <XAxis dataKey="mes" tick={{ fontSize: 12, fill: "#000" }} />
              <YAxis tick={{ fontSize: 12, fill: "#000" }} />
              <Tooltip
                formatter={(v: number) => fmt(v) + " Ton"}
                labelFormatter={(v, p) => {
                  const r = p?.[0]?.payload as any;
                  const dt = r?.dataRecorde ? new Date(r.dataRecorde + "T00:00").toLocaleDateString("pt-BR") : "";
                  return dt ? `${v} — recorde em ${dt}` : String(v);
                }}
              />
              <Bar dataKey="recorde" fill="#F2B807">
                <LabelList dataKey="recorde" position="top" style={{ fontSize: 13, fontWeight: 700, fill: "#000" }} formatter={(v: number) => (v ? fmt(v) : "")} />
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </section>
    </div>
  );
}

/* -------------------- DASHBOARD POR TURNO -------------------- */
function TurnoDashboard({ entries }: { entries: Entry[] }) {
  const hoje = new Date().toISOString().slice(0, 10);
  const seteDiasAtras = new Date(Date.now() - 6 * 86400000).toISOString().slice(0, 10);
  const inicioMes = hoje.slice(0, 8) + "01";
  const inicioAno = hoje.slice(0, 4) + "-01-01";

  const [dtIni, setDtIni] = useState(seteDiasAtras);
  const [dtFim, setDtFim] = useState(hoje);
  const [categoriasSel, setCategoriasSel] = useState<string[]>([]);

  const presets = [
    { label: "Diária", ini: hoje },
    { label: "Últimos 7 dias", ini: seteDiasAtras },
    { label: "Mensal", ini: inicioMes },
    { label: "Anual", ini: inicioAno },
  ];

  const filtered = useMemo(
    () =>
      entries.filter((e) => {
        if (categoriasSel.length && !categoriasSel.includes(e.categoria)) return false;
        if (dtIni && e.data < dtIni) return false;
        if (dtFim && e.data > dtFim) return false;
        return true;
      }),
    [entries, categoriasSel, dtIni, dtFim]
  );

  // Turnos que aparecem no período (Sem turno só entra se houver lançamento antigo sem turno)
  const turnos = useMemo(() => {
    const list: string[] = [...TURNOS];
    if (filtered.some((e) => !e.turno)) list.push(SEM_TURNO);
    return list;
  }, [filtered]);

  const dias = useMemo(() => Array.from(new Set(filtered.map((e) => e.data))).sort(), [filtered]);

  const resumo = useMemo(() => {
    const total = filtered.reduce((s, e) => s + e.qteTon, 0);
    return turnos.map((t) => {
      const doTurno = filtered.filter((e) => turnoLabel(e.turno) === t);
      const soma = doTurno.reduce((s, e) => s + e.qteTon, 0);
      const porDia = new Map<string, number>();
      for (const e of doTurno) porDia.set(e.data, (porDia.get(e.data) ?? 0) + e.qteTon);
      let recorde = 0;
      let dataRecorde = "";
      for (const [d, v] of porDia) if (v > recorde) { recorde = v; dataRecorde = d; }
      return {
        turno: t,
        total: soma,
        pct: total > 0 ? (soma / total) * 100 : 0,
        media: porDia.size ? soma / porDia.size : 0,
        dias: porDia.size,
        recorde,
        dataRecorde,
      };
    });
  }, [filtered, turnos]);

  // Produção diária: barras agrupadas por data, uma por turno
  const porDia = useMemo(
    () =>
      dias.map((d) => {
        const row: Record<string, string | number> = { data: d };
        for (const e of filtered) {
          if (e.data !== d) continue;
          const t = turnoLabel(e.turno);
          row[t] = ((row[t] as number) ?? 0) + e.qteTon;
        }
        return row;
      }),
    [filtered, dias]
  );

  // Produção por turno empilhada por categoria
  const porCategoria = useMemo(
    () =>
      turnos.map((t) => {
        const row: Record<string, string | number> = { turno: t };
        let total = 0;
        for (const e of filtered) {
          if (turnoLabel(e.turno) !== t) continue;
          row[e.categoria] = ((row[e.categoria] as number) ?? 0) + e.qteTon;
          total += e.qteTon;
        }
        row.__total = total;
        return row;
      }),
    [filtered, turnos]
  );
  const catsAtivas = CATEGORIAS.filter((c) => porCategoria.some((r) => r[c]));

  // Tabela produto × turno
  const porProduto = useMemo(() => {
    const map = new Map<string, Record<string, number>>();
    for (const e of filtered) {
      const key = `${e.categoria}||${e.produto}`;
      const row = map.get(key) ?? {};
      const t = turnoLabel(e.turno);
      row[t] = (row[t] ?? 0) + e.qteTon;
      row.__total = (row.__total ?? 0) + e.qteTon;
      map.set(key, row);
    }
    return Array.from(map.entries())
      .map(([key, vals]) => {
        const [categoria, produto] = key.split("||");
        return { categoria, produto, vals };
      })
      .sort((a, b) => b.vals.__total - a.vals.__total);
  }, [filtered]);

  const totalGeral = resumo.reduce((s, r) => s + r.total, 0);
  const semTurnoQtd = filtered.filter((e) => !e.turno).length;
  const dataBR = (v: string) => new Date(v + "T00:00").toLocaleDateString("pt-BR");
  const btn = (on: boolean) =>
    `rounded-md px-3 py-1 text-xs font-medium transition ${on ? "bg-primary text-primary-foreground" : "bg-muted text-muted-foreground hover:text-foreground"}`;
  const chip = (on: boolean) =>
    `rounded-full px-3 py-1 text-xs font-medium transition ${on ? "bg-primary text-primary-foreground" : "bg-muted text-muted-foreground hover:text-foreground"}`;

  return (
    <div className="space-y-6">
      <div className="space-y-3 rounded-xl border bg-card p-4 shadow-sm">
        <div className="flex flex-wrap items-center gap-2">
          <span className="text-xs font-semibold uppercase text-muted-foreground">Período rápido</span>
          {presets.map((p) => (
            <button key={p.label} onClick={() => { setDtIni(p.ini); setDtFim(hoje); }} className={btn(dtIni === p.ini && dtFim === hoje)}>
              {p.label}
            </button>
          ))}
        </div>
        <div className="flex flex-wrap items-end gap-3">
          <label className="flex items-center gap-2 text-sm">
            <span className="text-xs text-muted-foreground">De</span>
            <input type="date" value={dtIni} onChange={(e) => setDtIni(e.target.value)} className={inputCls + " w-40"} />
          </label>
          <label className="flex items-center gap-2 text-sm">
            <span className="text-xs text-muted-foreground">Até</span>
            <input type="date" value={dtFim} onChange={(e) => setDtFim(e.target.value)} className={inputCls + " w-40"} />
          </label>
        </div>
        <div className="flex flex-wrap items-center gap-2 border-t pt-3">
          <span className="text-xs font-semibold uppercase text-muted-foreground">Categoria</span>
          <button onClick={() => setCategoriasSel([])} className={chip(categoriasSel.length === 0)}>Todas</button>
          {CATEGORIAS.map((c) => (
            <button
              key={c}
              onClick={() => setCategoriasSel((prev) => (prev.includes(c) ? prev.filter((x) => x !== c) : [...prev, c]))}
              className={chip(categoriasSel.includes(c))}
            >
              {c}
            </button>
          ))}
        </div>
      </div>

      {semTurnoQtd > 0 && (
        <p className="rounded-lg border border-amber-300 bg-amber-50 px-4 py-2 text-xs text-amber-900">
          {semTurnoQtd} lançamento(s) deste período estão sem turno e aparecem como "{SEM_TURNO}". Edite-os em <strong>Lançar Dados</strong> para incluir o turno.
        </p>
      )}

      {filtered.length === 0 ? (
        <div className="rounded-xl border-2 border-dashed bg-card p-16 text-center">
          <p className="text-sm text-muted-foreground">Nenhum lançamento no período selecionado.</p>
        </div>
      ) : (
        <>
          <div className={`grid gap-4 sm:grid-cols-2 ${turnos.length > 3 ? "lg:grid-cols-4" : "lg:grid-cols-3"}`}>
            {resumo.map((r) => (
              <div key={r.turno} className="overflow-hidden rounded-xl border bg-card shadow-sm">
                <div className="h-1" style={{ background: TURNO_COLORS[r.turno] }} />
                <div className="p-5">
                  <p className="text-xs font-medium uppercase tracking-wide text-muted-foreground">
                    {r.turno === SEM_TURNO ? SEM_TURNO : `Turno ${r.turno}`}
                  </p>
                  <p className="mt-1 text-4xl font-bold tabular-nums text-foreground">{fmt(r.total)}</p>
                  <p className="text-xs text-muted-foreground">Ton · {fmt(r.pct)}% do total</p>
                  <div className="mt-3 grid grid-cols-2 gap-2 border-t pt-3 text-xs">
                    <div>
                      <p className="text-muted-foreground">Média/dia</p>
                      <p className="font-semibold tabular-nums">{fmt(r.media)} Ton</p>
                    </div>
                    <div>
                      <p className="text-muted-foreground">Melhor dia</p>
                      <p className="font-semibold tabular-nums">{r.recorde ? `${fmt(r.recorde)} (${dataBR(r.dataRecorde).slice(0, 5)})` : "—"}</p>
                    </div>
                  </div>
                </div>
              </div>
            ))}
          </div>

          <section className="rounded-xl border bg-card p-6 shadow-sm">
            <img src={logo} alt="Nutrimilho" className="mx-auto mb-3 h-8 w-auto" />
            <h2 className="mb-1 text-center text-lg font-bold uppercase tracking-wide text-foreground">Produção Diária por Turno</h2>
            <p className="mb-4 text-center text-xs text-muted-foreground">Cada cor representa um turno · Total do período: {fmt(totalGeral)} Ton</p>
            <div className="w-full overflow-x-auto">
              <div style={{ minWidth: Math.max(760, dias.length * turnos.length * 34), height: 420 }}>
                <ResponsiveContainer>
                  <BarChart data={porDia} margin={{ top: 28, right: 16, left: 0, bottom: 20 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
                    <XAxis dataKey="data" tick={{ fontSize: 12, fill: "#000" }} tickFormatter={dataBR} />
                    <YAxis tick={{ fontSize: 12, fill: "#000" }} />
                    <Tooltip formatter={(v: number) => fmt(v) + " Ton"} labelFormatter={(v) => dataBR(String(v))} />
                    <Legend wrapperStyle={{ fontSize: 12 }} formatter={(v) => (v === SEM_TURNO ? v : `Turno ${v}`)} />
                    {turnos.map((t) => (
                      <Bar key={t} dataKey={t} fill={TURNO_COLORS[t]}>
                        <LabelList dataKey={t} position="top" style={{ fontSize: 11, fontWeight: 600, fill: "#000" }} formatter={(v: number) => (v ? fmt(v) : "")} />
                      </Bar>
                    ))}
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </section>

          <section className="rounded-xl border bg-card p-6 shadow-sm">
            <img src={logo} alt="Nutrimilho" className="mx-auto mb-3 h-8 w-auto" />
            <h2 className="mb-1 text-center text-lg font-bold uppercase tracking-wide text-foreground">Total por Turno e Categoria</h2>
            <p className="mb-4 text-center text-xs text-muted-foreground">Soma do período selecionado</p>
            <div style={{ width: "100%", height: 400 }}>
              <ResponsiveContainer>
                <BarChart data={porCategoria} margin={{ top: 34, right: 16, left: 0, bottom: 20 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
                  <XAxis dataKey="turno" tick={{ fontSize: 13, fill: "#000", fontWeight: 600 }} tickFormatter={(v) => (v === SEM_TURNO ? v : `Turno ${v}`)} />
                  <YAxis tick={{ fontSize: 12, fill: "#000" }} />
                  <Tooltip formatter={(v: number) => fmt(v) + " Ton"} labelFormatter={(v) => (v === SEM_TURNO ? v : `Turno ${v}`)} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  {catsAtivas.map((c) => (
                    <Bar key={c} dataKey={c} stackId="a" fill={CAT_COLORS[c]} maxBarSize={120}>
                      <LabelList dataKey={c} position="center" style={{ fontSize: 12, fill: "#000" }} formatter={(v: number) => (v ? fmt(v) : "")} />
                    </Bar>
                  ))}
                  <Bar dataKey="__total" fill="transparent" legendType="none" maxBarSize={120}>
                    <LabelList dataKey="__total" position="top" style={{ fontSize: 13, fontWeight: 700, fill: "#000" }} formatter={(v: number) => (v ? fmt(v) : "")} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </section>

          <section className="rounded-xl border bg-card p-6 shadow-sm">
            <h2 className="mb-4 text-center text-lg font-bold uppercase tracking-wide text-foreground">Produto × Turno (Ton)</h2>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b text-left text-xs uppercase text-muted-foreground">
                    <th className="py-2 pr-4">Categoria</th>
                    <th className="py-2 pr-4">Produto</th>
                    {turnos.map((t) => (
                      <th key={t} className="py-2 pr-4 text-right">{t === SEM_TURNO ? t : `Turno ${t}`}</th>
                    ))}
                    <th className="py-2 text-right">Total</th>
                  </tr>
                </thead>
                <tbody>
                  {porProduto.map((r) => (
                    <tr key={r.categoria + r.produto} className="border-b last:border-0 hover:bg-muted/50">
                      <td className="py-2 pr-4">
                        <span className="inline-block h-2 w-2 rounded-full" style={{ background: CAT_COLORS[r.categoria] || "#999" }} /> {r.categoria}
                      </td>
                      <td className="py-2 pr-4">{r.produto}</td>
                      {turnos.map((t) => (
                        <td key={t} className="py-2 pr-4 text-right tabular-nums">{r.vals[t] ? fmt(r.vals[t]) : "—"}</td>
                      ))}
                      <td className="py-2 text-right font-semibold tabular-nums">{fmt(r.vals.__total)}</td>
                    </tr>
                  ))}
                  <tr className="border-t-2 font-semibold">
                    <td className="py-2 pr-4" colSpan={2}>Total</td>
                    {resumo.map((r) => (
                      <td key={r.turno} className="py-2 pr-4 text-right tabular-nums">{fmt(r.total)}</td>
                    ))}
                    <td className="py-2 text-right tabular-nums">{fmt(totalGeral)}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </section>
        </>
      )}
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

function Stat({ label, value, tint }: { label: string; value: string; tint: "green" | "yellow" | "leaf" | "blue" }) {
  const bg = { green: "bg-brand-green", yellow: "bg-brand-yellow", leaf: "bg-brand-leaf", blue: "bg-sky-500" }[tint];
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
