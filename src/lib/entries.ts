import { createServerFn } from "@tanstack/react-start";
import { desc, eq } from "drizzle-orm";
import { z } from "zod";

import { db } from "./db/client";
import { entries as entriesTable } from "./db/schema";

const entryInput = z.object({
  data: z.string(),
  categoria: z.string(),
  produto: z.string(),
  qteTon: z.number(),
});

function toEntry(row: typeof entriesTable.$inferSelect) {
  return {
    id: row.id,
    data: row.data,
    categoria: row.categoria,
    produto: row.produto,
    qteTon: Number(row.qteTon),
  };
}

export const listEntries = createServerFn({ method: "GET" }).handler(async () => {
  const rows = await db.select().from(entriesTable).orderBy(desc(entriesTable.data));
  return rows.map(toEntry);
});

function toRow(entry: z.infer<typeof entryInput>) {
  return { ...entry, qteTon: String(entry.qteTon) };
}

export const addEntry = createServerFn({ method: "POST" })
  .validator(entryInput)
  .handler(async ({ data }) => {
    const [row] = await db.insert(entriesTable).values(toRow(data)).returning();
    return toEntry(row);
  });

export const addEntries = createServerFn({ method: "POST" })
  .validator(z.array(entryInput))
  .handler(async ({ data }) => {
    if (data.length === 0) return [];
    const rows = await db.insert(entriesTable).values(data.map(toRow)).returning();
    return rows.map(toEntry);
  });

export const deleteEntry = createServerFn({ method: "POST" })
  .validator(z.object({ id: z.string() }))
  .handler(async ({ data }) => {
    await db.delete(entriesTable).where(eq(entriesTable.id, data.id));
  });

export const clearEntries = createServerFn({ method: "POST" }).handler(async () => {
  await db.delete(entriesTable);
});
