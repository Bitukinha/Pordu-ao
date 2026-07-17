import { pgTable, uuid, date, text, numeric } from "drizzle-orm/pg-core";

export const entries = pgTable("entries", {
  id: uuid("id").primaryKey().defaultRandom(),
  data: date("data", { mode: "string" }).notNull(),
  categoria: text("categoria").notNull(),
  produto: text("produto").notNull(),
  qteTon: numeric("qte_ton", { precision: 12, scale: 2 }).notNull(),
});
