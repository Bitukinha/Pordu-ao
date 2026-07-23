import { pgTable, uuid, date, text, numeric, timestamp } from "drizzle-orm/pg-core";

export const entries = pgTable("entries", {
  id: uuid("id").primaryKey().defaultRandom(),
  data: date("data", { mode: "string" }).notNull(),
  categoria: text("categoria").notNull(),
  produto: text("produto").notNull(),
  qteTon: numeric("qte_ton", { precision: 12, scale: 2 }).notNull(),
});

export const orders = pgTable("orders", {
  id: uuid("id").primaryKey().defaultRandom(),
  produto: text("produto").notNull(),
  embalagem: text("embalagem").notNull(), // "big_bag" | "saco"
  quantidadeAlvo: numeric("quantidade_alvo", { precision: 12, scale: 2 }).notNull(),
  createdAt: timestamp("created_at", { withTimezone: true }).notNull().defaultNow(),
});

export const orderLogs = pgTable("order_logs", {
  id: uuid("id").primaryKey().defaultRandom(),
  orderId: uuid("order_id")
    .notNull()
    .references(() => orders.id, { onDelete: "cascade" }),
  data: date("data", { mode: "string" }).notNull(),
  turno: text("turno").notNull(),
  quantidade: numeric("quantidade", { precision: 12, scale: 2 }).notNull(),
  createdAt: timestamp("created_at", { withTimezone: true }).notNull().defaultNow(),
});
