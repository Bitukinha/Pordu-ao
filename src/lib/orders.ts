import { createServerFn } from "@tanstack/react-start";
import { asc, desc, eq } from "drizzle-orm";
import { z } from "zod";

import { db } from "./db/client";
import { orderLogs as orderLogsTable, orders as ordersTable } from "./db/schema";

const orderInput = z.object({
  produto: z.string(),
  embalagem: z.enum(["big_bag", "saco"]),
  quantidadeAlvo: z.number(),
});

const orderLogInput = z.object({
  orderId: z.string(),
  data: z.string(),
  turno: z.string(),
  quantidade: z.number(),
});

function toOrder(row: typeof ordersTable.$inferSelect) {
  return {
    id: row.id,
    produto: row.produto,
    embalagem: row.embalagem as "big_bag" | "saco",
    quantidadeAlvo: Number(row.quantidadeAlvo),
    createdAt: row.createdAt.toISOString(),
  };
}

function toOrderLog(row: typeof orderLogsTable.$inferSelect) {
  return {
    id: row.id,
    orderId: row.orderId,
    data: row.data,
    turno: row.turno,
    quantidade: Number(row.quantidade),
  };
}

export const listOrders = createServerFn({ method: "GET" }).handler(async () => {
  const [orderRows, logRows] = await Promise.all([
    db.select().from(ordersTable).orderBy(desc(ordersTable.createdAt)),
    db.select().from(orderLogsTable).orderBy(asc(orderLogsTable.data)),
  ]);
  return {
    orders: orderRows.map(toOrder),
    logs: logRows.map(toOrderLog),
  };
});

export const addOrder = createServerFn({ method: "POST" })
  .validator(orderInput)
  .handler(async ({ data }) => {
    const [row] = await db
      .insert(ordersTable)
      .values({ ...data, quantidadeAlvo: String(data.quantidadeAlvo) })
      .returning();
    return toOrder(row);
  });

export const deleteOrder = createServerFn({ method: "POST" })
  .validator(z.object({ id: z.string() }))
  .handler(async ({ data }) => {
    await db.delete(ordersTable).where(eq(ordersTable.id, data.id));
  });

export const addOrderLog = createServerFn({ method: "POST" })
  .validator(orderLogInput)
  .handler(async ({ data }) => {
    const [row] = await db
      .insert(orderLogsTable)
      .values({ ...data, quantidade: String(data.quantidade) })
      .returning();
    return toOrderLog(row);
  });

export const deleteOrderLog = createServerFn({ method: "POST" })
  .validator(z.object({ id: z.string() }))
  .handler(async ({ data }) => {
    await db.delete(orderLogsTable).where(eq(orderLogsTable.id, data.id));
  });
