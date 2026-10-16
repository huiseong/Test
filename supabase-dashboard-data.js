import { selectRows } from './supabase-client.js';

export async function loadDashboardData() {
  const rows = await selectRows(
    'sales',
    '?select=sale_date,company,country,category,product,channel,qty,price,sales,profit,marketing,defect,ontime&order=sale_date.asc'
  );

  return rows.map((row) => ({
    date: row.sale_date,
    company: row.company,
    country: row.country,
    category: row.category,
    product: row.product,
    channel: row.channel,
    qty: Number(row.qty),
    price: Number(row.price),
    sales: Number(row.sales),
    profit: Number(row.profit),
    marketing: Number(row.marketing),
    defect: Number(row.defect),
    ontime: Number(row.ontime),
  }));
}
