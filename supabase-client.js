import { SUPABASE_ANON_KEY, SUPABASE_URL } from './supabase-config.js';

export async function selectRows(table, query = '') {
  const response = await fetch(`${SUPABASE_URL}/rest/v1/${encodeURIComponent(table)}${query}`, {
    headers: {
      apikey: SUPABASE_ANON_KEY,
      Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
    },
  });

  if (!response.ok) {
    throw new Error(`Supabase request failed (${response.status}): ${await response.text()}`);
  }

  return response.json();
}

export { SUPABASE_URL };
