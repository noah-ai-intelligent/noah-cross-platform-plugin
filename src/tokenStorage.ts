/** Session token persistence for the taskpane.
 *
 * Deliberately uses `localStorage` (device-local, sandboxed to this
 * origin) and NOT `Office.context.document.settings` / `roamingSettings` —
 * those persist INSIDE the Word/Excel/PowerPoint file itself, so a refresh
 * token stored there would travel with the document to anyone it's shared
 * with. `localStorage` stays on the device, sharing the same threat model as
 * the SPA (`noah-frontend-v2/src/api/chat.ts`).
 */

const ACCESS_TOKEN_KEY = "noah.access_token";
const REFRESH_TOKEN_KEY = "noah.refresh_token";

export interface TokenPair {
  access_token: string;
  refresh_token: string;
}

function getStorage() {
  return {
    setItem: async (k: string, v: string) => localStorage.setItem(k, v),
    getItem: async (k: string) => localStorage.getItem(k),
    removeItem: async (k: string) => localStorage.removeItem(k),
  };
}

export async function saveTokens(tokens: TokenPair): Promise<void> {
  const s = getStorage();
  await s.setItem(ACCESS_TOKEN_KEY, tokens.access_token);
  await s.setItem(REFRESH_TOKEN_KEY, tokens.refresh_token);
}

export async function getAccessToken(): Promise<string | null> {
  return getStorage().getItem(ACCESS_TOKEN_KEY);
}

export async function getRefreshToken(): Promise<string | null> {
  return getStorage().getItem(REFRESH_TOKEN_KEY);
}

export async function clearTokens(): Promise<void> {
  const s = getStorage();
  await s.removeItem(ACCESS_TOKEN_KEY);
  await s.removeItem(REFRESH_TOKEN_KEY);
}
