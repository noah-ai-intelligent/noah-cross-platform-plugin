import { API_BASE_URL, CLIENT_PLATFORM } from "./config";
import { getAccessToken, getRefreshToken, saveTokens, clearTokens } from "./tokenStorage";

export class ApiError extends Error {
  constructor(
    public status: number,
    message: string,
    public body?: any,
    public retryAfter?: number
  ) {
    super(message);
    this.name = "ApiError";
  }
}

export function timeHeaders(): Record<string, string> {
  const now = new Date();
  return {
    "X-User-Time": now.toISOString(),
    "X-User-Timezone": Intl.DateTimeFormat().resolvedOptions().timeZone,
    "X-User-UTC-Offset": String(-now.getTimezoneOffset()),
  };
}

let refreshPromise: Promise<string | null> | null = null;

export async function request<T>(path: string, init: RequestInit = {}): Promise<T> {
  let token = await getAccessToken();
  let res = await fetch(`${API_BASE_URL}${path}`, {
    ...init,
    headers: {
      "Content-Type": "application/json",
      "X-Client-Platform": CLIENT_PLATFORM,
      "ngrok-skip-browser-warning": "69420",
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
      ...timeHeaders(),
      ...(init.headers ?? {}),
    },
  });

  if (res.status === 401 && token && path !== "/auth/refresh") {
    if (!refreshPromise) {
      refreshPromise = (async () => {
        const refreshToken = await getRefreshToken();
        if (!refreshToken) return null;
        try {
          const refreshRes = await fetch(`${API_BASE_URL}/auth/refresh`, {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
              "X-Client-Platform": CLIENT_PLATFORM,
              "ngrok-skip-browser-warning": "69420",
              ...timeHeaders(),
            },
            body: JSON.stringify({ refresh_token: refreshToken }),
          });

          if (refreshRes.ok) {
            const data = await refreshRes.json();
            await saveTokens({
              access_token: data.access_token,
              refresh_token: data.refresh_token,
            });
            return data.access_token;
          } else {
            await clearTokens();
            return null;
          }
        } catch {
          await clearTokens();
          return null;
        } finally {
          refreshPromise = null;
        }
      })();
    }

    const newToken = await refreshPromise;
    if (newToken) {
      token = newToken;
      res = await fetch(`${API_BASE_URL}${path}`, {
        ...init,
        headers: {
          "Content-Type": "application/json",
          "X-Client-Platform": CLIENT_PLATFORM,
          "ngrok-skip-browser-warning": "69420",
          Authorization: `Bearer ${token}`,
          ...timeHeaders(),
          ...(init.headers ?? {}),
        },
      });
    }
  }

  if (!res.ok) {
    let detail: any;
    const raw = await res.text().catch(() => "");
    try {
      detail = raw ? JSON.parse(raw) : undefined;
    } catch {
      detail = raw;
    }

    let errorMessage = `${res.status} ${res.statusText}`;
    if (typeof detail === "string") {
      errorMessage = detail;
    } else if (detail && typeof detail === "object") {
      const msg = detail.detail?.message || detail.detail || detail.message;
      if (msg) {
        if (typeof msg === "string") {
          errorMessage = msg;
        } else {
          try {
            errorMessage = JSON.stringify(msg);
          } catch {
            errorMessage = String(msg);
          }
        }
      }
    }

    const retryAfter = res.headers.get("retry-after")
      ? parseInt(res.headers.get("retry-after")!, 10)
      : undefined;

    throw new ApiError(res.status, errorMessage, detail, retryAfter);
  }

  if (res.status === 204) return {} as T;

  const contentType = res.headers.get("content-type") ?? "";
  if (contentType.includes("application/json")) {
    return res.json() as Promise<T>;
  } else {
    return res.blob() as unknown as Promise<T>;
  }
}
