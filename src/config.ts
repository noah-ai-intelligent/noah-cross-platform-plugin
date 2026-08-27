/** Deployment-time constants. Vite inlines `import.meta.env.*` at build time —
 * set these via `.env.production` / `.env.development` (not committed), same
 * convention as noah-frontend-v2. */

import { currentHost } from "./document/capabilities";

export const API_BASE_URL: string =
  (import.meta.env.VITE_NOAH_API_URL as string | undefined) ??
  "https://noah.enpointe.io/api/v1";

export const CLIENT_PLATFORM = currentHost().startsWith("Google") ? "google_addon" : "office_addin";

/** This add-in's own hosted callback page — must be registered as an
 * Authorized Redirect URI on both the Google OAuth client and the
 * Microsoft/Azure AD app backing Noah's SSO (see the implementation plan,
 * "SSO app registration for public/worldwide use"). */
export const APP_BASE_URL: string =
  (import.meta.env.VITE_NOAH_OFFICE_URL as string | undefined) ??
  (window.location.origin.includes("googleusercontent.com")
    ? "https://noah-office.enpointe.io"
    : window.location.origin);

export const REDIRECT_URI: string =
  (import.meta.env.VITE_NOAH_REDIRECT_URI as string | undefined) ??
  `${APP_BASE_URL}/login/login.html`;
