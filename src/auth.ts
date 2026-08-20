/** Taskpane-side half of sign-in: opens the login dialog (`src/login`) via
 * the Office Dialog API and resolves once it messages back a token pair. */

import { getAccessToken, saveTokens } from "./tokenStorage";

export type AuthResult =
  | { status: "signed-in" }
  | { status: "cancelled" }
  | { status: "error"; message: string };

export async function isSignedIn(): Promise<boolean> {
  return (await getAccessToken()) !== null;
}

export function signIn(): Promise<AuthResult> {
  return new Promise((resolve) => {
    if (!window.Office?.context?.ui) {
      window.location.href = "/login";
      return resolve({ status: "cancelled" });
    }

    window.Office.context.ui.displayDialogAsync(
      `${window.location.origin}/login`,
      { height: 60, width: 30, promptBeforeOpen: false },
      (asyncResult) => {
        if (asyncResult.status === window.Office?.AsyncResultStatus.Failed) {
          resolve({
            status: "error",
            message:
              asyncResult.error?.message ?? "Could not open the sign-in window.",
          });
          return;
        }

        const dialog = asyncResult.value;

        dialog.addEventHandler(
          window.Office?.EventType.DialogMessageReceived,
          async (arg) => {
            const message = arg as Office.DialogParentMessageReceivedEventArgs;
            dialog.close();
            try {
              const payload = JSON.parse(message.message) as {
                ok: boolean;
                tokens?: { access_token: string; refresh_token: string };
                orgId?: string;
              };
              if (payload.ok && payload.tokens) {
                await saveTokens(payload.tokens);
                if (payload.orgId) {
                  localStorage.setItem("noah_selected_org", payload.orgId);
                }
                resolve({ status: "signed-in" });
              } else {
                resolve({ status: "error", message: "Sign-in failed." });
              }
            } catch {
              resolve({ status: "error", message: "Sign-in failed." });
            }
          }
        );

        dialog.addEventHandler(window.Office?.EventType.DialogEventReceived, (arg) => {
          const event = arg as { error: number };
          dialog.close();
          // 12006 = user closed the dialog.
          if (event.error === 12006) {
            resolve({ status: "cancelled" });
          } else {
            resolve({ status: "error", message: `Dialog closed (${event.error}).` });
          }
        });
      }
    );
  });
}
