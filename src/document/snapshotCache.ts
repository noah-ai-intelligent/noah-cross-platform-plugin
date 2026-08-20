/** Sending the selection once per conversation instead of once per question.
 *
 * The backend caches an uploaded selection under a content hash and lets the
 * client send that id on later turns. The obvious implementation — reproduce
 * the server's hash here — is a trap: it is a sha256 over the server's own
 * canonical JSON encoding, and matching that byte-for-byte in TypeScript AND
 * again in Apps Script is a bug farm that fails silently (a mismatch just
 * means the cache never hits, and nobody notices except the latency).
 *
 * So we never compute it. We upload once, keep the id the server gave us
 * against OUR own cheap fingerprint of the same selection, and send that id
 * next time. The fingerprint only has to be stable within this client.
 *
 * `409 context_required` is the designed recovery: the server is telling us it
 * no longer holds an id we claimed. That is a normal branch — the cache is
 * allowed to be wrong — not an error to show anyone.
 */

import { Selection } from "./selection";
import { fingerprint } from "./anchor";

/** In-memory only, and deliberately so. It is keyed to a document the user has
 * open right now; surviving a taskpane reload would buy one avoided upload and
 * risk pairing a stale id with a changed selection. */
const idsByFingerprint = new Map<string, string>();

/** The local key for a selection — its content, ignoring where it came from,
 * which mirrors what the server excludes from its own digest (a region read
 * twice from the same place is the same snapshot). */
export function localFingerprint(selection: Selection): string {
  return fingerprint({
    grid: selection.grid ? { rows: selection.grid.rows } : null,
    text: selection.text ? selection.text.blocks.map((b) => [b.type, b.text]) : null,
    shapes: selection.shapes
      ? selection.shapes.slides.map((s) => [s.title, s.body, s.notes])
      : null,
  });
}

export function remember(selection: Selection, snapshotId: string): void {
  if (!snapshotId) return;
  idsByFingerprint.set(localFingerprint(selection), snapshotId);
}

export function lookup(selection: Selection): string | undefined {
  return idsByFingerprint.get(localFingerprint(selection));
}

/** Called when the server says it no longer has an id we sent. Drops every
 * mapping pointing at it, so the next ask uploads instead of claiming it
 * again. */
export function forget(snapshotId: string): void {
  for (const [key, value] of idsByFingerprint) {
    if (value === snapshotId) idsByFingerprint.delete(key);
  }
}

export function clear(): void {
  idsByFingerprint.clear();
}
