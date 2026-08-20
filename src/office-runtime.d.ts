/** Ambient declaration for the bit of `OfficeRuntime` we use.
 *
 * `@types/office-js` mentions `OfficeRuntime.auth` in a doc comment but does
 * not declare the global — the declarations live in a separate
 * `@types/office-runtime` package. Rather than take a second types dependency
 * for one API, this declares exactly the surface `tokenStorage.ts` calls.
 *
 * Keep it minimal on purpose: a hand-written declaration that claims more than
 * we use is a lie the compiler will believe. If we start using `OfficeRuntime`
 * more widely, replace this file with the real types package.
 */

declare namespace OfficeRuntime {
  /** Device-local, sandboxed to this add-in. Deliberately NOT
   * `Office.context.document.settings`, which persists inside the document and
   * would travel with the file — see the note in `tokenStorage.ts`. */
  const storage: {
    getItem(key: string): Promise<string | null>;
    setItem(key: string, value: string): Promise<void>;
    removeItem(key: string): Promise<void>;
    clear(): Promise<void>;
    getItems(keys: string[]): Promise<{ [key: string]: string | null }>;
    setItems(items: { [key: string]: string }): Promise<void>;
    removeItems(keys: string[]): Promise<void>;
  };
}
