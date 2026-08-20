/** Ribbon FunctionFile. The manifest currently only wires the "Ask Noah"
 * button as a ShowTaskpane action (no ExecuteFunction ribbon commands), so
 * this file has nothing to associate yet — it exists because Office requires
 * a valid FunctionFile URL per host in the manifest. Add
 * `Office.actions.associate("someId", handler)` here if a ribbon quick-action
 * (e.g. "Insert last answer") is added later. */

Office.onReady();
