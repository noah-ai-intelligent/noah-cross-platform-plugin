/** Driving one background ask from start to answer.
 *
 * The synchronous `/ask` route can't suspend, so the backend never offers the
 * agent the document read tools on it. `/ask/start` can — the turn parks a
 * request, we perform it in the host, post the result, and the turn continues.
 * That round trip is the entire reason this loop exists.
 *
 * Poll-based for now. C6 swaps the wait for a pushed stream, and deliberately
 * without changing anything else: the stream carries the identical payload
 * this poll returns, so the fulfil logic below is the same either way. That
 * property is what makes the stream an optimisation rather than a second
 * implementation.
 */

import { JobRequest, JobState, getJob, postJobContext } from "../addonClient";
import { Host } from "../document/capabilities";
import { fulfil } from "../document/readOps";

export interface RunOptions {
  host: Host;
  /** Shown while the agent is reading — "Reading Sheet2!A1:D40". The user
   * should always be able to see what Noah asked their document for. */
  onActivity?: (label: string) => void;
  /** Called as the answer streams in from the server. */
  onAnswerUpdate?: (answer: any) => void;
  signal?: { cancelled: boolean };
}

/** Poll cadence. Starts tight because most turns land in a few seconds, then
 * backs off so a long report isn't polled fifty times. Reset to the floor
 * whenever the job asks us for something — during a read exchange the latency
 * that matters is ours, not the model's. */
const FIRST_WAIT_MS = 700;
const MAX_WAIT_MS = 3000;
const STEP_MS = 400;
/** A turn that never finishes shouldn't hold the taskpane forever. The job
 * outlives this on the server; the user can ask again. */
const OVERALL_TIMEOUT_MS = 2 * 60 * 1000;

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/** Perform every read the turn is waiting on, and post the results together.
 *
 * All of them in one POST rather than one call each: the turn is blocked on
 * the slowest, and a second round trip per request would add latency for
 * nothing.
 */
async function fulfilRequests(
  jobId: string,
  requests: JobRequest[],
  options: RunOptions
): Promise<void> {
  const reads = requests.filter((r) => r.tool !== "edit_document");
  if (!reads.length) return;

  options.onActivity?.(reads[0].human_label || "Reading the document…");

  const results = [];
  for (const request of reads) {
    // Sequentially, not in parallel: Excel.run batches don't interleave
    // safely, and the agent's read budget means there are only ever a
    // handful.
    results.push(await fulfil(request, options.host));
  }
  await postJobContext(jobId, results);
}

export async function runJob(jobId: string, options: RunOptions): Promise<JobState> {
  const deadline = Date.now() + OVERALL_TIMEOUT_MS;
  let wait = FIRST_WAIT_MS;

  while (Date.now() < deadline) {
    if (options.signal?.cancelled) {
      return { status: "error", error: "Cancelled." };
    }

    const job = await getJob(jobId);

    if (job.answer && options.onAnswerUpdate) {
      options.onAnswerUpdate(job.answer);
    }

    if (job.status === "done" || job.status === "error") return job;

    if (job.status === "running" && job.activity) {
      options.onActivity?.(job.activity);
    }

    if (job.status === "needs_context" && job.requests?.length) {
      await fulfilRequests(jobId, job.requests, options);
      // The turn is unblocked now, so go straight back to a tight poll rather
      // than carrying the backed-off delay into the next phase.
      wait = FIRST_WAIT_MS;
      continue;
    }

    await sleep(wait);
    wait = Math.min(wait + STEP_MS, MAX_WAIT_MS);
  }

  return {
    status: "error",
    error: "That took longer than expected. Ask again, or check back shortly.",
  };
}
