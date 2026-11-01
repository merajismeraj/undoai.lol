/**
 * The counter.
 *
 * The only module on the page allowed to touch the network, and all it can say
 * is "one more". A POST carries no body at all — no file name, no size, no
 * content, nothing about who sent it — so the number on the page costs a
 * document exactly nothing in privacy. `test/no-network.test.ts` enforces that
 * this file is the only one with a `fetch` in it, and that it imports nothing
 * from the document pipeline.
 *
 * Every call fails soft. A counter is decoration; cleaning a file is the
 * product, and the product must work with the endpoint down, blocked by an
 * extension, or offline.
 */

const ENDPOINT = "/api/count";

type CountResponse = { count: number | null };

async function ask(method: "GET" | "POST"): Promise<number | null> {
  try {
    const response = await fetch(ENDPOINT, { method, cache: "no-store" });
    if (!response.ok) return null;
    const body = (await response.json()) as CountResponse;
    return typeof body.count === "number" ? body.count : null;
  } catch {
    return null;
  }
}

/** The running total, or null if the counter isn't answering. */
export function readCount(): Promise<number | null> {
  return ask("GET");
}

/** One more document cleaned. Returns the new total, or null. */
export function recordCleaned(): Promise<number | null> {
  return ask("POST");
}
