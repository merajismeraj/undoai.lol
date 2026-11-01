import assert from "node:assert/strict";
import { readFile, readdir } from "node:fs/promises";
import path from "node:path";
import test from "node:test";

/**
 * The promise on the page is that a document never leaves the machine. That is
 * only worth as much as it is checkable, so it is checked here: no module that
 * ships to the browser may reference a network API, and the project may not
 * grow a runtime dependency that could smuggle one in.
 *
 * Both tests read the source rather than the behaviour on purpose — an
 * exfiltrating call that only fires on a rare document would pass every
 * functional test in the suite and fail this one.
 */

const SRC = new URL("../src/", import.meta.url).pathname;

const NETWORK_APIS =
  /\b(fetch|XMLHttpRequest|WebSocket|EventSource|RTCPeerConnection|importScripts)\s*\(|navigator\.(sendBeacon|serviceWorker|connection)|new\s+(Worker|SharedWorker|BroadcastChannel)\b/;

async function sourceFiles(dir: string): Promise<string[]> {
  const found: string[] = [];
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) found.push(...(await sourceFiles(full)));
    else if (entry.name.endsWith(".ts")) found.push(full);
  }
  return found;
}

test("nothing that ships to the browser can reach the network", async () => {
  const files = await sourceFiles(SRC);
  assert.ok(files.length >= 5, "expected to find the source modules");

  for (const file of files) {
    const source = await readFile(file, "utf8");
    for (const [index, line] of source.split("\n").entries()) {
      // Comments and the regexes that *detect* AI tool names inside a document
      // are prose and patterns, not calls.
      const code = line.replace(/\/\/.*$/, "").replace(/^\s*\*.*$/, "");
      assert.doesNotMatch(
        code,
        NETWORK_APIS,
        `${path.relative(SRC, file)}:${index + 1} reaches the network`,
      );
    }
  }
});

test("the project has no runtime dependencies", async () => {
  const manifest = JSON.parse(
    await readFile(new URL("../package.json", import.meta.url), "utf8"),
  ) as { dependencies?: Record<string, string> };

  assert.deepEqual(
    manifest.dependencies ?? {},
    {},
    "a runtime dependency can ship code this audit never sees",
  );
});
