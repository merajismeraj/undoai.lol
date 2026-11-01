import assert from "node:assert/strict";
import { readFile, readdir } from "node:fs/promises";
import path from "node:path";
import test from "node:test";

/**
 * The promise on the page is that a document never leaves the machine. That is
 * only worth as much as it is checkable, so it is checked here rather than
 * asserted in prose.
 *
 * One module may talk to the network — the counter, which sends an empty POST
 * and gets a number back. Everything else, the whole document pipeline
 * included, must have no way to reach anything at all. These tests read the
 * source rather than the behaviour on purpose: an exfiltrating call that only
 * fires on a rare document would pass every functional test in the suite and
 * fail this one.
 */

const SRC = new URL("../src/", import.meta.url).pathname;

/** The only file allowed a `fetch`, and the only thing it is allowed to hit. */
const COUNTER = "counter.ts";
const COUNTER_ENDPOINT = '"/api/count"';

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

/** Source with comments stripped — prose about `fetch`, and the regexes that
    detect AI tool names inside a document, are not calls. */
async function codeOf(file: string): Promise<string[]> {
  const source = await readFile(file, "utf8");
  return source
    .split("\n")
    .map((line) => line.replace(/\/\/.*$/, "").replace(/^\s*\*.*$/, ""));
}

test("only the counter can reach the network", async () => {
  const files = await sourceFiles(SRC);
  assert.ok(files.length >= 6, "expected to find the source modules");

  for (const file of files) {
    const name = path.relative(SRC, file);
    for (const [index, line] of (await codeOf(file)).entries()) {
      if (!NETWORK_APIS.test(line)) continue;
      assert.equal(name, COUNTER, `${name}:${index + 1} reaches the network`);
    }
  }
});

test("the document pipeline has no way to reach anything", async () => {
  for (const file of await sourceFiles(path.join(SRC, "ooxml"))) {
    const name = path.relative(SRC, file);
    for (const [index, line] of (await codeOf(file)).entries()) {
      assert.doesNotMatch(line, NETWORK_APIS, `${name}:${index + 1} reaches the network`);
      assert.doesNotMatch(
        line,
        /^\s*import\s[^;]*from\s+["'](?!\.)/,
        `${name}:${index + 1} imports something that isn't ours`,
      );
    }
  }
});

test("the counter can only ever say 'one more'", async () => {
  // Comments stripped first: prose is full of apostrophes and slashes, and a
  // scan that reads them would be measuring the wrong thing.
  const code = (await codeOf(path.join(SRC, COUNTER))).join("\n");
  const source = await readFile(path.join(SRC, COUNTER), "utf8");

  // It may hold exactly one address, and that address is our own endpoint.
  const urls = code.match(/(["'])(?:\/|https?:)[^"'\n]*\1/g) ?? [];
  assert.deepEqual(urls, [COUNTER_ENDPOINT], "the counter points somewhere unexpected");

  // No body means nothing about the document can ride along.
  assert.doesNotMatch(source, /\bbody\s*:/, "the counter sends a request body");
  assert.doesNotMatch(source, /from\s+["']\.\/ooxml/, "the counter can see the document pipeline");
});

test("nothing third-party ships to the browser", async () => {
  for (const file of await sourceFiles(SRC)) {
    const name = path.relative(SRC, file);
    for (const [index, line] of (await codeOf(file)).entries()) {
      assert.doesNotMatch(
        line,
        /^\s*import\s[^;]*from\s+["'](?!\.)/,
        `${name}:${index + 1} pulls in a dependency the audit never sees`,
      );
    }
  }
});
