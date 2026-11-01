/**
 * The page: a file input, three checkboxes, and a readout.
 *
 * Plain DOM rather than a framework — there is one interaction here (a file
 * goes in, a report and a download come out), and a framework would be more
 * code than the tool it is wrapping. Everything user-supplied (file names,
 * property values pulled out of the document) is set with `textContent`, never
 * `innerHTML`: a document is untrusted input, and a filename is a fine place to
 * hide a `<script>`.
 */

import { DEFAULT_OPTIONS, MAX_BYTES, cleanDocument, type CleanOptions, type CleanResult } from "./ooxml/clean";
import { supportsZip } from "./ooxml/zip";
import type { Group } from "./ooxml/text";

const GROUP_TITLES: Record<Group, string> = {
  typography: "Typography",
  invisible: "Invisible characters",
  symbols: "Symbols and emoji",
  leftovers: "Chat and markdown leftovers",
  revisions: "Comments and revisions",
  metadata: "Metadata and fingerprints",
  review: "Left for you to decide",
};

function must<T extends HTMLElement>(id: string): T {
  const node = document.getElementById(id);
  if (!node) throw new Error(`Missing #${id} in the page`);
  return node as T;
}

const dropZone = must<HTMLLabelElement>("drop");
const dropLabel = must<HTMLSpanElement>("drop-label");
const fileInput = must<HTMLInputElement>("file");
const results = must<HTMLDivElement>("results");
const authorInput = must<HTMLInputElement>("author");
const toggles = {
  metadata: must<HTMLInputElement>("opt-metadata"),
  revisions: must<HTMLInputElement>("opt-revisions"),
  leftovers: must<HTMLInputElement>("opt-leftovers"),
};

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
  text?: string,
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function formatBytes(bytes: number): string {
  return bytes < 1024 * 1024
    ? `${Math.max(1, Math.round(bytes / 1024))} KB`
    : `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function currentOptions(): CleanOptions {
  return {
    ...DEFAULT_OPTIONS,
    metadata: toggles.metadata.checked,
    revisions: toggles.revisions.checked,
    leftovers: toggles.leftovers.checked,
    author: authorInput.value,
  };
}

/* ── running files ───────────────────────────────────────────────────────── */

let busy = false;

async function run(files: File[]) {
  if (files.length === 0 || busy) return;

  // Raw deflate lands in every current browser; an old one would otherwise fail
  // somewhere deep inside the zip reader with nothing useful to say.
  if (!supportsZip()) {
    results.prepend(
      status(
        files[0].name,
        "This browser can't rebuild a .docx. Use a current Chrome, Edge, Firefox or Safari.",
        true,
      ),
    );
    return;
  }

  busy = true;
  dropLabel.textContent = "cleaning…";
  const options = currentOptions();

  for (const file of files) {
    const card = status(file.name, "cleaning…");
    results.prepend(card);

    try {
      const result = await cleanDocument(file, options);
      card.replaceWith(report(result));
    } catch (err) {
      const message = err instanceof Error ? err.message : "Could not read that file.";
      card.replaceWith(status(file.name, message, true));
    }
  }

  busy = false;
  dropLabel.textContent = "drop another, or click to pick one";
}

function status(name: string, message: string, failed = false): HTMLElement {
  const card = el("section", "result");
  const line = el("p", failed ? "status failed" : "status");
  line.append(el("span", undefined, `${name} — `), el("span", undefined, message));
  card.append(line);
  return card;
}

/* ── the readout ─────────────────────────────────────────────────────────── */

function report(result: CleanResult): HTMLElement {
  const card = el("section", "result");

  const head = el("div", "result-head");
  const heading = el("div");
  heading.append(el("p", "result-name", result.fileName));

  const changes = result.findings
    .filter((f) => f.group !== "review")
    .reduce((sum, f) => sum + f.count, 0);
  heading.append(
    el(
      "p",
      "result-meta",
      `${result.kind} · ${changes} change${changes === 1 ? "" : "s"} · ${formatBytes(
        result.originalBytes,
      )} → ${formatBytes(result.cleanedBytes)}`,
    ),
  );

  const button = el("button", "button", "Download");
  button.type = "button";
  button.addEventListener("click", () => download(result));

  head.append(heading, button);
  card.append(head);

  const body = el("div", "result-body");

  if (changes === 0) {
    body.append(
      el(
        "p",
        "result-meta",
        "Nothing mechanical to strip — this file already reads like something a person typed.",
      ),
    );
  } else {
    const groups = el("div", "groups");
    for (const group of new Set(result.findings.map((f) => f.group))) {
      const block = el("div");
      block.append(el("h3", "group-title", GROUP_TITLES[group]));

      const list = el("ul", "findings");
      for (const finding of result.findings.filter((f) => f.group === group)) {
        const item = el("li", "finding");
        item.append(el("span", undefined, finding.label), el("b", undefined, String(finding.count)));
        list.append(item);
      }

      block.append(list);
      groups.append(block);
    }
    body.append(groups);
  }

  if (result.notes.length > 0) {
    const block = el("div");
    block.append(el("h3", "group-title", "What was in there"));
    const list = el("ul", "notes");
    for (const note of result.notes) list.append(el("li", undefined, note));
    block.append(list);
    body.append(block);
  }

  if (result.cliches.length > 0) {
    const block = el("div");
    block.append(el("h3", "group-title", "Wording nobody can fix but you"));
    block.append(
      el(
        "p",
        "result-meta",
        "Left exactly as written. These are the words that still read as machine-made after every character is fixed.",
      ),
    );

    const chips = el("div", "chips");
    for (const cliche of result.cliches) {
      chips.append(el("span", "chip", `${cliche.label} · ${cliche.count}`));
    }
    block.append(chips);
    body.append(block);
  }

  card.append(body);
  return card;
}

function download(result: CleanResult) {
  const url = URL.createObjectURL(result.blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = result.fileName;
  link.click();
  // Revoking in the same tick cancels the download in some browsers — the click
  // only queues it.
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

/* ── wiring ──────────────────────────────────────────────────────────────── */

fileInput.addEventListener("change", () => {
  // Copy the list out before resetting the input: `files` is live, and clearing
  // the value empties the very list we are about to read. Resetting is what
  // lets the same file be run again after changing an option.
  const files = [...(fileInput.files ?? [])];
  fileInput.value = "";
  void run(files);
});

dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("dragging");
});

dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragging"));

dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("dragging");
  void run([...(event.dataTransfer?.files ?? [])]);
});

// The size limit lives in one place — the module — and the page repeats it.
const hint = document.querySelector(".drop-hint");
if (hint) hint.textContent = `.docx or .pptx, up to ${formatBytes(MAX_BYTES)}`;
