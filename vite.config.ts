import { defineConfig } from "vite";

/**
 * No framework, no runtime dependencies: the whole tool is four modules of
 * plain TypeScript plus a page that wires them to a file input. Vite is here to
 * bundle and nothing else, so the built output is a handful of static files
 * that will sit behind any CDN — or none.
 */
export default defineConfig({
  base: "./",
  build: { target: "es2022", sourcemap: true },
});
