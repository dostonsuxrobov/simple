import { createHash } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { defineConfig, type Plugin } from "vite";

/**
 * WordCanvas ships its 28 fallback fonts as base64 data URLs inside its code,
 * once in the editor bundle and once more in its worker (about 15 MB each).
 * Every window then keeps that text in memory twice on top of the decoded
 * fonts: an empty Documents window measured 162 MB with them and 112 MB
 * without. After the build, each embedded font becomes a file in
 * assets/wc-fonts and the code refers to it by URL; the fonts load exactly as
 * before (the engine fetches the URL either way).
 */
function fontsAsFiles(): Plugin {
  const pattern = /"data:font\/(ttf|otf|woff2?);base64,([A-Za-z0-9+/=]+)"/g;
  let outDir = "dist";
  return {
    name: "simple-fonts-as-files",
    apply: "build",
    configResolved(config) {
      outDir = path.resolve(config.root, config.build.outDir);
    },
    closeBundle() {
      const assets = path.join(outDir, "assets");
      if (!fs.existsSync(assets)) return;
      const fontsDir = path.join(assets, "wc-fonts");
      const written = new Set<string>();
      for (const name of fs.readdirSync(assets)) {
        if (!name.endsWith(".js")) continue;
        const file = path.join(assets, name);
        const source = fs.readFileSync(file, "utf8");
        if (!source.includes("data:font/")) continue;
        // A worker resolves a relative URL against its own script, a page chunk against the chunk.
        const base = /^worker-/.test(name) ? "self.location.href" : "import.meta.url";
        const next = source.replace(pattern, (_all, kind: string, base64: string) => {
          const bytes = Buffer.from(base64, "base64");
          const fontName = `${createHash("sha256").update(bytes).digest("hex").slice(0, 16)}.${kind}`;
          if (!written.has(fontName)) {
            fs.mkdirSync(fontsDir, { recursive: true });
            fs.writeFileSync(path.join(fontsDir, fontName), bytes);
            written.add(fontName);
          }
          return `new URL("wc-fonts/${fontName}",${base}).href`;
        });
        if (next !== source) fs.writeFileSync(file, next);
      }
      // An embedded font left behind would silently bring the memory cost back.
      for (const name of fs.readdirSync(assets)) {
        if (name.endsWith(".js") && fs.readFileSync(path.join(assets, name), "utf8").includes("data:font/")) {
          throw new Error(`${name} still embeds a font as a data URL; update fontsAsFiles in vite.config.ts.`);
        }
      }
    },
  };
}

export default defineConfig({
  base: "./",
  plugins: [fontsAsFiles()],
  build: {
    target: "es2022",
    sourcemap: false,
    chunkSizeWarningLimit: 18000,
  },
});
