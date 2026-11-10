#!/usr/bin/env node
/**
 * V4.0.0 optional Artifact Tool renderer for render_preview.py.
 *
 * This script is invoked only when the caller supplies a Node runtime and
 * @oai/artifact-tool package directory. It does not form part of the ordinary
 * Python dependency path.
 */

import fs from "node:fs/promises";
import path from "node:path";
import { createRequire } from "node:module";
import { parseArgs } from "node:util";
import { pathToFileURL } from "node:url";

const SKILL_VERSION = "4.0.0";

const { values } = parseArgs({
  options: {
    input: { type: "string" },
    "output-dir": { type: "string" },
    "node-modules": { type: "string" },
    scale: { type: "string", default: "1" },
  },
});

if (!values.input || !values["output-dir"] || !values["node-modules"]) {
  console.error(
    "Usage: node artifact_tool_render.mjs --input <pptx> --output-dir <dir> --node-modules <dir> [--scale 1]"
  );
  process.exit(2);
}

const inputPath = path.resolve(values.input);
const outputDir = path.resolve(values["output-dir"]);
const nodeModules = path.resolve(values["node-modules"]);
const scale = Number.parseFloat(values.scale);
if (!Number.isFinite(scale) || scale <= 0) {
  console.error("--scale must be a positive number");
  process.exit(2);
}

const requireFromRuntime = createRequire(path.join(nodeModules, "noop.cjs"));
const artifactToolEntry = requireFromRuntime.resolve("@oai/artifact-tool");
const { FileBlob, PresentationFile } = await import(
  pathToFileURL(artifactToolEntry).href
);

await fs.mkdir(outputDir, { recursive: true });
const presentation = await PresentationFile.importPptx(await FileBlob.load(inputPath));
const slides = presentation.slides.items;
const paths = [];

for (let index = 0; index < slides.length; index += 1) {
  const target = path.join(outputDir, `slide_${String(index + 1).padStart(2, "0")}.png`);
  const preview = await presentation.export({
    slide: slides[index],
    format: "png",
    scale,
  });
  await fs.writeFile(target, Buffer.from(await preview.arrayBuffer()));
  paths.push(target);
}

console.log(JSON.stringify({ skillVersion: SKILL_VERSION, slideCount: slides.length, paths }, null, 2));
