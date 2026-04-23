import { readFile, writeFile } from "node:fs/promises";

const templatePath = new URL("../manifest.template.xml", import.meta.url);
const baseUrl = (process.argv[2] || "https://localhost:3000").replace(/\/+$/, "");
const outputPath = process.argv[3]
  ? new URL(`file://${process.argv[3]}`)
  : new URL("../manifest.xml", import.meta.url);

const template = await readFile(templatePath, "utf8");
const rendered = template.replaceAll("__ADDIN_BASE_URL__", baseUrl);
await writeFile(outputPath, rendered, "utf8");
console.log(`Rendered manifest to ${outputPath.pathname} for ${baseUrl}`);
