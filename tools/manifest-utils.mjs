import fs from "node:fs/promises";
import path from "node:path";

function normalizeHost(host) {
  if (!host) {
    return "https://localhost:5173";
  }
  return host.endsWith("/") ? host.slice(0, -1) : host;
}

export async function buildManifest({ host, output }) {
  const rootDir = process.cwd();
  const templatePath = path.resolve(rootDir, "manifest.template.xml");
  const template = await fs.readFile(templatePath, "utf8");
  const resolvedHost = normalizeHost(host);
  const taskpaneUrl = `${resolvedHost}/taskpane.html`;

  const content = template
    .replace(/{{HOST_ORIGIN}}/g, resolvedHost)
    .replace(/{{TASKPANE_URL}}/g, taskpaneUrl);

  const outputPath = path.resolve(rootDir, output);
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await fs.writeFile(outputPath, content, "utf8");
  return outputPath;
}

export function parseArgs(argv) {
  const args = new Map();
  for (const arg of argv) {
    if (!arg.startsWith("--")) {
      continue;
    }
    const [key, value] = arg.slice(2).split("=");
    args.set(key, value ?? "");
  }
  return args;
}
