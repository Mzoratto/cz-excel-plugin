import fs from "node:fs/promises";
import path from "node:path";

function parseArgs(argv) {
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

async function build() {
  const args = parseArgs(process.argv.slice(2));
  const host = (args.get("host") || "https://localhost:5173").replace(/\/$/, "");
  const output = args.get("output") || "manifest.generated.xml";
  const templatePath = path.resolve(process.cwd(), "manifest.template.xml");
  const template = await fs.readFile(templatePath, "utf8");

  const taskpaneUrl = `${host}/taskpane.html`;
  const content = template
    .replace(/{{HOST_ORIGIN}}/g, host)
    .replace(/{{TASKPANE_URL}}/g, taskpaneUrl);

  const outputPath = path.resolve(process.cwd(), output);
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await fs.writeFile(outputPath, content, "utf8");
  console.log(`✅ Manifest generated at ${outputPath}`);
}

build().catch((error) => {
  console.error(`❌ Manifest generation failed: ${error.message}`);
  process.exitCode = 1;
});
