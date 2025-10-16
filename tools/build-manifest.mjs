import { buildManifest, parseArgs } from "./manifest-utils.mjs";

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const host = args.get("host");
  const output = args.get("output") || "manifest.generated.xml";

  try {
    const outputPath = await buildManifest({ host, output });
    console.log(`✅ Manifest generated at ${outputPath}`);
  } catch (error) {
    console.error(`❌ Manifest generation failed: ${error instanceof Error ? error.message : String(error)}`);
    process.exitCode = 1;
  }
}

main();
