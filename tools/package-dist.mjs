import fs from "node:fs";
import fsPromises from "node:fs/promises";
import path from "node:path";
import archiver from "archiver";

function formatTimestamp(date) {
  return date.toISOString().replace(/[:.]/g, "-");
}

async function ensureFileExists(filePath, label) {
  try {
    await fsPromises.access(filePath, fs.constants.F_OK);
  } catch {
    throw new Error(`Expected ${label} at ${filePath} – run the build first.`);
  }
}

async function createArchive() {
  const rootDir = process.cwd();
  const distDir = path.resolve(rootDir, "dist");
  const manifestPath = path.resolve(rootDir, "manifest.xml");
  const releaseDir = path.resolve(rootDir, "release");

  await ensureFileExists(distDir, "dist folder");
  await ensureFileExists(manifestPath, "manifest.xml");
  await fsPromises.mkdir(releaseDir, { recursive: true });

  const existingEntries = await fsPromises.readdir(releaseDir, { withFileTypes: true });
  await Promise.all(
    existingEntries
      .filter((entry) => entry.isFile() && entry.name.endsWith(".zip"))
      .map((entry) => fsPromises.unlink(path.join(releaseDir, entry.name)))
  );

  const timestamp = formatTimestamp(new Date());
  const archiveName = `cz-excel-copilot-${timestamp}.zip`;
  const outputPath = path.join(releaseDir, archiveName);

  const output = fs.createWriteStream(outputPath);
  const archive = archiver("zip", { zlib: { level: 9 } });

  archive.on("warning", (error) => {
    if (error.code === "ENOENT") {
      console.warn(error.message);
    } else {
      throw error;
    }
  });

  archive.on("error", (error) => {
    throw error;
  });

  archive.pipe(output);
  archive.file(manifestPath, { name: "manifest.xml" });
  archive.directory(distDir, "dist");

  await archive.finalize();

  return outputPath;
}

createArchive()
  .then((outputPath) => {
    console.log(`✅ Package ready: ${outputPath}`);
  })
  .catch((error) => {
    console.error(`❌ Packaging failed: ${error.message}`);
    process.exitCode = 1;
  });
