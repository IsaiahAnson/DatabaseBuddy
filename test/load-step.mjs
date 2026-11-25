// Loads a Pipedream step file outside Pipedream, for unit tests.
// Package imports are replaced with empty stubs and defineComponent returns its argument,
// so the test gets the plain component object and can call its methods.
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { fileURLToPath, pathToFileURL } from "node:url";
import path from "node:path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const workDir = mkdtempSync(path.join(tmpdir(), "database-buddy-test-"));
process.on("exit", () => rmSync(workDir, { recursive: true, force: true }));

export async function loadStep(fileName) {
  const source = readFileSync(path.join(root, "steps", fileName), "utf8");
  const stubbed = source.replace(
    /^import\s+(.+?)\s+from\s+["'][^"']+["'];?\s*$/gm,
    (_, binding) => `const ${binding} = {};`
  );
  const target = path.join(workDir, fileName.replace(/\.js$/, ".mjs"));
  writeFileSync(target, `const defineComponent = (component) => component;\n${stubbed}`);
  const module = await import(pathToFileURL(target).href);
  return module.default;
}

// Runs a component method with `this` set to the component's methods plus the given props.
export function bind(component, props = {}) {
  return { ...component.methods, ...props };
}
