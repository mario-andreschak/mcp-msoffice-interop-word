import fs from "node:fs";
import path from "node:path";
function roots(): string[] {
  const configured = process.env.WORD_ALLOWED_ROOTS;
  const values: unknown = configured ? JSON.parse(configured) : [process.cwd()];
  if (
    !Array.isArray(values) ||
    !values.length ||
    values.some((v) => typeof v !== "string" || !path.isAbsolute(v))
  )
    throw new Error(
      "WORD_ALLOWED_ROOTS must be a nonempty JSON array of absolute local directories.",
    );
  return values.map((v) => fs.realpathSync(v));
}
function within(candidate: string): boolean {
  return roots().some((root) => {
    const rel = path.relative(root, candidate);
    return (
      rel === "" ||
      (!path.isAbsolute(rel) &&
        rel !== ".." &&
        !rel.startsWith(".." + path.sep))
    );
  });
}
function validate(value: string): void {
  if (
    !path.isAbsolute(value) ||
    value.includes("\0") ||
    value.startsWith("\\\\") ||
    value.startsWith("//") ||
    (process.platform === "win32" && value.slice(2).includes(":"))
  )
    throw new Error(
      "Use an absolute local path within WORD_ALLOWED_ROOTS; network/device paths and alternate streams are unsupported.",
    );
}
export function confineRead(value: string): string {
  validate(value);
  const real = fs.realpathSync(value);
  if (!within(real) || !fs.statSync(real).isFile())
    throw new Error(
      "File is outside WORD_ALLOWED_ROOTS or is not a regular file.",
    );
  return real;
}
export function confineWrite(value: string): string {
  validate(value);
  const parent = fs.realpathSync(path.dirname(value));
  const target = path.join(parent, path.basename(value));
  if (!within(target))
    throw new Error("Destination is outside WORD_ALLOWED_ROOTS.");
  if (fs.existsSync(target))
    throw new Error(
      "Destination already exists. Use saveActiveDocument for an existing active document, or choose a new path.",
    );
  return target;
}
