import { WordService } from "./word-service.js";
import { methods, type WordMethod } from "./methods.js";
const service = new WordService();
const jsonResults = new Set([
  "getSelectionText",
  "getSelectionInfo",
  "findAndReplace",
]);
process.on(
  "message",
  (job: { id: number; method: WordMethod; args: unknown[] }) => {
    if (
      !Number.isSafeInteger(job?.id) ||
      !methods.includes(job.method) ||
      !Array.isArray(job.args)
    )
      return;
    try {
      // Never await or serialize native COM proxies: their property traps are not Promise/JSON objects.
      const result = (service[job.method] as (...args: any[]) => unknown).apply(
        service,
        job.args,
      );
      const value = jsonResults.has(job.method) ? result : null;
      if (Buffer.byteLength(JSON.stringify(value)) > 1024 * 1024)
        throw new Error("Word result exceeds the 1MiB limit.");
      process.send?.({ id: job.id, result: value });
    } catch (error) {
      process.send?.({
        id: job.id,
        error:
          error instanceof Error
            ? error.message.slice(0, 4096)
            : "Word operation failed.",
      });
    }
  },
);
process.on("disconnect", () => process.exit(0));
