import { fork, type ChildProcess } from "node:child_process";
import { AsyncLocalStorage } from "node:async_hooks";
import { fileURLToPath } from "node:url";
import { methods, type WordMethod } from "./methods.js";
export const callSignal = new AsyncLocalStorage<AbortSignal | undefined>();
interface Job {
  id: number;
  method: WordMethod;
  args: unknown[];
  resolve(value: any): void;
  reject(error: Error): void;
  timer: ReturnType<typeof setTimeout>;
  signal?: AbortSignal;
  abort(): void;
}
export class WordController {
  #child?: ChildProcess;
  #active?: Job;
  #queue: Job[] = [];
  #next = 0;
  #stopped = false;
  constructor(
    private readonly options: { worker?: string; timeoutMs?: number } = {},
  ) {}
  invoke(
    method: WordMethod,
    args: unknown[],
    signal = callSignal.getStore(),
  ): Promise<any> {
    if (this.#stopped)
      return Promise.reject(
        new Error(
          "Word worker stopped; restart the server before further document operations.",
        ),
      );
    if (!methods.includes(method))
      return Promise.reject(new Error("Unknown Word operation."));
    if (Buffer.byteLength(JSON.stringify(args)) > 1024 * 1024)
      return Promise.reject(new Error("Word arguments exceed the 1MiB limit."));
    if (this.#queue.length >= 32)
      return Promise.reject(new Error("Word operation queue is full."));
    signal?.throwIfAborted();
    return new Promise((resolve, reject) => {
      const job: Job = {
        id: ++this.#next,
        method,
        args,
        resolve,
        reject,
        timer: setTimeout(
          () =>
            this.#cancel(
              job,
              "Word operation deadline exceeded; document outcome may be uncertain.",
            ),
          this.options.timeoutMs ?? 30000,
        ),
        signal,
        abort: () =>
          this.#cancel(
            job,
            "Word operation cancelled; an active document operation may have completed.",
          ),
      };
      signal?.addEventListener("abort", job.abort, { once: true });
      this.#queue.push(job);
      this.#pump();
    });
  }
  #settle(job: Job, error?: Error, result?: unknown) {
    clearTimeout(job.timer);
    job.signal?.removeEventListener("abort", job.abort);
    if (error) job.reject(error);
    else job.resolve(result);
  }
  #cancel(job: Job, message: string) {
    if (this.#active === job) {
      this.#fail(new Error(message));
      return;
    }
    const index = this.#queue.indexOf(job);
    if (index >= 0) {
      this.#queue.splice(index, 1);
      this.#settle(job, new Error(message));
    }
  }
  #fail(error: Error) {
    this.#stopped = true;
    this.#child?.kill();
    if (this.#active) this.#settle(this.#active, error);
    this.#active = undefined;
    for (const job of this.#queue) this.#settle(job, error);
    this.#queue = [];
  }
  #pump() {
    if (this.#active || this.#stopped) return;
    const job = this.#queue.shift();
    if (!job) return;
    this.#active = job;
    if (!this.#child) {
      const child = fork(
        this.options.worker ??
          fileURLToPath(new URL("./worker.js", import.meta.url)),
        [],
        {
          stdio: ["ignore", "ignore", "ignore", "ipc"],
          execArgv: ["--max-old-space-size=128"],
        },
      );
      this.#child = child;
      child.on("error", () =>
        this.#fail(new Error("Unable to start the Word worker.")),
      );
      child.on("exit", () => {
        if (!this.#stopped)
          this.#fail(
            new Error(
              "Word worker exited unexpectedly; document outcome may be uncertain.",
            ),
          );
      });
      child.on("message", (message: any) => {
        const active = this.#active;
        if (!active || message?.id !== active.id) return;
        this.#active = undefined;
        this.#settle(
          active,
          message.error
            ? new Error(String(message.error).slice(0, 4096))
            : undefined,
          message.result,
        );
        this.#pump();
      });
    }
    this.#child.send(
      { id: job.id, method: job.method, args: job.args },
      (error) => {
        if (error) this.#fail(new Error("Unable to send Word operation."));
      },
    );
  }
  async close(): Promise<void> {
    const child = this.#child;
    this.#fail(new Error("Word server is shutting down."));
    if (!child || child.exitCode !== null || child.signalCode !== null) return;
    await new Promise<void>((resolve) => {
      const timer = setTimeout(() => {
        child.kill("SIGKILL");
        resolve();
      }, 2000);
      child.once("exit", () => {
        clearTimeout(timer);
        resolve();
      });
    });
  }
}
export const controller = new WordController();
export const wordService = new Proxy(
  {} as Record<WordMethod, (...args: any[]) => Promise<any>>,
  {
    get: (_target, key) =>
      typeof key === "string" && methods.includes(key as WordMethod)
        ? (...args: unknown[]) => controller.invoke(key as WordMethod, args)
        : undefined,
  },
);
