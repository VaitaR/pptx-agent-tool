import fs from "node:fs";
import path from "node:path";
import { defaultSpecPath, validateDeckSpec } from "./deck-spec.js";
import { renderDeck } from "./render-pptx.js";
import { VerifyExecutionError } from "./verify-pptx.js";
import { verifyDeck } from "./verify-pptx.js";

type CommandName = "validate-spec" | "render" | "verify" | "check";

type ParsedArgs = {
  command: CommandName;
  flags: Record<string, string | boolean>;
};

const EXIT_CODES = {
  ok: 0,
  invalidSpec: 2,
  environment: 3,
  render: 4,
  verify: 5,
  usage: 64
} as const;

type InvalidSpecPayload = {
  command: CommandName;
  specPath: string;
  valid: false;
  errors: string[];
  warnings: string[];
};

type RuntimeErrorPayload = {
  command: CommandName | "unknown";
  specPath?: string;
  ok: false;
  exitCode: number;
  error: {
    kind: "environment" | "render" | "verify" | "usage";
    message: string;
  };
};

function parseArgs(argv: string[]): ParsedArgs {
  const [commandRaw, ...rest] = argv;
  if (!commandRaw || !["validate-spec", "render", "verify", "check"].includes(commandRaw)) {
    throw new Error("Usage: deckgen <validate-spec|render|verify|check> [--spec path] [--out path] [--pptx path] [--artifacts-dir path] [--report-json path] [--json]");
  }

  const flags: Record<string, string | boolean> = {};
  for (let index = 0; index < rest.length; index += 1) {
    const token = rest[index];
    if (!token.startsWith("--")) {
      throw new Error(`Unexpected argument: ${token}`);
    }
    const key = token.slice(2);
    const next = rest[index + 1];
    if (!next || next.startsWith("--")) {
      flags[key] = true;
      continue;
    }
    flags[key] = next;
    index += 1;
  }

  return { command: commandRaw as CommandName, flags };
}

function tryParseArgs(argv: string[]): ParsedArgs | null {
  try {
    return parseArgs(argv);
  } catch {
    return null;
  }
}

function flagString(flags: Record<string, string | boolean>, key: string): string | undefined {
  const value = flags[key];
  return typeof value === "string" ? value : undefined;
}

function flagBoolean(flags: Record<string, string | boolean>, key: string): boolean {
  return flags[key] === true;
}

function writeJsonReport(reportPath: string | undefined, payload: unknown): void {
  if (!reportPath) return;
  const absolutePath = path.resolve(process.cwd(), reportPath);
  fs.mkdirSync(path.dirname(absolutePath), { recursive: true });
  fs.writeFileSync(absolutePath, `${JSON.stringify(payload, null, 2)}\n`);
}

function printJson(payload: unknown): void {
  console.log(JSON.stringify(payload, null, 2));
}

function readRawSpec(specPath: string): unknown {
  return JSON.parse(fs.readFileSync(specPath, "utf8")) as unknown;
}

function exitInvalidSpec(
  command: CommandName,
  specPath: string,
  reportPath: string | undefined,
  jsonStdout: boolean,
  errors: string[]
): never {
  const payload: InvalidSpecPayload = {
    command,
    specPath,
    valid: false,
    errors,
    warnings: []
  };
  writeJsonReport(reportPath, payload);
  if (jsonStdout) printJson(payload);
  console.error(`Invalid deck spec:\n- ${errors.join("\n- ")}`);
  process.exit(EXIT_CODES.invalidSpec);
}

async function run(): Promise<void> {
  const { command, flags } = parseArgs(process.argv.slice(2));
  const specPath = path.resolve(process.cwd(), flagString(flags, "spec") ?? defaultSpecPath());
  const reportPath = flagString(flags, "report-json");
  const jsonStdout = flagBoolean(flags, "json");
  let rawSpec: unknown;

  try {
    rawSpec = readRawSpec(specPath);
  } catch (error) {
    exitInvalidSpec(command, specPath, reportPath, jsonStdout, [
      `Unable to read or parse spec JSON: ${error instanceof Error ? error.message : String(error)}`
    ]);
  }

  if (command === "validate-spec") {
    const report = validateDeckSpec(rawSpec, { baseDir: path.dirname(specPath) });
    const payload = {
      command,
      specPath,
      ...report
    };
    writeJsonReport(reportPath, payload);
    if (jsonStdout) printJson(payload);
    if (!report.valid) {
      console.error(`Invalid deck spec:\n- ${report.errors.join("\n- ")}`);
      process.exit(EXIT_CODES.invalidSpec);
    }
    if (!jsonStdout) {
      console.log(`Spec is valid: ${path.relative(process.cwd(), specPath)}`);
    }
    return;
  }

  const specReport = validateDeckSpec(rawSpec, { baseDir: path.dirname(specPath) });
  if (!specReport.valid) {
    const payload = {
      command,
      specPath,
      ...specReport
    };
    writeJsonReport(reportPath, payload);
    if (jsonStdout) printJson(payload);
    console.error(`Invalid deck spec:\n- ${specReport.errors.join("\n- ")}`);
    process.exit(EXIT_CODES.invalidSpec);
  }

  if (command === "render") {
    const payload = await renderDeck({
      specPath,
      outputPath: flagString(flags, "out"),
      artifactsDir: flagString(flags, "artifacts-dir"),
      quiet: jsonStdout
    });
    writeJsonReport(reportPath, payload);
    if (jsonStdout) printJson(payload);
    return;
  }

  if (command === "verify") {
    const payload = verifyDeck({
      specPath,
      pptxPath: flagString(flags, "pptx"),
      artifactsDir: flagString(flags, "artifacts-dir"),
      verbose: !jsonStdout
    });
    writeJsonReport(reportPath, payload);
    if (jsonStdout) printJson(payload);
    if (!payload.passed) {
      process.exit(EXIT_CODES.verify);
    }
    return;
  }

  const renderReport = await renderDeck({
    specPath,
    outputPath: flagString(flags, "out"),
    artifactsDir: flagString(flags, "artifacts-dir"),
    quiet: jsonStdout
  });
  const verifyReport = verifyDeck({
    specPath,
    pptxPath: renderReport.pptxPath,
    artifactsDir: renderReport.artifactsDir,
    verbose: !jsonStdout
  });
  const payload = {
    command,
    specPath,
    render: renderReport,
    verify: verifyReport
  };
  writeJsonReport(reportPath, payload);
  if (jsonStdout) printJson(payload);
  if (!verifyReport.passed) {
    process.exit(EXIT_CODES.verify);
  }
}

run().catch((error) => {
  const parsed = tryParseArgs(process.argv.slice(2));
  const reportPath = parsed ? flagString(parsed.flags, "report-json") : undefined;
  const jsonStdout = parsed ? flagBoolean(parsed.flags, "json") : false;
  const specPath = parsed ? path.resolve(process.cwd(), flagString(parsed.flags, "spec") ?? defaultSpecPath()) : undefined;

  let exitCode: number = EXIT_CODES.render;
  let kind: RuntimeErrorPayload["error"]["kind"] = "render";
  if (error instanceof VerifyExecutionError && error.kind === "environment") {
    exitCode = EXIT_CODES.environment;
    kind = "environment";
  }
  if (!parsed) {
    exitCode = EXIT_CODES.usage;
    kind = "usage";
  }

  const payload: RuntimeErrorPayload = {
    command: parsed?.command ?? "unknown",
    specPath,
    ok: false,
    exitCode,
    error: {
      kind,
      message: error instanceof Error ? error.message : String(error)
    }
  };

  writeJsonReport(reportPath, payload);
  if (jsonStdout) {
    printJson(payload);
  }
  console.error(payload.error.message);
  process.exit(exitCode);
});