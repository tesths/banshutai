import * as engineRaw from "../shared/ppt-engine.mjs";
import { downloadBufferInBrowser, loadBundledTemplateBytes } from "./runtime";

export type PptSummary = {
  lineCount: number;
  pageCount: number;
  didShrink: boolean;
  didSplit: boolean;
  splitCount: number;
};

export type GenerateStage = "generating" | "downloading";

export type GenerateResult = {
  outputPath: string;
  summary: PptSummary;
};

type EngineModule = {
  analyzeTemplateInput(args: {
    templateBytes: Uint8Array;
    lines: string[];
  }): Promise<PptSummary>;
  generatePresentationBuffer(args: {
    templateBytes: Uint8Array;
    lines: string[];
    accentColor?: string | null;
  }): Promise<{
    outputBuffer: Uint8Array;
    summary: PptSummary;
  }>;
  parseInputLines(inputText: string): string[];
  assertAccentColor(accentColor: string): void;
};

const engine = engineRaw as unknown as EngineModule;

export function parseInputLines(inputText: string): string[] {
  return engine.parseInputLines(inputText);
}

export function normalizeAccentInput(value: string): string {
  return value.trim().replace(/^#/, "").toUpperCase();
}

export function sanitizeAccentInput(value: string): string {
  return normalizeAccentInput(value).replace(/[^0-9A-F]/g, "").slice(0, 6);
}

export function isAccentColorValid(value: string): boolean {
  return /^[0-9A-F]{6}$/.test(value);
}

export async function loadDefaultTemplateBytes(
  baseUrl = import.meta.env.BASE_URL || "/"
): Promise<Uint8Array> {
  return loadBundledTemplateBytes(baseUrl);
}

export async function analyzeInput(
  templateBytes: Uint8Array,
  inputText: string
): Promise<PptSummary> {
  return engine.analyzeTemplateInput({
    templateBytes,
    lines: parseInputLines(inputText)
  });
}

export async function generateAndSavePresentation({
  templateBytes,
  inputText,
  accentColor,
  defaultFileName = "板书台-导出结果.pptx",
  onStageChange
}: {
  templateBytes: Uint8Array;
  inputText: string;
  accentColor: string | null;
  defaultFileName?: string;
  onStageChange?: (stage: GenerateStage) => void;
}): Promise<GenerateResult> {
  const lines = parseInputLines(inputText);
  if (lines.length === 0) {
    throw new Error("请先粘贴至少一行内容");
  }

  if (accentColor) {
    engine.assertAccentColor(accentColor);
  }

  onStageChange?.("generating");
  const { outputBuffer, summary } = await engine.generatePresentationBuffer({
    templateBytes,
    lines,
    accentColor
  });

  onStageChange?.("downloading");
  const outputPath = await downloadBufferInBrowser(outputBuffer, defaultFileName);

  return {
    outputPath,
    summary
  };
}
