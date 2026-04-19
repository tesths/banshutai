import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("./runtime", () => ({
  downloadBufferInBrowser: vi.fn(),
  loadBundledTemplateBytes: vi.fn()
}));

vi.mock("../shared/ppt-engine.mjs", () => ({
  analyzeTemplateInput: vi.fn(),
  generatePresentationBuffer: vi.fn(),
  parseInputLines: (inputText: string) => {
    const lines = inputText.replace(/\r\n/g, "\n").split("\n");
    if (lines[lines.length - 1] === "") {
      lines.pop();
    }
    return lines;
  },
  assertAccentColor: vi.fn()
}));

import {
  analyzeInput,
  generateAndSavePresentation,
  isAccentColorValid,
  loadDefaultTemplateBytes,
  normalizeAccentInput,
  parseInputLines,
  sanitizeAccentInput
} from "./pptWorkflow";
import { downloadBufferInBrowser, loadBundledTemplateBytes } from "./runtime";
import * as engine from "../shared/ppt-engine.mjs";

const runtimeMock = vi.mocked({ downloadBufferInBrowser, loadBundledTemplateBytes });
const engineMock = vi.mocked(engine);

describe("pptWorkflow", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    runtimeMock.loadBundledTemplateBytes.mockResolvedValue(new Uint8Array([1, 2, 3]));
    runtimeMock.downloadBufferInBrowser.mockResolvedValue("demo-output.pptx");
    engineMock.analyzeTemplateInput.mockResolvedValue({
      lineCount: 2,
      pageCount: 2,
      didShrink: false,
      didSplit: false,
      splitCount: 0
    });
    engineMock.generatePresentationBuffer.mockResolvedValue({
      outputBuffer: new Uint8Array([7, 8, 9]),
      summary: {
        lineCount: 2,
        pageCount: 2,
        didShrink: false,
        didSplit: false,
        splitCount: 0
      }
    });
  });

  it("normalizes and sanitizes accent values for browser input", () => {
    expect(normalizeAccentInput(" #ff66aa ")).toBe("FF66AA");
    expect(sanitizeAccentInput("#ff66zz00")).toBe("FF6600");
    expect(isAccentColorValid("FF6600")).toBe(true);
    expect(isAccentColorValid("FF66")).toBe(false);
  });

  it("parses multi-line input and removes a trailing empty line", () => {
    expect(parseInputLines("第一页\r\n第二页\r\n")).toEqual(["第一页", "第二页"]);
    expect(parseInputLines("  前后空格  \n")).toEqual(["  前后空格  "]);
  });

  it("loads the bundled template bytes through the runtime adapter", async () => {
    await expect(loadDefaultTemplateBytes("/preview/")).resolves.toEqual(new Uint8Array([1, 2, 3]));
    expect(runtimeMock.loadBundledTemplateBytes).toHaveBeenCalledWith("/preview/");
  });

  it("analyzes input with parsed lines", async () => {
    const templateBytes = new Uint8Array([1, 2, 3]);

    await expect(analyzeInput(templateBytes, "第一页\n第二页")).resolves.toMatchObject({
      pageCount: 2
    });
    expect(engineMock.analyzeTemplateInput).toHaveBeenCalledWith({
      templateBytes,
      lines: ["第一页", "第二页"]
    });
  });

  it("generates and downloads through the workflow pipeline", async () => {
    const templateBytes = new Uint8Array([1, 2, 3]);
    const stages: Array<"generating" | "downloading"> = [];

    const result = await generateAndSavePresentation({
      templateBytes,
      inputText: "第一页\n第二页",
      accentColor: "FF6600",
      defaultFileName: "demo-output.pptx",
      onStageChange: (stage) => {
        stages.push(stage);
      }
    });

    expect(engineMock.assertAccentColor).toHaveBeenCalledWith("FF6600");
    expect(engineMock.generatePresentationBuffer).toHaveBeenCalledWith({
      templateBytes,
      lines: ["第一页", "第二页"],
      accentColor: "FF6600"
    });
    expect(runtimeMock.downloadBufferInBrowser).toHaveBeenCalledWith(
      new Uint8Array([7, 8, 9]),
      "demo-output.pptx"
    );
    expect(stages).toEqual(["generating", "downloading"]);
    expect(result.outputPath).toBe("demo-output.pptx");
  });

  it("rejects empty input before generating", async () => {
    await expect(
      generateAndSavePresentation({
        templateBytes: new Uint8Array([1, 2, 3]),
        inputText: "",
        accentColor: null
      })
    ).rejects.toThrow("请先粘贴至少一行内容");

    expect(engineMock.generatePresentationBuffer).not.toHaveBeenCalled();
    expect(runtimeMock.downloadBufferInBrowser).not.toHaveBeenCalled();
  });
});
