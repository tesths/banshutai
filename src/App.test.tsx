import { fireEvent, render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { beforeEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import * as workflow from "./lib/pptWorkflow";

vi.mock("./lib/pptWorkflow", () => ({
  analyzeInput: vi.fn(),
  generateAndSavePresentation: vi.fn(),
  isAccentColorValid: (value: string) => /^[0-9A-F]{6}$/.test(value),
  loadDefaultTemplateBytes: vi.fn(),
  normalizeAccentInput: (value: string) => value.trim().replace(/^#/, "").toUpperCase(),
  parseInputLines: (inputText: string) => {
    const lines = inputText.replace(/\r\n/g, "\n").split("\n");
    if (lines[lines.length - 1] === "") {
      lines.pop();
    }
    return lines;
  },
  sanitizeAccentInput: (value: string) =>
    value
      .trim()
      .replace(/^#/, "")
      .toUpperCase()
      .replace(/[^0-9A-F]/g, "")
      .slice(0, 6)
}));

const workflowMock = vi.mocked(workflow);

function makeSummary(pageCount: number) {
  return {
    lineCount: pageCount,
    pageCount,
    didShrink: false,
    didSplit: false,
    splitCount: 0
  };
}

describe("App", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    workflowMock.loadDefaultTemplateBytes.mockResolvedValue(new Uint8Array([1, 2, 3]));
    workflowMock.analyzeInput.mockImplementation(async (_templateBytes, inputText) => {
      const lines = workflow.parseInputLines(inputText);
      return makeSummary(lines.length);
    });
    workflowMock.generateAndSavePresentation.mockResolvedValue({
      outputPath: "/tmp/demo-output.pptx",
      summary: makeSummary(2)
    } as any);
  });

  async function waitForTemplateReady() {
    return screen.findByText(/模板\s*已加载/);
  }

  it("uses download messaging and the template default accent", async () => {
    const user = userEvent.setup();
    render(<App />);

    await waitForTemplateReady();
    expect(workflowMock.loadDefaultTemplateBytes).toHaveBeenCalledTimes(1);
    expect(screen.getByRole("button", { name: "生成并下载" })).toBeInTheDocument();
    expect(screen.getByRole("heading", { name: /板书台，教师黑板贴生成器/ })).toBeInTheDocument();
    expect(screen.queryByText("浏览器版")).not.toBeInTheDocument();
    expect(screen.getByRole("heading", { name: "模板预览" })).toBeInTheDocument();

    await user.type(screen.getByPlaceholderText(/每行一页，直接粘贴/), "第一页\n第二页");

    expect(await screen.findByText("已预估 2 页")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "生成并下载" })).toBeEnabled();

    await user.click(screen.getByRole("button", { name: "自定义" }));
    const accentInput = screen.getByPlaceholderText("FF6600");
    await user.clear(accentInput);
    await user.type(accentInput, "#ff66zz00");

    expect(accentInput).toHaveValue("FF6600");

    await user.click(screen.getByRole("button", { name: "生成并下载" }));

    await waitFor(() => {
      expect(workflowMock.generateAndSavePresentation).toHaveBeenCalledWith(
        expect.objectContaining({
          templateBytes: expect.any(Uint8Array),
          inputText: "第一页\n第二页",
          accentColor: "FF6600",
          defaultFileName: "板书台-导出结果.pptx",
          onStageChange: expect.any(Function)
        })
      );
    });

    expect(await screen.findByText("已生成并下载 demo-output.pptx")).toBeInTheDocument();
    expect(await screen.findByText("最近结果：demo-output.pptx")).toBeInTheDocument();
  });

  it("renders teacher-focused support sections for scanning and search", async () => {
    render(<App />);

    await waitForTemplateReady();

    expect(screen.getByRole("heading", { name: "适合哪些教学场景" })).toBeInTheDocument();
    expect(screen.getByRole("heading", { name: "三步生成课堂黑板贴" })).toBeInTheDocument();
    expect(screen.getByRole("heading", { name: "常见问题" })).toBeInTheDocument();
    expect(screen.getByText("语文课堂标题、词语卡片、板书关键词")).toBeInTheDocument();
    expect(screen.getByText("把每一行文字粘贴进输入框")).toBeInTheDocument();
    expect(screen.getByText("浏览器端直接生成并下载 `.pptx` 文件，不需要安装桌面软件。")).toBeInTheDocument();
  });

  it("surfaces template loading failures before the workspace becomes usable", async () => {
    workflowMock.loadDefaultTemplateBytes.mockRejectedValueOnce(new Error("网络错误"));

    render(<App />);

    await screen.findByRole("heading", { name: "模板加载失败" });
    expect(screen.getByText("网络错误")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "生成并下载" })).toBeDisabled();
  });

  it("renders the template default accent state", async () => {
    render(<App />);

    await waitForTemplateReady();
    expect(screen.getByRole("button", { name: "生成并下载" })).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "打开输出位置" })).not.toBeInTheDocument();
    expect(screen.queryByRole("heading", { name: "摘要" })).not.toBeInTheDocument();

    const accentInput = screen.getByPlaceholderText("FF6600") as HTMLInputElement;
    const colorPicker = screen.getByLabelText("原生取色器") as HTMLInputElement;
    expect(accentInput).toHaveValue("4E95D9");
    expect(colorPicker.value.toLowerCase()).toBe("#4e95d9");
    expect(screen.getByAltText("默认模板效果参考图")).toBeInTheDocument();
    expect(screen.getByText("原版 PPT 参考图")).toBeInTheDocument();
    expect(screen.getByText("只展示模板样式参考，不实时替换输入文字。")).toBeInTheDocument();
  });

  it("keeps preset colors in sync across the preview, hex input, and color picker", async () => {
    const user = userEvent.setup();
    render(<App />);

    await waitForTemplateReady();
    await user.click(screen.getByRole("button", { name: "自定义" }));
    await user.click(screen.getByRole("button", { name: "常用色 #4E95D9" }));

    const hexInput = screen.getByPlaceholderText("FF6600");
    const colorPicker = screen.getByLabelText("原生取色器") as HTMLInputElement;

    expect(hexInput).toHaveValue("4E95D9");
    expect(colorPicker.value.toLowerCase()).toBe("#4e95d9");
    expect(screen.getByText("#4E95D9")).toBeInTheDocument();
  });

  it("syncs the native color picker into the hex input", async () => {
    render(<App />);

    await waitForTemplateReady();
    await userEvent.click(screen.getByRole("button", { name: "自定义" }));

    const colorPicker = screen.getByLabelText("原生取色器") as HTMLInputElement;
    fireEvent.change(colorPicker, { target: { value: "#3b7cff" } });

    expect(screen.getByPlaceholderText("FF6600")).toHaveValue("3B7CFF");
    expect(colorPicker.value.toLowerCase()).toBe("#3b7cff");
    expect(screen.getByText("#3B7CFF")).toBeInTheDocument();
  });

  it("syncs the hex input into the native color picker", async () => {
    const user = userEvent.setup();
    render(<App />);

    await waitForTemplateReady();
    await user.click(screen.getByRole("button", { name: "自定义" }));

    const hexInput = screen.getByPlaceholderText("FF6600");
    await user.clear(hexInput);
    await user.type(hexInput, "1A8CF0");

    const colorPicker = screen.getByLabelText("原生取色器") as HTMLInputElement;
    expect(hexInput).toHaveValue("1A8CF0");
    expect(colorPicker.value.toLowerCase()).toBe("#1a8cf0");
    expect(screen.getByText("#1A8CF0")).toBeInTheDocument();
  });

  it("describes the custom accent as only covering the blue outline layer", async () => {
    render(<App />);

    await waitForTemplateReady();
    await userEvent.click(screen.getByRole("button", { name: "自定义" }));

    expect(screen.getByText("仅覆盖蓝色描边层。")).toBeInTheDocument();
  });

  it("blocks generation while a custom accent is incomplete", async () => {
    const user = userEvent.setup();
    render(<App />);

    await waitForTemplateReady();
    await user.type(screen.getByPlaceholderText(/每行一页，直接粘贴/), "第一页");
    await user.click(screen.getByRole("button", { name: "自定义" }));

    const hexInput = screen.getByPlaceholderText("FF6600");
    await user.clear(hexInput);
    await user.type(hexInput, "12");

    expect(screen.getByText("请输入 6 位十六进制颜色，例如 FF6600")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "生成并下载" })).toBeDisabled();
  });

  it("surfaces generation failures inside the workflow rail", async () => {
    const user = userEvent.setup();
    workflowMock.generateAndSavePresentation.mockRejectedValueOnce(new Error("导出失败"));
    render(<App />);

    await waitForTemplateReady();
    await user.type(screen.getByPlaceholderText(/每行一页，直接粘贴/), "第一页");
    await user.click(screen.getByRole("button", { name: "生成并下载" }));

    expect(await screen.findByRole("heading", { name: "生成失败" })).toBeInTheDocument();
    expect(screen.getByText("导出失败")).toBeInTheDocument();
  });
});
