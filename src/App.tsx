import { useEffect, useRef, useState, type CSSProperties, type ChangeEvent } from "react";
import {
  analyzeInput,
  generateAndSavePresentation,
  isAccentColorValid,
  loadDefaultTemplateBytes,
  normalizeAccentInput,
  parseInputLines,
  sanitizeAccentInput,
  type GenerateStage,
  type PptSummary
} from "./lib/pptWorkflow";

const DEFAULT_ACCENT_HEX = "4E95D9";
const DEFAULT_SHELL_ACCENT = `#${DEFAULT_ACCENT_HEX}`;
const DEFAULT_OUTPUT_NAME = "板书台-导出结果.pptx";
const ANALYSIS_DELAY_MS = 180;
const CUSTOM_ACCENT_COPY = "仅覆盖蓝色描边层。";
const TEMPLATE_PREVIEW_SRC = `${(import.meta.env.BASE_URL || "/").replace(/\/?$/, "/")}default-template-preview.png`;
const ACCENT_PRESETS = [
  { name: "蓝", hex: "4E95D9" },
  { name: "橙", hex: "D86A36" },
  { name: "青", hex: "0F9ED5" },
  { name: "绿", hex: "4F7A55" },
  { name: "红", hex: "B24B3F" },
  { name: "紫", hex: "8A3FFC" }
] as const;
const HERO_PILLS = ["浏览器端生成", "课堂黑板贴排版", "PPTX 直接下载"] as const;
const SCENARIO_TAGS = ["语文课堂标题、词语卡片、板书关键词", "英语短句、班会口号、小组展示标题", "备课时批量整理一行一页的课堂用字"] as const;
const QUICK_STEPS = [
  "把每一行文字粘贴进输入框",
  "确认页数预估，按需调整描边颜色",
  "点击生成并下载，直接拿到课堂可用的 PPTX 文件"
] as const;
const FAQ_ITEMS = [
  {
    question: "板书台适合做什么？",
    answer: "适合教师快速制作课堂标题、黑板贴、词语卡片和板书关键词，不需要从空白 PPT 反复复制页面。"
  },
  {
    question: "需要安装软件吗？",
    answer: "浏览器端直接生成并下载 `.pptx` 文件，不需要安装桌面软件。"
  },
  {
    question: "可以改颜色吗？",
    answer: "可以。默认保留模板原始描边，也可以切换到自定义颜色，只覆盖蓝色描边层。"
  }
] as const;

const EMPTY_SUMMARY: PptSummary = {
  lineCount: 0,
  pageCount: 0,
  didShrink: false,
  didSplit: false,
  splitCount: 0
};

function formatLastSaved(path: string | null): string {
  if (!path) {
    return "尚未生成";
  }

  return path.split(/[/\\]/).pop() || path;
}

function readErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }

  return "发生未知错误";
}

export default function App() {
  const [inputText, setInputText] = useState("");
  const [templateBytes, setTemplateBytes] = useState<Uint8Array | null>(null);
  const [templateError, setTemplateError] = useState<string | null>(null);
  const [isLoadingTemplate, setIsLoadingTemplate] = useState(true);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisError, setAnalysisError] = useState<string | null>(null);
  const [summary, setSummary] = useState<PptSummary>(EMPTY_SUMMARY);
  const [accentMode, setAccentMode] = useState<"default" | "custom">("default");
  const [accentInput, setAccentInput] = useState(DEFAULT_ACCENT_HEX);
  const [isGenerating, setIsGenerating] = useState(false);
  const [generateError, setGenerateError] = useState<string | null>(null);
  const [savedPath, setSavedPath] = useState<string | null>(null);
  const [statusMessage, setStatusMessage] = useState("正在加载默认模板");

  const templateLoadToken = useRef(0);
  const analysisToken = useRef(0);

  const inputLines = parseInputLines(inputText);
  const normalizedAccent = normalizeAccentInput(accentInput);
  const hasValidCustomAccent =
    accentMode !== "custom" || (normalizedAccent.length === 6 && isAccentColorValid(normalizedAccent));
  const accentForEngine = accentMode === "custom" && hasValidCustomAccent ? normalizedAccent : null;
  const shellAccent = accentMode === "custom" && hasValidCustomAccent ? `#${normalizedAccent}` : DEFAULT_SHELL_ACCENT;
  const accentModeDisabled = accentMode !== "custom" || isGenerating;
  const primaryActionBusyLabel = isGenerating ? "正在生成并下载…" : "生成并下载";
  const canGenerate =
    Boolean(templateBytes) &&
    inputLines.length > 0 &&
    !isLoadingTemplate &&
    !isGenerating &&
    hasValidCustomAccent;
  const accentError =
    accentMode === "custom" && !hasValidCustomAccent ? "请输入 6 位十六进制颜色，例如 FF6600" : null;
  const templateStateLabel = templateError
    ? "加载失败"
    : isLoadingTemplate
      ? "加载中"
      : templateBytes
        ? "已加载"
        : "未就绪";
  const workflowError = templateError || analysisError || generateError;
  const workflowHint =
    workflowError ||
    (inputLines.length > 0
      ? `已识别 ${inputLines.length} 行文本，预计输出 ${summary.pageCount} 页。`
      : templateBytes
        ? "一行一页，粘贴后会自动预估并直接下载。"
        : "默认模板正在加载。");
  const statusTone = workflowError ? "error" : isLoadingTemplate || isAnalyzing || isGenerating ? "busy" : "ready";

  useEffect(() => {
    void reloadTemplate();
  }, []);

  useEffect(() => {
    if (!templateBytes) {
      return;
    }

    const token = ++analysisToken.current;
    setIsAnalyzing(true);
    setAnalysisError(null);

    const handle = window.setTimeout(() => {
      void analyzeInput(templateBytes, inputText)
        .then((nextSummary) => {
          if (analysisToken.current !== token) {
            return;
          }

          setSummary(nextSummary);
          setStatusMessage(`已预估 ${nextSummary.pageCount} 页`);
        })
        .catch((error) => {
          if (analysisToken.current !== token) {
            return;
          }

          setAnalysisError(readErrorMessage(error));
          setStatusMessage("预估失败");
        })
        .finally(() => {
          if (analysisToken.current === token) {
            setIsAnalyzing(false);
          }
        });
    }, ANALYSIS_DELAY_MS);

    return () => {
      window.clearTimeout(handle);
    };
  }, [templateBytes, inputText]);

  async function reloadTemplate() {
    const token = ++templateLoadToken.current;
    setIsLoadingTemplate(true);
    setTemplateError(null);
    setStatusMessage("正在加载默认模板");

    try {
      const bytes = await loadDefaultTemplateBytes();
      if (templateLoadToken.current !== token) {
        return;
      }

      setTemplateBytes(bytes);
      setStatusMessage("默认模板已就绪");
    } catch (error) {
      if (templateLoadToken.current !== token) {
        return;
      }

      setTemplateBytes(null);
      setTemplateError(readErrorMessage(error));
      setStatusMessage("模板加载失败");
    } finally {
      if (templateLoadToken.current === token) {
        setIsLoadingTemplate(false);
      }
    }
  }

  function handleAccentChange(event: ChangeEvent<HTMLInputElement>) {
    setAccentMode("custom");
    setAccentInput(sanitizeAccentInput(event.target.value));
    setGenerateError(null);
  }

  function handleAccentPickerChange(event: ChangeEvent<HTMLInputElement>) {
    setAccentMode("custom");
    setAccentInput(sanitizeAccentInput(event.target.value));
    setGenerateError(null);
  }

  function handleAccentPresetSelect(hex: string) {
    setAccentMode("custom");
    setAccentInput(hex);
    setGenerateError(null);
  }

  async function handleGenerate() {
    if (!templateBytes) {
      setGenerateError("默认模板还没有加载完成");
      return;
    }

    if (inputLines.length === 0) {
      setGenerateError("请先粘贴至少一行内容");
      return;
    }

    if (!hasValidCustomAccent) {
      setGenerateError("自定义颜色还未填写完整");
      return;
    }

    setGenerateError(null);
    setIsGenerating(true);
    setStatusMessage("正在生成并下载");

    try {
      const result = await generateAndSavePresentation({
        templateBytes,
        inputText,
        accentColor: accentForEngine,
        defaultFileName: DEFAULT_OUTPUT_NAME,
        onStageChange: (stage: GenerateStage) => {
          if (stage === "generating") {
            setIsGenerating(true);
            setStatusMessage("正在生成 PPTX");
          }

          if (stage === "downloading") {
            setIsGenerating(true);
            setStatusMessage("正在下载文件");
          }
        }
      });

      setSummary(result.summary);
      setSavedPath(result.outputPath);
      setStatusMessage(`已生成并下载 ${formatLastSaved(result.outputPath)}`);
    } catch (error) {
      setGenerateError(readErrorMessage(error));
      setStatusMessage("生成失败");
    } finally {
      setIsGenerating(false);
    }
  }

  const appStyle = {
    "--accent": shellAccent
  } as CSSProperties;

  return (
    <div className="app-shell" style={appStyle}>
      <main className="app-frame">
        <header className="shell-head">
          <div className="shell-brand">
            <p className="shell-eyebrow">教师板书工具</p>
            <h1>板书台，教师黑板贴生成器</h1>
            <p className="shell-copy">
              把课堂标题、板书关键词和词语卡片按一行一页排进模板，浏览器端直接生成并下载 PPTX 文件，适合语文、英语、班会等教学场景。
            </p>
          </div>

          <div className="hero-pills" aria-label="工具特点">
            {HERO_PILLS.map((pill) => (
              <span key={pill} className="hero-pill">
                {pill}
              </span>
            ))}
          </div>
        </header>

        <section className="status-strip" aria-live="polite">
          <div className="status-strip-copy">
            <span className="status-dot" data-state={statusTone} />
            <div>
              <h2 className="status-title">{statusMessage}</h2>
              <p className="status-subtitle">{workflowHint}</p>
            </div>
          </div>

          <div className="status-strip-side">
            <div className="status-meta" aria-label="当前状态">
              <span className="status-pill">模板 {templateStateLabel}</span>
              <span className="status-pill">结果 {formatLastSaved(savedPath)}</span>
            </div>

            <button
              className="ghost-button"
              type="button"
              onClick={() => {
                void reloadTemplate();
              }}
              disabled={isLoadingTemplate || isGenerating}
            >
              {isLoadingTemplate ? "重新加载中…" : "重新加载模板"}
            </button>
          </div>
        </section>

        <section className="workspace">
          <section className="composer panel" aria-labelledby="input-panel-title">
            <div className="panel-head">
              <div>
                <p className="panel-kicker">输入内容</p>
                <h2 id="input-panel-title">粘贴文本</h2>
              </div>
            </div>

            <label className="paste-box">
              <span className="sr-only">多行文本输入</span>
              <textarea
                value={inputText}
                onChange={(event) => {
                  setInputText(event.target.value);
                  setGenerateError(null);
                }}
                placeholder={`每行一页，直接粘贴。\n\n例如：\n借助注释\n分析人物\n理解情节`}
                spellCheck={false}
                autoCapitalize="off"
                autoComplete="off"
                autoCorrect="off"
              />
            </label>

            <p className="composer-note">按行生成，不会改动原文，浏览器会直接下载 `.pptx` 结果。</p>

            <section className="actions-panel" aria-labelledby="download-panel-title">
              <div className="panel-head">
                <div>
                  <p className="panel-kicker">生成</p>
                  <h2 id="download-panel-title">下载输出</h2>
                </div>
              </div>

              <div className="actions">
                <button className="primary-button" type="button" onClick={handleGenerate} disabled={!canGenerate}>
                  {primaryActionBusyLabel}
                </button>
              </div>

              <p className="field-help">当前 {inputLines.length} 行，预计生成 {summary.pageCount} 页。</p>
              <p className="output-path">最近结果：{formatLastSaved(savedPath)}</p>
            </section>
          </section>

          <aside className="rail">
            <section className="panel preview-panel" aria-labelledby="preview-panel-title">
              <div className="panel-head">
                <div>
                  <p className="panel-kicker">最终效果</p>
                  <h2 id="preview-panel-title">模板预览</h2>
                </div>
                <p className="preview-meta">原版 PPT 参考图</p>
              </div>

              <figure className="preview-figure">
                <img className="preview-image" src={TEMPLATE_PREVIEW_SRC} alt="默认模板效果参考图" />
                <figcaption className="preview-caption">
                  只展示模板样式参考，不实时替换输入文字。
                </figcaption>
              </figure>
            </section>

            <section className="panel" aria-labelledby="accent-panel-title">
              <div className="panel-head">
                <div>
                  <p className="panel-kicker">描边颜色</p>
                  <h2 id="accent-panel-title">描边颜色</h2>
                </div>
              </div>

              <div className="segmented" aria-label="Accent 模式">
                <button
                  type="button"
                  className={accentMode === "default" ? "segment active" : "segment"}
                  onClick={() => {
                    setAccentMode("default");
                    setGenerateError(null);
                  }}
                >
                  默认
                </button>
                <button
                  type="button"
                  className={accentMode === "custom" ? "segment active" : "segment"}
                  onClick={() => {
                    setAccentMode("custom");
                    setGenerateError(null);
                  }}
                >
                  自定义
                </button>
              </div>

              <div className="accent-preview">
                <span className="accent-swatch" style={{ backgroundColor: shellAccent }} />
                <div>
                  <strong>
                    {accentMode === "custom" && hasValidCustomAccent ? `#${normalizedAccent}` : "模板默认蓝色"}
                  </strong>
                  <p>{accentMode === "custom" ? CUSTOM_ACCENT_COPY : "保留模板原始描边。"}</p>
                </div>
              </div>

              <div className="accent-tools">
                <label className="native-picker">
                  <span className="native-picker-label">原生取色器</span>
                  <input
                    type="color"
                    value={shellAccent}
                    onChange={handleAccentPickerChange}
                    disabled={accentModeDisabled}
                  />
                </label>

                <div className="preset-group">
                  <p className="preset-label">常用色预设</p>
                  <div className="preset-list">
                    {ACCENT_PRESETS.map((preset) => {
                      const isActive = accentMode === "custom" && normalizedAccent === preset.hex;

                      return (
                        <button
                          key={preset.hex}
                          type="button"
                          className={isActive ? "preset-button active" : "preset-button"}
                          onClick={() => {
                            handleAccentPresetSelect(preset.hex);
                          }}
                          disabled={accentModeDisabled}
                          aria-label={`常用色 #${preset.hex}`}
                        >
                          <span className="preset-swatch" style={{ backgroundColor: `#${preset.hex}` }} />
                          <span className="preset-name">{preset.name}</span>
                        </button>
                      );
                    })}
                  </div>
                </div>
              </div>

              <label className="hex-field" data-invalid={Boolean(accentError)}>
                <span className="hex-prefix">#</span>
                <input
                  type="text"
                  inputMode="text"
                  maxLength={6}
                  value={accentInput}
                  onChange={handleAccentChange}
                  disabled={accentMode !== "custom" || isGenerating}
                  placeholder="FF6600"
                  aria-invalid={Boolean(accentError)}
                />
              </label>

              {accentError ? (
                <p className="field-error">{accentError}</p>
              ) : (
                <p className="field-help">输入 6 位十六进制颜色。</p>
              )}
            </section>
          </aside>
        </section>

        <section className="support-grid" aria-label="工具说明">
          <article className="support-card" aria-labelledby="scenario-title">
            <p className="panel-kicker">适用场景</p>
            <h2 id="scenario-title">适合哪些教学场景</h2>
            <p className="support-copy">
              这是一款教师板书工具，适合把短句、标题、词语和课堂重点快速整理成一页一条的黑板贴或投屏内容。
            </p>
            <ul className="support-list">
              {SCENARIO_TAGS.map((item) => (
                <li key={item}>{item}</li>
              ))}
            </ul>
          </article>

          <article className="support-card" aria-labelledby="steps-title">
            <p className="panel-kicker">使用方式</p>
            <h2 id="steps-title">三步生成课堂黑板贴</h2>
            <ol className="support-steps">
              {QUICK_STEPS.map((step) => (
                <li key={step}>{step}</li>
              ))}
            </ol>
          </article>

          <article className="support-card" aria-labelledby="faq-title">
            <p className="panel-kicker">常见问题</p>
            <h2 id="faq-title">常见问题</h2>
            <dl className="faq-list">
              {FAQ_ITEMS.map((item) => (
                <div key={item.question} className="faq-item">
                  <dt>{item.question}</dt>
                  <dd>{item.answer}</dd>
                </div>
              ))}
            </dl>
          </article>
        </section>

        <footer className="site-foot">
          <p>面向教师的黑板贴生成器，同时覆盖文本转 PPTX、课堂板书整理和批量页面生成场景。</p>
        </footer>
      </main>
    </div>
  );
}
