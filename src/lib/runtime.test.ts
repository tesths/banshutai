import { afterEach, describe, expect, it, vi } from "vitest";

import {
  downloadBufferInBrowser,
  fetchBundledTemplateInBrowser,
  getBundledTemplateUrl
} from "./runtime";

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe("runtime", () => {
  it("resolves the bundled template URL from the configured Vite base", () => {
    expect(getBundledTemplateUrl("./")).toBe("./default-template.pptx");
    expect(getBundledTemplateUrl("/docs/")).toBe("/docs/default-template.pptx");
    expect(getBundledTemplateUrl("/docs")).toBe("/docs/default-template.pptx");
  });

  it("fetches the bundled template from a Vite base-aware path", async () => {
    const bytes = new Uint8Array([1, 2, 3]);
    const fetchMock = vi.fn().mockResolvedValue({
      ok: true,
      arrayBuffer: async () => bytes.buffer
    });
    vi.stubGlobal("fetch", fetchMock);

    await expect(fetchBundledTemplateInBrowser("/docs/")).resolves.toEqual(bytes);
    expect(fetchMock).toHaveBeenCalledWith("/docs/default-template.pptx");
  });

  it("throws when the browser cannot load the bundled template", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue({
        ok: false
      })
    );

    await expect(fetchBundledTemplateInBrowser("./")).rejects.toThrow("默认模板加载失败");
  });

  it("creates a Blob download for generated pptx output", async () => {
    const click = vi.fn();
    const anchor = { click, href: "", download: "" } as unknown as HTMLAnchorElement;
    vi.spyOn(document, "createElement").mockReturnValue(anchor);
    const createObjectUrl = vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:test");
    const revokeObjectUrl = vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => {});

    await expect(downloadBufferInBrowser(new Uint8Array([9]), "demo.pptx")).resolves.toBe(
      "demo.pptx"
    );

    expect(anchor.download).toBe("demo.pptx");
    expect(click).toHaveBeenCalled();
    expect(createObjectUrl.mock.calls[0]?.[0]).toBeInstanceOf(Blob);
    expect((createObjectUrl.mock.calls[0]?.[0] as Blob).type).toBe(
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    expect(revokeObjectUrl).toHaveBeenCalledWith("blob:test");
  });
});
