export function getBundledTemplateUrl(baseUrl = import.meta.env.BASE_URL || "/"): string {
  const normalizedBaseUrl = baseUrl.endsWith("/") ? baseUrl : `${baseUrl}/`;
  return `${normalizedBaseUrl}default-template.pptx`;
}

export async function fetchBundledTemplateInBrowser(
  baseUrl = import.meta.env.BASE_URL || "/"
): Promise<Uint8Array> {
  const response = await fetch(getBundledTemplateUrl(baseUrl));
  if (!response.ok) {
    throw new Error("默认模板加载失败");
  }

  return new Uint8Array(await response.arrayBuffer());
}

export async function loadBundledTemplateBytes(
  baseUrl = import.meta.env.BASE_URL || "/"
): Promise<Uint8Array> {
  return fetchBundledTemplateInBrowser(baseUrl);
}

export async function downloadBufferInBrowser(
  buffer: Uint8Array,
  fileName: string
): Promise<string> {
  const blobBytes = buffer.slice();
  const blob = new Blob([blobBytes], {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
  });
  const href = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = href;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(href);
  return fileName;
}
