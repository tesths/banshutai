// @vitest-environment node

import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { buildRobotsTxt, buildSitemapXml, normalizeSiteUrl, resolveSiteUrl } from "../../scripts/seo-assets.mjs";

describe("seo assets", () => {
  it("normalizes configured site URLs before generating assets", () => {
    expect(normalizeSiteUrl("https://banshutai.example.com/")).toBe("https://banshutai.example.com");
    expect(normalizeSiteUrl("banshutai.vercel.app")).toBe("https://banshutai.vercel.app");
    expect(normalizeSiteUrl("")).toBeNull();
  });

  it("falls back to the inferred Vercel production domain when no site URL is configured", () => {
    expect(resolveSiteUrl({ npm_package_name: "banshutai" })).toBe("https://banshutai.dididigu.com");
  });

  it("builds robots text with a sitemap when a site URL exists", () => {
    expect(buildRobotsTxt("https://banshutai.example.com")).toContain(
      "Sitemap: https://banshutai.example.com/sitemap.xml"
    );
    expect(buildRobotsTxt(null)).not.toContain("Sitemap:");
  });

  it("builds a single-page sitemap with absolute URLs", () => {
    expect(buildSitemapXml("https://banshutai.example.com")).toContain(
      "<loc>https://banshutai.example.com/</loc>"
    );
  });

  it("declares teacher-focused metadata in index.html", async () => {
    const html = await readFile(resolve(process.cwd(), "index.html"), "utf8");

    expect(html).toContain("教师黑板贴生成器");
    expect(html).toContain("文本转 PPTX");
    expect(html).toContain("\"FAQPage\"");
    expect(html).toContain("\"HowTo\"");
    expect(html).toContain("\"WebSite\"");
    expect(html).toContain("https://banshutai.dididigu.com/default-template-preview.png");
    expect(html).toContain("<link rel=\"canonical\" href=\"https://banshutai.dididigu.com/\" />");
    expect(html).toContain("<meta property=\"og:url\" content=\"https://banshutai.dididigu.com/\" />");
    expect(html).toContain("og:image");
    expect(html).toContain("twitter:image:alt");
  });
});
