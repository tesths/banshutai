import { mkdir, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { pathToFileURL } from "node:url";

const DEFAULT_SITE_URL = "https://banshutai.dididigu.com";

export function normalizeSiteUrl(value) {
  const trimmed = value?.trim();

  if (!trimmed) {
    return null;
  }

  const withProtocol = /^https?:\/\//i.test(trimmed) ? trimmed : `https://${trimmed}`;

  try {
    const url = new URL(withProtocol);
    const normalizedPath = url.pathname === "/" ? "" : url.pathname.replace(/\/$/, "");

    return `${url.origin}${normalizedPath}`;
  } catch {
    return null;
  }
}

export function resolveSiteUrl(env = process.env) {
  const configuredSiteUrl = normalizeSiteUrl(
    env.SITE_URL ||
      env.VITE_SITE_URL ||
      env.VERCEL_PROJECT_PRODUCTION_URL ||
      env.VERCEL_URL ||
      env.URL ||
      env.DEPLOY_PRIME_URL ||
      ""
  );

  if (configuredSiteUrl) {
    return configuredSiteUrl;
  }

  return normalizeSiteUrl(DEFAULT_SITE_URL);
}

export function buildRobotsTxt(siteUrl) {
  const lines = ["User-agent: *", "Allow: /"];

  if (siteUrl) {
    lines.push("", `Sitemap: ${siteUrl}/sitemap.xml`);
  }

  return `${lines.join("\n")}\n`;
}

export function buildSitemapXml(siteUrl) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
  <url>
    <loc>${siteUrl}/</loc>
    <changefreq>weekly</changefreq>
    <priority>1.0</priority>
  </url>
</urlset>
`;
}

export async function writeSeoAssets(outputDir, env = process.env) {
  const siteUrl = resolveSiteUrl(env);
  const destination = resolve(outputDir);

  await mkdir(destination, { recursive: true });
  await writeFile(resolve(destination, "robots.txt"), buildRobotsTxt(siteUrl), "utf8");

  if (siteUrl) {
    await writeFile(resolve(destination, "sitemap.xml"), buildSitemapXml(siteUrl), "utf8");
  }
}

async function main() {
  const outputDir = process.argv[2] || "dist";
  await writeSeoAssets(outputDir);
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
}
