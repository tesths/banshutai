# 板书台

面向教师的黑板贴生成器。把多行文本粘贴进页面后，工具会按“一行一页”自动排入模板，预估页数，并直接导出 `pptx` 文件。

![板书台生成效果预览](./docs/assets/generated-preview.png)

## 功能概览

- 浏览器端运行，不依赖桌面壳层
- 默认加载 `public/default-template.pptx`
- 多行文本一行一页，保留空行与首尾空格
- 长文本优先缩字，不够再自动拆分页
- 三层叠字同步更新，并保持居中布局
- 网页内直接显示模板最终效果预览
- 可自定义蓝色描边层颜色
- 生成后直接下载 `.pptx`
- 构建时自动输出 `robots.txt` 和 `sitemap.xml`

## 目录说明

```text
.
├── docs/
│   ├── assets/                     README 预览图
│   ├── templates/                  原始模板备份
│   └── DESIGN.md                   视觉设计参考
├── public/
│   ├── default-template.pptx       运行时默认模板
│   ├── favicon.svg
│   └── site.webmanifest
├── src/
│   ├── lib/                        浏览器工作流与运行时适配
│   ├── shared/                     PPTX XML 引擎
│   ├── App.tsx                     页面交互
│   └── styles.css                  页面样式
├── test/                           Node 集成测试
├── .gitignore
├── package.json
├── vercel.json
└── vite.config.ts
```

## 开发环境

- Node.js `24+`
- npm `10+`

## 本地开发

```bash
npm install
npm run dev
```

默认地址：

- `http://127.0.0.1:1420/`

常用命令：

```bash
npm test
npm run typecheck
npm run build
npm run preview
```

如果部署到正式域名，建议在构建环境里设置：

```bash
SITE_URL=https://你的正式域名
```

这样构建产物里的 `sitemap.xml` 和 `robots.txt` 会使用正式站点地址。

## 模板与文档

- `public/default-template.pptx` 是网页实际生成时使用的默认模板，必须保留
- `docs/templates/original-template-backup.pptx` 是原始模板备份，换电脑后也可以继续调整模板
- `docs/DESIGN.md` 记录了当前页面的视觉方向

如果要调整导出文字区域，优先修改模板第一页的三层文本框，并保持三层叠放关系。

## 部署到 Vercel

这个仓库已经是标准的 Vite 静态项目，直接把 GitHub 仓库导入 Vercel 即可。

1. 把代码推送到 GitHub。
2. 登录 Vercel，点击 `Add New` -> `Project`。
3. 选择这个仓库：`tesths/banshutai`。
4. 构建设置确认如下：

```text
Framework Preset: Vite
Install Command: npm install
Build Command: npm run build
Output Directory: dist
```

5. 点击 `Deploy`。

仓库里已经包含 `vercel.json`，Vercel 连接 GitHub 后可以直接按上述配置部署。

## 部署前建议

```bash
npm test
npm run typecheck
npm run build
```
