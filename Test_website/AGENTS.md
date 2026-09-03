# AGENTS.md

## Project Overview

BioMechanism AI 是一个通过 HTML5、CSS 和原生 JavaScript 实现的静态网站，用于展示 AI 与高通量数据分析在生物分子机制解析中的应用。整体 UI 参考苹果官网的简约产品叙事风格：大留白、强标题、克制配色和高质量科研视觉。

## Folder Structure

- `index.html`：页面主体、内容结构和语义化 HTML。
- `styles.css`：全站样式、设计 token、响应式规则和动效。
- `script.js`：滚动淡入和机制图谱切换交互。
- `public/assets/`：站点运行时可访问的图片资源。
- `docs/`：设计概念、说明或开发参考资料。

## Build And Run Commands

```powershell
npm install
npm run dev
npm run build
```

## Testing Workflow

1. 修改后先运行 `npm run build` 检查静态站打包。
2. 若改动样式或交互，运行 `npm run dev` 并在浏览器检查桌面和移动端布局。
3. 重点检查首页首屏、机制图谱按钮、CTA、移动端文本换行和图片裁切。

## Coding Conventions

- 保持中文内容清晰、专业、简洁。
- 页面内容优先保持在 `index.html` 中；机制图谱切换数据位于 `script.js`。
- CSS 使用已有变量和组件规则，避免无关的大规模重构。
- 卡片圆角保持在 `8px` 或以下，整体风格保持简约、科研可信。

## Important Files

- `public/assets/hero-biomolecular-network.png` 是首页 hero 的核心视觉资产。
- `docs/concept-biomolecular-ai-site.png` 是本轮视觉方向参考图。
