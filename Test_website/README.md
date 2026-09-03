# BioMechanism AI HTML5 网站

## 快速运行

Windows 下可以直接双击：

```powershell
open_site.bat
```

它会启动本地服务并自动打开 `http://127.0.0.1:5173/`。

也可以手动运行：

```powershell
npm install
npm run dev
```

浏览器打开终端输出的本地地址，默认是 `http://127.0.0.1:5173/`。

## 项目说明

这是一个通过 HTML5、CSS 和原生 JavaScript 实现的简约科研网站，整体视觉参考苹果官网的产品式叙事：大留白、强标题、克制配色、精细排版和少量高质量科学视觉。内容覆盖蛋白互作、药物开发、受体研究、信号轴解析、数据到机制流程、可解释机制图谱和验证导向输出。

## 常用命令

```powershell
npm run dev
npm run build
npm run preview
```

## GitHub Pages 在线预览

网站通过仓库根目录下的 `.github/workflows/website-pages.yml` 自动构建和部署。

开发流程：

```text
dev 分支更新 Test_website
→ GitHub Actions
→ npm ci
→ npm run build
→ Test_website/dist
→ GitHub Pages
```

Vite 配置 `vite.config.js` 使用相对资源路径，因此构建结果可以部署在 GitHub Pages 的项目子路径下。

首次启用时，请在 GitHub 仓库中进入：

```text
Settings → Pages → Build and deployment → Source → GitHub Actions
```

完成这一次设置后，后续 `dev` 分支中 `Test_website` 的修改会自动触发构建和部署。

## 目录结构

- `index.html`：HTML5 页面主体。
- `styles.css`：苹果式简约视觉、响应式布局和动效。
- `script.js`：滚动淡入和机制图谱切换交互。
- `public/assets/hero-biomolecular-network.png`：首页科学视觉资产。
- `docs/concept-apple-style-html5.png`：本次苹果式 HTML5 视觉概念参考图。
- `vite.config.js`：Vite 构建及 GitHub Pages 路径配置。

## 依赖

仅使用 Vite 作为本地开发和构建工具。页面运行时不依赖 React、Vue 或其他前端框架。
