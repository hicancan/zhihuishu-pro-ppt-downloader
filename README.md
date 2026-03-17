# 🚀 Zhihuishu-Pro-PPT-Downloader | 智慧树 PPT 资源全量提取利器

[![Python](https://img.shields.io/badge/Python-3.11%2B-blue?logo=python)](https://www.python.org/)
[![uv](https://img.shields.io/badge/managed%20by-uv-arc.svg)](https://github.com/astral-sh/uv)
[![Tampermonkey](https://img.shields.io/badge/Tampermonkey-v5.0%2B-green?logo=tampermonkey)](https://www.tampermonkey.net/)
[![Platform](https://img.shields.io/badge/Platform-Zhihuishu%20Pro-orange)](https://ai-smart-course-student-pro.zhihuishu.com/)

> **基于“分离式架构”的极致提取方案。不翻页、不点开、直接调用站内 API，实现 100% 精准的课程 PPT 资源获取与审计。**

---

## 🌟 核心特性

*   **⚡ 刀刀精准的提取逻辑**：直接调用智慧树内部 Vue 运行时 API (`getListNodeResourcesWithStatus`)，无需模拟点击，无需解析 DOM。
*   **🌍 全版本兼容**：同时支持智慧树旧版学习页 (`learnPage`) 和最新的单课学习页 (`singleCourse`)。
*   **📊 离线审计系统**：通过 Python 构建的 Manifest (清单) 审计系统，解决批量下载中“有没有漏”的终极痛点。
*   **🔬 深度对账 (Reconciliation)**：支持将导出的清单与本地下载目录进行自动化对比，自动生成缺失/冗余报告。
*   **🛡️ 身份验证穿透**：利用油猴脚本寄生在已登录的浏览器上下文中，完美绕过复杂的签名校验。

---

## 🏗️ 架构设计：分离式架构 (Split Architecture)

本项目采用“前端取数，后端审计”的策略：

1.  **[提取端] Tampermonkey 脚本**：运行在浏览器内，负责利用已登录的 Cookie 和站点内部 API，全量扫描知识点并导出 `resource-manifest.json`。
2.  **[审计端] Python CLI 工具**：在本地运行，负责解析 Manifest 清单，进行去重、统计、对账，并指导用户完成最终的资源归档。

---

## 🚀 快速上手

### 1. 环境准备
确保已安装 [uv](https://github.com/astral-sh/uv) 和 [Tampermonkey](https://www.tampermonkey.net/)。

```powershell
# 初始化 Python 环境
uv sync
```

### 2. 前端提取 (浏览器)
1. 安装脚本：`tampermonkey/zhihuishu-pro-ppt-downloader.user.js`。
2. 登录智慧树，进入任意课程学习页。
3. 在脚本面板点击 **“从头遍历整门课 PPT”** 或 **“导出整门课资源 JSON”**。

# 本地审计与全自动下载 (命令行)
将导出的 JSON 文件放入 `tampermonkey/` 目录，执行以下命令：

```powershell
# 1. 查看清单摘要（资源数、PPT 分布、去重统计）
uv run zhihuishu-pro-ppt-downloader manifest-summary "tampermonkey/your-manifest.json"

# 2. 全自动批量下载（直接读取清单 URL 下载到本地）
uv run zhihuishu-pro-ppt-downloader manifest-download "tampermonkey/your-manifest.json" --downloads-dir "./downloads"

#### 4. 转换 PDF (CLI)
按照课程顺序将 PPT 转换为 PDF。

**合并为一个带书签的文件：**
```powershell
uv run zhihuishu-pro-ppt-downloader manifest-export-pdf "path/to/manifest.json" --downloads-dir "./downloads" --output "course_merged.pdf"
```

**分散转换为独立文件：**
```powershell
uv run zhihuishu-pro-ppt-downloader manifest-export-pdf "path/to/manifest.json" --downloads-dir "./downloads" --individual --output-dir "./pdfs"
```

#### 5. 自动化对账 (CLI)
```powershell
uv run zhihuishu-pro-ppt-downloader manifest-reconcile "tampermonkey/your-manifest.json" --downloads-dir "./downloads"
```

---

## 📂 项目结构

| 目录/文件 | 说明 |
| :--- | :--- |
| `tampermonkey/` | 核心油猴脚本：Vue 运行时劫持与 API 调度逻辑 |
| `src/` | Python CLI 源码：Manifest 解析与自动化对账系统 |
| `doc/` | 深度探索文档：包含站点逆向报告与新布局分析 |
| `GEMINI.md` | AI 指令上下文：为 Gemini CLI 提供的专家级引导 |

---

## 📜 技术说明

### 为什么不使用纯 Python 爬虫？
智慧树的 API 带有复杂的请求签名 (`secretStr`)。通过油猴脚本直接调用页面已挂载的 API 对象，可以省去重写签名算法的巨大维护压力，实现“以最小成本达成最高精度”。

---

## 🤝 贡献与反馈
如果你发现新的页面结构或 API 变更，欢迎提交 Issue 或 Pull Request。

**鸣谢：** 本项目致力于提升数字化学习资源的获取效率。请在遵守相关法律法规及平台协议的前提下使用。
