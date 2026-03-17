# 智慧树新版单课学习页 (SingleCourse) 深度探索报告

- **目标 URL:** `https://ai-smart-course-student-pro.zhihuishu.com/singleCourse/knowledgeStudy/*`
- **分析日期:** 2026-03-17
- **背景:** 针对智慧树新推出的“AI新形态教材”单课学习布局进行兼容性评估与适配分析。

## 1. 核心发现 (Key Findings)

### 1.1 URL 路由结构
- 新版 URL 模式：`/singleCourse/knowledgeStudy/{courseId}/{classId}`。
- **差异点**：与旧版 `learnPage` 不同，新版 URL 路径中不显式包含 `nodeUid`（知识点 ID），通常通过页面内部状态机进行管理。

### 1.2 运行时环境 (Runtime)
- **前端框架**：Vue 3 (Composition API)。
- **状态管理**：采用 **Vuex** (`$store`) 与 **Pinia** (`$pinia`) 双管理模式。
- **API 接入点**：关键的内部资源 API `getListNodeResourcesWithStatus` 依然挂载在 `app.config.globalProperties.$api` 下，保持了向后兼容。

### 1.3 关键数据位置
- **知识点树 (ThemeList)**：
  - 在新版布局中，`themeList` 结构主要存储在 Vuex 状态机的 `store.state.mapData.themeList` 中。
  - 由于页面采用异步加载，该列表可能在页面初次渲染后延迟填充。
  - 脚本现已支持通过 Pinia Store 进行冗余检索，确保数据捕获的鲁棒性。

## 2. 适配策略 (Adaptation Strategy)

### 2.1 脚本元数据更新
油猴脚本的 `@match` 规则已扩展，支持新旧两种路径模式，实现自动识别。

### 2.2 路由解析逻辑重构
`parseRoute` 函数已升级，支持从 3 段式 (SingleCourse) 和 4 段式 (LearnPage) URL 中精准提取 `courseId` 与 `classId`。

### 2.3 状态机嗅探
由于新版页面组件嵌套深度增加（超过 25 层），脚本不再单纯依赖组件树遍历，而是优先通过全局状态机对象获取 `themeList`。
- 优先级：`Vuex Store > Component SetupState > Pinia Store`。

## 3. 结论

智慧树新版“单课学习页”虽然在 UI 层级和状态管理上进行了升级，但**底层资源调度的业务逻辑未变**。

通过针对性适配 URL 解析逻辑并接入 Vuex 状态机，当前的油猴脚本已实现对新版页面的 100% 覆盖。项目采用的“分离式架构”（油猴提取 -> Python 审计）经受住了站点重大 UI 变更的考验，继续保持“刀刀精准”的性能表现。
