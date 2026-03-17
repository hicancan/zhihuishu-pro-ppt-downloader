# 智慧树课程 PPT 批量下载脚本逆向分析报告

- 生成时间：2026-03-17
- 分析对象：[zhihuishu-batch-ppt-downloader.user.js](D:/code/github/hicancan/zhihuishu-downloader/tampermonkey/zhihuishu-batch-ppt-downloader.user.js)
- 分析方式：MCP 页面运行时检查、MCP Network 抓包、脚本静态审查、课程页实测
- 课程样本：`数学建模2025-2026(2)`

## 1. 结论摘要

当前批量脚本已经从“逐页打开课程内容，再读当前页 DOM 资源卡”升级为“读取课程节点树后，直接调用站内内部 API 拉取每个节点的资源列表”。  
在“只下载 PPT/PPTX”这个目标下，这个思路明显比旧方案更稳，也更接近“刀刀精准”。

但是，逆向过程中确认到一处必须修正的边界：

- `themeList` 中存在 `resourceCount = 0`，但内部 API 实际仍会返回资源的节点。
- 这说明“先用 `resourceCount > 0` 过滤节点，再去查资源”并不绝对可靠。
- 当前仓库中的脚本已修正为“扫描整门课全部知识点”，不再依赖 `resourceCount` 预过滤。

修正位置：

- 上下文识别与 API 化遍历入口： [zhihuishu-batch-ppt-downloader.user.js](D:/code/github/hicancan/zhihuishu-downloader/tampermonkey/zhihuishu-batch-ppt-downloader.user.js#L68)
- 节点级 API 拉取： [zhihuishu-batch-ppt-downloader.user.js](D:/code/github/hicancan/zhihuishu-downloader/tampermonkey/zhihuishu-batch-ppt-downloader.user.js#L254)
- 旧缓存隔离： [zhihuishu-batch-ppt-downloader.user.js](D:/code/github/hicancan/zhihuishu-downloader/tampermonkey/zhihuishu-batch-ppt-downloader.user.js#L661)
- 全量节点扫描修正： [zhihuishu-batch-ppt-downloader.user.js](D:/code/github/hicancan/zhihuishu-downloader/tampermonkey/zhihuishu-batch-ppt-downloader.user.js#L369)

## 2. 逆向目标

本次逆向重点验证三个问题：

1. 脚本识别 PPT 的链路是否真正依赖“打开每个知识点页面”。
2. 当前过滤规则是否会漏掉真实 PPT，或者把非 PPT 误识别为 PPT。
3. 站点内部接口是否允许在“当前只打开一个课程页”的前提下，直接查询其他知识点资源。

## 3. 课程页运行时结构

MCP 直接检查页面运行时后，确认脚本可以稳定拿到以下信息：

- 课程页 URL 中自带 `courseId / nodeUid / classId`
- Vue 运行时里存在完整的 `themeList`
- Vue 运行时 `app.config.globalProperties.$api` 暴露了课程内部资源接口

当前脚本的运行时识别方式是：

1. 从 URL 解析 `courseId / classId / 当前 nodeUid`
2. 从 Vue 组件树中找到持有 `themeList` 的课程组件
3. 展开 `themeList -> subThemeList -> knowledgeList`
4. 得到整门课全部知识点
5. 对每个知识点调用内部资源接口，再做 PPT 过滤

这条链路已经不再依赖“手动或自动切换到那个知识点页面再读取 DOM”。

## 4. 内部资源接口链路

### 4.1 关键接口

MCP 运行时和脚本静态检查都确认，课程资源的关键入口是：

- `app.config.globalProperties.$api.getListNodeResourcesWithStatus(...)`

脚本实际调用位置：

- [zhihuishu-batch-ppt-downloader.user.js](D:/code/github/hicancan/zhihuishu-downloader/tampermonkey/zhihuishu-batch-ppt-downloader.user.js#L254)

### 4.2 抓包结果

MCP 抓到的真实请求目标是：

- `POST https://kg-ai-run.zhihuishu.com/run/gateway/t/stu/resources/list-knowledge-resource`

抓包里可见的真实请求体不是明文的 `courseId/classId/knowledgeId`，而是：

```json
{
  "secretStr": "...",
  "date": 1773742120906
}
```

同时请求里还有站点自带的头，例如：

- `xqjzxhiz`
- 登录态 Cookie
- `origin/referer`

这说明：

- 站点前端不会把查询参数原样发出去
- 页面运行时会先做一次签名或加密，再发真实请求
- 所以“完全脱离课程页，用外部裸 HTTP 脚本直接调接口”并不稳

## 5. 不打开目标页面，能不能直接拿到 PPT

答案是：能。

MCP 实测时，浏览器当前停留在别的知识点页面，但直接从当前页面运行时调用：

```js
api.getListNodeResourcesWithStatus({
  courseId,
  classId,
  knowledgeId: "1894301155266269184"
})
```

仍然成功返回了 `经典案例2-森林救火` 的完整资源列表，其中包含：

- `森林救火问题（简单优化模型）.pptx`
- 直链位于 `file.zhihuishu.com`

因此，结论非常明确：

- 不需要真的打开每个知识点页面
- 需要一个已经登录的课程页，作为“调用内部 API 的执行上下文”

## 6. 当前课程里的资源类型特征

### 6.1 当前视频页样本

MCP 在 `多目标优化模型` 这个当前页上直接查到的资源有：

- 1 个视频资源：`resourcesDataType = 22`
- 2 个外链资源：`resourcesDataType = 23`
- 当前页没有 PPT

这说明脚本面板里显示 `当前页 PPT 数：0` 在这种页面是正确结果，不是识别失败。

### 6.2 森林救火页样本

MCP 对 `经典案例2-森林救火` 的接口返回是：

- 1 个 PPT：`resourcesType = 1`，`resourcesDataType = 11`，`resourcesSuffix = pptx`
- 1 个教材：`resourcesDataType = 21`
- 1 个视频：`resourcesDataType = 22`
- 2 个外链：`resourcesDataType = 23`

其中 PPT 原始直链为：

- `https://file.zhihuishu.com/.../e20b525ad8ba4e5db8185787134bd195.pptx`

这个样本验证了当前 PPT 过滤思路的核心假设：

- PPT 的后缀确实是 `ppt/pptx`
- PPT 的下载域名确实是 `file.zhihuishu.com`
- 视频与教材都可以用 `resourcesDataType` 排除

## 7. 精准性评估

## 7.1 识别准确的部分

当前脚本对 PPT 的核心判断是：

- 从节点资源接口取全量资源
- 只保留后缀为 `ppt/pptx` 的资源
- 排除教材 `21` 和视频 `22`
- 仅对实际可下载直链发起下载

这比“根据资源卡外观/DOM 类名猜类型”要精准得多。  
对目前这个课程样本而言，这套逻辑是成立的。

## 7.2 已确认的非绝对精准点

唯一确认到的真实边界是：

- `resourceCount = 0` 并不等于“接口一定没有资源”

MCP 抽到的反例：

- 节点：`偏微分方程模型`
- `themeList.resourceCount = 0`
- 但接口实际返回了 3 个资源

这 3 个资源虽然不是 PPT，只是：

- 1 个视频
- 2 个外链

但它足以证明：  
如果脚本继续使用 `resourceCount > 0` 作为扫描前过滤条件，就存在理论漏抓风险。

所以本仓库脚本已改为：

- 不再预先跳过 `resourceCount = 0` 的节点
- 直接扫描整门课所有知识点

## 7.3 当前精度判断

修正后，当前脚本可以评价为：

- 对这门课的 PPT 批量发现链路，已经达到“高精度”
- 方法论上比旧版“逐页打开再读 DOM”更稳定
- 但还不能宣称对未来所有课程、所有资源结构都“绝对零漏抓”

原因在于它仍然依赖三个前提：

1. 站点继续保留现有内部 API
2. PPT 继续以 `ppt/pptx` 后缀暴露
3. 页面运行时仍能正常生成 `secretStr`

## 8. 当前脚本的剩余风险

### 8.1 站点接口变更风险

如果站点未来修改：

- `getListNodeResourcesWithStatus` 的调用方式
- `response` 数据结构
- `secretStr` 生成逻辑

则脚本会失效。

### 8.2 资源命名方式变化风险

如果未来出现：

- 下载直链没有 `ppt/pptx` 后缀
- 后缀不在 `resourcesSuffix`、文件名和 URL 任一处暴露

那么当前过滤规则会漏掉这类 PPT。

### 8.3 浏览器环境依赖

当前方案虽然不需要逐页打开节点，但仍然依赖：

- 用户已登录智慧树
- 浏览器课程页打开着
- 油猴脚本能访问页面运行时

## 9. 建议

为了让脚本更接近“可证明的精准”，建议后续再补两项：

1. 增加调试导出模式  
   每次批量运行后，把每个节点的原始资源元数据导出为 JSON，便于后验核对。  
   该能力现已在脚本中实现，可直接导出“整门课资源 JSON 清单”。

2. 增加异常资源白名单日志  
   对“后缀不明、但 `resourcesType = 1` 的文件资源”单独记日志，而不是直接忽略。

## 9.1 基于导出清单的实测验证

在导出文件 [数学建模2025-2026(2) - resource-manifest.json](D:/code/github/hicancan/zhihuishu-downloader/tampermonkey/数学建模2025-2026(2)%20-%20resource-manifest.json) 上再次做了后验核对，结果如下：

- 总知识点数：74
- 实际扫描知识点数：74
- 实际资源总数：241
- 被脚本判定为 PPT 的资源总数：19
- `resourceCount = 0` 的节点数：3
- `resourceCount = 0` 但接口实际返回资源的节点数：1
- `resourceCount = 0` 但存在 PPT 的节点数：0
- 接口报错节点数：0

对 19 个已识别 PPT 的后验统计结果：

- 全部 `resourcesType = 1`
- 全部 `resourcesDataType = 11`
- 全部后缀为 `pptx`
- 全部下载域名来自 `file.zhihuishu.com`
- 没有发现“脚本未判为 PPT，但看起来像 PPT”的可疑资源
- 没有发现重复 `resourcesUid`
- 没有发现重复下载链接

这说明：  
在这门课程样本上，当前脚本对 PPT 的识别结果与导出清单完全一致，没有发现漏判或误判。

## 10. 最终结论

这版批量脚本的路线已经选对了：

- 用课程页运行时拿节点树
- 用站内内部 API 查节点资源
- 只下载 PPT/PPTX

相比早期“切页 + 读 DOM”的方案，它已经明显更稳，也更接近真正的逆向自动化工具。

但严格地说，它原先并不是“刀刀精准”，因为：

- 它曾依赖 `resourceCount > 0` 预过滤知识点

这一点已经在本仓库中修正。  
修正后，这个脚本在当前课程样本上的 PPT 批量抓取精度，可以评估为：

- `高精度，可用于实际使用`
- `仍建议保留调试导出能力，以应对后续站点变更`
