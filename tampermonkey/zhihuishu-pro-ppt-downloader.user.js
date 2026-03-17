// ==UserScript==
// @name         Zhihuishu Pro PPT Downloader
// @namespace    https://github.com/hicancan/zhihuishu-downloader
// @version      0.4.0
// @description  Discover and download all original PPT/PPTX files in the current Zhihuishu course without opening every knowledge page.
// @author       hicancan
// @match        https://ai-smart-course-student-pro.zhihuishu.com/learnPage/*
// @match        https://ai-smart-course-student-pro.zhihuishu.com/singleCourse/knowledgeStudy/*
// @run-at       document-idle
// @grant        GM_download
// @grant        GM_notification
// @connect      file.zhihuishu.com
// ==/UserScript==

(function () {
  "use strict";

  const SCRIPT_ID = "zhs-batch-ppt-downloader";
  const TASK_KEY_PREFIX = "zhs-batch-ppt-downloader";
  const TASK_VERSION = 3;
  const COMPONENT_MAX_DEPTH = 32;
  const PPT_SUFFIXES = new Set(["ppt", "pptx"]);
  const BOOT_RETRY_DELAY_MS = 1500;
  const BETWEEN_NODE_DELAY_MS = 350;
  const BETWEEN_DOWNLOAD_DELAY_MS = 900;

  const runtime = {
    initialized: false,
    loopStarted: false,
    exportRunning: false,
  };

  if (window.__ZHS_BATCH_PPT_DOWNLOADER__) {
    return;
  }
  window.__ZHS_BATCH_PPT_DOWNLOADER__ = true;

  boot();

  function boot() {
    injectStyle();
    ensureUI();
    renderStatus("等待课程上下文加载...");
    waitForContext()
      .then((context) => {
        const legacyTask = loadAnyTask(context.courseId, context.classId);
        if (legacyTask && legacyTask.version !== TASK_VERSION) {
          renderStatus("检测到旧版任务缓存，已忽略，请重新开始遍历。");
        }
        updatePanel(context);
        bindTaskLifecycle(context);
      })
      .catch((err) => {
        console.warn(`[${SCRIPT_ID}] boot retry`, err);
        renderStatus("当前页面还没进入课程学习页，稍后自动重试。");
        window.setTimeout(boot, BOOT_RETRY_DELAY_MS);
      });
  }

  async function waitForContext(maxAttempts = 60) {
    for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
      const context = buildContext();
      if (context.ok) {
        return context;
      }
      await sleep(500);
    }
    throw new Error("课程页面上下文未就绪");
  }

  function buildContext() {
    const route = parseRoute();
    const app = document.querySelector("#app")?.__vue_app__;
    const root = app?._container?._vnode?.component || app?._instance;
    const api = app?.config?.globalProperties?.$api || null;
    const store = app?.config?.globalProperties?.$store || null;

    if (!route.courseId || !route.classId || !root || !api?.getListNodeResourcesWithStatus) {
      return { ok: false };
    }

    const courseInst = findCourseInstance(root);
    let themeList = Array.isArray(readInstanceValue(courseInst, "themeList"))
      ? readInstanceValue(courseInst, "themeList")
      : [];

    // Fallback to Vuex store if themeList is empty (common in new singleCourse layout)
    if (!themeList.length && store?.state?.mapData?.themeList) {
      themeList = store.state.mapData.themeList;
    }

    const previewData = simplifyPreview(readInstanceValue(courseInst, "previewData") || {});
    const knowledgeNodes =
      flattenKnowledgeNodes(themeList).length > 0
        ? flattenKnowledgeNodes(themeList)
        : extractNodesFromLegacyTask(loadAnyTask(route.courseId, route.classId));
    
    // Attempt to get nodeUid from store if missing from URL
    const effectiveNodeUid = route.nodeUid || store?.state?.knowledgeId || "";

    const currentNode =
      knowledgeNodes.find((item) => item.nodeUid === effectiveNodeUid) ||
      buildCurrentFallback(effectiveNodeUid, previewData);
    const currentPagePptResources = collectCurrentPagePpts(root, previewData);
    const mapName =
      String(previewData.mapName || "").trim() ||
      String(readInstanceValue(courseInst, "mapName") || "").trim() ||
      String(store?.state?.subjectName || "").trim() ||
      String(document.querySelector("[class*='course'] h1, [class*='course'] h2, h1")?.textContent || "").trim() ||
      "当前课程";

    if (!knowledgeNodes.length) {
      return { ok: false };
    }

    return {
      ok: true,
      app,
      api,
      courseId: route.courseId,
      classId: route.classId,
      nodeUid: effectiveNodeUid,
      mapName,
      previewData,
      knowledgeNodes,
      currentNode,
      currentPagePptResources,
    };
  }

  function bindTaskLifecycle(context) {
    if (runtime.initialized) {
      return;
    }
    runtime.initialized = true;

    const startBtn = document.getElementById(`${SCRIPT_ID}-start`);
    const stopBtn = document.getElementById(`${SCRIPT_ID}-stop`);
    const exportBtn = document.getElementById(`${SCRIPT_ID}-export`);

    startBtn?.addEventListener("click", async () => {
      if (runtime.exportRunning) {
        notify("正在导出资源清单，请等待导出结束", true);
        return;
      }
      const fresh = await waitForContext();
      const task = createTask(fresh);
      saveTask(task);
      renderStatus(`已建立任务，共 ${task.nodes.length} 个知识点，准备直接查资源接口。`);
      updatePanel(fresh, task);
      notify(`已开始遍历课程资源，共 ${task.nodes.length} 个知识点`);
      void runTaskLoop();
    });

    stopBtn?.addEventListener("click", async () => {
      const fresh = await waitForContext();
      const task = loadTask(fresh.courseId, fresh.classId);
      if (!task) {
        notify("当前没有运行中的新任务");
        return;
      }
      task.active = false;
      task.updatedAt = Date.now();
      saveTask(task);
      renderStatus("已停止遍历任务");
      updatePanel(fresh, task);
      notify("已停止遍历任务");
    });

    exportBtn?.addEventListener("click", async () => {
      if (runtime.loopStarted) {
        notify("请先停止批量下载任务，再导出资源清单", true);
        return;
      }
      if (runtime.exportRunning) {
        notify("资源清单导出已在进行中");
        return;
      }
      const fresh = await waitForContext();
      void exportCourseManifest(fresh);
    });

    const task = loadTask(context.courseId, context.classId);
    if (task?.active) {
      renderStatus("检测到未完成任务，正在继续拉取课程资源...");
      updatePanel(context, task);
      void runTaskLoop();
    }
  }

  async function runTaskLoop() {
    if (runtime.loopStarted) {
      return;
    }
    runtime.loopStarted = true;

    try {
      const context = await waitForContext();
      let task = loadTask(context.courseId, context.classId);
      if (!task?.active) {
        updatePanel(context, task);
        return;
      }

      while (task?.active) {
        const node = task.nodes[task.currentIndex];
        if (!node) {
          finishTask(task, context, "整门课 PPT 已全部检索完成");
          break;
        }

        task.lastNodeUid = node.nodeUid;
        task.lastNodeName = node.knowledgeName;
        task.updatedAt = Date.now();
        saveTask(task);
        updatePanel(context, task);
        renderStatus(`正在检索：${node.knowledgeName}`);

        let resources = [];
        try {
          resources = await fetchNodePpts(context, node);
        } catch (error) {
          console.error(`[${SCRIPT_ID}] fetchNodePpts`, error);
          task.failedDownloads.push({
            nodeUid: node.nodeUid,
            knowledgeName: node.knowledgeName,
            resourcesName: "",
            reason: `资源查询失败：${String(error?.message || error)}`,
          });
        }

        task.lastNodePptCount = resources.length;
        task.foundPptCount += resources.length;

        for (const resource of resources) {
          const resourceKey = buildResourceKey(node, resource);
          if (task.downloadedResourceKeys[resourceKey]) {
            continue;
          }

          try {
            const filename = buildPptFilename(node, resource);
            await triggerDownload(resource.downloadUrl, filename);
            task.downloadedResourceKeys[resourceKey] = true;
            task.downloadedCount += 1;
            task.updatedAt = Date.now();
            saveTask(task);
            renderStatus(`已发起下载：${resource.resourcesName}`);
          } catch (error) {
            console.error(`[${SCRIPT_ID}] triggerDownload`, error);
            task.failedDownloads.push({
              nodeUid: node.nodeUid,
              knowledgeName: node.knowledgeName,
              resourcesName: resource.resourcesName,
              reason: String(error?.message || error),
            });
            task.updatedAt = Date.now();
            saveTask(task);
            notify(`下载失败：${resource.resourcesName}`, true);
          }

          await sleep(BETWEEN_DOWNLOAD_DELAY_MS);
        }

        task.finishedNodeUids[node.nodeUid] = true;
        task.currentIndex += 1;
        task.updatedAt = Date.now();
        saveTask(task);
        updatePanel(context, task);

        task = loadTask(context.courseId, context.classId);
        if (!task?.active) {
          renderStatus("任务已停止");
          break;
        }

        await sleep(BETWEEN_NODE_DELAY_MS);
      }
    } catch (error) {
      console.error(`[${SCRIPT_ID}]`, error);
      notify(error?.message || "遍历任务失败", true);
    } finally {
      runtime.loopStarted = false;
    }
  }

  async function exportCourseManifest(context) {
    runtime.exportRunning = true;
    try {
      renderStatus(`正在导出整门课资源清单，共 ${context.knowledgeNodes.length} 个知识点...`);
      updatePanel(context);

      const manifest = {
        generatedAt: new Date().toISOString(),
        generator: {
          scriptId: SCRIPT_ID,
          version: "0.4.0",
        },
        course: {
          mapName: context.mapName,
          courseId: context.courseId,
          classId: context.classId,
          currentNodeUid: context.nodeUid,
          currentNodeName: context.currentNode?.knowledgeName || "",
        },
        rules: {
          pptSuffixes: Array.from(PPT_SUFFIXES).sort(),
          notes: [
            "资源清单来自课程页内部 API，而不是当前页面 DOM。",
            "isPpt 为脚本当前判定结果，rawResourceCount 为 themeList 中声明的资源数。",
          ],
        },
        stats: {
          totalNodes: context.knowledgeNodes.length,
          scannedNodes: 0,
          totalResources: 0,
          totalPpts: 0,
          zeroCountNodes: 0,
          zeroCountButActualResourcesNodes: 0,
          zeroCountButPptNodes: 0,
          errorCount: 0,
        },
        zeroCountButActualResourcesNodes: [],
        zeroCountButPptNodes: [],
        errors: [],
        nodes: [],
      };

      for (let index = 0; index < context.knowledgeNodes.length; index += 1) {
        const node = context.knowledgeNodes[index];
        renderStatus(`正在导出资源清单：${node.knowledgeName} (${index + 1}/${context.knowledgeNodes.length})`);

        try {
          const resourceGroups = await fetchNodeResourceGroups(context, node);
          const exportResources = resourceGroups.map((group, resourceIndex) => buildExportResourceRecord(group, resourceIndex));
          const pptResources = exportResources.filter((item) => item.isPpt);

          manifest.stats.scannedNodes += 1;
          manifest.stats.totalResources += exportResources.length;
          manifest.stats.totalPpts += pptResources.length;
          if (Number(node.resourceCount || 0) === 0) {
            manifest.stats.zeroCountNodes += 1;
            if (exportResources.length > 0) {
              manifest.stats.zeroCountButActualResourcesNodes += 1;
              manifest.zeroCountButActualResourcesNodes.push({
                nodeUid: node.nodeUid,
                knowledgeName: node.knowledgeName,
                actualResourceCount: exportResources.length,
              });
            }
            if (pptResources.length > 0) {
              manifest.stats.zeroCountButPptNodes += 1;
              manifest.zeroCountButPptNodes.push({
                nodeUid: node.nodeUid,
                knowledgeName: node.knowledgeName,
                pptCount: pptResources.length,
              });
            }
          }

          manifest.nodes.push({
            index,
            themeName: node.themeName,
            subThemeName: node.subThemeName,
            knowledgeName: node.knowledgeName,
            nodeUid: node.nodeUid,
            rawResourceCount: Number(node.resourceCount || 0),
            actualResourceCount: exportResources.length,
            pptCount: pptResources.length,
            resources: exportResources,
          });
        } catch (error) {
          manifest.stats.errorCount += 1;
          manifest.errors.push({
            index,
            nodeUid: node.nodeUid,
            knowledgeName: node.knowledgeName,
            error: String(error?.message || error),
          });
        }

        await sleep(BETWEEN_NODE_DELAY_MS);
      }

      const filename = buildManifestFilename(context);
      downloadTextFile(filename, JSON.stringify(manifest, null, 2), "application/json");
      renderStatus(`资源清单导出完成：${filename}`);
      updatePanel(context);
      notify(`已导出整门课资源清单：${filename}`);
    } catch (error) {
      console.error(`[${SCRIPT_ID}] exportCourseManifest`, error);
      renderStatus("资源清单导出失败");
      notify(`资源清单导出失败：${String(error?.message || error)}`, true);
    } finally {
      runtime.exportRunning = false;
      updatePanel(context);
    }
  }

  async function fetchNodeResourceGroups(context, node) {
    const response = await context.api.getListNodeResourcesWithStatus({
      courseId: context.courseId,
      classId: context.classId,
      knowledgeId: node.nodeUid,
    });
    return extractResourceGroups(response);
  }

  async function fetchNodePpts(context, node) {
    const resourceGroups = await fetchNodeResourceGroups(context, node);
    const normalized = resourceGroups
      .map((item) => normalizeResource(item.resourcesDetail || item))
      .filter((item) => item.isPpt);

    return dedupeResources(normalized).sort((a, b) => {
      if (a.sorted !== b.sorted) {
        return a.sorted - b.sorted;
      }
      return a.resourcesName.localeCompare(b.resourcesName, "zh-CN");
    });
  }

  function extractResourceGroups(response) {
    if (Array.isArray(response?.data?.resourceList)) {
      return response.data.resourceList;
    }
    if (Array.isArray(response?.rt?.resourceList)) {
      return response.rt.resourceList;
    }
    if (Array.isArray(response?.resourceList)) {
      return response.resourceList;
    }
    if (Array.isArray(response?.rt?.resourcesDetail)) {
      return response.rt.resourcesDetail.map((item) => ({ resourcesDetail: item }));
    }
    if (Array.isArray(response?.data?.resourcesDetail)) {
      return response.data.resourcesDetail.map((item) => ({ resourcesDetail: item }));
    }
    return [];
  }

  function finishTask(task, context, message) {
    task.active = false;
    task.updatedAt = Date.now();
    saveTask(task);
    renderStatus(message);
    updatePanel(context, task);
    notify(message);
  }

  async function triggerDownload(url, filename) {
    if (!url) {
      throw new Error("缺少下载地址");
    }
    try {
      await pageFetchDownload(url, filename);
      return;
    } catch (error) {
      console.warn(`[${SCRIPT_ID}] pageFetchDownload failed`, error);
    }
    try {
      directOpenDownload(url);
      return;
    } catch (error) {
      console.warn(`[${SCRIPT_ID}] directOpenDownload failed`, error);
    }
    if (typeof GM_download === "function") {
      await new Promise((resolve, reject) => {
        GM_download({
          url,
          name: filename,
          saveAs: false,
          onload: resolve,
          onerror: reject,
          ontimeout: reject,
        });
      });
      return;
    }
    throw new Error("当前环境不支持下载");
  }

  async function pageFetchDownload(url, filename) {
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(`下载失败：${response.status}`);
    }
    const blob = await response.blob();
    const blobUrl = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = blobUrl;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(() => URL.revokeObjectURL(blobUrl), 3000);
  }

  function directOpenDownload(url) {
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.target = "_blank";
    anchor.rel = "noopener noreferrer";
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
  }

  function createTask(context) {
    clearTask(context.courseId, context.classId);
    return {
      version: TASK_VERSION,
      active: true,
      startedAt: Date.now(),
      updatedAt: Date.now(),
      courseId: context.courseId,
      classId: context.classId,
      mapName: context.mapName,
      // Some nodes report resourceCount = 0 but still return resources from the API.
      // Scan the full knowledge list to avoid missing edge-case PPTs.
      nodes: context.knowledgeNodes.slice(),
      currentIndex: 0,
      lastNodeUid: "",
      lastNodeName: context.currentNode?.knowledgeName || "",
      lastNodePptCount: context.currentPagePptResources.length,
      finishedNodeUids: {},
      downloadedResourceKeys: {},
      downloadedCount: 0,
      foundPptCount: 0,
      failedDownloads: [],
    };
  }

  function buildExportResourceRecord(group, resourceIndex) {
    const detail = group?.resourcesDetail || group || {};
    const normalized = normalizeResource(detail);
    return {
      index: resourceIndex,
      resourcesUid: normalized.resourcesUid,
      resourcesName: normalized.resourcesName,
      resourcesType: normalized.resourcesType,
      resourcesDataType: normalized.resourcesDataType,
      resourcesSuffix: normalized.resourcesSuffix,
      resourcesUrl: normalized.resourcesUrl,
      fileUrl: normalized.fileUrl,
      downloadUrl: normalized.downloadUrl,
      fileName: normalized.fileName,
      sorted: normalized.sorted,
      active: normalized.active,
      isPpt: normalized.isPpt,
      fromFileDomain: normalized.fromFileDomain,
      hasBookDetail: !!group?.resourcesBookDetail,
      hasVideoDetail: !!group?.resourcesVideoDetail,
      hasQuoteDetail: !!group?.resourcesQuoteDetail,
      quotePptPageSeqs: Array.isArray(group?.resourcesQuoteDetail?.pptPageSeqs)
        ? group.resourcesQuoteDetail.pptPageSeqs.slice()
        : [],
      rawDetail: pickResourceDetail(detail),
    };
  }

  function pickResourceDetail(detail) {
    return {
      id: Number(detail?.id || 0),
      resourcesUid: String(detail?.resourcesUid || ""),
      resourcesName: String(detail?.resourcesName || ""),
      resourcesType: Number(detail?.resourcesType || 0),
      resourcesDataType: Number(detail?.resourcesDataType || 0),
      resourcesDistributeType: Number(detail?.resourcesDistributeType || 0),
      resourcesLocalType: Number(detail?.resourcesLocalType || 0),
      resourcesFileId: String(detail?.resourcesFileId || ""),
      resourcesFileName: String(detail?.resourcesFileName || ""),
      resourcesCover: String(detail?.resourcesCover || ""),
      resourcesUrl: String(detail?.resourcesUrl || ""),
      resourcesSize: String(detail?.resourcesSize || ""),
      resourcesSuffix: String(detail?.resourcesSuffix || ""),
      sorted: Number(detail?.sorted || 0),
      specificSorted: Number(detail?.specificSorted || 0),
    };
  }

  function parseRoute() {
    const segments = String(location.pathname || "")
      .split("/")
      .filter(Boolean);
    
    // Legacy: /learnPage/{courseId}/{nodeUid}/{classId}
    if (segments[0] === "learnPage") {
      return {
        type: "legacy",
        courseId: String(segments[1] || ""),
        nodeUid: String(segments[2] || ""),
        classId: String(segments[3] || ""),
      };
    }
    
    // New: /singleCourse/knowledgeStudy/{courseId}/{classId}
    if (segments[0] === "singleCourse") {
      return {
        type: "new",
        courseId: String(segments[2] || ""),
        nodeUid: "", // Often missing in URL for new layout
        classId: String(segments[3] || ""),
      };
    }

    return { courseId: "", nodeUid: "", classId: "" };
  }

  function findCourseInstance(root) {
    return findComponent(root, (inst) => {
      const themeList = readInstanceValue(inst, "themeList");
      return Array.isArray(themeList) && themeList.length > 0;
    });
  }

  function findComponent(root, predicate) {
    let found = null;
    traverseComponents(root, (inst) => {
      if (predicate(inst)) {
        found = inst;
        return false;
      }
      return true;
    }, COMPONENT_MAX_DEPTH);
    return found;
  }

  function traverseComponents(root, visitor, maxDepth = COMPONENT_MAX_DEPTH, seen = new Set()) {
    const walkInst = (inst, depth) => {
      if (!inst || seen.has(inst) || depth > maxDepth) {
        return;
      }
      seen.add(inst);
      const shouldContinue = visitor(inst);
      if (shouldContinue === false) {
        return;
      }

      const childNodes = [];
      if (inst.subTree) {
        childNodes.push(inst.subTree);
      }
      if (inst.suspense?.activeBranch) {
        childNodes.push(inst.suspense.activeBranch);
      }
      if (inst.child) {
        childNodes.push(inst.child);
      }
      if (Array.isArray(inst.children)) {
        childNodes.push(...inst.children);
      }

      for (const child of childNodes) {
        walkNode(child, depth + 1);
      }
    };

    const walkNode = (node, depth) => {
      if (!node || depth > maxDepth) {
        return;
      }
      if (Array.isArray(node)) {
        node.forEach((child) => walkNode(child, depth));
        return;
      }
      if (node.component) {
        walkInst(node.component, depth);
      }
      if (node.suspense?.activeBranch) {
        walkNode(node.suspense.activeBranch, depth + 1);
      }
      if (Array.isArray(node.children)) {
        node.children.forEach((child) => walkNode(child, depth + 1));
      } else if (node.children && typeof node.children === "object") {
        walkNode(node.children, depth + 1);
      }
      if (Array.isArray(node.dynamicChildren)) {
        node.dynamicChildren.forEach((child) => walkNode(child, depth + 1));
      }
    };

    walkInst(root, 0);
  }

  function readInstanceValue(inst, key) {
    if (!inst) {
      return undefined;
    }
    const sources = [inst.setupState, inst.data, inst.ctx, inst.exposed];
    for (const source of sources) {
      if (source && Object.prototype.hasOwnProperty.call(source, key)) {
        return source[key];
      }
    }
    if (inst.proxy && key in inst.proxy) {
      return inst.proxy[key];
    }
    return undefined;
  }

  function flattenKnowledgeNodes(themeList) {
    const result = [];
    const seen = new Set();
    for (const theme of themeList || []) {
      for (const subTheme of theme.subThemeList || []) {
        for (const knowledge of subTheme.knowledgeList || []) {
          const nodeUid = String(knowledge.knowledgeId || knowledge.nodeUid || "");
          if (!nodeUid || seen.has(nodeUid)) {
            continue;
          }
          seen.add(nodeUid);
          result.push({
            themeId: String(theme.themeId || ""),
            themeName: String(theme.themeName || ""),
            subThemeId: String(subTheme.themeId || ""),
            subThemeName: String(subTheme.themeName || ""),
            knowledgeId: nodeUid,
            nodeUid,
            knowledgeName: String(knowledge.knowledgeName || ""),
            resourceCount: Number(knowledge.resourceCount || 0),
          });
        }
      }
    }
    return result;
  }

  function extractNodesFromLegacyTask(task) {
    if (!task || !Array.isArray(task.nodes)) {
      return [];
    }
    return task.nodes
      .map((item) => ({
        themeId: String(item.themeId || ""),
        themeName: String(item.themeName || ""),
        subThemeId: String(item.subThemeId || ""),
        subThemeName: String(item.subThemeName || ""),
        knowledgeId: String(item.knowledgeId || item.nodeUid || ""),
        nodeUid: String(item.nodeUid || item.knowledgeId || ""),
        knowledgeName: String(item.knowledgeName || ""),
        resourceCount: Number(item.resourceCount || 0),
      }))
      .filter((item) => item.nodeUid);
  }

  function collectCurrentPagePpts(root, previewData) {
    const items = [];
    traverseComponents(
      root,
      (inst) => {
        const cardData = getCardData(inst);
        if (cardData) {
          items.push(normalizeResource(cardData, { active: !!inst.props?.active }));
        }
        return true;
      },
      COMPONENT_MAX_DEPTH
    );

    if (previewData.resourcesUid && previewData.previewType === "ppt") {
      items.push(
        normalizeResource(previewData, {
          active: true,
          resourcesType: 1,
          resourcesDataType: 11,
        })
      );
    }

    return dedupeResources(items).filter((item) => item.isPpt);
  }

  function getCardData(inst) {
    const cardData = inst?.props?.cardData;
    if (!cardData || typeof cardData !== "object") {
      return null;
    }
    if (!cardData.resourcesUid && !cardData.resourcesName && !cardData.resourcesUrl && !cardData.fileUrl && !cardData.fileName) {
      return null;
    }
    return cardData;
  }

  function normalizeResource(card, extra = {}) {
    const resourcesUrl = String(card.resourcesUrl || "");
    const fileUrl = String(card.fileUrl || "");
    const downloadUrl = fileUrl || resourcesUrl;
    const fileName = String(card.fileName || "");
    const resourcesName = String(card.resourcesName || fileName || guessFilenameFromUrl(downloadUrl) || "未命名PPT");
    const suffixFromCard = String(card.resourcesSuffix || "").toLowerCase();
    const suffixFromUrl = getUrlSuffix(downloadUrl);
    const suffixFromName = getNameSuffix(fileName || resourcesName);
    const suffix = suffixFromCard || suffixFromUrl || suffixFromName;
    const resourcesType = Number(extra.resourcesType ?? card.resourcesType ?? 0);
    const resourcesDataType = Number(extra.resourcesDataType ?? card.resourcesDataType ?? 0);
    const fromFileDomain = /^https?:\/\/file\.zhihuishu\.com/i.test(downloadUrl);
    const hasPptSuffix = PPT_SUFFIXES.has(suffix);
    const isPpt =
      !!downloadUrl &&
      hasPptSuffix &&
      (fromFileDomain || resourcesType === 1) &&
      resourcesDataType !== 21 &&
      resourcesDataType !== 22;

    return {
      resourcesUid: String(card.resourcesUid || downloadUrl || resourcesName || ""),
      resourcesName,
      resourcesUrl,
      fileUrl,
      downloadUrl,
      fileName,
      resourcesType,
      resourcesDataType,
      resourcesSuffix: suffix,
      sorted: Number(card.sorted || 0),
      active: !!extra.active,
      fromFileDomain,
      isPpt,
    };
  }

  function dedupeResources(items) {
    const seen = new Set();
    const result = [];
    for (const item of items) {
      const key = `${item.resourcesUid}::${item.downloadUrl}`;
      if (!item.downloadUrl || seen.has(key)) {
        continue;
      }
      seen.add(key);
      result.push(item);
    }
    return result;
  }

  function buildCurrentFallback(nodeUid, previewData) {
    const headingText =
      String(document.querySelector("h1")?.textContent || "").trim() ||
      String(document.querySelector("h2")?.textContent || "").trim();
    const knowledgeName =
      String(previewData.nodeName || "").trim() ||
      headingText;
    if (!knowledgeName) {
      return null;
    }
    return {
      themeName: "",
      subThemeName: "",
      knowledgeName,
      nodeUid: String(nodeUid || ""),
    };
  }

  function simplifyPreview(preview) {
    return {
      mapName: String(preview.mapName || ""),
      nodeName: String(preview.nodeName || ""),
      previewType: String(preview.previewType || ""),
      resourcesName: String(preview.resourcesName || ""),
      resourcesUid: String(preview.resourcesUid || ""),
      resourcesUrl: String(preview.resourcesUrl || ""),
      fileUrl: String(preview.fileUrl || ""),
      fileName: String(preview.fileName || ""),
      resourcesSuffix: String(preview.resourcesSuffix || ""),
    };
  }

  function getTaskKey(courseId, classId) {
    return `${TASK_KEY_PREFIX}:${courseId}:${classId}`;
  }

  function loadAnyTask(courseId, classId) {
    try {
      const raw = localStorage.getItem(getTaskKey(courseId, classId));
      return raw ? JSON.parse(raw) : null;
    } catch (error) {
      return null;
    }
  }

  function loadTask(courseId, classId) {
    const task = loadAnyTask(courseId, classId);
    return task?.version === TASK_VERSION ? task : null;
  }

  function saveTask(task) {
    localStorage.setItem(getTaskKey(task.courseId, task.classId), JSON.stringify(task));
  }

  function clearTask(courseId, classId) {
    localStorage.removeItem(getTaskKey(courseId, classId));
  }

  function ensureUI() {
    if (document.getElementById(`${SCRIPT_ID}-panel`)) {
      return;
    }

    const panel = document.createElement("div");
    panel.id = `${SCRIPT_ID}-panel`;
    panel.innerHTML = `
      <div class="${SCRIPT_ID}-title">课程 PPT 批量下载</div>
      <div id="${SCRIPT_ID}-status" class="${SCRIPT_ID}-status">等待页面数据加载...</div>
      <div id="${SCRIPT_ID}-summary" class="${SCRIPT_ID}-summary"></div>
      <div class="${SCRIPT_ID}-actions">
        <button id="${SCRIPT_ID}-start" class="${SCRIPT_ID}-btn primary">从头遍历整门课 PPT</button>
        <button id="${SCRIPT_ID}-stop" class="${SCRIPT_ID}-btn">停止</button>
      </div>
      <div class="${SCRIPT_ID}-actions ${SCRIPT_ID}-actions-secondary">
        <button id="${SCRIPT_ID}-export" class="${SCRIPT_ID}-btn ${SCRIPT_ID}-btn-export">导出整门课资源 JSON</button>
      </div>
    `;
    document.body.appendChild(panel);
  }

  function updatePanel(context, task = null) {
    const summary = document.getElementById(`${SCRIPT_ID}-summary`);
    if (!summary || !context?.ok) {
      return;
    }

    const effectiveTask = task || loadTask(context.courseId, context.classId);
    const legacyTask = !effectiveTask ? loadAnyTask(context.courseId, context.classId) : null;
    const totalNodes = effectiveTask?.nodes?.length || context.knowledgeNodes.length;
    const finishedNodes = effectiveTask ? Object.keys(effectiveTask.finishedNodeUids || {}).length : 0;
    const downloadedCount = Number(effectiveTask?.downloadedCount || 0);
    const failedCount = Number(effectiveTask?.failedDownloads?.length || 0);
    const foundPptCount = Number(effectiveTask?.foundPptCount || 0);
    const lastNodeName = effectiveTask?.lastNodeName || context.currentNode?.knowledgeName || "";
    const lastNodePptCount = Number(effectiveTask?.lastNodePptCount ?? context.currentPagePptResources.length);
    summary.innerHTML = `
      <div>课程：${escapeHtml(context.mapName || "当前课程")}</div>
      <div>当前知识点：${escapeHtml(context.currentNode?.knowledgeName || "")}</div>
      <div>当前页 PPT 数：${context.currentPagePptResources.length}</div>
      <div>最近检索节点：${escapeHtml(lastNodeName)}</div>
      <div>该节点 PPT 数：${lastNodePptCount}</div>
      <div>遍历进度：${finishedNodes} / ${totalNodes}</div>
      <div>累计发现 PPT：${foundPptCount}</div>
      <div>已发起下载：${downloadedCount}</div>
      <div>失败数量：${failedCount}</div>
      <div>任务状态：${effectiveTask?.active ? "运行中" : "未运行"}</div>
      <div>调试导出：${runtime.exportRunning ? "导出中" : "就绪"}</div>
      <div>${legacyTask && legacyTask.version !== TASK_VERSION ? "旧缓存：已忽略，点击开始会重建任务" : "资源获取方式：直接调用课程内部 API"}</div>
    `;
  }

  function renderStatus(message) {
    const status = document.getElementById(`${SCRIPT_ID}-status`);
    if (status) {
      status.textContent = message;
    }
  }

  function notify(message, isError = false) {
    if (typeof GM_notification === "function") {
      GM_notification({
        title: isError ? "课程 PPT 批量下载器错误" : "课程 PPT 批量下载器",
        text: message,
        timeout: 2400,
      });
      return;
    }
    console[isError ? "error" : "log"](`[${SCRIPT_ID}] ${message}`);
  }

  function injectStyle() {
    if (document.getElementById(`${SCRIPT_ID}-style`)) {
      return;
    }
    const style = document.createElement("style");
    style.id = `${SCRIPT_ID}-style`;
    style.textContent = `
      #${SCRIPT_ID}-panel {
        position: fixed;
        right: 20px;
        bottom: 20px;
        z-index: 2147483646;
        width: 380px;
        border-radius: 16px;
        background: rgba(15, 23, 42, 0.96);
        color: #fff;
        box-shadow: 0 20px 56px rgba(15, 23, 42, 0.35);
        padding: 16px;
        font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
      }

      .${SCRIPT_ID}-title {
        font-size: 16px;
        font-weight: 700;
        margin-bottom: 10px;
      }

      .${SCRIPT_ID}-status,
      .${SCRIPT_ID}-summary {
        font-size: 12px;
        line-height: 1.7;
      }

      .${SCRIPT_ID}-status {
        color: #cbd5e1;
        margin-bottom: 10px;
      }

      .${SCRIPT_ID}-summary {
        display: flex;
        flex-direction: column;
        gap: 3px;
        margin-bottom: 12px;
      }

      .${SCRIPT_ID}-actions {
        display: flex;
        gap: 8px;
      }

      .${SCRIPT_ID}-actions-secondary {
        margin-top: 8px;
      }

      .${SCRIPT_ID}-btn {
        flex: 1;
        border: none;
        border-radius: 10px;
        padding: 10px 12px;
        font-size: 12px;
        font-weight: 700;
        cursor: pointer;
      }

      .${SCRIPT_ID}-btn.primary {
        background: #60a5fa;
        color: #0f172a;
      }

      .${SCRIPT_ID}-btn-export {
        background: #1e293b;
        color: #e2e8f0;
        border: 1px solid rgba(148, 163, 184, 0.35);
      }
    `;
    document.head.appendChild(style);
  }

  function sanitizeFilename(name) {
    return String(name || "downloaded_ppt").replace(/[<>:\"/\\\\|?*]+/g, "_").trim();
  }

  function buildPptFilename(node, resource) {
    const baseName = sanitizeFilename(
      [node?.themeName, node?.subThemeName, node?.knowledgeName, resource?.resourcesName]
        .filter(Boolean)
        .join(" - ")
    );
    if (/\.(ppt|pptx)$/i.test(baseName)) {
      return baseName;
    }
    return `${baseName}.${resource?.resourcesSuffix || "pptx"}`;
  }

  function buildManifestFilename(context) {
    const baseName = sanitizeFilename([context.mapName, "resource-manifest"].filter(Boolean).join(" - "));
    return `${baseName}.json`;
  }

  function buildResourceKey(node, resource) {
    return `${node?.nodeUid || ""}::${resource?.resourcesUid || ""}::${resource?.downloadUrl || ""}`;
  }

  function getUrlSuffix(url) {
    if (!url) {
      return "";
    }
    try {
      const pathname = new URL(url, location.href).pathname || "";
      const filename = pathname.split("/").pop() || "";
      const dotIndex = filename.lastIndexOf(".");
      return dotIndex >= 0 ? filename.slice(dotIndex + 1).toLowerCase() : "";
    } catch (error) {
      return "";
    }
  }

  function getNameSuffix(name) {
    const value = String(name || "").trim();
    const dotIndex = value.lastIndexOf(".");
    return dotIndex >= 0 ? value.slice(dotIndex + 1).toLowerCase() : "";
  }

  function guessFilenameFromUrl(url) {
    try {
      const pathname = new URL(url, location.href).pathname || "";
      return decodeURIComponent(pathname.split("/").pop() || "");
    } catch (error) {
      return "";
    }
  }

  function downloadTextFile(filename, content, mimeType = "text/plain;charset=utf-8") {
    const blob = new Blob([content], { type: mimeType });
    const blobUrl = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = blobUrl;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(() => URL.revokeObjectURL(blobUrl), 3000);
  }

  function escapeHtml(value) {
    return String(value || "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll("\"", "&quot;")
      .replaceAll("'", "&#39;");
  }

  function sleep(ms) {
    return new Promise((resolve) => window.setTimeout(resolve, ms));
  }
})();
