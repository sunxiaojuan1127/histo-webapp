/* =========================================================================
 * app.js — 分组直方图分析 · 网页版 UI 层
 * 依赖: lib/xlsx.full.min.js, lib/plotly.min.js, histo-core.js
 * ========================================================================= */
(function () {
  "use strict";

  var C = window.HistoCore;

  var PALETTES = {
    "清爽蓝 (默认)": ["#506EF5", "#F5A623", "#2FA84F", "#E8536B",
                     "#8E6BE8", "#00A6B2", "#D96A2B", "#6B7B8C"],
    "科学经典": ["#1F77B4", "#FF7F0E", "#2CA02C", "#D62728",
                 "#9467BD", "#8C564B", "#E377C2", "#7F7F7F"],
    "柔和马卡龙": ["#7EB8DA", "#F6B99D", "#A8D8A0", "#E8A0B4",
                   "#B8A8D8", "#8CC5C5", "#D8C088", "#A0A8B8"],
    "高对比": ["#0D3B66", "#F95738", "#00B4A0", "#FFC145",
               "#721E87", "#00A6ED", "#D7263D", "#1B998B"],
  };

  function $(id) { return document.getElementById(id); }
  function radioVal(name) {
    var el = document.querySelector('input[name="' + name + '"]:checked');
    return el ? el.value : null;
  }
  function fmtCount(x) { return x.toLocaleString("en-US"); }
  function num(input) {
    if (!input || input.value === "" || input.value == null) return null;
    var n = parseFloat(input.value);
    return Number.isFinite(n) ? n : null;
  }
  var toastTimer = null;
  function toast(msg, ms) {
    var t = $("toast");
    t.textContent = msg;
    t.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.classList.remove("show"); }, ms || 3200);
  }

  var S = {
    rows: null, headers: null, labels: null, counts: null,
    rowIdx: null, colIdx: null,
    metricIdxs: [], metricIdx: -1,
    sorted: [], selected: [],
    name: "", sheet: "", nGroups: 0,
    stepTouched: false,
    groupRenameMap: {}, headerRenameMap: {},
    plotTitle: "", plotTitleEdited: false,
    binSegments: [],
  };
  var renderTimer = null;

  /* ================= 工具: 名称重命名映射 ================= */
  function parseRenameMap(text) {    /* 解析 "原名=新名" 格式, 每行一条; 跳过空行与注释行(#).
       返回对象. 同一原名多次出现以最后一次为准.   */
    var m = {};
    if (!text) return m;
    String(text).split(/\r?\n/).forEach(function (line) {
      var t = line.replace(/^\s+|\s+$/g, "");
      if (!t || t.charAt(0) === "#") return;
      var eq = t.indexOf("=");
      if (eq < 0) return;
      var k = t.slice(0, eq).replace(/^\s+|\s+$/g, "");
      var v = t.slice(eq + 1).replace(/^\s+|\s+$/g, "");
      if (k) m[k] = v;
    });
    return m;
  }
  function applyRename(name, map) {
    if (!map || !name) return name;
    var v = map[name];
    /* 空值/空白值视为无映射, 仍返回原名 (方便 diffRename 写出 "E-8=" 然后用户清空即可重置) */
    if (v == null || String(v).replace(/\s+/g, "") === "") return name;
    return v;
  }
  function diffRename(origList, curMap) {
    /* 生成 "原名=新名" 文本, 新名=空 表示删除映射, 永远保存全部原列表. */
    return origList.map(function (k) { return k + "=" + (curMap[k] || ""); }).join("\n");
  }

  /* ================= 工具: 分段步长解析/格式化 ================= */
  function _fmtNum(x) {
    /* 输出简洁数字: 不强制科学计数, 自动去尾零 */
    if (Math.abs(x) < 1e-12) return "0";
    var s = (+(+x).toPrecision(8)).toString();
    return s;
  }
  function parseBinSegments(text) {
    /* 解析 "起始 结束 步长" 每行一段; 空格/逗号/制表符 都行; 空行与 # 注释跳过 */
    var segs = [];
    if (!text) return segs;
    String(text).split(/\r?\n/).forEach(function (line) {
      var t = line.replace(/^\s+|\s+$/g, "");
      if (!t || t.charAt(0) === "#") return;
      var parts = t.split(/[\s,]+/).filter(Boolean);
      if (parts.length < 3) return;
      var a = parseFloat(parts[0]);
      var b = parseFloat(parts[1]);
      var w = parseFloat(parts[2]);
      if (!Number.isFinite(a) || !Number.isFinite(b) || !Number.isFinite(w)) return;
      if (!(w > 0) || !(b > a)) return;
      segs.push({ start: a, end: b, step: w });
    });
    segs.sort(function (x, y) { return x.start - y.start; });
    return segs;
  }
  function formatBinSegments(segs) {
    if (!Array.isArray(segs) || !segs.length) return "";
    return segs.map(function (s) {
      return _fmtNum(s.start) + "  " + _fmtNum(s.end) + "  " + _fmtNum(s.step);
    }).join("\n");
  }
  var SEG_TEMPLATES = {
    /* 用户示例 (R13 需求里提到的): 1-2 用 0.2, 2-20 用 2 */
    "user": "1  2  0.2\n2  20  2",
    "fine_low": "0.01  0.1  0.01\n0.1  1  0.1\n1  10  1\n10  100  10",
    "fluo_0_10": "0  0.5  0.05\n0.5  2  0.1\n2  5  0.5\n5  10  1",
    "wide_0_100": "0  1  0.1\n1  5  0.5\n5  20  1\n20  100  5",
    "log3_3": "0.001  0.01  0.001\n0.01  0.1  0.01\n0.1  1  0.1\n1  10  1\n10  100  10\n100  1000  100",
  };

  /* ================= 控件可见性 ================= */
  function updateControlVisibility() {
    var kind = $("kindSelect").value;
    var isHist = kind.indexOf("hist") === 0;
    var facet = $("facetCb").checked;
    $("binSec").style.display = isHist ? "" : "none";
    $("curveSec").style.display = isHist ? "" : "none";
    $("facetOpts").style.display = (facet && isHist) ? "" : "none";
    $("axLabel").textContent = isHist ? "数值轴范围（X 轴）" : "数值轴范围（Y 轴）";
    $("axLogCb").innerHTML = (isHist ? "X 轴用对数刻度" : "Y 轴用对数刻度") +
      ' <span class="hint">（值跨数量级如 0.2~60 时勾上）</span>';
    $("ySec").style.display = isHist ? "" : "none";

    var binMode = radioVal("binMode");
    $("stepRow").style.display = binMode === "step" ? "" : "none";
    $("piecesRow").style.display = binMode === "pieces" ? "" : "none";
    $("countRow").style.display = binMode === "count" ? "" : "none";
    $("smoothBw").disabled = radioVal("showMode") === "bars";

    var axMode = radioVal("axMode");
    $("clipRow").style.display = axMode === "clip" ? "" : "none";
    $("manualRow").style.display = axMode === "manual" ? "" : "none";
    $("colorRow").style.display = $("sameColorCb").checked ? "" : "none";

    /* 分段步长叠加提示: 仅当选了"分段步长" + 多组 + 还未启用 Facet 时显示 */
    var piecesOn = radioVal("binMode") === "pieces";
    var multiGroup = (S.selected && S.selected.length >= 2);
    var facetOn = $("facetCb").checked;
    var overlayHint = $("piecesOverlayHint");
    if (overlayHint) overlayHint.style.display = (piecesOn && multiGroup && !facetOn) ? "" : "none";

    /* Facet 裁切滑杆: 仅在 Facet 勾选 + 子图模式显示 */
    var facetClipRow = $("facetClipRow");
    if (facetClipRow) facetClipRow.style.display = (facetOn && piecesOn === false) ? "" : "none";
  }

  /* ================= 数据加载 ================= */
  function loadFile(file) {
    var rd = new FileReader();
    rd.onload = function (e) {
      try { loadWorkbook(e.target.result, file.name); }
      catch (err) { toast("读取失败: " + err.message); }
    };
    rd.readAsArrayBuffer(file);
  }
  function loadSample() {
    fetch("sample_grouped.xlsx")
      .then(function (r) { if (!r.ok) throw new Error("示例文件下载失败"); return r.arrayBuffer(); })
      .then(function (buf) { loadWorkbook(buf, "sample_grouped.xlsx"); })
      .catch(function (e) { toast(e.message); });
  }
  function loadWorkbook(buffer, fname) {
    var t0 = performance.now();
    var wb = C.parseWorkbook(window.XLSX, buffer);
    var ms = Math.round(performance.now() - t0);
    if (wb.headers.length < 2) throw new Error("表头少于 2 列，无法分析");
    S.rows = wb.rows; S.headers = wb.headers;
    S.name = fname; S.sheet = wb.sheetName; S.stepTouched = false;

    var gi = C.findGroupIdx(S.headers);
    if (gi.rowIdx !== null && gi.colIdx !== null) {
      S.rowIdx = gi.rowIdx; S.colIdx = gi.colIdx;
      $("manualGroupSec").style.display = "none";
    } else {
      S.rowIdx = null; S.colIdx = null;
      $("manualGroupSec").style.display = "";
      fillManualSelect($("rowSelect"));
      fillManualSelect($("colSelect"));
    }
    buildGroups();

    $("fileState").textContent = "已加载: " + fname;
    $("fileState").className = "badge";
    $("fileInfo").textContent = "✔ " + fname + " · 工作表「" + S.sheet + "」· " +
      fmtCount(S.rows.length) + " 行 × " + S.headers.length + " 列 · 解析 " + ms + "ms";
    $("emptyCard").style.display = "none";
    $("chartCard").style.display = "";
    updateControlVisibility();
    scheduleRender();
  }
  function fillManualSelect(sel) {
    sel.innerHTML = "";
    var ph = document.createElement("option");
    ph.value = ""; ph.textContent = "（请选择）";
    sel.appendChild(ph);
    S.headers.forEach(function (h, i) {
      var o = document.createElement("option");
      o.value = i; o.textContent = h || ("（无表头列 " + (i + 1) + "）");
      sel.appendChild(o);
    });
  }
  function applyManualGrouping() {
    var r = parseInt($("rowSelect").value, 10);
    var c = parseInt($("colSelect").value, 10);
    if (isNaN(r) || isNaN(c) || r === c) return;
    S.rowIdx = r; S.colIdx = c;
    buildGroups();
    scheduleRender();
  }

  function buildGroups() {
    if (S.rowIdx === null || S.colIdx === null) return;
    S.labels = C.buildLabels(S.rows, S.rowIdx, S.colIdx);
    var cm = C.groupCounts(S.labels);
    S.counts = cm;
    S.sorted = C.sortLabels(Object.keys(cm));
    S.nGroups = S.sorted.length;

    var excl = {};
    excl[S.rowIdx] = 1; excl[S.colIdx] = 1;
    S.metricIdxs = C.numericMetricIdxs(S.rows, S.headers, excl);
    if (!S.metricIdxs.length) toast("没有找到任何数值列，无法绘制");
    S.metricIdx = S.metricIdxs.length ? S.metricIdxs[0] : -1;
    fillMetricSelect();
    S.selected = S.sorted.slice();
    renderGroupList();
    syncRenameTextarea();
    syncHdrRenameTextarea();
    $("groupCount").textContent = "共 " + S.nGroups + " 组";
  }
  function fillMetricSelect() {
    var sel = $("metricSelect");
    sel.innerHTML = "";
    S.metricIdxs.forEach(function (ci) {
      var o = document.createElement("option");
      o.value = ci;
      var orig = S.headers[ci];
      var renamed = applyRename(orig, S.headerRenameMap);
      o.textContent = renamed !== orig ? renamed + "  (" + orig + ")" : orig;
      sel.appendChild(o);
    });
    if (S.metricIdx >= 0) sel.value = String(S.metricIdx);
    /* 同步指标列的 "原名=新名" 文本到 textarea, 当已存在的改名随填充一起呈现 */
    syncHdrRenameTextarea();
  }

  /* ================= 分组列表 ================= */
  function renderGroupList() {
    var box = $("groupList");
    box.innerHTML = "";
    var kw = $("groupSearch").value.trim().toLowerCase();
    /* 支持模糊匹配: 输入"对照"也能匹配 "对照组A" (应用改名后) */
    var kwRenamed = kw;
    S.sorted.forEach(function (g) {
      var disp = applyRename(g, S.groupRenameMap);
      var kwMatch = !kw ||
        g.toLowerCase().indexOf(kw) >= 0 ||
        (disp !== g && disp.toLowerCase().indexOf(kw) >= 0);
      if (!kwMatch) return;
      var label = document.createElement("label");
      label.className = "gi";
      label._groupKey = g;
      var cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = S.selected.indexOf(g) >= 0;
      var span = document.createElement("span");
      span.textContent = disp !== g ? (disp + "  (" + g + ")") : g;
      var n = document.createElement("span");
      n.className = "n";
      n.textContent = "n=" + fmtCount(S.counts[g]);
      label.appendChild(cb); label.appendChild(span); label.appendChild(n);
      box.appendChild(label);
      cb.addEventListener("change", function () {
        var i2 = S.selected.indexOf(g);
        if (cb.checked && i2 < 0) S.selected.push(g);
        else if (!cb.checked && i2 >= 0) S.selected.splice(i2, 1);
        S.stepTouched = true;
        scheduleRender();
      });
    });
  }
  function selectAllOn() { S.selected = S.sorted.slice(); renderGroupList(); scheduleRender(); }
  function selectAllOff() { S.selected = []; renderGroupList(); scheduleRender(); }

  /* renderGroupList 之后同步显示文本（不重渲染 checkbox 状态, 只更新 label） */
  function refreshGroupListLabels() {
    var box = $("groupList");
    var kw = $("groupSearch").value.trim().toLowerCase();
    var visibleIdx = 0;
    Array.from(box.children).forEach(function (lab) {
      if (lab.className !== "gi") return;
      /* 找该 label 对应的原 group */
      var g = lab._groupKey;
      if (!g) return;
      var disp = applyRename(g, S.groupRenameMap);
      var span = lab.querySelector("span");
      if (span && !span.classList.contains("n")) {
        span.textContent = disp !== g ? (disp + "  (" + g + ")") : g;
      }
    });
  }
  /* 把每个 DOM label 与对应 group key 关联（renderGroupList 里加） */
  /* 同步 "原名=新名" 文本到 textarea，用于一开面板就看到所有可选原名 */
  function syncRenameTextarea() {
    if (S.sorted && S.sorted.length) {
      $("renameText").value = diffRename(S.sorted, S.groupRenameMap);
    }
  }
  function syncHdrRenameTextarea() {
    if (S.headers && S.metricIdxs && S.metricIdxs.length) {
      var candidates = S.metricIdxs.map(function (ci) { return S.headers[ci]; });
      $("hdrRenameText").value = diffRename(candidates, S.headerRenameMap);
    }
  }

  /* ================= 配置 ================= */
  function readAxisCfg(isHist) {
    var mode = radioVal("axMode");
    return {
      mode: mode,
      clipPct: mode === "clip" ? (parseFloat($("clipPct").value) || 1) : 0,
      min: mode === "manual" ? num($("axMin")) : null,
      max: mode === "manual" ? num($("axMax")) : null,
      log: $("axLogCb").checked,
      yMin: isHist ? num($("yMin")) : null,
      yMax: isHist ? num($("yMax")) : null,
      yLog: isHist ? $("yLogCb").checked : false,
    };
  }
  function defaultPlotTitle(cfg) {
    var extra = cfg.facet
      ? "（子图模式" + (cfg.isHist && cfg.binMode === "pieces" ? " · 分段步长" : "") + "）"
      : (cfg.isHist && cfg.binMode === "pieces" ? "（分段步长）" : "");
    return cfg.metricName + " — 按 Row-Column 分组" + extra;
  }
  function readCfg() {
    var kind = $("kindSelect").value;
    var isHist = kind.indexOf("hist") === 0;
    var metricName = (S.metricIdx >= 0) ? applyRename(S.headers[S.metricIdx], S.headerRenameMap) : "";
    var cfg = {
      kind: kind, isHist: isHist,
      metricName: metricName,
      binMode: radioVal("binMode"),
      binStep: num($("binStep")) || 0.1,
      nBins: parseInt($("binCount").value, 10) || 50,
      binSegments: parseBinSegments($("binSegText").value),
      showMode: isHist ? radioVal("showMode") : "bars",
      bw: parseFloat($("smoothBw").value) || 1,
      smoothLevel: parseFloat($("curveSmoothLevel").value) || 30,
      stack: radioVal("stackMode") === "stack",
      facet: $("facetCb").checked,
      facetAuto: $("facetAutoStepCb").checked,
      facetClipAuto: $("facetClipAutoCb").checked,
      facetClipLo: parseFloat($("facetClipLo").value) || 0.5,
      facetClipHi: parseFloat($("facetClipHi").value) || 99.5,
      facetCols: Math.max(1, Math.min(4, parseInt(radioVal("facetCols"), 10) || 3)),
      facetH: parseInt($("facetH").value, 10) || 250,
      axis: readAxisCfg(isHist),
      palette: PALETTES[$("paletteSelect").value] || PALETTES["清爽蓝 (默认)"],
      sameColor: $("sameColorCb").checked,
      customColor: $("colorPick").value,
      showStat: $("statCb").checked,
    };
    /* 标题: 用户编辑过就用用户输入, 否则用默认 */
    cfg.titleDefault = defaultPlotTitle(cfg);
    cfg.title = (S.plotTitleEdited && S.plotTitle !== "") ? S.plotTitle : cfg.titleDefault;
    return cfg;
  }

  /* 合并多组数值、算轴范围（raw / clip / manual） */
  function concatAll(entries) {
    var a = [];
    entries.forEach(function (e) { for (var i = 0; i < e.vals.length; i++) a.push(e.vals[i]); });
    return a;
  }
  function axisRange(allV, ax) {
    var mn = Infinity, mx = -Infinity, i;
    for (i = 0; i < allV.length; i++) {
      if (allV[i] < mn) mn = allV[i];
      if (allV[i] > mx) mx = allV[i];
    }
    if (!allV.length) return { vmin: 0, vmax: 1, rawMin: 0, rawMax: 1 };
    var vmin = mn, vmax = mx;
    if (ax.clipPct > 0 && allV.length > 20) {
      vmin = C.percentile(allV, ax.clipPct);
      vmax = C.percentile(allV, 100 - ax.clipPct);
    }
    if (ax.min != null) vmin = ax.min;
    if (ax.max != null) vmax = ax.max;
    if (!(vmin < vmax)) { vmin -= 0.5; vmax += 0.5; }
    return { vmin: vmin, vmax: vmax, rawMin: mn, rawMax: mx };
  }

  /* 只保留 >0 的数值（用于对数轴） */
  function posFilter(e) {
    return { label: e.label, vals: e.vals.filter(function (v) { return v > 0; }) };
  }

  /* ================= 直方图 trace ================= */
  function addHistTraces(data, e, color, vmin, vmax, step, kind, showMode, bw,
                         isFacet, xaxis, yaxis) {
    var vals = e.vals;
    var ntxt = fmtCount(vals.length);
    var tr = {
      type: "histogram", x: vals,
      name: e.label + " (n=" + ntxt + ")",
      legendgroup: e.label,
      marker: { color: color },
      opacity: 0.72,
      hovertemplate: "<b>" + e.label + "</b><br>区间: %{x}<br>" +
        (kind === "hist_freq" ? "频数: " : "占比: ") + "%{y}<br>" +
        "n_total=" + ntxt + "<extra></extra>",
    };
    if (kind === "hist_percent" || kind === "hist_cum") tr.histnorm = "percent";
    if (kind === "hist_cum") tr.cumulative = { enabled: true };
    tr.xbins = { start: vmin, end: vmax, size: Math.max(step, 1e-9) };
    if (showMode === "curve") tr.visible = false;
    if (isFacet) { if (xaxis) tr.xaxis = xaxis; if (yaxis) tr.yaxis = yaxis; }
    data.push(tr);

    if (showMode === "both" || showMode === "curve") {
      var kv = C.kdeCurveForHist(vals, vmin, vmax, step, kind, bw);
      if (kv) {
        var sc = {
          type: "scatter", mode: "lines", x: kv.x, y: kv.y,
          name: showMode === "curve"
            ? e.label + " (n=" + ntxt + ") 曲线"
            : e.label + " 平滑曲线",
          legendgroup: e.label,
          showlegend: showMode === "curve",
          line: { color: color, width: showMode === "curve" ? 4 : 3 },
          hovertemplate: "%{x:.4g}<br>%{y:.4g}<extra></extra>",
        };
        if (isFacet) { if (xaxis) sc.xaxis = xaxis; if (yaxis) sc.yaxis = yaxis; }
        data.push(sc);
      }
    }
  }

  /* ================= 不等宽分箱: Bar trace ================= */
  /* 与 addHistTraces 同语义，但用 edges 数组 + Bar trace 渲染非均匀柱宽 */
  function addPiecewiseBars(data, e, color, edges, kind, showMode, bw,
                            isFacet, xaxis, yaxis, smoothLevel) {
    var vals = e.vals;
    var ntxt = fmtCount(vals.length);
    var hb = C.histogramByEdges(vals, edges, kind);
    if (!hb) return;
    /* 曲线原始点 = bin 中心 (hb.x) + 柱高 (hb.y) */
    var rawPoints = hb.x.map(function (xv, i) { return { x: xv, y: hb.y[i] }; });
    /* smoothLevel: 0~100, 0=折线 (贴柱顶), 100=Catmull-Rom 最大插值 */
    var smoothStep = Math.max(1, Math.round(smoothLevel / 100 * 24) + 1);
    var curvePts = C.catmullRomSpline(rawPoints, smoothStep);
    var curveX = curvePts.map(function (p) { return p.x; });
    var curveY = curvePts.map(function (p) { return p.y; });
    var showMarkers = smoothStep <= 1;
    var tr = {
      type: "bar",
      x: hb.x,
      y: hb.y,
      width: hb.widths,
      name: e.label + " (n=" + ntxt + ")",
      legendgroup: e.label,
      marker: { color: color, line: { width: 0 } },
      opacity: 0.78,
      hovertemplate: "<b>" + e.label + "</b><br>区间: %{x}<br>" +
        (kind === "hist_freq" ? "频数: " : "占比: ") + "%{y}<br>" +
        "n_total=" + ntxt + "<extra></extra>",
    };
    if (showMode === "curve") tr.visible = false;
    if (isFacet) { if (xaxis) tr.xaxis = xaxis; if (yaxis) tr.yaxis = yaxis; }
    data.push(tr);

    if (showMode === "both" || showMode === "curve") {
      /* piecewise 模式曲线: 用 Catmull-Rom 插值 (smoothLevel 控制插值密度),
       * 0=折线贴柱顶, 100=最平滑; 默认 30 让肉眼看起来既贴柱又有平滑感. */
      var sc = {
        type: "scatter",
        mode: showMarkers ? "lines+markers" : "lines",
        x: curveX, y: curveY,
        name: showMode === "curve"
          ? e.label + " (n=" + ntxt + ") 曲线"
          : e.label + " 曲线",
        legendgroup: e.label,
        showlegend: showMode === "curve",
        line: { color: color, width: showMode === "curve" ? 3.5 : 2.5, shape: "linear" },
        marker: showMarkers
          ? { color: color, size: showMode === "curve" ? 6 : 4,
              line: { color: "#fff", width: 1 } }
          : undefined,
        connectgaps: false,
        hovertemplate: "<b>" + e.label + "</b><br>x=%{x:.4g}<br>" +
          (kind === "hist_freq" ? "频数: " : "占比: ") + "%{y:.3g}<br>" +
          "n_total=" + ntxt + "<extra></extra>",
      };
      if (isFacet) { if (xaxis) sc.xaxis = xaxis; if (yaxis) sc.yaxis = yaxis; }
      data.push(sc);
    }
  }

  /* ================= 主绘图 ================= */
  function buildPlot(entries, cfg) {
    var kind = cfg.kind, isHist = cfg.isHist;
    var valid = entries.filter(function (e) { return e.vals.length > 0; });
    if (!valid.length) return { data: [], layout: null, error: "选中的分组没有数值" };

    var allV = concatAll(valid);
    var ax = axisRange(allV, cfg.axis);

    var xLog = isHist && cfg.axis.log && ax.vmin > 0;
    var yLogNum = (!isHist) && cfg.axis.log && ax.vmin > 0;   // violin/box 数值轴对数
    var yLog = isHist && cfg.axis.yLog;                        // hist 柱高轴对数

    // 对数轴需要剔除 <=0 的数据
    var plotVals = valid;
    if (xLog || yLogNum) {
      plotVals = valid.map(posFilter).filter(function (e) { return e.vals.length > 0; });
      if (!plotVals.length)
        return { data: [], layout: null, error: "对数刻度下没有 >0 的数值。请关闭对数刻度，或检查数据是否含 0/负数。" };
      ax = axisRange(concatAll(plotVals), cfg.axis);
      if (ax.vmin <= 0) ax.vmin = Math.max(ax.vmin, 1e-9);
    }

    var colors = plotVals.map(function (_, i) {
      return cfg.sameColor ? cfg.customColor : cfg.palette[i % cfg.palette.length];
    });

    var yTitleByKind = {
      hist_freq: "频数（细胞个数）", hist_percent: "百分比 (%)",
      hist_cum: "累计百分比 (%)", violin: cfg.metricName, box: cfg.metricName,
    };
    var yTitle = yTitleByKind[kind];

    // ------- 步长 -------
    var globalStep = null, perGroupStep = null;
    var globalEdges = null, perGroupEdges = null;  /* 分段步长模式 */
    if (isHist) {
      if (cfg.binMode === "step") {
        globalStep = Math.max(cfg.binStep, 1e-9);
      } else if (cfg.binMode === "count") {
        globalStep = (ax.vmax - ax.vmin) / Math.max(cfg.nBins, 1);
      } else if (cfg.binMode === "pieces") {
        globalEdges = C.piecewiseBinEdges(allV, cfg.binSegments);
        if (globalEdges) {
          /* 用第一段宽度作 KDE 窗口 (兜底) */
          globalStep = globalEdges[1] - globalEdges[0];
        } else {
          toast("分段步长无效或无段，已自动改用 IQR 推荐步长");
          globalStep = C.suggestStep(allV);
        }
      } else {
        globalStep = C.suggestStep(allV);
      }
      if (cfg.facet && cfg.facetAuto && cfg.binMode !== "step" && cfg.binMode !== "pieces") {
        perGroupStep = plotVals.map(function (e) { return C.suggestStep(e.vals); });
      }
      if (cfg.facet && cfg.facetAuto && cfg.binMode === "pieces" && globalEdges) {
        perGroupEdges = plotVals.map(function (e) {
          return C.piecewiseBinEdges(e.vals, cfg.binSegments) || globalEdges;
        });
      }
    }

    var data = [], layout = {};

    /* ------------ Y 轴范围 helper：只填入有的端点，null 交给 plotly autorange ------------ */
    function setRange(axisObj, lo, hi, logFlag) {
      if (lo != null && logFlag) lo = Math.log10(Math.max(lo, 1e-9));
      if (hi != null && logFlag) hi = Math.log10(Math.max(hi, 1e-9));
      if (lo != null && hi != null && lo >= hi) return;
      if (lo == null && hi == null) return;
      axisObj.range = [lo != null ? lo : null, hi != null ? hi : null];
    }

    if (!cfg.facet) {
      /* ============ 叠加模式 ============ */
      var xaxis = { showgrid: true, gridcolor: "#eee" };
      var yaxis = { showgrid: true, gridcolor: "#eee" };
      if (isHist) {
        xaxis.title = cfg.metricName + (xLog ? " (log 刻度)" : "");
        xaxis.range = [ax.vmin, ax.vmax];
        if (xLog) { xaxis.type = "log"; xaxis.range = [Math.max(ax.vmin, 1e-9), ax.vmax]; }
        yaxis.title = yTitle;
        if (kind === "hist_percent" || kind === "hist_cum") yaxis.ticksuffix = "%";
        if (yLog) yaxis.type = "log";
        setRange(yaxis, cfg.axis.yMin, cfg.axis.yMax, yLog);

        for (var i = 0; i < plotVals.length; i++) {
          if (cfg.binMode === "pieces" && globalEdges) {
            addPiecewiseBars(data, plotVals[i], colors[i], globalEdges,
                             kind, cfg.showMode, cfg.bw, false, "", "",
                             cfg.smoothLevel);
          } else {
            addHistTraces(data, plotVals[i], colors[i], ax.vmin, ax.vmax,
                          globalStep, kind, cfg.showMode, cfg.bw, false);
          }
        }
      } else {
        xaxis.title = "分组 (Row-Column)";
        xaxis.showgrid = false;
        yaxis.title = cfg.metricName;
        if (yLogNum) yaxis.type = "log";
        // 数值轴范围（auto / clip / manual / 手动端点）
        var yLo = null, yHi = null;
        if (cfg.axis.mode === "clip") { yLo = ax.vmin; yHi = ax.vmax; }
        else if (cfg.axis.mode === "manual") { yLo = cfg.axis.min; yHi = cfg.axis.max; }
        setRange(yaxis, yLo, yHi, yLogNum);

        for (var j = 0; j < plotVals.length; j++) {
          var e = plotVals[j], color = colors[j];
          if (kind === "violin") {
            var vt = { type: "violin", y: e.vals, name: e.label, legendgroup: e.label,
                       line: { color: color }, fillcolor: C.hexToRgba(color, 0.65),
                       opacity: 0.9, points: false, box: { visible: true },
                       meanline: { visible: true },
                       hovertemplate: "<b>" + e.label + "</b><br>数值: %{y:.4g}<extra></extra>" };
            if (plotVals.length === 1) vt.side = "positive";
            data.push(vt);
          } else {
            data.push({ type: "box", y: e.vals, name: e.label, legendgroup: e.label,
                        marker: { color: color }, fillcolor: C.hexToRgba(color, 0.5),
                        line: { color: color }, boxpoints: false,
                        hovertemplate: "<b>" + e.label + "</b><br>数值: %{y:.4g}<extra></extra>" });
          }
        }
      }
      layout.xaxis = xaxis; layout.yaxis = yaxis;
      layout.barmode = isHist ? (cfg.stack ? "stack" : "overlay") : "group";
      layout.bargap = isHist ? 0.03 : 0.25;
      layout.margin = { l: 66, r: 22, t: 78, b: 60 };

      if (isHist) {
        var st;
        if (cfg.binMode === "step") st = "步长 = " + globalStep.toPrecision(4);
        else if (cfg.binMode === "count") st = "柱数 = " + cfg.nBins + "（≈步长 " + globalStep.toPrecision(4) + "）";
        else if (cfg.binMode === "pieces") st = "分段步长: " + cfg.binSegments.map(function (s) {
          return "[" + _fmtNum(s.start) + "~" + _fmtNum(s.end) + "] 每 " + _fmtNum(s.step);
        }).join(" · ");
        else st = "自动步长 (IQR 法则) = " + globalStep.toPrecision(4);
        layout.annotations = [{
          text: st, xref: "paper", yref: "paper", x: 1, y: 1.14,
          xanchor: "right", showarrow: false, font: { size: 11.5, color: "#506EF5" },
        }];
      }
    } else {
      /* ============ Facet 子图模式 ============ */
      var n = plotVals.length;
      var ncols = Math.min(cfg.facetCols, n);
      var nrows = Math.ceil(n / ncols);
      layout.grid = { rows: nrows, columns: ncols, pattern: "independent",
                      roworder: "top to bottom" };
      layout.barmode = isHist ? (cfg.stack ? "stack" : "overlay") : "group";
      layout.bargap = isHist ? 0.03 : 0.25;
      layout.height = Math.max(340, nrows * cfg.facetH + 130);
      layout.margin = { l: 64, r: 18, t: 96, b: 46 };
      var anns = [];

      /* Facet 横坐标显示策略:
       * 每张子图都显示 X 轴 tick 数字 (用户在子图里读数);
       * 分组名/n 用 plotly annotation 表示, 避免与 X 轴 title 冲突.   */
      var lastRowStart = (nrows - 1) * ncols;
      var anns = (cfg.facet && perGroupStep)
        ? [{ text: "每组步长 = " + perGroupStep.map(function (s) { return s.toPrecision(4); }).join(" / "),
             xref: "paper", yref: "paper",
             x: 1, y: 1.06, xanchor: "right", showarrow: false,
             font: { size: 10.5, color: "#506EF5" } }]
        : [];
      if (cfg.facet && perGroupEdges) {
        anns.push({
          text: "分段步长: " + cfg.binSegments.map(function (s) {
            return "[" + _fmtNum(s.start) + "~" + _fmtNum(s.end) + "]×" + _fmtNum(s.step);
          }).join(" · "),
          xref: "paper", yref: "paper", x: 1, y: 1.06,
          xanchor: "right", showarrow: false, font: { size: 10.5, color: "#506EF5" },
        });
      }
      for (var k = 0; k < n; k++) {
        var ek = plotVals[k], ck = colors[k];
        var xaKey = k === 0 ? "xaxis" : "xaxis" + (k + 1);
        var yaKey = k === 0 ? "yaxis" : "yaxis" + (k + 1);
        var refX = k === 0 ? "x" : "x" + (k + 1);
        var refY = k === 0 ? "y" : "y" + (k + 1);
        var xa = { showgrid: true, gridcolor: "#eee", showline: true, linecolor: "#999",
                   ticks: "outside", tickcolor: "#999", tickfont: { size: 9.5 } };
        var ya = { showgrid: true, gridcolor: "#eee", showline: true, linecolor: "#999",
                   ticks: "outside", tickcolor: "#999", tickfont: { size: 10 } };

        /* 计算本子图 X 轴显示范围（subMin/subMax）
         * 优先级:
         *   1) 分段步长模式 + perGroupEdges     → 用各自 bounds 的首尾
         *   2) 分段步长模式 + globalEdges        → 用 globalEdges 首尾 (统一)
         *   3) 自动/固定步长/柱数 + facetClipAuto → 用各组 P低~P高 (裁 outlier)
         *   4) 其他                              → 用 rawMin/rawMax
         */
        var subMin, subMax, subEdges = null;
        var clippedExcused = "";  /* annotation 显示 "已自动排除 N 个极端值" */
        if (isHist && cfg.binMode === "pieces") {
          if (perGroupEdges && perGroupEdges[k] && perGroupEdges[k].length >= 2) {
            subEdges = perGroupEdges[k]; subMin = subEdges[0]; subMax = subEdges[subEdges.length - 1];
          } else if (globalEdges && globalEdges.length >= 2) {
            subEdges = globalEdges; subMin = globalEdges[0]; subMax = globalEdges[globalEdges.length - 1];
          } else {
            subMin = Math.min.apply(null, ek.vals); subMax = Math.max.apply(null, ek.vals);
          }
        } else {
          subMin = Math.min.apply(null, ek.vals);
          subMax = Math.max.apply(null, ek.vals);
          /* Facet 子图自动按主体范围裁切 (hist + 非分段): 把 subMin/subMax 收缩到 P低~P高,
           * 让 outlier 不会拉扁主体分布。但默认步长仍是全量数据按 IQR 算的. */
          if (isHist && cfg.facetClipAuto && ek.vals.length > 20) {
            var clipLo = cfg.facetClipLo, clipHi = cfg.facetClipHi;
            var pLo = C.percentile(ek.vals, clipLo);
            var pHi = C.percentile(ek.vals, clipHi);
            var nBelow = 0, nAbove = 0;
            for (var vi = 0; vi < ek.vals.length; vi++) {
              if (ek.vals[vi] < pLo) nBelow++;
              else if (ek.vals[vi] > pHi) nAbove++;
            }
            subMin = pLo; subMax = pHi;
            if ((nBelow + nAbove) > 0) {
              clippedExcused = "  裁掉 " + nBelow + "/" + nAbove + " 个极值";
            }
          }
        }
        if (!(subMax > subMin)) subMax = subMin + 1;

        if (isHist) {
          var subStep = globalStep;
          if (perGroupStep) subStep = perGroupStep[k];
          else if (cfg.binMode === "count") subStep = (subMax - subMin) / Math.max(cfg.nBins, 1);
          if (subEdges) {
            addPiecewiseBars(data, ek, ck, subEdges, kind,
                             cfg.showMode, cfg.bw, true, refX, refY,
                             cfg.smoothLevel);
          } else {
            addHistTraces(data, ek, ck, subMin, subMax, subStep, kind,
                          cfg.showMode, cfg.bw, true, refX, refY);
          }
          var cleanLabel = applyRename(ek.label, S.groupRenameMap);
          xa.title = { text: cfg.metricName + (xLog ? " (log)" : ""),
                       font: { size: 11 } };
          xa.showticklabels = true;
          xa.range = [subMin, subMax];
          if (xLog) { xa.type = "log"; xa.range = [Math.max(subMin, 1e-9), subMax]; }
          if (kind === "hist_percent" || kind === "hist_cum") {
            ya.ticksuffix = "%";
            ya.range = [0, 100];
          }
          if (yLog) ya.type = "log";
          else if (cfg.axis.yMin != null || cfg.axis.yMax != null)
            setRange(ya, cfg.axis.yMin, cfg.axis.yMax, false);
          /* 分组名/n 当作图内 annotation 写在子图左上角 (paper 坐标基于子图 grid) */
          if (cfg.facet) {
            anns.push({ text: cleanLabel + "  (n=" + fmtCount(ek.vals.length) + ")" + clippedExcused,
                        xref: refX, yref: refY,
                        x: subMin, y: (kind === "hist_percent" || kind === "hist_cum") ? 95 : (yLog ? Math.log10(Math.max(subMax, 1)) : subMax),
                        xanchor: "left", yanchor: "top", showarrow: false,
                        font: { size: 11, color: "#333" },
                        bgcolor: "rgba(255,255,255,0.85)" });
          }
        } else {
          var isViolin = kind === "violin";
          var tr = isViolin
            ? { type: "violin", y: ek.vals, name: ek.label, legendgroup: ek.label,
                line: { color: ck }, fillcolor: C.hexToRgba(ck, 0.65), opacity: 0.9,
                points: false, box: { visible: true }, meanline: { visible: true },
                showlegend: false, xaxis: refX, yaxis: refY,
                hovertemplate: "<b>" + ek.label + "</b><br>数值: %{y:.4g}<extra></extra>" }
            : { type: "box", y: ek.vals, name: ek.label, legendgroup: ek.label,
                marker: { color: ck }, fillcolor: C.hexToRgba(ck, 0.5),
                line: { color: ck }, boxpoints: false, showlegend: false,
                xaxis: refX, yaxis: refY,
                hovertemplate: "<b>" + ek.label + "</b><br>数值: %{y:.4g}<extra></extra>" };
          data.push(tr);
          xa.showticklabels = false;
          xa.showgrid = false;
          var inLastRow2 = k >= lastRowStart;
          xa.title = { text: applyRename(ek.label, S.groupRenameMap) +
                         (inLastRow2 ? "  (n=" + fmtCount(ek.vals.length) + ")" : ""),
                       font: { size: 11.5 } };
          if (yLogNum) ya.type = "log";
          var l0 = null, h0 = null;
          if (cfg.axis.mode === "clip") { l0 = ax.vmin; h0 = ax.vmax; }
          else if (cfg.axis.mode === "manual") { l0 = cfg.axis.min; h0 = cfg.axis.max; }
          setRange(ya, l0, h0, yLogNum);
        }
        layout[xaKey] = xa; layout[yaKey] = ya;
      }
      if (anns.length) layout.annotations = anns;
      layout.showlegend = false;
    }

    layout.title = { text: cfg.title,
                     font: { size: 15 } };
    layout.template = "plotly_white";
    layout.plot_bgcolor = "#fff";
    layout.font = { size: 12.5 };
    layout.showlegend = !cfg.facet;
    if (!cfg.facet) layout.legend = { title: { text: "分组" }, font: { size: 12 } };
    return { data: data, layout: layout };
  }

  /* ================= 渲染 ================= */
  function scheduleRender() {
    clearTimeout(renderTimer);
    renderTimer = setTimeout(render, 120);
  }
  function render() {
    if (!S.rows || S.rowIdx === null || S.colIdx === null) return;
    if (S.metricIdx < 0) { toast("没有可用数值列"); return; }
    var cfg = readCfg();
    console.log("[render] facetCb=" + $("facetCb").checked + " cfg.facet=" + cfg.facet +
      " binMode=" + cfg.binMode + " n=" + S.selected.length);
    var entries = C.extractByGroup(S.rows, S.labels, S.metricIdx, S.selected);
    /* 应用分组名重命名映射 (仅展示用, 内部 select 仍用原名) */
    entries.forEach(function (e) { e.label = applyRename(e.label, S.groupRenameMap); });
    /* 同步 selected 显示文本 (多选框 + 全选按钮 */
    refreshGroupListLabels();
    var fig = buildPlot(entries, cfg);
    if (fig.error) { toast(fig.error); return; }

    /* 诊断面板: ?debug=1 时显示, 包含 facetCb 真实状态 / cfg.facet / grid 布局 / 当前标题,
     * 帮助排查 "勾了 checkbox 但不生效" 类问题 */
    if (/[?&]debug=1\b/.test(location.search)) {
      var ds = $("diagStatus");
      ds.style.display = "";
      var fcb = $("facetCb"), fi = fcb ? fcb.checked : null;
      var grid = (fig.layout && fig.layout.grid) ? fig.layout.grid.rows + "x" + fig.layout.grid.columns : "无";
      var hist = (window.__diagLog && window.__diagLog.length > 1)
        ? "<br>📜 facetCb 变化轨迹: " + window.__diagLog.map(function (e) {
            return "t=" + e.t + "ms:<b>" + e.facetCb + "</b>";
          }).join(" → ")
        : "";
      ds.innerHTML =
        "🔍 <b>调试</b> &nbsp;·&nbsp; " +
        "facetCb DOM:" + (fcb ? "<b>" + fi + "</b>" : "未找到") + " &nbsp;|&nbsp; " +
        "cfg.facet:<b>" + cfg.facet + "</b> &nbsp;|&nbsp; " +
        "binMode:<b>" + cfg.binMode + "</b> &nbsp;|&nbsp; " +
        "选中组:<b>" + entries.length + "</b> &nbsp;|&nbsp; " +
        "布局:<b>" + grid + "</b> &nbsp;|&nbsp; " +
        "标题:<b>" + (fig.layout && fig.layout.title && fig.layout.title.text ? fig.layout.title.text : "(无)") + "</b>" +
        hist;
      window.__diag = { facetCb: fi, cfgFacet: cfg.facet, binMode: cfg.binMode, n: entries.length, grid: grid,
                        title: (fig.layout && fig.layout.title && fig.layout.title.text) || "" };
    } else {
      $("diagStatus").style.display = "none";
    }

    Plotly.react($("plot"), fig.data, fig.layout, { responsive: true, displaylogo: false });

    // 摘要条
    var strip = $("sumStrip");
    strip.innerHTML = "";
    entries.forEach(function (e) {
      if (!e.vals.length) return;
      var chip = document.createElement("span");
      chip.className = "chip";
      chip.textContent = e.label + "  n=" + fmtCount(e.vals.length);
      strip.appendChild(chip);
    });

    // 统计表
    var show = cfg.showStat && entries.some(function (e) { return e.vals.length; });
    $("statCard").style.display = show ? "" : "none";
    if (show) {
      var rows = C.statRows(entries.filter(function (e) { return e.vals.length; }));
      var html = "<table class='stat'><thead><tr><th>分组</th><th>n</th><th>均值</th>" +
        "<th>中位数</th><th>标准差</th><th>最小</th><th>最大</th></tr></thead><tbody>";
      rows.forEach(function (r) {
        html += "<tr><td>" + r.label + "</td><td>" + fmtCount(r.n) + "</td><td>" + r.mean +
          "</td><td>" + r.median + "</td><td>" + r.std + "</td><td>" + r.min +
          "</td><td>" + r.max + "</td></tr>";
      });
      html += "</tbody></table><div class='hint' style='margin-top:6px'>指标: " +
        cfg.metricName + "</div>";
      $("statBox").innerHTML = html;
    }

    // 轴范围提示
    var allV = [];
    entries.forEach(function (e) { for (var i = 0; i < e.vals.length; i++) allV.push(e.vals[i]); });
    if (allV.length) {
      var mm = axisRange(allV, cfg.axis);
      var s = (cfg.isHist ? "X" : "Y") + "轴 数据范围 " + mm.rawMin.toPrecision(4) + " ~ " +
        mm.rawMax.toPrecision(4);
      if (cfg.axis.mode === "clip")
        s += "（裁切 " + cfg.axis.clipPct + "% 后 " + mm.vmin.toPrecision(4) + " ~ " + mm.vmax.toPrecision(4) + "）";
      else if (cfg.axis.mode === "manual")
        s += "（手动 " + mm.vmin.toPrecision(4) + " ~ " + mm.vmax.toPrecision(4) + "）";
      $("axRangeHint").textContent = s;
    }
  }

  /* ================= 事件 ================= */
  function bind() {
    var dz = $("dropZone"), fi = $("fileInput");
    dz.addEventListener("click", function () { fi.click(); });
    dz.addEventListener("dragover", function (e) { e.preventDefault(); dz.classList.add("over"); });
    dz.addEventListener("dragleave", function () { dz.classList.remove("over"); });
    dz.addEventListener("drop", function (e) {
      e.preventDefault(); dz.classList.remove("over");
      if (e.dataTransfer.files && e.dataTransfer.files.length) loadFile(e.dataTransfer.files[0]);
    });
    fi.addEventListener("change", function () {
      if (fi.files && fi.files.length) loadFile(fi.files[0]);
      fi.value = "";
    });
    $("loadSampleBtn").addEventListener("click", loadSample);
    $("allBtn").addEventListener("click", selectAllOn);
    $("noneBtn").addEventListener("click", selectAllOff);
    $("groupSearch").addEventListener("input", renderGroupList);
    $("rowSelect").addEventListener("change", applyManualGrouping);
    $("colSelect").addEventListener("change", applyManualGrouping);

    $("metricSelect").addEventListener("change", function () {
      S.metricIdx = parseInt($("metricSelect").value, 10);
      S.stepTouched = true;
      scheduleRender();
    });
    $("kindSelect").addEventListener("change", function () {
      updateControlVisibility(); scheduleRender();
    });

    document.querySelectorAll('input[name="binMode"]').forEach(function (el) {
      el.addEventListener("change", function () {
        if (radioVal("binMode") === "step" && !S.stepTouched && S.rows) {
          var vals = [];
          S.selected.forEach(function (g) {
            var ex = C.extractByGroup(S.rows, S.labels, S.metricIdx, [g]);
            if (ex.length && ex[0].vals.length) vals = vals.concat(ex[0].vals);
          });
          if (vals.length) $("binStep").value = C.suggestStep(vals).toPrecision(4);
        }
        updateControlVisibility(); scheduleRender();
      });
    });
    $("binStep").addEventListener("input", function () { S.stepTouched = true; scheduleRender(); });
    $("suggestBtn").addEventListener("click", function () {
      if (!S.rows) { toast("请先上传数据"); return; }
      var vals = [];
      S.selected.forEach(function (g) {
        var ex = C.extractByGroup(S.rows, S.labels, S.metricIdx, [g]);
        if (ex.length && ex[0].vals.length) vals = vals.concat(ex[0].vals);
      });
      if (!vals.length) { toast("当前分组没有数值"); return; }
      var st = C.suggestStep(vals);
      $("binStep").value = st.toPrecision(4);
      S.stepTouched = true;
      toast("已按当前选中的 " + S.selected.length + " 组算出建议步长: " + st.toPrecision(4));
      scheduleRender();
    });
    $("binCount").addEventListener("input", function () {
      $("binCountVal").textContent = $("binCount").value; scheduleRender();
    });
    $("smoothBw").addEventListener("input", function () {
      $("smoothBwVal").textContent = parseFloat($("smoothBw").value).toFixed(2);
      scheduleRender();
    });
    $("curveSmoothLevel").addEventListener("input", function () {
      $("curveSmoothLevelVal").textContent = $("curveSmoothLevel").value;
      scheduleRender();
    });
    $("clipPct").addEventListener("input", function () {
      $("clipPctVal").textContent = $("clipPct").value + "%"; scheduleRender();
    });
    $("facetH").addEventListener("input", function () {
      $("facetHVal").textContent = $("facetH").value; scheduleRender();
    });
    $("facetCb").addEventListener("change", function () {
      console.log("[facetCb] change → checked=" + $("facetCb").checked);
      updateControlVisibility(); scheduleRender();
    });
    $("axLogCb").addEventListener("change", scheduleRender);
    $("axMin").addEventListener("change", scheduleRender);
    $("axMax").addEventListener("change", scheduleRender);
    $("yMin").addEventListener("change", scheduleRender);
    $("yMax").addEventListener("change", scheduleRender);
    $("yLogCb").addEventListener("change", scheduleRender);
    $("sameColorCb").addEventListener("change", function () {
      updateControlVisibility(); scheduleRender();
    });
    $("colorPick").addEventListener("input", scheduleRender);
    $("statCb").addEventListener("change", scheduleRender);

    /* 分组重命名 */
    $("renameApplyBtn").addEventListener("click", function () {
      S.groupRenameMap = parseRenameMap($("renameText").value);
      var n = Object.keys(S.groupRenameMap).filter(function (k) { return S.groupRenameMap[k]; }).length;
      $("renameApplied").textContent = n ? "已应用 " + n + " 个改名" : "已清空所有改名";
      renderGroupList();
      scheduleRender();
    });
    $("renameResetBtn").addEventListener("click", function () {
      S.groupRenameMap = {};
      $("renameText").value = diffRename(S.sorted, {});
      $("renameApplied").textContent = "已重置";
      renderGroupList();
      scheduleRender();
    });

    /* 指标列重命名 */
    $("hdrRenameApplyBtn").addEventListener("click", function () {
      S.headerRenameMap = parseRenameMap($("hdrRenameText").value);
      var n = Object.keys(S.headerRenameMap).filter(function (k) { return S.headerRenameMap[k]; }).length;
      $("hdrRenameApplied").textContent = n ? "已应用 " + n + " 个改名" : "已清空所有改名";
      fillMetricSelect();
      scheduleRender();
    });
    $("hdrRenameResetBtn").addEventListener("click", function () {
      S.headerRenameMap = {};
      var candidates = (S.metricIdxs || []).map(function (ci) { return S.headers[ci]; });
      $("hdrRenameText").value = diffRename(candidates, {});
      $("hdrRenameApplied").textContent = "已重置";
      fillMetricSelect();
      scheduleRender();
    });

    /* 图表标题编辑 */
    $("plotTitleApplyBtn").addEventListener("click", function () {
      var v = $("plotTitleInput").value;
      S.plotTitle = v;
      S.plotTitleEdited = (v.trim().length > 0);
      $("plotTitleApplied").textContent = S.plotTitleEdited ? "已应用自定义标题" : "已清空（用默认）";
      scheduleRender();
    });
    $("plotTitleResetBtn").addEventListener("click", function () {
      $("plotTitleInput").value = "";
      S.plotTitle = "";
      S.plotTitleEdited = false;
      $("plotTitleApplied").textContent = "已重置为默认";
      scheduleRender();
    });
    $("plotTitleInput").addEventListener("keydown", function (e) {
      if (e.key === "Enter") $("plotTitleApplyBtn").click();
    });

    /* 分段步长编辑 */
    $("binSegApplyBtn").addEventListener("click", function () {
      S.binSegments = parseBinSegments($("binSegText").value);
      $("binSegText").value = formatBinSegments(S.binSegments);
      var n = S.binSegments.length;
      if (n === 0) {
        $("binSegHint").textContent = "无有效段（请填「起始 结束 步长」每行一段）";
        $("binSegHint").style.color = "#b45309";
      } else {
        $("binSegHint").textContent = "已应用 " + n + " 段: " +
          S.binSegments.map(function (s) { return "[" + _fmtNum(s.start) + "~" + _fmtNum(s.end) + "]×" + _fmtNum(s.step); }).join(" · ");
        $("binSegHint").style.color = "#16a34a";
      }
      updateControlVisibility();
      scheduleRender();
    });
    $("binSegResetBtn").addEventListener("click", function () {
      $("binSegText").value = "";
      S.binSegments = [];
      $("binSegHint").textContent = "已清空";
      $("binSegHint").style.color = "#66728c";
      updateControlVisibility();
      scheduleRender();
    });
    $("binSegTplSel").addEventListener("change", function () {
      var k = $("binSegTplSel").value;
      if (!k) return;
      $("binSegText").value = SEG_TEMPLATES[k] || "";
      $("binSegTplSel").value = "";
      $("binSegApplyBtn").click();
    });

    $("facetAutoStepCb").addEventListener("change", scheduleRender);
    $("facetH").addEventListener("input", function () {
      $("facetHVal").textContent = $("facetH").value;
      scheduleRender();
    });
    document.querySelectorAll('input[name="facetCols"]').forEach(function (el) {
      el.addEventListener("change", scheduleRender);
    });

    /* Facet 自动裁切 X 轴 */
    $("facetClipAutoCb").addEventListener("change", function () {
      updateControlVisibility(); scheduleRender();
    });
    $("facetClipLo").addEventListener("input", function () {
      $("facetClipLoVal").textContent = parseFloat(this.value).toFixed(1) + "%";
      scheduleRender();
    });
    $("facetClipHi").addEventListener("input", function () {
      $("facetClipHiVal").textContent = parseFloat(this.value).toFixed(1) + "%";
      scheduleRender();
    });
    /* binMode 切换也要更新 UI (提示) */
    document.querySelectorAll('input[name="binMode"]').forEach(function (el) {
      el.addEventListener("change", function () { updateControlVisibility(); scheduleRender(); });
    });
    /* 已选分组变化也要刷新提示（注意初始空 NodeList 用委托即可，应绑定到 document 上） */
    $("groupList").addEventListener("change", function () { updateControlVisibility(); scheduleRender(); });

    ["showMode", "stackMode", "axMode"].forEach(function (name) {
      document.querySelectorAll('input[name="' + name + '"]').forEach(function (el) {
        el.addEventListener("change", function () {
          updateControlVisibility(); scheduleRender();
        });
      });
    });
  }

  /* ================= 启动 ================= */
  var ps = $("paletteSelect");
  Object.keys(PALETTES).forEach(function (k) {
    var o = document.createElement("option");
    o.value = k; o.textContent = k;
    ps.appendChild(o);
  });
  bind();
  updateControlVisibility();
  // 测试/演示入口: 打开 ?sample=1 自动载入示例数据, ?facet=1 自动开子图, ?cols=N 自动设每行张数
  if (/[?&]sample=1/.test(location.search)) loadSample();
  if (/[?&]facet=1/.test(location.search)) {
    $("facetCb").checked = true;
    updateControlVisibility();
  }
  var colsMatch = location.search.match(/[?&]cols=(\d)/);
  if (colsMatch) {
    var c = colsMatch[1];
    var el = document.querySelector('input[name="facetCols"][value="' + c + '"]');
    if (el) el.checked = true;
  }
  /* URL: ?clipOff=1 关闭 Facet 自动裁剪 */
  if (/[?&]clipOff=1/.test(location.search)) {
    $("facetClipAutoCb").checked = false;
  }
  /* URL 参数: ?rename=E-7=对照组A&rename=F-10=处理组B */
  var renameMatches = location.search.match(/[?&]rename=([^&]+)/g);
  if (renameMatches) {
    var txt = renameMatches.map(function (m) {
      return decodeURIComponent(m.replace(/^[?&]rename=/, ""));
    }).join("\n");
    $("renameText").value = txt;
    $("renameApplyBtn").click();
  }
  var hdrRenameMatch = location.search.match(/[?&]hdrRename=([^&]+)/);
  if (hdrRenameMatch) {
    $("hdrRenameText").value = decodeURIComponent(hdrRenameMatch[1]);
    $("hdrRenameApplyBtn").click();
  }
  /* URL 参数: ?title=我的标题 (直接应用, 不需要点按钮) */
  var titleMatch = location.search.match(/[?&]title=([^&]+)/);
  if (titleMatch) {
    $("plotTitleInput").value = decodeURIComponent(titleMatch[1]);
    $("plotTitleApplyBtn").click();
  }
  /* URL 参数: ?seg=0+1+0.1%0A1+20+2 (加载并应用分段步长, 自动切到 pieces 模式) */
  var segMatch = location.search.match(/[?&]seg=([^&]+)/);
  if (segMatch) {
    $("binSegText").value = decodeURIComponent(segMatch[1]);
    var piecesRadio = document.querySelector('input[name="binMode"][value="pieces"]');
    if (piecesRadio) piecesRadio.checked = true;
    updateControlVisibility();
    $("binSegApplyBtn").click();
  }
  /* URL 参数: ?smooth=0~100 设置 piecewise 曲线平滑度 */
  var smoothMatch = location.search.match(/[?&]smooth=(\d{1,3})/);
  if (smoothMatch) {
    var sv = parseInt(smoothMatch[1], 10);
    if (sv >= 0 && sv <= 100) {
      $("curveSmoothLevel").value = sv;
      $("curveSmoothLevelVal").textContent = sv;
    }
  }
  /* 自检模式: ?selftest=1 加载后 1.5s 自动点击 facetCb, 模拟用户手动点击, 验证响应 */
  if (/[?&]selftest=1/.test(location.search)) {
    setTimeout(function () {
      console.log("[selftest] before click, facetCb.checked=" + $("facetCb").checked);
      $("facetCb").click();
      console.log("[selftest] after click, facetCb.checked=" + $("facetCb").checked);
    }, 1500);
  }
  /* 实时诊断: ?debug=1 时每秒采样 facetCb, 记录状态变化轨迹到 window.__diagLog */
  if (/[?&]debug=1\b/.test(location.search)) {
    window.__diagLog = [{ t: 0, facetCb: $("facetCb").checked }];
    var t0 = Date.now();
    setInterval(function () {
      var cur = $("facetCb").checked;
      var last = window.__diagLog[window.__diagLog.length - 1];
      if (cur !== last.facetCb) {
        window.__diagLog.push({ t: Date.now() - t0, facetCb: cur });
        console.log("[diag] facetCb changed to " + cur + " at t=" + (Date.now() - t0) + "ms");
      }
    }, 100);
  }
  /* 步骤引导: ?wizard=1 显示三步操作提示横幅, 引导用户重现并诊断 */
  if (/[?&]wizard=1/.test(location.search)) {
    setTimeout(function () {
      var w = document.createElement("div");
      w.id = "wizardBanner";
      w.innerHTML =
        "<div style='position:fixed;top:0;left:0;right:0;background:#FAAD14;color:#fff;padding:10px 16px;z-index:9999;font-size:13px;line-height:1.6;box-shadow:0 2px 8px rgba(0,0,0,0.2)'>" +
        "🧪 <b>诊断模式</b> &nbsp;|&nbsp; " +
        "1) 等示例数据加载完（约 1-2 秒）<br>" +
        "2) 在左侧 <b>「子图模式 (FACET)」</b> 区，手动勾选 <b>「每个分组一张独立子图」</b> checkbox<br>" +
        "3) 看下面黄色诊断条的变化（会显示 facetCb 真实状态 + cfg.facet + 布局 + 标题）<br>" +
        "4) 截图发给我（包含黄色诊断条）" +
        "</div>";
      document.body.appendChild(w);
    }, 200);
  }
})();
