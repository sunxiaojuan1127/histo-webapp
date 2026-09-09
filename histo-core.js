/* =========================================================================
 * histo-core.js — 分组直方图分析核心算法（纯逻辑，无 DOM 依赖）
 * 由 Python 版 plot_utils.py 移植而来，可在浏览器与 Node 中运行。
 * 导出对象: HistoCore
 * ========================================================================= */
(function (root, factory) {
  if (typeof module !== "undefined" && module.exports) {
    module.exports = factory();
  } else {
    root.HistoCore = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  /* ---------------- 基础工具 ---------------- */

  function isPlainNum(v) {
    return typeof v === "number" && Number.isFinite(v);
  }

  /** 单元格 -> 数值；无法转换返回 null。'12.3'/' 5 ' 之类可转，'E'/'N/A' 不可。 */
  function numOrNull(v) {
    if (v == null) return null;
    if (isPlainNum(v)) return v;
    if (typeof v === "string") {
      var s = v.trim();
      if (s === "") return null;
      if (/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(s)) {
        var n = Number(s);
        return Number.isFinite(n) ? n : null;
      }
    }
    return null;
  }

  function fmtCell(v) {
    if (v == null) return "";
    if (isPlainNum(v)) {
      var s = String(v);
      return s.indexOf(".") >= 0 && Number.isInteger(v) ? String(Math.trunc(v)) : s;
    }
    return String(v).trim();
  }

  /** numpy 风格线性插值分位数，p 为 0~100 */
  function percentile(arr, p) {
    var n = arr.length;
    if (!n) return 0;
    var a = Array.prototype.slice.call(arr).sort(function (x, y) { return x - y; });
    var idx = (n - 1) * (p / 100);
    var lo = Math.floor(idx), hi = Math.ceil(idx);
    if (lo === hi) return a[lo];
    return a[lo] + (idx - lo) * (a[hi] - a[lo]);
  }

  function sum(arr) { var s = 0, i; for (i = 0; i < arr.length; i++) s += arr[i]; return s; }
  function mean(arr) { return arr.length ? sum(arr) / arr.length : 0; }

  /** 样本标准差 (ddof=1)，对应 numpy std(ddof=1) */
  function stddev(arr) {
    var n = arr.length;
    if (n < 2) return 0;
    var m = mean(arr), s = 0, i, d;
    for (i = 0; i < n; i++) { d = arr[i] - m; s += d * d; }
    return Math.sqrt(s / (n - 1));
  }

  function hexToRgba(hex, alpha) {
    var h = hex.replace("#", "");
    if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
    var r = parseInt(h.slice(0, 2), 16),
        g = parseInt(h.slice(2, 4), 16),
        b = parseInt(h.slice(4, 6), 16);
    if ([r, g, b].some(Number.isNaN)) return hex;
    return "rgba(" + r + "," + g + "," + b + "," + alpha + ")";
  }

  function linspace(a, b, n) {
    var out = new Array(n), i;
    if (n === 1) return [a];
    for (i = 0; i < n; i++) out[i] = a + ((b - a) * i) / (n - 1);
    return out;
  }

  /* ---------------- Excel 解析 ---------------- */

  /**
   * 解析 Excel。XLSX 为 SheetJS 对象；buffer 为 ArrayBuffer。
   * 返回 { headers, rows, sheetName, nRows }：
   *   headers: 首行字符串数组；rows: 其余各行的原始单元格数组（无表头）
   */
  function parseWorkbook(XLSX, buffer) {
    var wb = XLSX.read(buffer, { type: "array", cellDates: false });
    if (!wb.SheetNames || !wb.SheetNames.length) throw new Error("文件里没有工作表");
    var name = wb.SheetNames[0];
    var ws = wb.Sheets[name];
    if (!ws || !ws["!ref"]) throw new Error("工作表是空的");
    var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false, raw: true });
    if (!aoa.length) throw new Error("工作表没有内容");
    var headers = (aoa[0] || []).map(function (h) {
      return h == null ? "" : String(h).trim();
    });
    var rows = [];
    for (var i = 1; i < aoa.length; i++) {
      var r = aoa[i];
      if (!r) continue;
      var has = false;
      for (var j = 0; j < r.length; j++) if (r[j] !== null && r[j] !== "") { has = true; break; }
      if (has) rows.push(r);
    }
    return { headers: headers, rows: rows, sheetName: name, nRows: rows.length };
  }

  /** 在表头里找 Row / Column 两列（忽略大小写与首尾空格）。找不到返回 null */
  function findGroupIdx(headers) {
    var ri = null, ci = null;
    for (var i = 0; i < headers.length; i++) {
      var h = headers[i].toLowerCase();
      if (ri === null && h === "row") ri = i;
      else if (ci === null && h === "column") ci = i;
    }
    return { rowIdx: ri, colIdx: ci };
  }

  /** 每行的分组标签 'E-7' 形式 */
  function buildLabels(rows, ri, ci) {
    var out = new Array(rows.length);
    for (var i = 0; i < rows.length; i++) {
      out[i] = fmtCell(rows[i][ri]) + "-" + fmtCell(rows[i][ci]);
    }
    return out;
  }

  /** 每个分组的行数: {label: count}，且按自然排序给出全部组名 */
  function groupCounts(labels) {
    var m = {}, i;
    for (i = 0; i < labels.length; i++) {
      var l = labels[i];
      if (l === "-") continue;
      m[l] = (m[l] || 0) + 1;
    }
    return m;
  }

  /** 自然排序: 'E-10' 排在 'E-9' 之后，同 Row 时按 Column 数字排 */
  function sortLabels(labels) {
    function key(label) {
      var p = String(label).split("-", 2);
      if (p.length === 2) {
        var c = parseFloat(p[1]);
        return [p[0], Number.isFinite(c) ? c : 0, label];
      }
      return [label, 0, label];
    }
    return labels.slice().sort(function (a, b) {
      var ka = key(a), kb = key(b);
      if (ka[0] < kb[0]) return -1;
      if (ka[0] > kb[0]) return 1;
      if (ka[1] < kb[1]) return -1;
      if (ka[1] > kb[1]) return 1;
      return ka[2] < kb[2] ? -1 : ka[2] > kb[2] ? 1 : 0;
    });
  }

  /** 除排除列外，至少含一个数值的列索引 */
  function numericMetricIdxs(rows, headers, exclIdxs) {
    var excl = exclIdxs || {};
    var ncol = headers.length;
    var metric = [];
    for (var c = 0; c < ncol; c++) {
      if (excl[c]) continue;
      var ok = false;
      for (var r = 0; r < rows.length && !ok; r++) {
        if (numOrNull(rows[r][c]) !== null) ok = true;
      }
      if (ok) metric.push(c);
    }
    return metric;
  }

  /**
   * 提取多个分组在某个指标列下的全部数值。
   * 返回 [{label, vals: number[]}, ...]（顺序 = wantedGroups 顺序）
   */
  function extractByGroup(rows, labels, metricIdx, wantedGroups) {
    var set = {};
    for (var g = 0; g < wantedGroups.length; g++) set[wantedGroups[g]] = true;
    var tmp = {}, i, l, v;
    for (i = 0; i < rows.length; i++) {
      l = labels[i];
      if (!set[l]) continue;
      v = numOrNull(rows[i][metricIdx]);
      if (v === null) continue;
      (tmp[l] = tmp[l] || []).push(v);
    }
    var out = [];
    for (i = 0; i < wantedGroups.length; i++) {
      out.push({ label: wantedGroups[i], vals: tmp[wantedGroups[i]] || [] });
    }
    return out;
  }

  /* ---------------- 步长建议 ---------------- */

  /**
   * Freedman-Diaconis / Scott 自动建议柱宽（与 Python 版一致）。
   */
  function suggestStep(values) {
    var v = values.filter(function (x) { return Number.isFinite(x); });
    var n = v.length;
    if (n < 2) return 0.5;
    var vmin = Infinity, vmax = -Infinity, i;
    for (i = 0; i < n; i++) {
      if (v[i] < vmin) vmin = v[i];
      if (v[i] > vmax) vmax = v[i];
    }
    var rng = vmax - vmin;
    if (rng <= 0) return Math.max(Math.abs(vmax) * 0.01, 1e-9);
    var iqr = percentile(v, 75) - percentile(v, 25);
    var fd;
    if (iqr > 0) fd = (2.0 * iqr) / Math.pow(n, 1 / 3);
    else fd = (rng * 3.5) / (Math.pow(n, 1 / 3) + 1);
    var s = stddev(v);
    var scott = s > 0 ? (3.5 * s) / Math.pow(n, 1 / 3) : fd;
    var step = Math.max(fd, scott);
    step = Math.max(step, 1e-9);
    step = Math.min(step, rng);
    return step;
  }

  /* ---------------- 正态 / KDE ---------------- */

  /* Abramowitz & Stegun 7.1.26, 误差 < 1.5e-7 */
  function _erf(x) {
    var sign = x < 0 ? -1 : 1;
    var ax = Math.abs(x);
    var t = 1 / (1 + 0.3275911 * ax);
    var y = 1 - Math.exp(-ax * ax) * t *
      (0.254829592 + t * (-0.284496736 + t * (1.421413741 + t * (-1.453152027 + t * 1.061405429))));
    return sign * y;
  }
  function normCdf(z) { return 0.5 * (1 + _erf(z * 0.7071067811865475)); }

  /**
   * 平滑曲线：Y 与直方图柱子严格同语义（峰值上限 100%）。
   * kind: hist_freq / hist_percent / hist_cum
   * 返回 {x: number[], y: number[], unit: string} 或 null
   */
  function kdeCurveForHist(vals, vmin, vmax, stepEff, kind, bw, nQuery, nBins) {
    nQuery = nQuery || 240;
    nBins = nBins || 220;
    var n = vals.length;
    if (n < 5) return null;
    var s = stddev(vals);
    if (!(s > 0)) return null;
    var h = Math.max(bw * s, 1e-12);
    var lo = Infinity, hi = -Infinity, i;
    for (i = 0; i < n; i++) {
      if (vals[i] < lo) lo = vals[i];
      if (vals[i] > hi) hi = vals[i];
    }
    var span = hi - lo;
    if (!(span > 0)) return null;
    // 数据加权直方图（细网格），用 bin 中心近似核密度
    var binW = span / nBins;
    var counts = new Array(nBins).fill(0);
    for (i = 0; i < n; i++) {
      var bi = Math.min(nBins - 1, Math.max(0, Math.floor((vals[i] - lo) / binW)));
      counts[bi]++;
    }
    var centers = new Array(nBins);
    for (i = 0; i < nBins; i++) centers[i] = lo + (i + 0.5) * binW;

    function cdfAt(x) {
      var ssum = 0;
      for (var b = 0; b < nBins; b++) {
        if (!counts[b]) continue;
        ssum += counts[b] * normCdf((x - centers[b]) / h);
      }
      return ssum / n;
    }
    var half = stepEff / 2;
    var xQ = linspace(vmin, vmax, nQuery);
    var yQ = new Array(nQuery);
    for (i = 0; i < nQuery; i++) {
      var x = xQ[i];
      if (kind === "hist_cum") {
        yQ[i] = cdfAt(x) * 100.0;
      } else {
        var up = cdfAt(x + half), dn = cdfAt(x - half);
        var pIn = Math.max(0, Math.min(1, up - dn));
        if (kind === "hist_freq") yQ[i] = pIn * n;
        else yQ[i] = pIn * 100.0;
      }
    }
    return {
      x: xQ, y: yQ,
      unit: kind === "hist_freq"
        ? "频数（窗口宽 = " + stepEff.toPrecision(4) + "）"
        : kind === "hist_percent"
          ? "百分比 %（窗口宽 = " + stepEff.toPrecision(4) + "）"
          : "累计百分比 %"
    };
  }

  /* ---------------- 分段步长（非均匀分箱） ---------------- */

  /**
   * 由多段 [start, end, step] 拼接出有序边界数组（含缺口自动填补）。
   *  - segments: [{start, end, step}, ...]，step > 0，end > start；顺序不限
   *  - 自动按 start 升序
   *  - 若相邻段之间有缺口，使用下一段的 step 反向填补
   *  - 超出 [vmin, vmax] 的部分自动裁掉
   *  - 末段尾巴补到 vmax
   *  - 无有效段 / 无法构成 2 个以上边界 → 返回 null（让上层回退到固定步长）
   */
  function piecewiseBinEdges(values, segments) {
    if (!Array.isArray(segments) || segments.length === 0) return null;
    var valid = [];
    for (var i = 0; i < segments.length; i++) {
      var s = segments[i];
      if (!s) continue;
      var a = +s.start, b = +s.end, w = +s.step;
      if (!Number.isFinite(a) || !Number.isFinite(b) || !Number.isFinite(w)) continue;
      if (!(w > 0) || !(b > a)) continue;
      valid.push({ start: a, end: b, step: w });
    }
    if (!valid.length) return null;
    valid.sort(function (x, y) { return x.start - y.start; });

    var vs = values.filter(function (v) { return Number.isFinite(v); });
    if (!vs.length) return null;
    var vmin = Infinity, vmax = -Infinity, k;
    for (k = 0; k < vs.length; k++) { if (vs[k] < vmin) vmin = vs[k]; if (vs[k] > vmax) vmax = vs[k]; }
    if (!(vmax > vmin)) return null;

    var edges = [];
    var EPS = 1e-9;
    var cursor = -Infinity;
    for (var si = 0; si < valid.length; si++) {
      var seg = valid[si];
      if (seg.end <= vmin - EPS || seg.start >= vmax + EPS) continue;
      var sStart = seg.start;
      var sEnd = seg.end;
      if (sEnd > vmax) sEnd = vmax;

      /* 用本段 step 反向填补 cursor → sStart 的缺口 */
      if (edges.length > 0 && sStart > cursor + EPS) {
        var xFill = cursor + seg.step;
        // 对齐：从 cursor 开始用 step 向右走到接近 sStart
        if (xFill < cursor + EPS) xFill = cursor + seg.step;
        while (xFill < sStart - EPS) {
          if (xFill > edges[edges.length - 1] + EPS) edges.push(+(+xFill).toFixed(12));
          xFill += seg.step;
        }
      }

      /* 起始边界（若尚未加入） */
      if (edges.length === 0 || sStart > edges[edges.length - 1] + EPS) {
        edges.push(+(+sStart).toFixed(12));
      }
      cursor = edges[edges.length - 1];

      /* 在本段内按 step 生成右边界 */
      var x = sStart + seg.step;
      while (x <= sEnd + EPS) {
        if (x > edges[edges.length - 1] + EPS) edges.push(+(+x).toFixed(12));
        x += seg.step;
      }
      cursor = edges[edges.length - 1];
    }
    if (edges.length === 0) return null;

    /* 末段补到 vmax */
    if (edges[edges.length - 1] < vmax - EPS) {
      edges.push(+(+vmax).toFixed(12));
    }

    /* 去重排序（防御性） */
    var uniq = [];
    for (var u = 0; u < edges.length; u++) {
      if (uniq.length === 0 || edges[u] > uniq[uniq.length - 1] + EPS) uniq.push(edges[u]);
    }
    return uniq.length >= 2 ? uniq : null;
  }

  /**
   * 已知边界 edges (length>=2)，统计落入每个 bin 的个数。
   *  - kind: hist_freq | hist_percent | hist_cum
   *  - 返回 { x: centers[], y: counts(累计/百分比后)[], widths[], rawCounts[] }
   *  - 上层可用 type:"bar" + x=centers + y=y + width=widths 渲染不等宽柱状图
   */
  function histogramByEdges(values, edges, kind) {
    if (!Array.isArray(edges) || edges.length < 2) return null;
    var nBin = edges.length - 1;
    var counts = new Array(nBin).fill(0);
    var vmin = edges[0], vmax = edges[edges.length - 1];
    var i;
    for (i = 0; i < values.length; i++) {
      var v = values[i];
      if (!(v >= vmin && v <= vmax)) continue;
      if (v === vmax) { counts[nBin - 1]++; continue; }  // 闭区间右端
      /* 二分找 bin: edges[lo] <= v < edges[lo+1] */
      var lo = 0, hi = nBin;
      while (hi - lo > 1) {
        var mid = (lo + hi) >> 1;
        if (edges[mid] <= v) lo = mid; else hi = mid;
      }
      counts[lo]++;
    }
    var centers = new Array(nBin);
    var widths = new Array(nBin);
    for (i = 0; i < nBin; i++) {
      centers[i] = (edges[i] + edges[i + 1]) / 2;
      widths[i] = edges[i + 1] - edges[i];
    }
    var total = values.length;
    var y;
    if (kind === "hist_freq") {
      y = counts.slice();
    } else if (kind === "hist_cum") {
      y = new Array(nBin);
      var sum = 0;
      for (i = 0; i < nBin; i++) { sum += counts[i]; y[i] = total > 0 ? +(sum * 100 / total).toFixed(6) : 0; }
    } else { /* hist_percent */
      y = new Array(nBin);
      for (i = 0; i < nBin; i++) y[i] = total > 0 ? +(counts[i] * 100 / total).toFixed(6) : 0;
    }
    return { x: centers, y: y, widths: widths, rawCounts: counts };
  }

  /* ---------------- 统计表 ---------------- */

  function statRows(entries) {
    var out = [];
    for (var i = 0; i < entries.length; i++) {
      var e = entries[i];
      if (!e.vals.length) { out.push({ label: e.label, n: 0 }); continue; }
      out.push({
        label: e.label,
        n: e.vals.length,
        mean: round4(mean(e.vals)),
        median: round4(percentile(e.vals, 50)),
        std: round4(stddev(e.vals)),
        min: round4(Math.min.apply(null, e.vals)),
        max: round4(Math.max.apply(null, e.vals)),
      });
    }
    return out;
  }
  function round4(x) {
    var r = Math.round(x * 10000) / 10000;
    return Object.is(r, -0) ? 0 : r;
  }

  /* ---------------- Catmull-Rom 样条插值 ---------------- */
  /* 用于"曲线平滑度"控件: 输入一组 (x,y) 点 (典型为 bin 中心 + 柱高),
   * 返回在相邻点之间均匀插入 smoothStep 个中间点的数组.
   * - points: [{x:number, y:number}]  (至少 2 个)
   * - smoothStep: int >= 1, 1=不插值 (折线), 越大越平滑
   * 边界点 p0 / p3 用端点复制 (即 endpoint 模式),
   * 适合 x 不等距 (piecewise bin 宽度不同) 的场景.                       */
  function catmullRomSpline(points, smoothStep) {
    if (!points || points.length === 0) return [];
    if (smoothStep <= 1 || points.length < 2) {
      return points.map(function (p) { return { x: p.x, y: p.y }; });
    }
    var n = points.length;
    var out = new Array((n - 1) * smoothStep + 1);
    var idx = 0;
    for (var i = 0; i < n - 1; i++) {
      var p0 = points[i - 1] || points[i];
      var p1 = points[i];
      var p2 = points[i + 1];
      var p3 = points[i + 2] || points[i + 1];
      for (var j = 0; j < smoothStep; j++) {
        var t = j / smoothStep;
        var t2 = t * t, t3 = t2 * t;
        var x = 0.5 * (
          (2 * p1.x) +
          (-p0.x + p2.x) * t +
          (2 * p0.x - 5 * p1.x + 4 * p2.x - p3.x) * t2 +
          (-p0.x + 3 * p1.x - 3 * p2.x + p3.x) * t3
        );
        var y = 0.5 * (
          (2 * p1.y) +
          (-p0.y + p2.y) * t +
          (2 * p0.y - 5 * p1.y + 4 * p2.y - p3.y) * t2 +
          (-p0.y + 3 * p1.y - 3 * p2.y + p3.y) * t3
        );
        out[idx++] = { x: x, y: y };
      }
    }
    out[idx++] = { x: points[n - 1].x, y: points[n - 1].y };
    return out;
  }

  /* ---------------- 命名空间 ---------------- */
  return {
    numOrNull: numOrNull,
    fmtCell: fmtCell,
    percentile: percentile,
    mean: mean,
    stddev: stddev,
    hexToRgba: hexToRgba,
    linspace: linspace,
    parseWorkbook: parseWorkbook,
    findGroupIdx: findGroupIdx,
    buildLabels: buildLabels,
    groupCounts: groupCounts,
    sortLabels: sortLabels,
    numericMetricIdxs: numericMetricIdxs,
    extractByGroup: extractByGroup,
    suggestStep: suggestStep,
    piecewiseBinEdges: piecewiseBinEdges,
    histogramByEdges: histogramByEdges,
    kdeCurveForHist: kdeCurveForHist,
    statRows: statRows,
    normCdf: normCdf,
    catmullRomSpline: catmullRomSpline,
  };
});
