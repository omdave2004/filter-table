/* Runs the real Filter Table bundle on the page with a small stand-in for the
   Power BI host, and drives the card and chart tiles from the report filters
   the visual sends, the way other visuals on a report page respond. */
(function () {
  "use strict";

  var TABLE = "Movements";
  var COLUMNS = [
    { name: "Destination", type: { text: true } },
    { name: "Category", type: { text: true } },
    { name: "Source", type: { text: true } },
    { name: "Status", type: { text: true } },
    { name: "Units", type: { numeric: true, integer: true }, format: "#,0" },
    { name: "Value (AUD)", type: { numeric: true }, format: "#,0.00" },
    { name: "Posted", type: { dateTime: true }, format: "dd MMM yyyy" }
  ].map(function (c, i) {
    return {
      displayName: c.name,
      queryName: TABLE + "." + c.name,
      expr: { source: { entity: TABLE }, ref: c.name },
      type: c.type,
      format: c.format,
      index: i,
      roles: { values: true }
    };
  });
  var IDX = { dest: 0, units: 4, value: 5 };
  var ROWS = window.OD_SAMPLE_ROWS || [];
  var PRESETS = ["Modern", "Corporate", "Minimal", "Midnight", "Ledger"];
  var reduceMotion = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  /* ---------- filter evaluation (Power BI JSON filters) ---------- */

  function colIndexFor(target) {
    if (!target) return -1;
    for (var i = 0; i < COLUMNS.length; i++) {
      if (COLUMNS[i].expr.ref === target.column) return i;
    }
    return -1;
  }

  function comparable(v) {
    if (typeof v === "number") return v;
    var s = String(v);
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return Date.parse(s.length === 10 ? s + "T00:00:00" : s);
    var n = Number(s.replace(/,/g, ""));
    return isFinite(n) ? n : s.toLowerCase();
  }

  function isBlank(v) { return v === null || v === undefined || v === ""; }

  function testCondition(cell, c) {
    var text = isBlank(cell) ? "" : String(cell).toLowerCase();
    var needle = String(c.value == null ? "" : c.value).toLowerCase();
    switch (c.operator) {
      case "Contains": return text.indexOf(needle) >= 0;
      case "DoesNotContain": return text.indexOf(needle) < 0;
      case "StartsWith": return text.indexOf(needle) === 0;
      case "DoesNotStartWith": return text.indexOf(needle) !== 0;
      case "Is": return text === needle;
      case "IsNot": return text !== needle;
      case "IsBlank": return isBlank(cell);
      case "IsNotBlank": return !isBlank(cell);
    }
    if (isBlank(cell)) return false;
    var a = comparable(cell), b = comparable(c.value);
    switch (c.operator) {
      case "GreaterThan": return a > b;
      case "GreaterThanOrEqual": return a >= b;
      case "LessThan": return a < b;
      case "LessThanOrEqual": return a <= b;
    }
    return true;
  }

  function rowPasses(row, filters) {
    for (var i = 0; i < filters.length; i++) {
      var f = filters[i];
      var ci = colIndexFor(f && f.target);
      if (ci < 0) continue;
      var cell = row[ci];
      if (f.operator === "In" && Array.isArray(f.values)) {
        var hit = f.values.some(function (v) {
          if (v === null) return isBlank(cell);
          if (typeof v === "number") return Number(cell) === v;
          return String(cell) === String(v);
        });
        if (!hit) return false;
      } else if (Array.isArray(f.conditions) && f.conditions.length) {
        var results = f.conditions.map(function (c) { return testCondition(cell, c); });
        var ok = f.logicalOperator === "Or"
          ? results.some(Boolean)
          : results.every(Boolean);
        if (!ok) return false;
      }
    }
    return true;
  }

  /* ---------- formatting ---------- */

  function fmtInt(n) { return Math.round(n).toLocaleString("en-AU"); }
  function fmtShort(n) {
    var a = Math.abs(n);
    if (a >= 1e6) return (n / 1e6).toFixed(2) + "M";
    if (a >= 1e4) return (n / 1e3).toFixed(1) + "K";
    return fmtInt(n);
  }
  function fmtDate(iso) {
    var d = new Date(String(iso).slice(0, 10) + "T00:00:00");
    if (isNaN(d)) return String(iso);
    return d.toLocaleDateString("en-AU", { day: "numeric", month: "short", year: "numeric" });
  }

  /* Plain-language summary of one JSON filter, for the "Filters on this page" line. */
  function describe(f) {
    var ci = colIndexFor(f && f.target);
    var name = ci >= 0 ? COLUMNS[ci].displayName : (f.target && f.target.column) || "Column";
    var isDate = ci >= 0 && COLUMNS[ci].type.dateTime;
    if (f.operator === "In" && Array.isArray(f.values)) {
      var vals = f.values.map(function (v) { return v === null ? "(Blank)" : String(v); });
      var shown = vals.slice(0, 3).join(", ");
      return name + " is " + shown + (vals.length > 3 ? " and " + (vals.length - 3) + " more" : "");
    }
    if (Array.isArray(f.conditions)) {
      var parts = f.conditions.map(function (c) {
        var v = isDate ? fmtDate(c.value) : (typeof c.value === "number" ? fmtInt(c.value) : "“" + c.value + "”");
        switch (c.operator) {
          case "Contains": return "contains " + v;
          case "GreaterThanOrEqual": return (isDate ? "from " : "at least ") + v;
          case "LessThanOrEqual": return "at most " + v;
          case "LessThan": return isDate ? "before " + v : "under " + v;
          case "GreaterThan": return "over " + v;
          default: return c.operator + " " + v;
        }
      });
      return name + " " + parts.join(" and ");
    }
    return name;
  }

  /* ---------- count-up for card values ---------- */

  function setNumber(el, to, format) {
    var from = Number(el.getAttribute("data-current") || to);
    el.setAttribute("data-current", String(to));
    if (reduceMotion || document.hidden || from === to || typeof requestAnimationFrame !== "function") {
      el.textContent = format(to);
      return;
    }
    var start = null, dur = 520, token = {};
    el.__countToken = token;
    function step(ts) {
      if (el.__countToken !== token) return;
      if (start === null) start = ts;
      var t = Math.min(1, (ts - start) / dur);
      var e = t === 1 ? 1 : 1 - Math.pow(2, -10 * t);
      el.textContent = format(from + (to - from) * e);
      if (t < 1) requestAnimationFrame(step);
    }
    requestAnimationFrame(step);
    // Lands on the exact value even if animation frames are paused.
    setTimeout(function () { if (el.__countToken === token) el.textContent = format(to); }, dur + 80);
  }

  /* ---------- report tiles ---------- */

  var SVGNS = "http://www.w3.org/2000/svg";
  function svgEl(tag, attrs) {
    var n = document.createElementNS(SVGNS, tag);
    for (var k in attrs) n.setAttribute(k, attrs[k]);
    return n;
  }

  /* Clustered bar chart drawn at the svg's real pixel size, with a value axis.
     Redrawn when its tile changes size; filter changes only rescale the bars. */
  function buildChart(svg) {
    var totals = {};
    ROWS.forEach(function (r) { totals[r[IDX.dest]] = (totals[r[IDX.dest]] || 0) + r[IDX.value]; });
    var order = Object.keys(totals).sort(function (a, b) { return totals[b] - totals[a]; });
    var max = totals[order[0]] || 1;
    var step = Math.pow(10, Math.floor(Math.log(max) / Math.LN10));
    var tick = max / step > 5 ? step : step / 2;
    var axisMax = Math.ceil(max / tick) * tick;
    var parts = {}, current = ROWS, drawnW = 0, drawnH = 0;

    function draw() {
      var box = svg.getBoundingClientRect();
      var W = Math.round(box.width), H = Math.round(box.height);
      // Too narrow to hold labels, bars and values: wait for a real size.
      if (W < 160 || H < 60 || (W === drawnW && H === drawnH)) return;
      drawnW = W; drawnH = H;
      while (svg.firstChild) svg.removeChild(svg.firstChild);
      svg.setAttribute("viewBox", "0 0 " + W + " " + H);
      var labelW = 34, valueW = 54, axisH = 22;
      var plotW = W - labelW - valueW - 8;
      var rowH = (H - axisH) / order.length;
      var barH = Math.max(12, Math.min(26, rowH * 0.56));
      for (var t = 0; t <= axisMax + 1e-6; t += tick) {
        var gx = labelW + plotW * t / axisMax;
        svg.appendChild(svgEl("line", { "class": "grid", x1: gx, x2: gx, y1: 0, y2: H - axisH }));
        var tl = svgEl("text", { "class": "axis", x: gx, y: H - 6, "text-anchor": t === 0 ? "start" : "middle" });
        tl.textContent = t === 0 ? "0" : (t >= 1e6 ? +(t / 1e6).toFixed(2) + "M" : +(t / 1e3).toFixed(1) + "K");
        svg.appendChild(tl);
      }
      order.forEach(function (key, i) {
        var cy = i * rowH + rowH / 2;
        var label = svgEl("text", { x: 0, y: cy + 4 });
        label.textContent = key;
        var w = plotW * totals[key] / max * (max / axisMax);
        var track = svgEl("rect", { "class": "track", x: labelW, y: cy - barH / 2, height: barH, width: w });
        var bar = svgEl("rect", { "class": "bar", x: labelW, y: cy - barH / 2, height: barH, width: w });
        var tip = svgEl("title", {});
        bar.appendChild(tip);
        var val = svgEl("text", { "class": "val", x: W, y: cy + 4, "text-anchor": "end" });
        svg.appendChild(label); svg.appendChild(track); svg.appendChild(bar); svg.appendChild(val);
        parts[key] = { bar: bar, val: val, tip: tip, total: totals[key] };
      });
      update(current, true);
    }

    function update(rows, instant) {
      current = rows;
      var sums = {};
      rows.forEach(function (r) { sums[r[IDX.dest]] = (sums[r[IDX.dest]] || 0) + r[IDX.value]; });
      order.forEach(function (key) {
        var p = parts[key];
        if (!p) return;
        var v = sums[key] || 0;
        // The bar keeps its full-total width and scales down, so only transform animates.
        if (instant) p.bar.style.transition = "none";
        p.bar.style.transform = "scaleX(" + (v / p.total) + ")";
        if (instant) { p.bar.getBoundingClientRect(); p.bar.style.transition = ""; }
        p.val.textContent = fmtShort(v);
        p.tip.textContent = key + ": $" + fmtInt(v) + " of $" + fmtInt(p.total);
      });
    }

    draw();
    if (typeof ResizeObserver === "function") new ResizeObserver(draw).observe(svg);
    else window.addEventListener("resize", draw);
    return function (rows) { update(rows, false); };
  }

  function bindTiles(root) {
    if (!root) return function () {};
    var q = function (sel) { return root.querySelector(sel); };
    var rowsEl = q("[data-kpi=rows]"), unitsEl = q("[data-kpi=units]"), valueEl = q("[data-kpi=value]");
    var rowsOf = q("[data-kpi-of=rows]"), unitsOf = q("[data-kpi-of=units]"), valueOf = q("[data-kpi-of=value]");
    var filtersEl = q("[data-filters-on]");
    var jsonEl = q("[data-filter-json]");
    var chartEl = q("svg[data-chart]");
    var chartUpdate = chartEl ? buildChart(chartEl) : null;
    var allUnits = 0, allValue = 0;
    ROWS.forEach(function (r) { allUnits += r[IDX.units]; allValue += r[IDX.value]; });

    return function render(filters) {
      var rows = filters.length ? ROWS.filter(function (r) { return rowPasses(r, filters); }) : ROWS;
      var units = 0, value = 0;
      rows.forEach(function (r) { units += r[IDX.units]; value += r[IDX.value]; });
      if (rowsEl) setNumber(rowsEl, rows.length, fmtInt);
      if (unitsEl) setNumber(unitsEl, units, fmtShort);
      if (valueEl) setNumber(valueEl, value, function (n) { return "$" + fmtShort(n); });
      var pct = function (part, whole) { return whole ? Math.round(100 * part / whole) + "% of all" : ""; };
      if (rowsOf) rowsOf.textContent = filters.length ? "of " + fmtInt(ROWS.length) + " rows" : "All rows";
      if (unitsOf) unitsOf.textContent = filters.length ? pct(units, allUnits) : "All rows";
      if (valueOf) valueOf.textContent = filters.length ? pct(value, allValue) : "All rows";
      if (chartUpdate) chartUpdate(rows);
      if (filtersEl) {
        filtersEl.textContent = "";
        var b = document.createElement("b");
        b.textContent = "Filters on this page: ";
        filtersEl.appendChild(b);
        filtersEl.appendChild(document.createTextNode(
          filters.length ? filters.map(describe).join("; ") : "none yet. Use the filters under the table's column headers."));
      }
      if (jsonEl) {
        var jsonTile = jsonEl.closest(".json-tile");
        if (jsonTile) jsonTile.classList.toggle("is-empty", !filters.length);
        jsonEl.textContent = filters.length
          ? JSON.stringify(filters.map(function (f) {
              var copy = {};
              Object.keys(f).forEach(function (k) { copy[k] = f[k]; });
              return copy;
            }), null, 2)
          : "Nothing sent yet. Pick values in any column filter and the payload appears here.";
      }
    };
  }

  /* Starting layout, saved the way an author's layout is saved. Each column
     gets the width its header and filter controls need (measured against the
     visual's own controls); spare room in a wide tile is shared out, and
     narrow screens scroll sideways inside the table. */
  var NEED = [168, 168, 168, 168, 150, 150, 200];
  function startLayout(total, hiddenNames) {
    var hidden = COLUMNS.filter(function (c) { return (hiddenNames || []).indexOf(c.displayName) >= 0; })
      .map(function (c) { return c.queryName; });
    var visible = COLUMNS.filter(function (c) { return hidden.indexOf(c.queryName) < 0; });
    var needSum = visible.reduce(function (a, c) { return a + NEED[c.index]; }, 0);
    // 70px covers the table's frame, padding and vertical scrollbar.
    var scale = Math.max(1, (total - 70) / needSum);
    var widths = {};
    COLUMNS.forEach(function (c) { widths[c.queryName] = Math.floor(NEED[c.index] * scale); });
    return JSON.stringify({ ft_layout_v2: JSON.stringify({ widths: widths, hidden: hidden }) });
  }

  /* ---------- host stand-in ---------- */

  function mount(el, opts) {
    opts = opts || {};
    var plugin = window.powerbi && window.powerbi.visuals && window.powerbi.visuals.plugins
      && window.powerbi.visuals.plugins.filterTableOm2024AAA;
    if (!plugin) {
      el.textContent = "The live demo could not start in this browser.";
      return null;
    }
    el.textContent = "";
    var inner = document.createElement("div");
    el.appendChild(inner);
    var preset = String(opts.preset || 0);
    function freshObjects() {
      return {
        tableStyle: { stylePreset: preset },
        proFeatures: { showTotals: true, enableExport: true },
        powerFeatures: {},
        layoutPersist: { state: startLayout(el.clientWidth, opts.hidden) }
      };
    }
    var objects = freshObjects();
    var columnObjects = {};
    var filters = [];
    var onFilters = opts.onFilters || function () {};
    var visual = null;
    var pending = null;

    function dataView() {
      var cols = COLUMNS.map(function (c) {
        var copy = {};
        Object.keys(c).forEach(function (k) { copy[k] = c[k]; });
        if (columnObjects[c.queryName]) copy.objects = columnObjects[c.queryName];
        return copy;
      });
      return {
        metadata: { columns: cols, objects: objects },
        table: { columns: cols, rows: ROWS }
      };
    }

    // The visual sizes its table to at least the viewport width, but its scroll
    // area is narrower by its frame and scrollbar. Passing the scroll area's
    // real width keeps the table from overflowing by those few pixels.
    function viewportWidth() {
      var sc = inner.querySelector(".ft-scroll");
      var inset = sc ? Math.max(0, inner.clientWidth - sc.clientWidth) : 0;
      return Math.max(0, inner.clientWidth - inset);
    }

    function update(type) {
      visual.update({
        type: type,
        viewport: { width: viewportWidth(), height: inner.clientHeight },
        dataViews: [dataView()],
        jsonFilters: filters.slice(),
        viewMode: 0,
        editMode: 0
      });
    }

    function scheduleUpdate(type) {
      if (pending) clearTimeout(pending);
      pending = setTimeout(function () { pending = null; update(type); }, 40);
    }

    function mergeInto(target, props) {
      Object.keys(props || {}).forEach(function (k) { target[k] = props[k]; });
    }

    var host = {
      eventService: { renderingStarted: function () {}, renderingFinished: function () {}, renderingFailed: function () {} },
      colorPalette: { isHighContrast: false, getColor: function () { return { value: "#1E5AA0" }; } },
      hostCapabilities: { allowInteractions: true },
      locale: "en-AU",
      createSelectionManager: function () {
        return {
          select: function () { return Promise.resolve([]); },
          clear: function () { return Promise.resolve(); },
          showContextMenu: function () { return Promise.resolve(); },
          registerOnSelectCallback: function () {},
          hasSelection: function () { return false; },
          getSelectionIds: function () { return []; }
        };
      },
      createSelectionIdBuilder: function () {
        return {
          withTable: function () { return this; },
          withCategory: function () { return this; },
          withMeasure: function () { return this; },
          createSelectionId: function () {
            return { equals: function () { return false; }, getKey: function () { return ""; }, getSelector: function () { return {}; } };
          }
        };
      },
      // Power BI stores formatting changes in the report and sends them back in the next update.
      persistProperties: function (changes) {
        ["merge", "replace"].forEach(function (kind) {
          (changes && changes[kind] || []).forEach(function (inst) {
            var sel = inst.selector;
            if (sel && sel.metadata) {
              var bag = columnObjects[sel.metadata] = columnObjects[sel.metadata] || {};
              bag[inst.objectName] = bag[inst.objectName] || {};
              mergeInto(bag[inst.objectName], inst.properties);
            } else if (!sel) {
              objects[inst.objectName] = objects[inst.objectName] || {};
              mergeInto(objects[inst.objectName], inst.properties);
            }
          });
        });
        (changes && changes.remove || []).forEach(function (inst) {
          if (!inst.selector && objects[inst.objectName]) {
            Object.keys(inst.properties || {}).forEach(function (k) { delete objects[inst.objectName][k]; });
          }
        });
        scheduleUpdate(2);
      },
      applyJsonFilter: function (f) {
        filters = Array.isArray(f) ? f.filter(Boolean) : (f ? [f] : []);
        onFilters(filters.slice());
      },
      downloadService: {
        exportVisualsContent: function (content, fileName) {
          var blob = new Blob([content], { type: "text/csv;charset=utf-8" });
          var a = document.createElement("a");
          a.href = URL.createObjectURL(blob);
          a.download = fileName || "filter-table-export.csv";
          document.body.appendChild(a);
          a.click();
          setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
          return Promise.resolve(true);
        }
      },
      launchUrl: function (u) { window.open(u, "_blank", "noopener"); },
      fetchMoreData: function () { return false; }
    };

    visual = plugin.create({ element: inner, host: host });
    update(2);
    update(4); // second pass once the scroll area exists to measure

    if (typeof ResizeObserver === "function") {
      var lastW = el.clientWidth, lastH = el.clientHeight;
      new ResizeObserver(function () {
        if (el.clientWidth === lastW && el.clientHeight === lastH) return;
        lastW = el.clientWidth; lastH = el.clientHeight;
        scheduleUpdate(4);
      }).observe(el);
    }

    /* Helpers the feature buttons use to drive the demo the way a reader or
       author would: clicking the visual's own controls, or setting what an
       author would set in the Format pane. */
    var wait = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
    function q(sel) { return inner.querySelector(sel); }
    function colIndex(name) {
      for (var i = 0; i < COLUMNS.length; i++) if (COLUMNS[i].displayName === name) return i;
      return -1;
    }
    function setInput(input, value) {
      input.focus();
      input.value = value;
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(new Event("change", { bubbles: true }));
    }

    var api = {
      setPreset: function (p) {
        preset = String(p);
        objects.tableStyle = objects.tableStyle || {};
        objects.tableStyle.stylePreset = preset;
        update(2);
      },
      // Back to the starting state: no filters, starting layout, current preset.
      reset: function () {
        if (pending) { clearTimeout(pending); pending = null; }
        try { if (visual && visual.destroy) visual.destroy(); } catch (e) {}
        inner.textContent = "";
        objects = freshObjects();
        columnObjects = {};
        filters = [];
        onFilters([]);
        visual = plugin.create({ element: inner, host: host });
        update(2);
        update(4);
        return wait(60);
      },
      // Author settings, as the Format pane would persist them.
      set: function (objectName, props) {
        objects[objectName] = objects[objectName] || {};
        mergeInto(objects[objectName], props);
        update(2);
        return wait(80);
      },
      openFilter: function (name) {
        var btn = q('.ft-ms-btn[data-ms-col="' + colIndex(name) + '"]');
        if (btn && !q('.ft-dropdown[data-dd-col="' + colIndex(name) + '"]')) btn.click();
        return wait(120);
      },
      tick: function (value) {
        var items = inner.querySelectorAll(".ft-dropdown .ft-dd-item");
        for (var i = 0; i < items.length; i++) {
          if (items[i].textContent.trim() === value && !items[i].classList.contains("selected")) { items[i].click(); break; }
        }
        return wait(150);
      },
      closePopups: function () {
        document.body.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
        document.body.dispatchEvent(new MouseEvent("click", { bubbles: true }));
        return wait(120);
      },
      setRange: function (name, min) {
        var cell = inner.querySelectorAll(".ft-fth")[visibleIndex(name)];
        var input = cell && cell.querySelector(".ft-range-input");
        if (input) setInput(input, String(min));
        return wait(400);
      },
      search: function (text) {
        var input = q(".ft-global-search input");
        if (input) setInput(input, text);
        return wait(400);
      },
      openColumnsMenu: function () {
        var buttons = inner.querySelectorAll("button");
        for (var i = 0; i < buttons.length; i++) {
          if (/^\s*Columns/.test(buttons[i].textContent)) { buttons[i].click(); break; }
        }
        return wait(120);
      },
      scrollTable: function (x) {
        var sc = q(".ft-scroll");
        if (sc) {
          sc.scrollTo({ left: x, behavior: reduceMotion ? "auto" : "smooth" });
          // Smooth scrolling does not run in a hidden or throttled tab; land there anyway.
          setTimeout(function () { if (sc.scrollLeft < 5) sc.scrollLeft = x; }, 700);
        }
        return wait(300);
      }
    };
    function visibleIndex(name) {
      var ths = inner.querySelectorAll(".ft-th");
      for (var i = 0; i < ths.length; i++) if (ths[i].textContent.indexOf(name) >= 0) return i;
      return -1;
    }
    return api;
  }

  window.ODReport = {
    presets: PRESETS,
    start: function (opts) {
      var render = bindTiles(opts.tiles);
      render([]);
      return mount(opts.mount, { preset: opts.preset, hidden: opts.hidden, onFilters: render });
    }
  };
})();
