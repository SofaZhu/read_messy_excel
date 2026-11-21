"use strict";

const byId = (id) => document.getElementById(id);
const fileInput = byId("file");
const fileTarget = byId("file-target");
const sheetSelect = byId("sheet");
const headerRow = byId("header-row");
const headerDepth = byId("header-depth");
const dropTail = byId("drop-tail");
const previewButton = byId("preview-button");
const exportButton = byId("export");
const applyTail = byId("apply-tail");
const restoreTail = byId("restore-tail");
let source = null;
let request = null;
let previewed = null;
let exportRequest = null;

const rowStates = {header: "表头", skipped: "已跳过", suggested: "建议删除", dropped: "已删除", kept: "已保留"};

const translations = {
  "No usable header candidates in the first 30 rows.": "前 30 行没有可用的表头候选。",
  "Fewer than three data rows support the top candidate.": "候选表头后可比较的明细不足三行。",
  "No type transition or merged header hierarchy supports the top candidate.": "未发现明确的类型变化或合并层级。",
  "The top header candidate scored below the minimum.": "首选候选的分数低于最低要求。",
  "The leading header candidates have similar scores.": "得分最高的候选彼此接近。",
  "summary label, numeric value, and missing detail field": "有汇总标签和数值，且缺少常见明细字段",
  "insufficient summary evidence": "汇总证据不足",
  "merged range crosses the proposed tail boundary": "合并区域跨越拟裁去的位置",
  "summary label was filled from a merged range": "汇总标签来自合并填充",
  "manual drop_tail override": "按指定尾行数删除",
  "drop_tail exceeds the number of data rows": "删除尾行数超过了明细总行数。",
  "Header extends beyond the worksheet": "指定的表头范围超出工作表。",
  "Selected header has no labels": "选定范围内没有表头标签。",
  "Worksheet range exceeds the 1,000,000-cell limit": "读取范围超过 100 万个单元格。",
  "A column name is too long to export to Excel": "有一个清理后的列名超过了 Excel 单元格的长度上限，请调整表头。",
};

function translate(text) {
  return translations[text] ?? text;
}

function element(tag, text) {
  const node = document.createElement(tag);
  if (text !== undefined) node.textContent = text;
  return node;
}

function setNotice(message, error = false) {
  const notice = byId("notice");
  notice.hidden = !message;
  notice.textContent = message;
  notice.classList.toggle("error", error);
}

function showSheets(sheets, index) {
  sheetSelect.replaceChildren(...sheets.map((name, number) => new Option(name, String(number))));
  if (!sheets.length) sheetSelect.add(new Option("选择文件后显示", ""));
  sheetSelect.value = sheets.length ? String(index) : "";
  sheetSelect.disabled = !sheets.length;
}

function resetParameters() {
  headerRow.value = "";
  headerDepth.value = "";
  headerDepth.disabled = true;
  headerDepth.placeholder = "自动";
  dropTail.value = "";
}

function setSource(next) {
  source = next;
  const label = byId("filename");
  label.textContent = source.demo ? "区域销售周报 · 内置示例" : source.file.name;
  label.title = label.textContent;
  showSheets([], 0);
  resetParameters();
  preview(0);
}

function useFile(file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    setNotice("请选择 .xlsx 文件。", true);
    return;
  }
  setSource({file});
}

function showTable(data, fallbackMessage) {
  const {report, shape, rows, columns} = data ?? {};
  byId("row-count").textContent = shape?.[0] ?? "—";
  byId("column-count").textContent = shape?.[1] ?? "—";
  byId("header-count").textContent = report?.header_rows.length || "—";
  byId("tail-count").textContent = report?.dropped_tail_rows_count ?? "—";
  byId("sheet-tag").textContent = report?.sheet_name ?? "";
  byId("sheet-tag").hidden = !report;
  byId("table-wrap").hidden = !columns;
  byId("empty").hidden = Boolean(columns);
  byId("empty").textContent = fallbackMessage ?? "载入示例或选择一个 .xlsx 文件开始。";
  const note = byId("preview-note");
  note.hidden = !shape;
  const first = report?.header_rows.at(-1) + 1;
  renderGrid("table", columns, rows?.map((values, index) => ({number: first + index, values})));
  if (shape) note.textContent = shape[0]
    ? `预览 ${rows.length} / ${shape[0]} 行 · ${columns.length} / ${shape[1]} 列`
    : "表头下没有明细行。";
}

function showSource(data) {
  byId("source-wrap").hidden = !data;
  byId("source-empty").hidden = Boolean(data);
  byId("source-note").hidden = !data;
  renderGrid("source", data?.columns, data?.rows);
  if (data) byId("source-note").textContent = `预览 ${data.rows.length} / ${data.shape[0]} 行 · ${data.columns.length} / ${data.shape[1]} 列`;
}

function renderGrid(id, columns, rows) {
  const tableHead = byId(`${id}-head`);
  const tableBody = byId(`${id}-body`);
  tableHead.replaceChildren();
  tableBody.replaceChildren();
  if (!columns) return;

  const head = element("tr");
  for (const name of ["原行号", ...columns]) {
    const cell = element("th", name);
    cell.scope = "col";
    cell.title = name;
    head.append(cell);
  }
  tableHead.append(head);
  const body = document.createDocumentFragment();
  let previous = null;
  for (const row of rows) {
    if (previous !== null && row.number > previous + 1) {
      const gap = element("tr");
      gap.className = "source-gap";
      const cell = element("td", `中间 ${row.number - previous - 1} 行未在预览中显示`);
      cell.colSpan = columns.length + 1;
      gap.append(cell);
      body.append(gap);
    }
    const tr = element("tr");
    tr.dataset.row = row.number;
    const number = element("th", row.number);
    number.scope = "row";
    if (row.kind && rowStates[row.kind]) {
      tr.className = `source-${row.kind}`;
      const state = element("span", rowStates[row.kind]);
      state.className = "row-state";
      number.append(state);
    }
    tr.append(number);
    for (const value of row.values) {
      const td = element("td", value ?? "—");
      if (value === null) {
        td.className = "missing";
        td.title = "空单元格";
      } else {
        td.title = value;
      }
      tr.append(td);
    }
    body.append(tr);
    previous = row.number;
  }
  tableBody.append(body);
}

function showReport(report) {
  const body = byId("report-body");
  const mode = byId("report-mode");
  body.replaceChildren();
  mode.hidden = !report;
  if (!report) {
    const empty = element("p", "解析后会显示原始 Excel 行号与合并区域。");
    empty.className = "quiet";
    body.append(empty);
    return;
  }
  mode.textContent = report.header_source === "manual" ? "手动表头" : "自动识别";
  for (const warning of report.warnings) {
    const line = element("p", translate(warning));
    line.className = "warning";
    body.append(line);
  }
  const list = element("dl");
  const coordinates = (values, total) => values.length
    ? values.join("、") + (total > values.length ? ` 等 ${total} 项` : "") : "无";
  const put = (label, value) => list.append(element("dt", label), element("dd", value));
  put("表头行号", report.header_rows.length ? report.header_rows.join("、") : "待确认");
  if (report.header_rows.length) {
    const interval = (range) => !range ? "无" : range[0] === range[1] ? String(range[0]) : range.join("–");
    put("跳过的标题行", interval(report.skipped_prefix));
    put("已填充合并区", coordinates(report.filled_merge_ranges, report.filled_merge_ranges_count));
    put("跳过的合并区", coordinates(report.skipped_merge_ranges, report.skipped_merge_ranges_count));
    put("裁去的空尾行", interval(report.trimmed_blank_tail));
    put("已删除尾行", coordinates(report.dropped_tail_rows, report.dropped_tail_rows_count));
  }
  body.append(list);
  if (report.candidates.length) {
    const details = element("details");
    details.open = !report.header_rows.length;
    const candidates = element("ul");
    for (const candidate of report.candidates) {
      const end = candidate.row + candidate.depth - 1;
      const range = end === candidate.row ? end : `${candidate.row}–${end}`;
      const item = element("li", `第 ${range} 行 · 得分 ${candidate.score.toFixed(2)} · 稳定度 ${candidate.type_stability.toFixed(2)} · 标签覆盖 ${candidate.label_coverage.toFixed(2)} · 类型变化 ${candidate.type_boundary.toFixed(2)}${candidate.layer_evidence ? " · 有合并层级" : ""}`);
      if (!report.header_rows.length) {
        const button = element("button", "用这个表头");
        button.type = "button";
        button.className = "candidate-button";
        button.addEventListener("click", () => {
          headerRow.value = candidate.row;
          headerDepth.disabled = false;
          headerDepth.value = candidate.depth;
          preview();
        });
        item.append(button);
      }
      candidates.append(item);
    }
    details.append(element("summary", "表头候选与评分"), candidates);
    body.append(details);
  }
  if (report.tail_decisions.length) {
    const details = element("details");
    details.open = report.suggested_tail_count > 0;
    const items = element("ul");
    for (const item of report.tail_decisions) {
      const line = element("li", `第 ${item.row} 行 · ${rowStates[item.action]} · ${translate(item.reason)}`);
      if (item.evidence?.length) {
        const labels = ["标签", "数值", "明细字段"];
        const evidence = item.evidence.map(({coordinate, value, filled}, index) =>
          `${coordinate} ${labels[index]} ${value?.trim() ? value : "空白"}${filled ? "（合并填充）" : ""}`
        );
        const description = element("p", evidence.join(" · "));
        description.className = "tail-evidence";
        line.append(description);
      }
      items.append(line);
    }
    const count = report.tail_decisions_count;
    details.append(element("summary", `尾行处理（${count}）`), items);
    if (count > report.tail_decisions.length) details.append(element("p", "仅显示前 30 项。"));
    body.append(details);
  }
}

function queryFor(index = Number(sheetSelect.value)) {
  const query = new URLSearchParams({sheet: index});
  if (source.demo) query.set("demo", "1");
  for (const [key, input] of [["header_row", headerRow], ["header_depth", headerDepth], ["drop_tail", dropTail]]) {
    if (!input.disabled && input.value !== "") query.set(key, input.valueAsNumber);
  }
  return query.toString();
}

function isCurrent(snapshot) {
  return Boolean(snapshot && snapshot === previewed && snapshot.query === queryFor()
    && [headerRow, headerDepth, dropTail].every(input => input.validity.valid));
}

function updateActions() {
  const current = isCurrent(previewed);
  byId("dirty-note").hidden = !previewed || current;
  exportButton.disabled = !current || Boolean(exportRequest);
  exportButton.textContent = exportRequest ? "正在导出…" : "导出完整 .xlsx";
  const suggest = previewed?.report.suggested_tail_count ?? 0;
  const dropped = previewed?.report.dropped_tail_rows_count ?? 0;
  byId("tail-review").hidden = !(suggest || dropped);
  applyTail.hidden = !suggest;
  applyTail.disabled = !current;
  restoreTail.hidden = !dropped;
  restoreTail.disabled = !current;
  if (suggest) byId("tail-message").textContent = `建议删除 ${suggest} 行。确认前这些行会保留在结果和导出文件中。`;
  else if (dropped) byId("tail-message").textContent = `结果已排除 ${dropped} 行尾行。`;
}

function sendWorkbook(route, query, file, signal) {
  return fetch(`/${route}?${query}`, {
    method: "POST",
    headers: {"Content-Type": "application/octet-stream"},
    body: file,
    signal,
  });
}

async function preview(index = Number(sheetSelect.value)) {
  if (!source) return;
  request?.abort();
  exportRequest?.abort();
  exportRequest = null;
  previewed = null;
  updateActions();
  const controller = new AbortController();
  request = controller;
  const query = queryFor(index);
  const file = source.file;
  previewButton.disabled = true;
  byId("results").setAttribute("aria-busy", "true");
  setNotice("正在读取工作簿…");
  showSource(null);
  showTable(null, "正在生成数据预览…");
  showReport(null);
  try {
    const response = await sendWorkbook("preview", query, file, controller.signal);
    const data = await response.json();
    if (request !== controller) return;
    if (data.sheets) showSheets(data.sheets, data.sheet ?? index);
    showSource(data.source);
    showTable(data, response.ok ? undefined : "应用有效的读取参数后，这里会显示明细。");
    showReport(data.report);
    if (response.ok) previewed = {query, file, report: data.report};
    setNotice(response.ok ? "" : translate(data.error ?? "无法读取工作簿。"), !response.ok);
  } catch (error) {
    if (request !== controller) return;
    setNotice("无法连接本地工作台。请检查启动命令是否仍在运行。", true);
    showTable(null);
  } finally {
    if (request === controller) {
      previewButton.disabled = false;
      byId("results").setAttribute("aria-busy", "false");
      updateActions();
    }
  }
}

async function downloadWorkbook() {
  const snapshot = previewed;
  if (!isCurrent(snapshot) || exportRequest) return;
  const controller = new AbortController();
  exportRequest = controller;
  updateActions();
  try {
    const response = await sendWorkbook("export", snapshot.query, snapshot.file, controller.signal);
    if (!response.ok) {
      const data = await response.json();
      if (isCurrent(snapshot)) setNotice(translate(data.error ?? "无法导出工作簿。"), true);
      return;
    }
    const blob = await response.blob();
    if (!isCurrent(snapshot)) return;
    const url = URL.createObjectURL(blob);
    const link = element("a");
    link.href = url;
    link.download = `${snapshot.file?.name.replace(/\.xlsx$/i, "") ?? "messy-excel-demo"}-cleaned.xlsx`;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30_000);
    setNotice("已生成完整工作表，请查看浏览器下载。");
  } catch (error) {
    if (!controller.signal.aborted && isCurrent(snapshot)) setNotice("无法导出工作簿，请重试。", true);
  } finally {
    if (exportRequest === controller) {
      exportRequest = null;
      updateActions();
    }
  }
}

headerRow.addEventListener("input", () => {
  headerDepth.disabled = headerRow.value === "";
  headerDepth.placeholder = headerDepth.disabled ? "自动" : "1";
});
byId("options").addEventListener("input", updateActions);
byId("options").addEventListener("submit", (event) => {
  event.preventDefault();
  preview();
});
applyTail.addEventListener("click", () => {
  if (!isCurrent(previewed)) return;
  dropTail.value = previewed.report.suggested_tail_count;
  preview();
});
byId("jump-tail").addEventListener("click", () => {
  const number = byId("source-body").querySelector("tr.source-suggested th, tr.source-dropped th");
  number?.scrollIntoView({block: "nearest", inline: "nearest"});
});
restoreTail.addEventListener("click", () => {
  if (!isCurrent(previewed)) return;
  dropTail.value = "";
  preview();
});
exportButton.addEventListener("click", downloadWorkbook);
sheetSelect.addEventListener("change", () => {
  resetParameters();
  preview();
});
fileInput.addEventListener("change", () => {
  useFile(fileInput.files[0]);
  fileInput.value = "";
});
fileTarget.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    fileInput.click();
  }
});
for (const event of ["dragover", "dragleave", "drop"]) {
  fileTarget.addEventListener(event, (action) => {
    action.preventDefault();
    fileTarget.classList.toggle("drag-over", event === "dragover");
    if (event === "drop") useFile(action.dataTransfer.files[0]);
  });
}
byId("demo").addEventListener("click", () => setSource({demo: true}));
setSource({demo: true});
