from __future__ import annotations

import argparse
import csv
import json
import mimetypes
import re
import subprocess
import sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse


BASE = Path(__file__).resolve().parent
OUTPUT_DIR = BASE / "skill_outputs"
SKILL_SCRIPT = BASE / "run_channel_activity_skill.py"
INPUT_WORKBOOK = BASE / "上海片区渠道活动后网点服务效果追踪Demo模拟数据.xlsx"

ARTIFACTS = {
    "input_workbook": INPUT_WORKBOOK,
    "email": OUTPUT_DIR / "scheduled_email_digest.md",
    "report": OUTPUT_DIR / "channel_activity_service_report.xlsx",
    "alerts": OUTPUT_DIR / "service_node_alerts.csv",
    "analysis": OUTPUT_DIR / "generated_analysis_table.csv",
    "wecom_messages": OUTPUT_DIR / "wecom_node_messages.md",
    "wecom_payloads": OUTPUT_DIR / "wecom_node_payloads.json",
    "stage_chart": OUTPUT_DIR / "weekly_charts" / "stage_distribution.png",
    "priority_chart": OUTPUT_DIR / "weekly_charts" / "priority_distribution.png",
    "material_chart": OUTPUT_DIR / "weekly_charts" / "material_needs_top5.png",
}


HTML = r"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>渠道活动服务追踪 Agent Demo 工作台</title>
  <style>
    :root {
      --ink: #14213d;
      --muted: #687385;
      --line: #d8dee8;
      --bg: #f6f8fb;
      --panel: #ffffff;
      --blue: #2864b4;
      --green: #198754;
      --amber: #b7791f;
      --red: #c2413f;
      --soft-blue: #eaf2ff;
      --soft-green: #e9f7ef;
      --soft-amber: #fff4dc;
      --soft-red: #fdeceb;
    }

    * { box-sizing: border-box; }
    body {
      margin: 0;
      color: var(--ink);
      background: var(--bg);
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
      letter-spacing: 0;
    }
    button, input, select { font: inherit; }
    .topbar {
      height: 64px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 0 28px;
      background: #ffffff;
      border-bottom: 1px solid var(--line);
      position: sticky;
      top: 0;
      z-index: 10;
    }
    .brand h1 {
      margin: 0;
      font-size: 20px;
      line-height: 1.2;
      font-weight: 760;
    }
    .brand p {
      margin: 5px 0 0;
      color: var(--muted);
      font-size: 13px;
    }
    .actions {
      display: flex;
      align-items: center;
      gap: 10px;
      flex-wrap: wrap;
      justify-content: flex-end;
    }
    .button {
      min-height: 36px;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: #fff;
      color: var(--ink);
      padding: 8px 12px;
      text-decoration: none;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 7px;
      white-space: nowrap;
    }
    .button.primary {
      background: var(--blue);
      border-color: var(--blue);
      color: #fff;
    }
    .button:disabled {
      opacity: .55;
      cursor: wait;
    }
    main {
      width: min(1380px, calc(100vw - 48px));
      margin: 22px auto 36px;
      display: grid;
      gap: 18px;
    }
    .section {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 18px;
    }
    .section-head {
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      gap: 16px;
      margin-bottom: 14px;
    }
    .section h2 {
      margin: 0;
      font-size: 17px;
      line-height: 1.3;
    }
    .section .hint {
      color: var(--muted);
      font-size: 13px;
      margin-top: 4px;
      line-height: 1.45;
    }
    .grid-2 {
      display: grid;
      grid-template-columns: minmax(0, 1.05fr) minmax(360px, .95fr);
      gap: 18px;
    }
    .metric-grid {
      display: grid;
      grid-template-columns: repeat(5, minmax(120px, 1fr));
      gap: 10px;
    }
    .metric {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 12px;
      min-height: 82px;
      background: #fff;
    }
    .metric .label {
      color: var(--muted);
      font-size: 12px;
      margin-bottom: 8px;
    }
    .metric .value {
      font-size: 27px;
      font-weight: 780;
      line-height: 1;
    }
    .metric .sub {
      margin-top: 8px;
      color: var(--muted);
      font-size: 12px;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .flow {
      display: grid;
      grid-template-columns: repeat(7, minmax(88px, 1fr));
      gap: 8px;
      align-items: stretch;
    }
    .step {
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fff;
      padding: 10px;
      min-height: 92px;
      position: relative;
    }
    .step .num {
      width: 24px;
      height: 24px;
      border-radius: 50%;
      display: inline-grid;
      place-items: center;
      font-size: 12px;
      background: #eef2f7;
      color: var(--muted);
      margin-bottom: 9px;
    }
    .step strong {
      display: block;
      font-size: 13px;
      line-height: 1.35;
    }
    .step span {
      display: block;
      color: var(--muted);
      font-size: 12px;
      margin-top: 5px;
      line-height: 1.35;
    }
    .step.done {
      border-color: #9ec5b3;
      background: var(--soft-green);
    }
    .step.active {
      border-color: #8fb3e8;
      background: var(--soft-blue);
      box-shadow: 0 0 0 3px rgba(40,100,180,.10);
    }
    .step.done .num { background: var(--green); color: #fff; }
    .step.active .num { background: var(--blue); color: #fff; }
    .status-line {
      color: var(--muted);
      font-size: 13px;
      margin-top: 12px;
      min-height: 20px;
    }
    .data-list {
      display: grid;
      gap: 10px;
    }
    .data-row {
      display: grid;
      grid-template-columns: 130px minmax(0, 1fr) 80px;
      gap: 12px;
      align-items: center;
      min-height: 44px;
      border-bottom: 1px solid #edf0f5;
      padding-bottom: 10px;
    }
    .data-row:last-child { border-bottom: none; padding-bottom: 0; }
    .data-row b { font-size: 13px; }
    .data-row span {
      min-width: 0;
      color: var(--muted);
      font-size: 12px;
      overflow: hidden;
      white-space: nowrap;
      text-overflow: ellipsis;
    }
    .pill {
      border-radius: 999px;
      padding: 4px 8px;
      font-size: 12px;
      text-align: center;
      border: 1px solid var(--line);
      color: var(--muted);
      background: #fff;
    }
    .pill.ok { color: var(--green); background: var(--soft-green); border-color: #b8dec9; }
    .pill.high { color: var(--red); background: var(--soft-red); border-color: #efb7b5; }
    .pill.mid { color: var(--amber); background: var(--soft-amber); border-color: #efd192; }
    .charts {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
    }
    .chart {
      border: 1px solid var(--line);
      border-radius: 8px;
      min-height: 170px;
      background: #fff;
      padding: 10px;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    .chart img {
      width: 100%;
      max-height: 220px;
      object-fit: contain;
      display: block;
    }
    .split {
      display: grid;
      grid-template-columns: minmax(360px, .9fr) minmax(0, 1.1fr);
      gap: 18px;
    }
    .phone {
      border: 1px solid #cfd6e3;
      border-radius: 8px;
      background: #f3f6fa;
      min-height: 560px;
      padding: 12px;
      display: flex;
      flex-direction: column;
      gap: 10px;
    }
    .phone-head {
      height: 34px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      color: var(--muted);
      font-size: 12px;
      border-bottom: 1px solid #dfe5ee;
      padding-bottom: 8px;
    }
    .messages {
      overflow: auto;
      display: flex;
      flex-direction: column;
      gap: 10px;
      max-height: 500px;
      padding-right: 4px;
    }
    .bubble {
      background: #fff;
      border: 1px solid #dfe5ee;
      border-radius: 8px;
      padding: 10px 12px;
      box-shadow: 0 1px 2px rgba(20, 33, 61, .04);
    }
    .bubble h3 {
      margin: 0 0 8px;
      font-size: 13px;
      line-height: 1.35;
    }
    .bubble p {
      margin: 5px 0;
      color: #2d3748;
      font-size: 12px;
      line-height: 1.45;
    }
    .table-wrap {
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: auto;
      background: #fff;
      max-height: 560px;
    }
    table {
      border-collapse: collapse;
      width: 100%;
      min-width: 920px;
      font-size: 12px;
    }
    th, td {
      border-bottom: 1px solid #edf0f5;
      padding: 9px 10px;
      text-align: left;
      vertical-align: top;
      line-height: 1.4;
    }
    th {
      background: #f7f9fc;
      color: #4a5568;
      position: sticky;
      top: 0;
      z-index: 2;
    }
    .email-box {
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fff;
      min-height: 360px;
      padding: 14px;
      white-space: pre-wrap;
      line-height: 1.55;
      font-size: 13px;
      overflow: auto;
      max-height: 520px;
    }
    .artifact-list {
      display: grid;
      gap: 10px;
    }
    .artifact {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 12px;
      min-height: 70px;
      background: #fff;
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 12px;
      align-items: center;
    }
    .artifact b {
      display: block;
      font-size: 13px;
      margin-bottom: 5px;
    }
    .artifact span {
      color: var(--muted);
      font-size: 12px;
      display: block;
      overflow: hidden;
      white-space: nowrap;
      text-overflow: ellipsis;
    }
    @media (max-width: 1100px) {
      .grid-2, .split { grid-template-columns: 1fr; }
      .metric-grid { grid-template-columns: repeat(3, minmax(120px, 1fr)); }
      .flow { grid-template-columns: repeat(2, minmax(120px, 1fr)); }
      .charts { grid-template-columns: 1fr; }
    }
    @media (max-width: 720px) {
      .topbar { height: auto; padding: 14px 16px; align-items: flex-start; gap: 12px; flex-direction: column; }
      main { width: calc(100vw - 24px); margin-top: 12px; }
      .metric-grid { grid-template-columns: repeat(2, minmax(120px, 1fr)); }
      .data-row { grid-template-columns: 1fr; gap: 4px; }
      .artifact { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <header class="topbar">
    <div class="brand">
      <h1>渠道活动服务追踪 Agent Demo 工作台</h1>
      <p>本地演示版：聚合到渠道/网点层级，不含个人客户 ID；企微与邮件发送为模拟流程。</p>
    </div>
    <div class="actions">
      <button class="button primary" id="runBtn">▶ 运行 Skill 分析</button>
      <button class="button" id="sendBtn">✦ 模拟企微发送</button>
      <a class="button" href="/artifact?name=report">下载 Excel 附件</a>
    </div>
  </header>

  <main>
    <section class="section">
      <div class="section-head">
        <div>
          <h2>运行链路</h2>
          <div class="hint">从模拟输入表开始，生成节点提醒、企业微信文案、邮件正文和 Excel 附件。</div>
        </div>
        <div class="pill ok" id="lastRun">等待读取</div>
      </div>
      <div class="flow" id="flow"></div>
      <div class="status-line" id="statusLine">正在载入最近一次运行结果...</div>
    </section>

    <section class="section">
      <div class="section-head">
        <div>
          <h2>本期看板</h2>
          <div class="hint">用于汇报时先讲结论，再展示由哪些明细支撑。</div>
        </div>
      </div>
      <div class="metric-grid" id="metrics"></div>
    </section>

    <section class="grid-2">
      <div class="section">
        <div class="section-head">
          <div>
            <h2>输入与输出</h2>
            <div class="hint">展示 skill 读取了哪些表，以及最后落地哪些交付物。</div>
          </div>
        </div>
        <div class="data-list" id="ioList"></div>
      </div>

      <div class="section">
        <div class="section-head">
          <div>
            <h2>数据概览</h2>
            <div class="hint">这些图会同步写入 Excel 附件，页面中直接预览 PNG。</div>
          </div>
        </div>
        <div class="charts">
          <div class="chart"><img id="chartStage" alt="节点提醒分布" /></div>
          <div class="chart"><img id="chartPriority" alt="优先级分布" /></div>
          <div class="chart"><img id="chartMaterial" alt="高频材料需求" /></div>
        </div>
      </div>
    </section>

    <section class="split">
      <div class="section">
        <div class="section-head">
          <div>
            <h2>企业微信提醒模拟</h2>
            <div class="hint">展示节点触发后推送给渠道销售的短消息，含保有量观察与材料来源。</div>
          </div>
        </div>
        <div class="phone">
          <div class="phone-head">
            <span>企业微信机器人｜上海片区渠道服务群</span>
            <span id="sentCount">0 条</span>
          </div>
          <div class="messages" id="messages"></div>
        </div>
      </div>

      <div class="section">
        <div class="section-head">
          <div>
            <h2>节点提醒明细</h2>
            <div class="hint">默认展示关注分较高的提醒，完整明细进入 Excel 附件。</div>
          </div>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>节点</th>
                <th>优先级</th>
                <th>渠道</th>
                <th>提醒信号</th>
                <th>建议动作</th>
                <th>来源</th>
              </tr>
            </thead>
            <tbody id="alertRows"></tbody>
          </table>
        </div>
      </div>
    </section>

    <section class="grid-2">
      <div class="section">
        <div class="section-head">
          <div>
            <h2>邮件正文预览</h2>
            <div class="hint">实际工作中可由定时任务发送，附件包含完整明细和图表数据。</div>
          </div>
        </div>
        <div class="email-box" id="emailBox"></div>
      </div>

      <div class="section">
        <div class="section-head">
          <div>
            <h2>演示交付物</h2>
            <div class="hint">汇报时可点击下载，说明每个文件在真实流程中的角色。</div>
          </div>
        </div>
        <div class="artifact-list" id="artifacts"></div>
      </div>
    </section>
  </main>

  <script>
    const steps = [
      ["读取活动表", "02_活动表"],
      ["匹配经营追踪", "03_经营追踪"],
      ["合并问卷反馈", "05_问卷反馈可选"],
      ["生成分析宽表", "标签/材料/建议"],
      ["计算节点提醒", "T+7/T+30/T+90"],
      ["生成企微文案", "节点即时提醒"],
      ["输出邮件附件", "周报邮件 + Excel"]
    ];
    let latest = null;
    let shownMessages = 0;

    function escapeHtml(value) {
      return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;");
    }

    function renderFlow(active = -1, doneCount = 0) {
      const flow = document.getElementById("flow");
      flow.innerHTML = steps.map((s, i) => {
        const cls = i === active ? "step active" : i < doneCount ? "step done" : "step";
        return `<div class="${cls}"><div class="num">${i + 1}</div><strong>${s[0]}</strong><span>${s[1]}</span></div>`;
      }).join("");
    }

    function metric(label, value, sub = "") {
      return `<div class="metric"><div class="label">${label}</div><div class="value">${escapeHtml(value)}</div><div class="sub">${escapeHtml(sub)}</div></div>`;
    }

    function renderMetrics(data) {
      const s = data.summary || {};
      const high = (s.priority_counts || {})["高"] || 0;
      const topMaterials = (data.material_top || []).slice(0, 3).map(x => x.name).join("、") || "-";
      document.getElementById("metrics").innerHTML = [
        metric("覆盖活动", s.events_read || 0, "上海片区"),
        metric("节点提醒", s.node_alerts_generated || 0, "T+7/T+30/T+90"),
        metric("高优先级", high, "需负责人确认"),
        metric("优秀案例", data.excellent_case_count || 0, "可沉淀复用"),
        metric("高频材料", data.material_top?.length || 0, topMaterials)
      ].join("");
    }

    function renderIO(data) {
      const inputSheets = data.summary?.input_sheets || [];
      const outputs = data.summary?.outputs || [];
      const rows = [
        ["输入工作簿", data.files.input_workbook, data.exists.input_workbook],
        ...inputSheets.map(name => [`读取表`, name, true]),
        ...outputs.map(name => [`输出文件`, name, true])
      ];
      document.getElementById("ioList").innerHTML = rows.map(row => `
        <div class="data-row">
          <b>${escapeHtml(row[0])}</b>
          <span title="${escapeHtml(row[1])}">${escapeHtml(row[1])}</span>
          <div class="pill ${row[2] ? "ok" : ""}">${row[2] ? "已就绪" : "缺失"}</div>
        </div>
      `).join("");
    }

    function renderCharts() {
      const stamp = Date.now();
      document.getElementById("chartStage").src = `/artifact?name=stage_chart&inline=1&t=${stamp}`;
      document.getElementById("chartPriority").src = `/artifact?name=priority_chart&inline=1&t=${stamp}`;
      document.getElementById("chartMaterial").src = `/artifact?name=material_chart&inline=1&t=${stamp}`;
    }

    function renderMessages(limit = shownMessages || 4) {
      const list = latest?.wecom_preview || [];
      shownMessages = Math.min(limit, list.length);
      document.getElementById("sentCount").textContent = `${shownMessages} / ${list.length} 条`;
      document.getElementById("messages").innerHTML = list.slice(0, shownMessages).map(msg => `
        <div class="bubble">
          <h3>${escapeHtml(msg.title)}</h3>
          ${msg.lines.map(line => `<p>${escapeHtml(line)}</p>`).join("")}
        </div>
      `).join("");
    }

    function renderAlerts(data) {
      const rows = (data.alert_preview || []).map(row => {
        const priorityClass = row["优先级"] === "高" ? "high" : row["优先级"] === "中" ? "mid" : "ok";
        return `<tr>
          <td>${escapeHtml(row["节点"])}</td>
          <td><span class="pill ${priorityClass}">${escapeHtml(row["优先级"])}</span></td>
          <td>${escapeHtml(row["渠道名称"])}</td>
          <td>${escapeHtml(row["提醒信号"])}</td>
          <td>${escapeHtml(row["建议动作"])}</td>
          <td>${escapeHtml(row["材料需求来源"])}</td>
        </tr>`;
      }).join("");
      document.getElementById("alertRows").innerHTML = rows;
    }

    function renderArtifacts(data) {
      const items = [
        ["模拟输入工作簿", "包含活动表、经营追踪、可选问卷反馈，用于演示数据来源。", "input_workbook"],
        ["邮件正文", "定期邮件正文预览，适合复制到真实邮件任务。", "email"],
        ["Excel 附件", "含提醒明细、活动汇总分析、Top 案例、图表数据。", "report"],
        ["企微 payload", "企业微信机器人 markdown 消息结构示例。", "wecom_payloads"],
        ["分析宽表 CSV", "由已有活动和保有量数据生成的标签、材料需求、建议动作。", "analysis"]
      ];
      document.getElementById("artifacts").innerHTML = items.map(item => `
        <div class="artifact">
          <div>
            <b>${item[0]}</b>
            <span>${item[1]}</span>
          </div>
          <a class="button" href="/artifact?name=${item[2]}">下载</a>
        </div>
      `).join("");
    }

    function renderAll(data) {
      latest = data;
      renderFlow(-1, steps.length);
      renderMetrics(data);
      renderIO(data);
      renderCharts();
      shownMessages = Math.min(shownMessages || 4, data.wecom_preview?.length || 0);
      renderMessages(shownMessages || 4);
      renderAlerts(data);
      renderArtifacts(data);
      document.getElementById("emailBox").textContent = data.email_preview || "暂无邮件输出。";
      document.getElementById("lastRun").textContent = data.exists.summary ? "最近输出已就绪" : "待运行";
      document.getElementById("statusLine").textContent = data.status || "已载入最近一次运行结果。";
    }

    async function loadStatus() {
      const res = await fetch("/api/status");
      const data = await res.json();
      renderAll(data);
    }

    async function runSkill() {
      const btn = document.getElementById("runBtn");
      btn.disabled = true;
      document.getElementById("lastRun").textContent = "运行中";
      let i = 0;
      renderFlow(0, 0);
      document.getElementById("statusLine").textContent = "开始运行 skill...";
      const timer = setInterval(() => {
        i = Math.min(i + 1, steps.length - 1);
        renderFlow(i, i);
        document.getElementById("statusLine").textContent = `处理中：${steps[i][0]}`;
      }, 520);
      try {
        const res = await fetch("/api/run", { method: "POST" });
        const data = await res.json();
        clearInterval(timer);
        renderFlow(-1, steps.length);
        renderAll(data);
        document.getElementById("statusLine").textContent = data.status || "Skill 运行完成。";
      } catch (err) {
        clearInterval(timer);
        document.getElementById("statusLine").textContent = `运行失败：${err}`;
      } finally {
        btn.disabled = false;
      }
    }

    function sendMore() {
      if (!latest) return;
      const total = latest.wecom_preview?.length || 0;
      renderMessages(Math.min((shownMessages || 0) + 3, total));
    }

    document.getElementById("runBtn").addEventListener("click", runSkill);
    document.getElementById("sendBtn").addEventListener("click", sendMore);
    renderFlow();
    loadStatus();
  </script>
</body>
</html>
"""


def read_text(path: Path, limit: int | None = None) -> str:
    if not path.exists():
        return ""
    text = path.read_text(encoding="utf-8")
    if limit and len(text) > limit:
        return text[:limit].rstrip() + "\n..."
    return text


def load_json(path: Path) -> dict:
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def read_csv_rows(path: Path, limit: int = 12) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    rows.sort(key=lambda r: float(r.get("关注分") or 0), reverse=True)
    return rows[:limit]


def material_top(path: Path, limit: int = 5) -> list[dict[str, int | str]]:
    if not path.exists():
        return []
    counts: dict[str, int] = {}
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            value = row.get("材料需求") or "-"
            counts[value] = counts.get(value, 0) + 1
    return [{"name": k, "count": v} for k, v in sorted(counts.items(), key=lambda item: item[1], reverse=True)[:limit]]


def excellent_case_count(path: Path) -> int:
    if not path.exists():
        return 0
    count = 0
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if row.get("服务效果标签") == "稳定提升" and row.get("活动规模类型") == "小场":
                count += 1
    return min(count, 3)


def wecom_preview(path: Path, limit: int = 12) -> list[dict[str, object]]:
    if not path.exists():
        return []
    payloads = json.loads(path.read_text(encoding="utf-8"))
    preview = []
    for payload in payloads[:limit]:
        content = payload["payload"]["markdown"]["content"]
        lines = [line.strip("> ").strip() for line in content.splitlines() if line.strip()]
        title = lines[0].replace("### ", "") if lines else payload.get("alert_id", "")
        title = re.sub(r"</?font[^>]*>", "", title)
        preview.append(
            {
                "title": title,
                "lines": lines[1:],
                "target_group": payload.get("target_group"),
                "priority": payload.get("priority"),
                "stage": payload.get("stage"),
            }
        )
    return preview


def status_payload(status: str = "已载入最近一次运行结果。") -> dict:
    summary = load_json(OUTPUT_DIR / "run_summary.json")
    return {
        "status": status,
        "summary": summary,
        "exists": {
            "summary": (OUTPUT_DIR / "run_summary.json").exists(),
            **{name: path.exists() for name, path in ARTIFACTS.items()},
        },
        "files": {name: str(path) for name, path in ARTIFACTS.items()},
        "email_preview": read_text(ARTIFACTS["email"], 6000),
        "alert_preview": read_csv_rows(ARTIFACTS["alerts"], 12),
        "material_top": material_top(ARTIFACTS["analysis"], 5),
        "excellent_case_count": excellent_case_count(ARTIFACTS["analysis"]),
        "wecom_preview": wecom_preview(ARTIFACTS["wecom_payloads"], 18),
    }


class DemoHandler(BaseHTTPRequestHandler):
    server_version = "ChannelActivityDemo/1.0"

    def log_message(self, fmt: str, *args: object) -> None:
        sys.stderr.write("%s - - [%s] %s\n" % (self.address_string(), self.log_date_time_string(), fmt % args))

    def send_json(self, payload: dict, code: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            body = HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if parsed.path == "/api/status":
            self.send_json(status_payload())
            return
        if parsed.path == "/artifact":
            self.serve_artifact(parsed.query)
            return
        self.send_error(404, "Not found")

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path != "/api/run":
            self.send_error(404, "Not found")
            return
        try:
            result = subprocess.run(
                [sys.executable, str(SKILL_SCRIPT)],
                cwd=str(BASE),
                check=True,
                text=True,
                capture_output=True,
                timeout=120,
            )
            payload = status_payload("Skill 运行完成，已刷新企微提醒、邮件正文和 Excel 附件。")
            payload["stdout"] = result.stdout[-3000:]
            self.send_json(payload)
        except subprocess.CalledProcessError as exc:
            self.send_json(
                {
                    "status": "Skill 运行失败。",
                    "stdout": exc.stdout,
                    "stderr": exc.stderr,
                },
                code=500,
            )
        except Exception as exc:
            self.send_json({"status": f"Skill 运行失败：{exc}"}, code=500)

    def serve_artifact(self, query: str) -> None:
        params = parse_qs(query)
        name = (params.get("name") or [""])[0]
        inline = (params.get("inline") or [""])[0] == "1"
        path = ARTIFACTS.get(name)
        if path is None or not path.exists() or not path.resolve().is_relative_to(BASE):
            self.send_error(404, "Artifact not found")
            return
        content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        data = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        disposition = "inline" if inline else "attachment"
        self.send_header("Content-Disposition", f'{disposition}; filename="{path.name}"')
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)


def main() -> None:
    parser = argparse.ArgumentParser(description="渠道活动服务追踪 Agent 本地 Demo 工作台")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", default=8001, type=int)
    args = parser.parse_args()
    server = ThreadingHTTPServer((args.host, args.port), DemoHandler)
    print(f"Demo workbench running at http://{args.host}:{args.port}/")
    print(f"Base directory: {BASE}")
    server.serve_forever()


if __name__ == "__main__":
    main()
