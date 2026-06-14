from __future__ import annotations

import asyncio
import html
import math
from datetime import datetime
from pathlib import Path

import pandas as pd
from playwright.async_api import async_playwright

import build_html_dashboard as dash


OUT_DIR = Path("outputs")
HTML_OUT = OUT_DIR / "negative_one_page_summary.html"
PNG_OUT = OUT_DIR / "negative_one_page_summary.png"


def pct(value: float, digits: int = 0) -> str:
    return f"{value * 100:.{digits}f}%"


def short(text: str, limit: int) -> str:
    normalized = " ".join(str(text or "").split())
    return normalized if len(normalized) <= limit else normalized[: limit - 1] + "…"


def parse_time(value: object) -> datetime | pd.NaT:
    text = str(value or "").strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%d/%m/%y %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return pd.NaT


def bar_rows(rows: list[dict], label_key: str, value_key: str, max_rows: int = 5) -> str:
    top = rows[:max_rows]
    max_value = max([r[value_key] for r in top] or [1])
    colors = ["#d80019", "#ff6b00", "#ffc400", "#28b36d", "#7a3fd1"]
    parts: list[str] = []
    for idx, row in enumerate(top):
        value = row[value_key]
        width = 100 * value / max_value if max_value else 0
        parts.append(
            f"""
            <div class="bar-row">
              <div class="rank">{idx + 1}</div>
              <div class="bar-label">{html.escape(str(row[label_key]))}</div>
              <div class="bar-track"><div class="bar-fill" style="width:{width:.1f}%;background:{colors[idx % len(colors)]}"></div></div>
              <div class="bar-value">{value}</div>
            </div>"""
        )
    return "\n".join(parts)


def timeline_svg(hour_counts: dict[int, int]) -> str:
    hours = list(range(18, 24))
    values = [hour_counts.get(hour, 0) for hour in hours]
    max_value = max(values or [1])
    width, height = 490, 178
    left, top, right, bottom = 38, 20, 18, 32
    plot_w = width - left - right
    plot_h = height - top - bottom
    points = []
    for idx, value in enumerate(values):
        x = left + (plot_w * idx / (len(hours) - 1))
        y = top + plot_h - (plot_h * value / max_value if max_value else 0)
        points.append((x, y, value, hours[idx]))
    area = " ".join(f"{x:.1f},{y:.1f}" for x, y, _, _ in points)
    base = f"{points[-1][0]:.1f},{top + plot_h:.1f} {points[0][0]:.1f},{top + plot_h:.1f}"
    circles = "\n".join(
        f'<circle cx="{x:.1f}" cy="{y:.1f}" r="5" fill="#d80019" stroke="#fff" stroke-width="2"/>'
        for x, y, _, _ in points
    )
    labels = "\n".join(
        f'<text x="{x:.1f}" y="{height - 10}" text-anchor="middle" class="axis">{hour}:00</text>'
        for x, _, _, hour in points
    )
    grid = "\n".join(
        f'<line x1="{left}" y1="{top + plot_h * i / 4:.1f}" x2="{width - right}" y2="{top + plot_h * i / 4:.1f}" class="grid"/>'
        for i in range(5)
    )
    peak = max(points, key=lambda item: item[2])
    return f"""
    <svg viewBox="0 0 {width} {height}" class="timeline" role="img" aria-label="Negative mentions by hour">
      <defs>
        <linearGradient id="areaFill" x1="0" x2="0" y1="0" y2="1">
          <stop offset="0%" stop-color="#d80019" stop-opacity=".34"/>
          <stop offset="100%" stop-color="#d80019" stop-opacity=".03"/>
        </linearGradient>
      </defs>
      {grid}
      <polygon points="{area} {base}" fill="url(#areaFill)"/>
      <polyline points="{area}" fill="none" stroke="#d80019" stroke-width="4" stroke-linejoin="round" stroke-linecap="round"/>
      {circles}
      {labels}
      <text x="{peak[0] + 14:.1f}" y="{max(18, peak[1] - 12):.1f}" class="peak">PEAK {peak[3]}:00</text>
    </svg>"""


def build_html() -> str:
    records = pd.concat([dash.load_original_records(), dash.load_new_records()], ignore_index=True)
    summary = dash.build_summary(records)
    neg = records[records["is_negative_final"] == "Yes"].copy()
    neg["dt"] = neg["post_time"].map(parse_time)

    platform_rows = summary["platforms"]
    top_platform = max(platform_rows, key=lambda row: row["negative"])
    fb = next(row for row in platform_rows if row["platform"] == "Facebook")
    xrow = next(row for row in platform_rows if row["platform"] == "X")
    peak_hour = int(neg.groupby(neg["dt"].dt.hour).size().idxmax())
    peak_date = neg.groupby(neg["dt"].dt.date).size().idxmax()
    hour_counts = neg.groupby(neg["dt"].dt.hour).size().to_dict()
    issue_total = max(summary["negative"], 1)
    leaders = summary["leaders"][:2]
    examples = neg.sort_values(["negative_score", "engagement"], ascending=False).head(4)

    channel_rows = [
        {"label": row["platform"], "value": row["negative"]}
        for row in platform_rows
        if row["negative"] > 0
    ]
    topic_rows = [
        {"label": "ช่องทางรับชม / ดูไม่ได้", "value": summary["access"]},
        {"label": "ซิม True / แพ็กเกจ", "value": summary["true_package"]},
        {"label": "ดูไม่ได้ / ต้องสมัคร", "value": summary["unable"]},
        {"label": "ถ้อยคำร้องเรียนตรง", "value": summary["complaint"]},
    ]
    leader_rows = [
        f"""
        <tr>
          <td>{html.escape(row["platform"])}</td>
          <td><strong>{html.escape(short(row["account"], 26))}</strong></td>
          <td>{row["negative_posts"]}</td>
          <td>{row["total_negative_engagement"]}</td>
          <td class="hot">{int(row["impact_score"])}</td>
        </tr>"""
        for row in leaders
    ]
    quote_cards = [
        f"""
        <div class="quote-card">
          <b>{html.escape(short(row.account, 26))}</b>
          <p>“{html.escape(short(row.message, 104))}”</p>
        </div>"""
        for row in examples.itertuples()
    ]

    negative_share_of_negative = summary["negative"] / summary["total"] if summary["total"] else 0
    fb_share = fb["negative"] / issue_total
    access_share = summary["access"] / issue_total
    true_share = summary["true_package"] / issue_total

    return f"""<!doctype html>
<html lang="th">
<head>
<meta charset="utf-8">
<title>Negative One Page Summary - True AF 2026</title>
<style>
  * {{ box-sizing: border-box; }}
  body {{ margin:0; background:#111; font-family:Tahoma, Arial, sans-serif; }}
  .page {{
    width:1600px; height:900px; overflow:hidden; color:#171717; background:#f7f7f7;
    border:10px solid #050505; position:relative;
  }}
  .hero {{
    height:116px; color:#fff; padding:15px 26px 12px 28px; position:relative;
    background:
      linear-gradient(90deg, rgba(0,0,0,.95), rgba(88,0,0,.86) 48%, rgba(214,0,30,.68)),
      linear-gradient(160deg, #050505 0%, #360006 52%, #b2001b 100%);
  }}
  .hero::after {{
    content:""; position:absolute; inset:0 0 auto auto; width:660px; height:116px;
    background:repeating-linear-gradient(86deg, rgba(255,255,255,.18) 0 2px, transparent 2px 36px);
    opacity:.23;
  }}
  .brand {{ display:flex; gap:22px; align-items:center; position:relative; z-index:1; }}
  .logo {{ font-size:46px; line-height:.9; color:#e60012; font-weight:900; letter-spacing:0; }}
  .logo span {{ display:block; color:#fff; font-size:15px; font-weight:700; margin-left:4px; }}
  h1 {{ margin:0; font-size:31px; line-height:1.12; letter-spacing:0; }}
  .sub {{ margin-top:7px; font-size:16px; color:#ffdce0; font-weight:700; }}
  .date {{ position:absolute; right:28px; top:18px; z-index:2; font-size:18px; font-weight:800; }}
  .strip {{
    height:30px; display:flex; align-items:center; gap:22px; padding:0 30px;
    background:#c60018; color:#fff; font-size:16px; font-weight:800;
  }}
  .grid {{ display:grid; gap:10px; padding:10px 12px 0; }}
  .top {{ grid-template-columns:1.35fr repeat(4, .78fr); height:164px; }}
  .middle {{ grid-template-columns:.9fr 1.02fr 1.28fr 1.5fr; height:300px; }}
  .bottom {{ grid-template-columns:1.45fr 1fr; height:246px; padding-bottom:10px; }}
  .panel {{
    background:#fff; border:2px solid #151515; border-radius:10px; padding:13px 14px;
    box-shadow:0 2px 0 rgba(0,0,0,.1); overflow:hidden;
  }}
  .panel h2 {{ margin:0 0 8px; color:#c60018; font-size:20px; line-height:1.1; }}
  .panel h3 {{ margin:0 0 8px; color:#222; font-size:16px; }}
  .summary ul {{ margin:4px 0 0 18px; padding:0; font-size:14px; line-height:1.3; }}
  .summary li {{ margin:2px 0; }}
  .kpi {{ display:flex; flex-direction:column; justify-content:center; align-items:center; text-align:center; }}
  .kpi .icon {{ width:56px; height:56px; border-radius:14px; background:#d80019; color:white; display:grid; place-items:center; font-size:30px; margin-bottom:5px; }}
  .kpi .label {{ font-size:15px; font-weight:900; color:#111; }}
  .kpi .value {{ font-size:42px; font-weight:900; color:#d80019; line-height:1; margin-top:6px; }}
  .kpi .note {{ color:#333; font-size:14px; font-weight:800; margin-top:5px; }}
  .peak-card {{ background:radial-gradient(circle at 30% 15%, #ff5a64 0, #ba0018 44%, #60000c 100%); color:#fff; border-color:#ff8290; }}
  .peak-card .value, .peak-card .note {{ color:#ffd84a; }}
  .donut-wrap {{ display:flex; align-items:center; justify-content:center; gap:18px; height:202px; }}
  .donut {{
    width:156px; height:156px; border-radius:50%;
    background:conic-gradient(#e60012 0 88.7%, #ffb400 88.7% 95%, #28b36d 95% 100%);
    display:grid; place-items:center;
  }}
  .donut::before {{ content:""; width:78px; height:78px; border-radius:50%; background:#fff; box-shadow:inset 0 0 0 18px #fff; }}
  .legend {{ display:grid; gap:12px; font-size:15px; font-weight:800; }}
  .dot {{ display:inline-block; width:14px; height:14px; border-radius:50%; margin-right:7px; vertical-align:middle; }}
  .meter {{ margin-top:10px; display:grid; gap:9px; }}
  .meter-row {{ display:grid; grid-template-columns:132px 1fr 48px; align-items:center; gap:8px; font-size:14px; font-weight:800; }}
  .meter-track {{ height:12px; background:#eceff3; border-radius:999px; overflow:hidden; }}
  .meter-fill {{ height:100%; background:#d80019; border-radius:999px; }}
  .bar-row {{ display:grid; grid-template-columns:26px 156px 1fr 52px; gap:8px; align-items:center; margin:7px 0; font-size:15px; }}
  .rank {{ width:24px; height:24px; border-radius:50%; background:#d80019; color:#fff; display:grid; place-items:center; font-weight:900; }}
  .bar-label {{ font-weight:800; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
  .bar-track {{ height:18px; background:#edf0f3; border-radius:999px; overflow:hidden; }}
  .bar-fill {{ height:100%; border-radius:999px; }}
  .bar-value {{ text-align:right; font-weight:900; color:#c60018; }}
  .axis {{ font:12px Tahoma, Arial; fill:#333; font-weight:700; }}
  .grid {{ stroke:#e5e7eb; stroke-width:1; }}
  .peak {{ font:700 15px Tahoma, Arial; fill:#d80019; }}
  table {{ width:100%; border-collapse:collapse; font-size:13px; }}
  th {{ text-align:left; background:#111827; color:#fff; padding:5px 6px; }}
  td {{ border-bottom:1px solid #e5e7eb; padding:5px 6px; vertical-align:top; }}
  .hot {{ color:#d80019; font-weight:900; }}
  .quotes {{ display:grid; grid-template-columns:1fr 1fr; gap:9px; }}
  .quote-card {{ border:1px solid #d8dde6; border-left:7px solid #d80019; border-radius:8px; padding:8px 10px; min-height:78px; }}
  .quote-card b {{ color:#d80019; font-size:14px; }}
  .quote-card p {{ margin:4px 0 0; font-size:12.3px; line-height:1.28; }}
  .actions {{ display:grid; grid-template-columns:38px 1fr; gap:9px; align-items:center; margin-top:8px; font-size:16px; }}
  .num {{ width:34px; height:34px; border-radius:50%; background:#d80019; color:white; display:grid; place-items:center; font-weight:900; font-size:18px; }}
  .footer {{ position:absolute; left:20px; right:20px; bottom:6px; color:#fff; font-size:12px; display:flex; justify-content:space-between; opacity:.9; }}
</style>
</head>
<body>
<div class="page">
  <div class="hero">
    <div class="brand">
      <div class="logo">true<span>together</span></div>
      <div>
        <h1>รายงานสรุปสถานการณ์: Negative Feedback<br>True Academy Fantasia 2026</h1>
        <div class="sub">หัวข้อหลัก: ไม่มีซิม True / ดู AF ไม่ได้ / ช่องทางรับชมเข้าถึงยาก</div>
      </div>
    </div>
    <div class="date">13-14 มิถุนายน 2569</div>
  </div>
  <div class="strip">
    <span>Social Listening: Facebook, X, IG, TikTok, YouTube</span>
    <span>Data period: 1-14 June 2026</span>
    <span>Negative One Page Summary</span>
  </div>

  <div class="grid top">
    <div class="panel summary">
      <h2>Executive Summary</h2>
      <ul>
        <li>Negative ที่เกี่ยวข้องมี <b>{summary["negative"]:,}</b> mentions จากทั้งหมด {summary["total"]:,} mentions ({pct(negative_share_of_negative, 1)} ของข้อมูลรวม)</li>
        <li>ปัญหาหลักไม่ใช่ตัวรายการ แต่เป็น friction ในการรับชมคอนเสิร์ตและเงื่อนไข True/dtac</li>
        <li>Facebook เป็นศูนย์กลางเสียงบ่นหลัก คำซ้ำมากคือ “ดูไม่ได้/ดูยาก/ต้องใช้”</li>
      </ul>
    </div>
    <div class="panel kpi">
      <div class="icon">!</div>
      <div class="label">Relevant Negative</div>
      <div class="value">{summary["negative"]:,}</div>
      <div class="note">mentions</div>
    </div>
    <div class="panel kpi">
      <div class="icon">☹</div>
      <div class="label">Access Issue</div>
      <div class="value">{pct(access_share)}</div>
      <div class="note">{summary["access"]:,} mentions</div>
    </div>
    <div class="panel kpi">
      <div class="icon">f</div>
      <div class="label">Platform สูงสุด</div>
      <div class="value">{pct(fb_share)}</div>
      <div class="note">{top_platform["platform"]}: {fb["negative"]} mentions</div>
    </div>
    <div class="panel kpi peak-card">
      <div class="icon">⏱</div>
      <div class="label">Peak Time</div>
      <div class="value">{peak_hour:02d}:00</div>
      <div class="note">วันที่ {peak_date.strftime("%d/%m/%Y")}</div>
    </div>
  </div>

  <div class="grid middle">
    <div class="panel">
      <h2>Issue Overview</h2>
      <div class="donut-wrap">
        <div class="donut"></div>
        <div class="legend">
          <div><span class="dot" style="background:#e60012"></span>Facebook {fb["negative"]}</div>
          <div><span class="dot" style="background:#ffb400"></span>X {xrow["negative"]}</div>
          <div><span class="dot" style="background:#28b36d"></span>Other {summary["negative"] - fb["negative"] - xrow["negative"]}</div>
        </div>
      </div>
      <div class="meter">
        <div class="meter-row"><span>True/package</span><div class="meter-track"><div class="meter-fill" style="width:{true_share * 100:.1f}%"></div></div><b>{pct(true_share)}</b></div>
      </div>
    </div>
    <div class="panel">
      <h2>Top Keywords</h2>
      {bar_rows(summary["keywords"], "keyword", "count", 6)}
    </div>
    <div class="panel">
      <h2>Hot Topics</h2>
      {bar_rows(topic_rows, "label", "value", 4)}
      <h3 style="margin-top:14px">Amplification Accounts</h3>
      <table><thead><tr><th>Channel</th><th>Account</th><th>Neg.</th><th>Eng.</th><th>Impact</th></tr></thead><tbody>{''.join(leader_rows)}</tbody></table>
    </div>
    <div class="panel">
      <h2>Timeline: Negative Mentions by Hour</h2>
      {timeline_svg(hour_counts)}
      <h3>Interpretation</h3>
      <p style="font-size:15px;line-height:1.42;margin:0">เสียงลบกระจุกในช่วงคอนเสิร์ต/Live โดยเฉพาะ {peak_hour:02d}:00 น. สะท้อนว่าปัญหาเกิดใน moment ที่ผู้ชมตั้งใจจะเข้าไปดูและพร้อมมีส่วนร่วม</p>
    </div>
  </div>

  <div class="grid bottom">
    <div class="panel">
      <h2>Voice of Customer</h2>
      <div class="quotes">{''.join(quote_cards)}</div>
    </div>
    <div class="panel">
      <h2>ข้อเสนอแนะสำหรับผู้บริหาร</h2>
      <div class="actions"><div class="num">1</div><div><b>Communication Clarity</b><br>ระบุเงื่อนไข “ดูได้/ดูไม่ได้/ต้องใช้ซิมหรือแพ็กเกจใด” ให้ชัดก่อนวัน Live</div></div>
      <div class="actions"><div class="num">2</div><div><b>UX/UI Optimization</b><br>ทำทางเข้า TrueVisions Now ให้ตรงและง่ายช่วง peak hour พร้อมข้อความช่วยเหลือ</div></div>
      <div class="actions"><div class="num">3</div><div><b>Support Escalation</b><br>ตั้ง war room real-time สำหรับปัญหา login, package, telco condition และ refund case</div></div>
    </div>
  </div>
  <div class="footer"><span>Source: Social Listening Dashboard | Generated from true_af_nlp_dashboard.html data model</span><span class="logo" style="font-size:26px">true</span></div>
</div>
</body>
</html>"""


async def render_png() -> None:
    OUT_DIR.mkdir(exist_ok=True)
    HTML_OUT.write_text(build_html(), encoding="utf-8")
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page(viewport={"width": 1600, "height": 900}, device_scale_factor=1)
        await page.goto(HTML_OUT.resolve().as_uri(), wait_until="networkidle")
        await page.screenshot(path=str(PNG_OUT), full_page=False)
        await browser.close()


if __name__ == "__main__":
    asyncio.run(render_png())
    print(f"Wrote {HTML_OUT}")
    print(f"Wrote {PNG_OUT}")
