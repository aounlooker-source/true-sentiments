from __future__ import annotations

import json
from collections import Counter
from pathlib import Path
from typing import Any

import pandas as pd

from build_nlp_dashboard import (
    ACCESS_TERMS,
    COMPLAINT_TERMS,
    NEGATIVE_TERMS,
    TEXT_COL,
    TRUE_PACKAGE_TERMS,
    hits,
)


INPUT_ORIGINAL = Path("aouy3.xlsx")
INPUT_AOUY_PAGE = Path("aouy.xlsx")
INPUT_2 = Path("aouy_2.xlsx")
OUTPUT_DIR = Path("outputs")
OUTPUT = OUTPUT_DIR / "true_af_nlp_dashboard.html"

PLATFORMS = ["Facebook", "X", "IG", "TikTok", "YouTube", "Pantip", "Other / Website"]
AOUY_PAGE_NAME = "อวยใส้แหก"

STRICT_NEGATIVE_TERMS = [
    "ดูไม่ได้",
    "ไม่ได้ดู",
    "หาช่องดูไม่ได้",
    "หาช่องดูยาก",
    "ดูยาก",
    "เข้าถึงยาก",
    "ยุ่งยาก",
    "ไม่อยาก",
    "ไม่ดู",
    "ไม่เคลียร์",
    "ไม่ชัดเจน",
    "เสียตังค์",
    "เสียเงิน",
    "จำกัด",
    "ไม่ใช่ลูกค้าทรู",
    "ต้องสมัคร",
    "ต้องใช้",
    "คืนตัง",
    "ไม่ผ่าน",
    "มั่ว",
    "เงียบ",
    "แย่",
    "ด่า",
    "บ่น",
    "ไม่ควร",
    "เข้าไม่ถึง",
    "ลำบาก",
    "ค้าง",
    "กระตุก",
    "ผิดพลาด",
    "ผิด",
    "เหี้ย",
    "เลวร้าย",
    "เซ็ง",
    "เซง",
    "อึน",
    "เสียดาย",
]

COMPLAINT_SIGNAL_TERMS = [
    "ดูไม่ได้",
    "ไม่ได้ดู",
    "หาช่องดูไม่ได้",
    "หาช่องดูยาก",
    "ดูยาก",
    "เข้าถึงยาก",
    "กดไม่ได้",
    "สมัครดูไม่ได้",
    "เข้าไม่ได้",
    "ไม่ใช่ลูกค้าทรู",
    "ไม่ได้ใช้ซิมทรู",
    "ไม่มีซิมทรู",
    "ต้องใช้ซิมทรู",
    "ต้องใช้ true",
    "ต้องใช้ทรู",
    "เสียความรู้สึก",
    "หลอก",
    "ไม่เคลียร์",
    "ไม่ชัดเจน",
    "ยุ่งยาก",
    "ลำบาก",
    "ค้าง",
    "กระตุก",
    "ระบบภาพ",
    "แย่",
    "เลวร้าย",
    "ด่า",
    "บ่น",
    "เซ็ง",
    "เซง",
    "อึน",
    "ไม่ดู",
    "ไม่อยากดู",
    "คืนตัง",
]

HARD_COMPLAINT_TERMS = [
    "ด่า",
    "บ่น",
    "แย่",
    "เลวร้าย",
    "เหี้ย",
    "เซ็ง",
    "เซง",
    "อึน",
    "เสียความรู้สึก",
    "หลอก",
    "ไม่เคลียร์",
    "ไม่ชัดเจน",
    "ผิดพลาด",
    "ค้าง",
    "กระตุก",
    "กดไม่ได้",
]

PROMO_CONTEXT_TERMS = [
    "ลุ้นรับ",
    "ร่วมกิจกรรม",
    "กติกา",
    "รางวัล",
    "ประกาศรายชื่อ",
    "ยืนยันรับสิทธิ์",
    "เงื่อนไขกิจกรรม",
    "คำตัดสิน",
    "ของรางวัล",
    "กด like",
    "share โพสต์",
]

STRICT_ACCESS_TERMS = [
    "ดูไม่ได้",
    "ไม่ได้ดู",
    "หาช่องดูไม่ได้",
    "หาช่องดูยาก",
    "ดูยาก",
    "เข้าถึงยาก",
    "ช่องทาง",
    "รับชม",
    "ถ่ายทอดสด",
    "ฟรีทีวี",
    "แอป",
    "app",
    "now",
    "รีรัน",
    "คอนเสิร์ต",
    "concert",
    "ดูสด",
]

STRICT_TRUE_TERMS = [
    "ทรู",
    "true",
    "dtac",
    "ดีเทค",
    "ซิม",
    "ลูกค้าทรู",
    "แพ็ก",
    "แพค",
    "package",
    "99",
    "199",
    "exclusive",
    "เสียตังค์",
    "เสียเงิน",
    "จ่าย",
    "สมัคร",
]


def clean_value(value: Any, default: str = "") -> str:
    if pd.isna(value):
        return default
    text = str(value).strip()
    return default if text in {"", "-", "nan", "NaN"} else text


def to_int(value: Any) -> int:
    if pd.isna(value):
        return 0
    try:
        return int(float(str(value).replace(",", "")))
    except ValueError:
        return 0


def normalize_platform(source: Any, url: Any = "") -> str:
    source_text = clean_value(source).lower()
    url_text = clean_value(url).lower()
    joined = f"{source_text} {url_text}"
    if "facebook" in joined or source_text == "fb":
        return "Facebook"
    if source_text == "x" or "twitter.com" in joined or "x.com" in joined:
        return "X"
    if "instagram" in joined or source_text in {"ig", "insta"}:
        return "IG"
    if "tiktok" in joined:
        return "TikTok"
    if "youtube" in joined or "youtu.be" in joined:
        return "YouTube"
    if "pantip" in joined or "forum" in joined:
        return "Pantip"
    return "Other / Website"


def sentiment_from_score(score: int, source_sentiment: str) -> str:
    if score >= 70:
        return "Very Negative"
    if score >= 40 or source_sentiment.lower() == "negative":
        return "Negative"
    if score >= 20:
        return "Mixed / Concern"
    return "Neutral / Other"


def smart_hits(text: str, terms: list[str]) -> list[str]:
    lower = text.lower()
    found: list[str] = []
    for term in terms:
        needle = term.lower()
        if needle in lower and term not in found:
            found.append(term)
    return found


def final_classification(text: str, source_sentiment: str) -> dict[str, Any]:
    negative_hits = smart_hits(text, STRICT_NEGATIVE_TERMS)
    access_hits = smart_hits(text, STRICT_ACCESS_TERMS)
    true_hits = smart_hits(text, STRICT_TRUE_TERMS)
    complaint_hits = smart_hits(text, COMPLAINT_TERMS + ["ค้าง", "กระตุก", "ผิดพลาด", "เหี้ย", "เลวร้าย"])
    unable_hits = smart_hits(text, ["ไม่ได้ดู", "ดูไม่ได้", "หาช่องดูไม่ได้", "หาช่องดูยาก", "ดูทางไหน", "ต้องสมัคร", "ไม่ใช่ลูกค้าทรู"])
    signal_hits = smart_hits(text, COMPLAINT_SIGNAL_TERMS)
    hard_complaint_hits = smart_hits(text, HARD_COMPLAINT_TERMS)
    promo_hits = smart_hits(text, PROMO_CONTEXT_TERMS)

    source_negative = source_sentiment.lower() == "negative"
    has_complaint_signal = bool(signal_hits or unable_hits or hard_complaint_hits)
    access_issue = bool(access_hits and has_complaint_signal)
    true_issue = bool(true_hits and has_complaint_signal)
    relevant_feedback = bool(unable_hits or access_issue or true_issue)

    # Promo/activity posts often contain words such as "package", "customer", or
    # "limited" but are not feedback about watching problems. Keep them out unless
    # they contain an explicit complaint signal.
    if promo_hits and not bool(signal_hits or unable_hits or hard_complaint_hits):
        negative_hits = []
        access_hits = []
        true_hits = []
        complaint_hits = []
        unable_hits = []
        access_issue = False
        true_issue = False
        relevant_feedback = False

    score = 0
    score += 14 * len(negative_hits)
    score += 8 * len(access_hits) if access_issue else 0
    score += 8 * len(true_hits) if true_issue else 0
    score += 10 * len(complaint_hits)
    score += 12 if unable_hits else 0
    if source_negative and relevant_feedback:
        score = max(score, 55)
    if not relevant_feedback:
        score = 0
    score = min(100, score)
    sentiment = sentiment_from_score(score, source_sentiment)
    if not relevant_feedback:
        sentiment = "Neutral / Other"

    return {
        "negative_score": score,
        "sentiment": sentiment,
        "negative_terms": ", ".join(negative_hits),
        "access_terms": ", ".join(access_hits) if access_issue else "",
        "true_package_terms": ", ".join(true_hits) if true_issue else "",
        "complaint_terms": ", ".join(complaint_hits),
        "is_access_complaint": "Yes" if access_issue else "No",
        "is_true_package_issue": "Yes" if true_issue else "No",
        "is_unable_to_watch": "Yes" if unable_hits else "No",
        "is_direct_complaint_language": "Yes" if complaint_hits else "No",
        "source_sentiment": source_sentiment or "Unknown",
        "is_relevant_feedback": "Yes" if relevant_feedback else "No",
        "is_negative_final": "Yes" if sentiment in {"Very Negative", "Negative"} and relevant_feedback else "No",
    }


def load_original_records() -> pd.DataFrame:
    df = pd.read_excel(INPUT_ORIGINAL, sheet_name="Comments_Data")
    rows: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        text = clean_value(row.get(TEXT_COL))
        cls = final_classification(text, "")
        rows.append(
            {
                "dataset": "Original Facebook comments",
                "platform": "Facebook",
                "account": clean_value(row.get("ชื่อผู้แสดงความคิดเห็น"), "UNKNOWN"),
                "message": text,
                "url": "",
                "post_url": "",
                "post_time": clean_value(row.get("เวลาที่แสดงความคิดเห็น")),
                "engagement": to_int(row.get("จำนวนกดไลค์และแสดงความรู้สึก")),
                "source_sentiment": "",
                "category": "Original comment export",
                **cls,
            }
        )
    return pd.DataFrame(rows)


def load_aouy_page_records() -> pd.DataFrame:
    df = pd.read_excel(INPUT_AOUY_PAGE, sheet_name="Comments_Data")
    account_col = df.columns[2]
    time_col = df.columns[4]
    engagement_col = df.columns[6]
    rows: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        text = clean_value(row.get(TEXT_COL))
        cls = final_classification(text, "")
        rows.append(
            {
                "dataset": AOUY_PAGE_NAME,
                "platform": "Facebook",
                "page": AOUY_PAGE_NAME,
                "account": clean_value(row.get(account_col), "UNKNOWN"),
                "message": text,
                "url": "",
                "post_url": "",
                "post_time": clean_value(row.get(time_col)),
                "engagement": to_int(row.get(engagement_col)),
                "source_sentiment": "",
                "category": "Page comment export",
                **cls,
            }
        )
    return pd.DataFrame(rows)


def load_new_records() -> pd.DataFrame:
    df = pd.read_excel(INPUT_2, sheet_name="all")
    rows: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        text = clean_value(row.get("Message"))
        source_sentiment = clean_value(row.get("Sentiment"), "Unknown")
        source = clean_value(row.get("Source"))
        url = clean_value(row.get("Direct URL")) or clean_value(row.get("Post URL"))
        platform = normalize_platform(source, url)
        cls = final_classification(text, source_sentiment)
        rows.append(
            {
                "dataset": "",
                "platform": platform,
                "account": clean_value(row.get("Account"), "UNKNOWN"),
                "message": text,
                "url": url,
                "post_url": clean_value(row.get("Post URL")),
                "post_time": clean_value(row.get("Post time")),
                "engagement": to_int(row.get("Engagement")),
                "source_sentiment": source_sentiment,
                "category": clean_value(row.get("Category")),
                "follower_count": to_int(row.get("Follower count")),
                **cls,
            }
        )
    return pd.DataFrame(rows)


def platform_summary(records: pd.DataFrame) -> list[dict[str, Any]]:
    rows = []
    for platform in PLATFORMS:
        df = records[records["platform"] == platform]
        total = len(df)
        negative = int((df["is_negative_final"] == "Yes").sum())
        very_negative = int((df["sentiment"] == "Very Negative").sum())
        access = int((df["is_access_complaint"] == "Yes").sum())
        true_pkg = int((df["is_true_package_issue"] == "Yes").sum())
        engagement = int(df.loc[df["is_negative_final"] == "Yes", "engagement"].sum()) if total else 0
        avg = round(float(df["negative_score"].mean()), 1) if total else 0
        rows.append(
            {
                "platform": platform,
                "total": total,
                "negative": negative,
                "negative_share": negative / total if total else 0,
                "very_negative": very_negative,
                "access": access,
                "true_package": true_pkg,
                "negative_engagement": engagement,
                "avg_score": avg,
            }
        )
    return rows


def top_keywords(df: pd.DataFrame, limit: int = 12) -> list[dict[str, Any]]:
    counter: Counter[str] = Counter()
    for text in df.loc[df["is_negative_final"] == "Yes", "negative_terms"].fillna(""):
        for term in [item.strip() for item in str(text).split(",") if item.strip()]:
            counter[term] += 1
    return [{"keyword": keyword, "count": count} for keyword, count in counter.most_common(limit)]


def account_leaders(records: pd.DataFrame, limit: int = 20) -> list[dict[str, Any]]:
    neg = records[records["is_negative_final"] == "Yes"].copy()
    if neg.empty:
        return []
    known = neg["account"].fillna("").astype(str).str.strip().str.lower()
    owned_accounts = {
        "trueaf thailand",
        "truevisions",
        "truevisionsofficial",
        "truevisions_now",
        "truemove h",
        "truemoveh",
        "dtac",
        "true corporation",
    }
    msg_lower = neg["message"].fillna("").astype(str).str.lower()
    msg_no_hashtags = msg_lower.str.replace(r"#\S+", "", regex=True)
    core_relevant = msg_no_hashtags.str.contains(
        "af2026|trueaf|true af|academy fantasia|truevisionsnow|truevisions now|นักล่าฝัน|คอนเสิร์ต|บ้าน af|บ้านปีนี้",
        regex=True,
        na=False,
    )
    hashtag_relevant_with_real_complaint = msg_lower.str.contains("af2026|trueaf|#af", regex=True, na=False) & (
        neg["negative_terms"].fillna("").astype(str).str.len() > 0
    )
    relevant = core_relevant | hashtag_relevant_with_real_complaint
    owned_like = known.str.startswith("true") | known.str.startswith("dtac")
    neg = neg[(known != "unknown") & (known != "-") & (~known.isin(owned_accounts)) & (~owned_like) & relevant]
    if neg.empty:
        return []
    grouped = (
        neg.groupby(["platform", "account"], dropna=False)
        .agg(
            negative_posts=("message", "count"),
            total_negative_engagement=("engagement", "sum"),
            avg_negative_score=("negative_score", "mean"),
            top_message=("message", lambda s: str(s.iloc[0])[:280]),
            url=("url", lambda s: str(s.iloc[0])),
        )
        .reset_index()
    )
    grouped["impact_score"] = (
        grouped["negative_posts"] * 10
        + grouped["total_negative_engagement"].clip(upper=5000) / 50
        + grouped["avg_negative_score"]
    )
    grouped = grouped.sort_values(["impact_score", "total_negative_engagement", "negative_posts"], ascending=False)
    grouped["avg_negative_score"] = grouped["avg_negative_score"].round(1)
    grouped["impact_score"] = grouped["impact_score"].round(1)
    return grouped.head(limit).to_dict("records")


def platform_payload(records: pd.DataFrame) -> dict[str, Any]:
    payload: dict[str, Any] = {}
    for platform in PLATFORMS:
        df = records[records["platform"] == platform].copy()
        neg = df[df["is_negative_final"] == "Yes"].sort_values(["negative_score", "engagement"], ascending=False)
        payload[platform] = {
            "keywords": top_keywords(df),
            "top_negative": neg.head(80)[
                [
                    "negative_score",
                    "sentiment",
                    "account",
                    "post_time",
                    "engagement",
                    "source_sentiment",
                    "is_access_complaint",
                    "is_true_package_issue",
                    "negative_terms",
                    "message",
                    "url",
                    "dataset",
                ]
            ].to_dict("records"),
            "leaders": account_leaders(df, 10),
        }
    return payload


def page_payload(records: pd.DataFrame) -> dict[str, Any]:
    neg = records[records["is_negative_final"] == "Yes"].sort_values(["negative_score", "engagement"], ascending=False)
    return {
        "summary": {
            "total": len(records),
            "negative": int((records["is_negative_final"] == "Yes").sum()),
            "negative_share": int((records["is_negative_final"] == "Yes").sum()) / len(records) if len(records) else 0,
            "very_negative": int((records["sentiment"] == "Very Negative").sum()),
            "access": int((records["is_access_complaint"] == "Yes").sum()),
            "true_package": int((records["is_true_package_issue"] == "Yes").sum()),
            "unable": int((records["is_unable_to_watch"] == "Yes").sum()),
            "complaint": int((records["is_direct_complaint_language"] == "Yes").sum()),
            "avg_score": round(float(records["negative_score"].mean()), 1) if len(records) else 0,
        },
        "keywords": top_keywords(records, 12),
        "leaders": account_leaders(records, 12),
        "top_negative": neg.head(80)[
            [
                "negative_score",
                "sentiment",
                "account",
                "post_time",
                "engagement",
                "source_sentiment",
                "is_access_complaint",
                "is_true_package_issue",
                "negative_terms",
                "message",
                "url",
                "dataset",
            ]
        ].to_dict("records"),
    }


def build_summary(records: pd.DataFrame) -> dict[str, Any]:
    total = len(records)
    negative = int((records["is_negative_final"] == "Yes").sum())
    very_negative = int((records["sentiment"] == "Very Negative").sum())
    access = int((records["is_access_complaint"] == "Yes").sum())
    true_pkg = int((records["is_true_package_issue"] == "Yes").sum())
    unable = int((records["is_unable_to_watch"] == "Yes").sum())
    complaint = int((records["is_direct_complaint_language"] == "Yes").sum())
    return {
        "total": total,
        "negative": negative,
        "negative_share": negative / total if total else 0,
        "very_negative": very_negative,
        "access": access,
        "true_package": true_pkg,
        "unable": unable,
        "complaint": complaint,
        "avg_score": round(float(records["negative_score"].mean()), 1),
        "platforms": platform_summary(records),
        "leaders": account_leaders(records, 25),
        "keywords": top_keywords(records, 20),
    }


def rows_for_table(records: pd.DataFrame, limit: int = 300) -> list[dict[str, Any]]:
    cols = [
        "dataset",
        "platform",
        "negative_score",
        "sentiment",
        "account",
        "post_time",
        "engagement",
        "source_sentiment",
        "is_access_complaint",
        "is_true_package_issue",
        "is_unable_to_watch",
        "negative_terms",
        "message",
        "url",
    ]
    return (
        records[records["is_negative_final"] == "Yes"]
        .sort_values(["negative_score", "engagement"], ascending=False)
        .head(limit)[cols]
        .to_dict("records")
    )


def build_html(data: dict[str, Any]) -> str:
    data_json = json.dumps(data, ensure_ascii=False)
    nav_buttons = "\n".join(
        [f'<button class="tab-btn" data-tab="{p}">{p}</button>' for p in PLATFORMS]
    )
    platform_sections = "\n".join(
        [
            f"""
      <section class="tab-panel platform-panel" id="tab-{p}">
        <div class="section-head">
          <h2>{p}</h2>
          <span class="muted" id="meta-{p}"></span>
        </div>
        <div class="kpis" id="kpis-{p}"></div>
        <div class="split">
          <div>
            <h3>Negative Keywords</h3>
            <div class="bars compact" id="keywords-{p}"></div>
          </div>
          <div>
            <h3>Accounts ที่นำกระแสเชิงลบ</h3>
            <div class="table-wrap small"><table><thead><tr><th>Account</th><th>Neg.</th><th>Eng.</th><th>Score</th><th>ตัวอย่าง</th></tr></thead><tbody id="leaders-{p}"></tbody></table></div>
          </div>
        </div>
        <h3>Top Negative Mentions</h3>
        <div class="table-wrap"><table><thead><tr><th>Score</th><th>Sentiment</th><th>Account</th><th>Eng.</th><th>Issue</th><th>Message</th></tr></thead><tbody id="rows-{p}"></tbody></table></div>
      </section>"""
            for p in PLATFORMS
        ]
    )

    return f"""<!doctype html>
<html lang="th">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Negative Feedback จากหัวข้อ ไม่มีซิม True ดู AF ไม่ได้</title>
  <style>
    :root {{
      --ink:#17212b; --muted:#64748b; --line:#d8e0e7; --page:#f4f7f9; --panel:#fff;
      --teal:#14716d; --red:#c2412d; --orange:#e1812c; --yellow:#d8a625; --blue:#315d85; --violet:#6d4bc2;
    }}
    * {{ box-sizing: border-box; }}
    body {{ margin:0; font-family:Tahoma, Arial, sans-serif; color:var(--ink); background:var(--page); }}
    header {{ background:linear-gradient(135deg,#143642 0%,#1f7a7a 100%); color:white; padding:28px 34px 22px; }}
    header h1 {{ margin:0; font-size:28px; line-height:1.2; letter-spacing:0; }}
    .date-range {{ margin-top:8px; font-size:14px; font-weight:700; color:#ffffff; opacity:.92; }}
    header p {{ max-width:1120px; margin:10px 0 0; color:#dceff0; line-height:1.65; }}
    main {{ padding:20px 28px 36px; max-width:1540px; margin:0 auto; }}
    .tabs {{ display:flex; flex-wrap:wrap; gap:8px; margin:0 0 16px; position:sticky; top:0; background:var(--page); padding:8px 0; z-index:5; }}
    .tab-btn {{ border:1px solid #bfd0da; background:#fff; color:#143642; border-radius:7px; padding:9px 12px; cursor:pointer; font-family:inherit; font-weight:700; }}
    .tab-btn.active {{ background:#143642; color:#fff; border-color:#143642; }}
    .tab-panel {{ display:none; }}
    .tab-panel.active {{ display:block; }}
    .kpis {{ display:grid; grid-template-columns:repeat(6,minmax(150px,1fr)); gap:12px; margin-bottom:18px; }}
    .kpi {{ background:var(--panel); border:1px solid var(--line); border-radius:8px; padding:13px 14px; box-shadow:0 10px 22px rgba(15,23,42,.05); }}
    .kpi span {{ display:block; font-size:12px; color:var(--muted); margin-bottom:8px; }}
    .kpi strong {{ font-size:22px; line-height:1.1; }}
    section {{ background:var(--panel); border:1px solid var(--line); border-radius:8px; padding:16px; box-shadow:0 10px 22px rgba(15,23,42,.05); margin-bottom:16px; }}
    h2 {{ font-size:18px; margin:0 0 12px; }}
    h3 {{ font-size:14px; margin:10px 0 10px; color:#143642; }}
    .section-head {{ display:flex; justify-content:space-between; align-items:flex-start; gap:12px; }}
    .grid {{ display:grid; grid-template-columns:1fr 1fr; gap:16px; align-items:start; }}
    .split {{ display:grid; grid-template-columns:.85fr 1.15fr; gap:16px; }}
    .bars {{ display:grid; gap:10px; }}
    .bar-row {{ display:grid; grid-template-columns:190px 1fr 86px; gap:12px; align-items:center; }}
    .compact .bar-row {{ grid-template-columns:130px 1fr 56px; }}
    .bar-track {{ height:18px; background:#edf2f6; border-radius:999px; overflow:hidden; }}
    .bar-fill {{ height:100%; border-radius:999px; background:var(--teal); }}
    .bar-label {{ font-size:13px; color:#334155; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }}
    .bar-value {{ text-align:right; color:#334155; font-weight:700; }}
    .insights {{ display:grid; gap:10px; line-height:1.55; color:#334155; }}
    .insights div {{ border-left:4px solid var(--teal); padding:8px 10px; background:#f7fbfb; border-radius:6px; }}
    table {{ width:100%; border-collapse:collapse; font-size:13px; }}
    th, td {{ border-bottom:1px solid #e2e8f0; padding:9px 10px; vertical-align:top; }}
    th {{ background:#143642; color:white; text-align:left; position:sticky; top:0; z-index:1; }}
    .table-wrap {{ max-height:560px; overflow:auto; border:1px solid var(--line); border-radius:8px; }}
    .table-wrap.small {{ max-height:360px; }}
    .score {{ font-weight:800; color:var(--red); white-space:nowrap; }}
    .tag {{ display:inline-flex; align-items:center; min-height:22px; padding:2px 8px; border-radius:999px; color:white; font-size:12px; margin:0 4px 4px 0; white-space:nowrap; }}
    .Very {{ background:var(--red); }} .Negative {{ background:var(--orange); }} .Mixed {{ background:var(--yellow); color:#1f2937; }} .Neutral {{ background:var(--blue); }}
    .issue {{ background:var(--violet); }}
    .quote {{ font-size:13px; line-height:1.6; max-width:760px; }}
    .muted {{ color:var(--muted); }}
    .controls {{ display:grid; grid-template-columns:1.4fr repeat(3,minmax(150px,.6fr)); gap:10px; margin:12px 0 16px; }}
    input, select {{ min-height:38px; border:1px solid #cbd5e1; border-radius:7px; padding:8px 10px; background:white; font-family:inherit; color:var(--ink); }}
    footer {{ color:var(--muted); font-size:12px; padding-top:8px; }}
    @media (max-width:1100px) {{ .kpis {{ grid-template-columns:repeat(3,1fr); }} .grid,.split {{ grid-template-columns:1fr; }} .controls {{ grid-template-columns:1fr 1fr; }} }}
    @media (max-width:700px) {{ header {{ padding:22px 18px; }} main {{ padding:16px; }} .kpis,.controls {{ grid-template-columns:1fr; }} .bar-row,.compact .bar-row {{ grid-template-columns:1fr; gap:6px; }} th,td {{ min-width:110px; }} }}
  </style>
</head>
<body>
  <header>
    <h1>Negative Feedback จากหัวข้อ ไม่มีซิม True ดู AF ไม่ได้</h1>
    <div class="date-range">Date of : 1 -14 June 2026</div>
    <p>เพจ comment อวยไส้แตกแหกไส้ฉีก และฐานข้อมูลประกอบ โดยแยกช่องทาง Facebook, X, IG, TikTok, YouTube, Pantip และ Other / Website พร้อมสรุปว่าช่องทางไหน negative สูง คำด่าเรื่องอะไร และ account ใดสร้างแรงกระเพื่อมเชิงลบ</p>
  </header>
  <main>
    <nav class="tabs">
      <button class="tab-btn active" data-tab="overview">Overview เดิม + รวมใหม่</button>
      <button class="tab-btn" data-tab="aouy-page">อวยใส้แหก</button>
      {nav_buttons}
      <button class="tab-btn" data-tab="summary">Data Insight Summary</button>
    </nav>

    <section class="tab-panel active" id="tab-overview">
      <div class="kpis" id="kpis"></div>
      <div class="grid">
        <section>
          <h2>Negative by Channel</h2>
          <div class="bars" id="platformBars"></div>
        </section>
        <section>
          <h2>Top Negative Keywords</h2>
          <div class="bars" id="keywordBars"></div>
        </section>
      </div>
      <section>
        <h2>Combined Negative Explorer</h2>
        <div class="controls">
          <input id="search" type="search" placeholder="ค้นหาคำ เช่น ดูไม่ได้, แพ็ก, ทรู, ระบบภาพ, สมัคร">
          <select id="platformFilter"><option value="">ทุกช่องทาง</option>{''.join([f'<option value="{p}">{p}</option>' for p in PLATFORMS])}</select>
          <select id="sentimentFilter"><option value="">ทุก sentiment</option><option>Very Negative</option><option>Negative</option><option>Mixed / Concern</option><option>Neutral / Other</option></select>
          <select id="issueFilter"><option value="">ทุกประเด็น</option><option value="is_access_complaint">ช่องทางรับชม</option><option value="is_true_package_issue">True / แพ็กเกจ</option><option value="is_unable_to_watch">ดูไม่ได้ / ต้องสมัคร</option></select>
        </div>
        <div class="table-wrap"><table><thead><tr><th>Platform</th><th>Score</th><th>Sentiment</th><th>Account</th><th>Eng.</th><th>Issue</th><th>Message</th></tr></thead><tbody id="combinedRows"></tbody></table></div>
      </section>
    </section>

    <section class="tab-panel" id="tab-aouy-page">
      <div class="section-head">
        <h2>เพจ อวยใส้แหก</h2>
        <span class="muted" id="aouyPageMeta"></span>
      </div>
      <div class="kpis" id="aouyPageKpis"></div>
      <div class="split">
        <div>
          <h3>Negative Keywords</h3>
          <div class="bars compact" id="aouyPageKeywords"></div>
        </div>
        <div>
          <h3>Accounts ที่มี negative feedback ในเพจนี้</h3>
          <div class="table-wrap small"><table><thead><tr><th>Account</th><th>Neg.</th><th>Eng.</th><th>Score</th><th>ตัวอย่าง</th></tr></thead><tbody id="aouyPageLeaders"></tbody></table></div>
        </div>
      </div>
      <h3>Top Negative Feedback จาก aouy.xlsx</h3>
      <div class="table-wrap"><table><thead><tr><th>Score</th><th>Sentiment</th><th>Account</th><th>Eng.</th><th>Issue</th><th>Message</th></tr></thead><tbody id="aouyPageRows"></tbody></table></div>
    </section>

    {platform_sections}

    <section class="tab-panel" id="tab-summary">
      <div class="grid">
        <section>
          <h2>Data Insight สรุปรวม</h2>
          <div class="insights" id="insightText"></div>
        </section>
        <section>
          <h2>Account ที่สร้าง Negative Impact สูงสุด</h2>
          <div class="table-wrap small"><table><thead><tr><th>Platform</th><th>Account</th><th>Neg.</th><th>Eng.</th><th>Impact</th><th>ตัวอย่าง</th></tr></thead><tbody id="summaryLeaders"></tbody></table></div>
        </section>
      </div>
      <section>
        <h2>Channel Detail Summary</h2>
        <div class="table-wrap"><table><thead><tr><th>Channel</th><th>Total</th><th>Negative</th><th>% Negative</th><th>Very Neg.</th><th>Access</th><th>True/Package</th><th>Neg. Eng.</th><th>Avg Score</th></tr></thead><tbody id="summaryChannels"></tbody></table></div>
      </section>
    </section>
    <footer>Source: เพจ comment อวยไส้แตกแหกไส้ฉีก + ฐานข้อมูลประกอบ. Negative final = source sentiment Negative หรือ rule-based NLP score เข้าข่าย Negative/Very Negative.</footer>
  </main>
  <script>
    const DATA = {data_json};
    const PLATFORMS = {json.dumps(PLATFORMS, ensure_ascii=False)};
    const colors = {{"Very Negative":"#c2412d","Negative":"#e1812c","Mixed / Concern":"#d8a625","Neutral / Other":"#315d85"}};
    const pct = n => `${{Math.round((n || 0) * 100)}}%`;
    const esc = value => String(value ?? "").replace(/[&<>"']/g, ch => ({{"&":"&amp;","<":"&lt;",">":"&gt;","\\"":"&quot;","'":"&#39;"}}[ch]));
    const cls = s => String(s).split(" ")[0].replace("/", "");
    const sentimentTag = label => `<span class="tag ${{cls(label)}}">${{esc(label)}}</span>`;
    const issueTags = row => {{
      const tags = [];
      if (row.is_access_complaint === "Yes") tags.push("ช่องทางรับชม");
      if (row.is_true_package_issue === "Yes") tags.push("True/แพ็กเกจ");
      if (row.is_unable_to_watch === "Yes") tags.push("ดูไม่ได้");
      return tags.map(t => `<span class="tag issue">${{esc(t)}}</span>`).join("");
    }};
    function barRows(items, maxValue, labelKey, valueKey, color="#14716d") {{
      return items.map(item => {{
        const value = item[valueKey] || 0;
        const share = maxValue ? value / maxValue : 0;
        return `<div class="bar-row"><div class="bar-label" title="${{esc(item[labelKey])}}">${{esc(item[labelKey])}}</div><div class="bar-track"><div class="bar-fill" style="width:${{pct(share)}};background:${{color}}"></div></div><div class="bar-value">${{value}}</div></div>`;
      }}).join("");
    }}
    function renderKpis() {{
      const s = DATA.summary;
      const cards = [
        ["Total mentions", s.total],
        ["Negative mentions", `${{s.negative}} (${{pct(s.negative_share)}})`],
        ["Very Negative", s.very_negative],
        ["Access issue", s.access],
        ["True/package issue", s.true_package],
        ["Avg negative score", s.avg_score],
      ];
      document.getElementById("kpis").innerHTML = cards.map(([label,value]) => `<div class="kpi"><span>${{label}}</span><strong>${{value}}</strong></div>`).join("");
    }}
    function renderOverview() {{
      const platforms = DATA.summary.platforms;
      document.getElementById("platformBars").innerHTML = barRows(platforms, Math.max(...platforms.map(p=>p.negative)), "platform", "negative", "#c2412d");
      document.getElementById("keywordBars").innerHTML = barRows(DATA.summary.keywords, Math.max(...DATA.summary.keywords.map(k=>k.count)), "keyword", "count", "#14716d");
      renderCombinedRows();
    }}
    function renderCombinedRows() {{
      const q = document.getElementById("search").value.trim().toLowerCase();
      const platform = document.getElementById("platformFilter").value;
      const sentiment = document.getElementById("sentimentFilter").value;
      const issue = document.getElementById("issueFilter").value;
      const rows = DATA.combinedRows.filter(row => {{
        const text = `${{row.message}} ${{row.negative_terms}} ${{row.account}}`.toLowerCase();
        return (!q || text.includes(q)) && (!platform || row.platform === platform) && (!sentiment || row.sentiment === sentiment) && (!issue || row[issue] === "Yes");
      }});
      document.getElementById("combinedRows").innerHTML = rows.map(row => `<tr><td>${{esc(row.platform)}}</td><td class="score">${{row.negative_score}}</td><td>${{sentimentTag(row.sentiment)}}</td><td><strong>${{esc(row.account)}}</strong><br><span class="muted">${{esc(row.post_time)}}</span></td><td>${{row.engagement}}</td><td>${{issueTags(row)}}</td><td class="quote">${{esc(row.message)}}${{row.url ? `<br><a href="${{esc(row.url)}}" target="_blank" rel="noreferrer">Link ต้นทาง</a>` : ""}}</td></tr>`).join("");
    }}
    function renderPlatform(platform) {{
      const summary = DATA.summary.platforms.find(p => p.platform === platform) || {{}};
      document.getElementById(`meta-${{platform}}`).textContent = `${{summary.total || 0}} mentions · ${{summary.negative || 0}} negative (${{pct(summary.negative_share)}}) · negative engagement ${{summary.negative_engagement || 0}}`;
      const platformCards = [
        ["Total mentions", summary.total || 0],
        ["Negative mentions", `${{summary.negative || 0}} (${{pct(summary.negative_share)}})`],
        ["Very Negative", summary.very_negative || 0],
        ["Access issue", summary.access || 0],
        ["True/package issue", summary.true_package || 0],
        ["Avg score", summary.avg_score || 0],
      ];
      document.getElementById(`kpis-${{platform}}`).innerHTML = platformCards.map(([label,value]) => `<div class="kpi"><span>${{label}}</span><strong>${{value}}</strong></div>`).join("");
      const payload = DATA.platforms[platform];
      document.getElementById(`keywords-${{platform}}`).innerHTML = barRows(payload.keywords, Math.max(1, ...payload.keywords.map(k=>k.count)), "keyword", "count", "#14716d");
      document.getElementById(`leaders-${{platform}}`).innerHTML = payload.leaders.map(row => `<tr><td><strong>${{esc(row.account)}}</strong><br><span class="muted">${{esc(row.platform)}}</span></td><td>${{row.negative_posts}}</td><td>${{row.total_negative_engagement}}</td><td class="score">${{row.avg_negative_score}}</td><td class="quote">${{esc(row.top_message)}}</td></tr>`).join("");
      document.getElementById(`rows-${{platform}}`).innerHTML = payload.top_negative.map(row => `<tr><td class="score">${{row.negative_score}}</td><td>${{sentimentTag(row.sentiment)}}</td><td><strong>${{esc(row.account)}}</strong><br><span class="muted">${{esc(row.post_time)}}</span></td><td>${{row.engagement}}</td><td>${{issueTags(row)}}</td><td class="quote">${{esc(row.message)}}${{row.url ? `<br><a href="${{esc(row.url)}}" target="_blank" rel="noreferrer">Link ต้นทาง</a>` : ""}}</td></tr>`).join("");
    }}
    function renderAouyPage() {{
      const payload = DATA.aouyPage;
      const s = payload.summary;
      document.getElementById("aouyPageMeta").textContent = `${{s.total}} comments · ${{s.negative}} relevant negative (${{pct(s.negative_share)}})`;
      const cards = [
        ["Total comments", s.total],
        ["Relevant negative", `${{s.negative}} (${{pct(s.negative_share)}})`],
        ["Very Negative", s.very_negative],
        ["Access issue", s.access],
        ["True/package issue", s.true_package],
        ["Avg score", s.avg_score],
      ];
      document.getElementById("aouyPageKpis").innerHTML = cards.map(([label,value]) => `<div class="kpi"><span>${{label}}</span><strong>${{value}}</strong></div>`).join("");
      document.getElementById("aouyPageKeywords").innerHTML = barRows(payload.keywords, Math.max(1, ...payload.keywords.map(k=>k.count)), "keyword", "count", "#14716d");
      document.getElementById("aouyPageLeaders").innerHTML = payload.leaders.map(row => `<tr><td><strong>${{esc(row.account)}}</strong></td><td>${{row.negative_posts}}</td><td>${{row.total_negative_engagement}}</td><td class="score">${{row.avg_negative_score}}</td><td class="quote">${{esc(row.top_message)}}</td></tr>`).join("");
      document.getElementById("aouyPageRows").innerHTML = payload.top_negative.map(row => `<tr><td class="score">${{row.negative_score}}</td><td>${{sentimentTag(row.sentiment)}}</td><td><strong>${{esc(row.account)}}</strong><br><span class="muted">${{esc(row.post_time)}}</span></td><td>${{row.engagement}}</td><td>${{issueTags(row)}}</td><td class="quote">${{esc(row.message)}}${{row.url ? `<br><a href="${{esc(row.url)}}" target="_blank" rel="noreferrer">Link ต้นทาง</a>` : ""}}</td></tr>`).join("");
    }}
    function renderSummary() {{
      const s = DATA.summary;
      const topChannel = [...s.platforms].sort((a,b)=>b.negative-a.negative)[0];
      const topRate = [...s.platforms].filter(p=>p.total>0).sort((a,b)=>b.negative_share-a.negative_share)[0];
      const topLeader = s.leaders[0] || {{}};
      document.getElementById("insightText").innerHTML = [
        `รวมทั้งหมด <strong>${{s.total}}</strong> mentions พบ negative <strong>${{s.negative}}</strong> mentions (${{pct(s.negative_share)}}) และ very negative <strong>${{s.very_negative}}</strong> mentions`,
        `ช่องทางที่มีจำนวน negative มากที่สุดคือ <strong>${{esc(topChannel.platform)}}</strong> จำนวน <strong>${{topChannel.negative}}</strong> mentions`,
        `ช่องทางที่มีอัตรา negative สูงสุดคือ <strong>${{esc(topRate.platform)}}</strong> ที่ <strong>${{pct(topRate.negative_share)}}</strong> ของ mentions ในช่องทางนั้น`,
        `ประเด็นหลักที่โดนต่อว่าคือช่องทางรับชม/ดูไม่ได้ <strong>${{s.access}}</strong> mentions และ True/ซิม/แพ็กเกจ/จ่ายเงิน <strong>${{s.true_package}}</strong> mentions`,
        `Account ที่มี negative impact สูงสุดคือ <strong>${{esc(topLeader.account || "-")}}</strong> บน <strong>${{esc(topLeader.platform || "-")}}</strong> จาก negative posts <strong>${{topLeader.negative_posts || 0}}</strong> และ engagement <strong>${{topLeader.total_negative_engagement || 0}}</strong>`
      ].map(t => `<div>${{t}}</div>`).join("");
      document.getElementById("summaryLeaders").innerHTML = s.leaders.map(row => `<tr><td>${{esc(row.platform)}}</td><td><strong>${{esc(row.account)}}</strong></td><td>${{row.negative_posts}}</td><td>${{row.total_negative_engagement}}</td><td class="score">${{row.impact_score}}</td><td class="quote">${{esc(row.top_message)}}</td></tr>`).join("");
      document.getElementById("summaryChannels").innerHTML = s.platforms.map(row => `<tr><td><strong>${{esc(row.platform)}}</strong></td><td>${{row.total}}</td><td>${{row.negative}}</td><td>${{pct(row.negative_share)}}</td><td>${{row.very_negative}}</td><td>${{row.access}}</td><td>${{row.true_package}}</td><td>${{row.negative_engagement}}</td><td>${{row.avg_score}}</td></tr>`).join("");
    }}
    document.querySelectorAll(".tab-btn").forEach(btn => btn.addEventListener("click", () => {{
      document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById(`tab-${{btn.dataset.tab}}`).classList.add("active");
    }}));
    ["search","platformFilter","sentimentFilter","issueFilter"].forEach(id => document.getElementById(id).addEventListener("input", renderCombinedRows));
    renderKpis(); renderOverview(); renderAouyPage(); PLATFORMS.forEach(renderPlatform); renderSummary();
  </script>
</body>
</html>
"""


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    original = load_original_records()
    aouy_page = load_aouy_page_records()
    new = load_new_records()
    records = pd.concat([original, new], ignore_index=True)
    data = {
        "summary": build_summary(records),
        "platforms": platform_payload(records),
        "combinedRows": rows_for_table(records),
        "aouyPage": page_payload(aouy_page),
    }
    OUTPUT.write_text(build_html(data), encoding="utf-8")
    print(f"Saved {OUTPUT}")
    print(f"Rows: {len(records)} (original={len(original)}, aouy_2={len(new)}, aouy_page={len(aouy_page)})")
    print(f"Negative: {data['summary']['negative']} ({data['summary']['negative_share']:.1%})")
    for row in data["summary"]["platforms"]:
        print(f"{row['platform']}: total={row['total']} negative={row['negative']} share={row['negative_share']:.1%}")


if __name__ == "__main__":
    main()
