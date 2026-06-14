---
name: true-af-negative-feedback-dashboard
description: Use this project skill when continuing or revising the True AF negative feedback analysis dashboard, PowerPoint summary, or Vercel deployment for the "ไม่มีซิม True ดู AF ไม่ได้" workstream.
---

# True AF Negative Feedback Dashboard

## Purpose

Continue the existing analysis for:

- Topic: `Negative Feedback จากหัวข้อ ไม่มีซิม True ดู AF ไม่ได้`
- Date range displayed on dashboard: `Date of : 1 -14 June 2026`
- Main live URL: `https://verceltrueaffeedback.vercel.app`
- Core output HTML: `outputs/true_af_nlp_dashboard.html`
- PowerPoint summary: `outputs/negative_feedback_af_summary.pptx`
- One-page negative image summary: `outputs/negative_one_page_summary.png`

The dashboard should focus only on negative feedback related to:

- ดู AF ไม่ได้
- ไม่มีซิม True / ไม่ได้ใช้ซิมทรู
- ต้องใช้ True / dtac / ซิม / แพ็กเกจ
- สมัครแล้วดูไม่ได้ / กดไม่ได้ / เข้าไม่ได้
- เงื่อนไขรับชมไม่ชัดเจน

Do not let generic promo posts, activity posts, ticket giveaway posts, or unrelated brand posts pollute the negative feedback view.

## Source Files

- `aouy3.xlsx`: current main Facebook comment export used in the combined overview.
- `aouy.xlsx`: separate page tab named `อวยใส้แหก`.
- `aouy_2.xlsx`: social/platform source workbook used internally, but do not show this filename in the dashboard UI.

Lock files such as `~$aouy.xlsx`, `~$aouy3.xlsx`, and `~$aouy_2.xlsx` are temporary Excel files and must be ignored.

## Important UI Wording

Dashboard header:

`Negative Feedback จากหัวข้อ ไม่มีซิม True ดู AF ไม่ได้`

Date line:

`Date of : 1 -14 June 2026`

Intro copy should not mention source filenames such as `aouy_2.xlsx`.

Use this wording:

`เพจ comment อวยไส้แตกแหกไส้ฉีก และฐานข้อมูลประกอบ โดยแยกช่องทาง Facebook, X, IG, TikTok, YouTube, Pantip และ Other / Website พร้อมสรุปว่าช่องทางไหน negative สูง คำด่าเรื่องอะไร และ account ใดสร้างแรงกระเพื่อมเชิงลบ`

Link label must be:

`Link ต้นทาง`

Do not use:

- `open source`
- `aouy_2.xlsx`
- `aouy_2 social listening`
- `social listening` in visible UI text

## Dashboard Tabs

Keep these tabs:

- `Overview เดิม + รวมใหม่`
- `อวยใส้แหก`
- `Facebook`
- `X`
- `IG`
- `TikTok`
- `YouTube`
- `Pantip`
- `Other / Website`
- `Data Insight Summary`

Every platform tab should include KPI cards:

- Total mentions
- Negative mentions
- Very Negative
- Access issue
- True/package issue
- Avg score

## Build Scripts

Primary scripts:

- `build_html_dashboard.py`: builds the dashboard HTML.
- `build_ppt_summary.py`: builds the PowerPoint summary.
- `build_vercel_site.py`: copies HTML/PPTX into a static Vercel site folder and adds the download button.
- `build_negative_one_page_image.py`: builds a 16:9 one-page PNG summary from the same data model.

Normal rebuild sequence:

```powershell
$env:PYTHONIOENCODING='utf-8'
python build_html_dashboard.py
python build_ppt_summary.py
python build_vercel_site.py
```

Build the one-page negative image summary:

```powershell
$env:PYTHONIOENCODING='utf-8'
python build_negative_one_page_image.py
```

Expected outputs:

- `outputs/negative_one_page_summary.html`
- `outputs/negative_one_page_summary.png`

The PNG is a `1600x900` executive one-page summary in the same red/black True-style visual direction as `negative.jpg`.

Deploy to Vercel:

```powershell
npx vercel --prod --yes --cwd outputs/vercel_true_af_feedback
```

After deploy, verify:

```powershell
(Invoke-WebRequest -Uri https://verceltrueaffeedback.vercel.app -UseBasicParsing).StatusCode
```

Expected status: `200`.

## Filtering Rules

The dashboard should show relevant negative feedback only. Exclude rows where the text is just promotion, contest mechanics, ticket giveaways, activity rules, or brand announcement content.

Examples of promo contexts to exclude unless there is an explicit complaint signal:

- `ลุ้นรับ`
- `ร่วมกิจกรรม`
- `กติกา`
- `รางวัล`
- `ประกาศรายชื่อ`
- `ยืนยันรับสิทธิ์`
- `เงื่อนไขกิจกรรม`
- `กด Like`
- `Share โพสต์`

Hard complaint signals include:

- `ดูไม่ได้`
- `ไม่ได้ดู`
- `หาช่องดูไม่ได้`
- `ดูยาก`
- `เข้าไม่ได้`
- `กดไม่ได้`
- `ไม่มีซิมทรู`
- `ไม่ได้ใช้ซิมทรู`
- `ต้องใช้ซิมทรู`
- `เสียความรู้สึก`
- `หลอก`
- `ไม่เคลียร์`
- `ไม่ชัดเจน`
- `ค้าง`
- `กระตุก`
- `ระบบภาพ`
- `แย่`
- `เลวร้าย`
- `เซ็ง`
- `ด่า`
- `บ่น`

## Current Expected Shape

After latest filtering, the combined overview was around:

- Total mentions: `3731`
- Relevant negative: `115`
- Facebook negative: `102`
- X negative: `8`
- IG negative: `3`
- TikTok negative: `2`

These numbers may change if source workbooks are updated.

## One-Page Negative Summary Image

Use this when asked to recreate, revise, or continue the single-image executive summary based on `outputs/true_af_nlp_dashboard.html`.

Reference example:

- `negative.jpg`

Generated files:

- `outputs/negative_one_page_summary.html`
- `outputs/negative_one_page_summary.png`

Source script:

- `build_negative_one_page_image.py`

Current one-page summary narrative:

- Headline: `รายงานสรุปสถานการณ์: Negative Feedback True Academy Fantasia 2026`
- Topic: `ไม่มีซิม True / ดู AF ไม่ได้ / ช่องทางรับชมเข้าถึงยาก`
- Relevant negative: `115` mentions
- Access issue: `109` mentions, displayed as about `95%`
- Platform สูงสุด: Facebook, `102` mentions, displayed as about `89%`
- Peak time: `20:00`, peak date `13/06/2026`
- Main pain points: access friction, True/dtac SIM or package condition, unable to watch, unclear communication
- Top keywords: `ดูไม่ได้`, `ดูยาก`, `ไม่ดู`, `ไม่ได้ดู`, `เสียดาย`, `ผิด`
- Recommended actions: communication clarity, UX/UI optimization for TrueVisions Now, support escalation/war room during Live

Design notes:

- Keep the output as a single 16:9 PNG, `1600x900`.
- Use a red/black True-like executive report style.
- Prioritize readability over squeezing every dashboard table into the image.
- Keep the visible copy focused on negative summary and executive action, not source workbook filenames.
- Use Playwright screenshot export from the generated HTML rather than manual image editing.

## Quality Checks

Before finalizing a revision:

1. Confirm `outputs/true_af_nlp_dashboard.html` exists.
2. Confirm JavaScript parses:

```powershell
@'
const fs = require('fs');
const vm = require('vm');
const html = fs.readFileSync('outputs/true_af_nlp_dashboard.html','utf8');
const scripts = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m=>m[1]);
for (const s of scripts) new vm.Script(s);
console.log('scripts parsed:', scripts.length);
'@ | node
```

3. Confirm unwanted text is not visible:

```powershell
Select-String -Path outputs/true_af_nlp_dashboard.html -Pattern 'aouy_2 social listening|aouy_2.xlsx|open source|TrueOnline ชวนลุ้นรับบัตร'
```

4. Confirm the Vercel production URL returns `200`.

For the one-page PNG, also confirm:

```powershell
@'
from PIL import Image
from pathlib import Path
path = Path('outputs/negative_one_page_summary.png')
print(path.exists(), path.stat().st_size)
print(Image.open(path).size)
'@ | python -
```

Expected dimensions: `(1600, 900)`.
