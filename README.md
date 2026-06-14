# true-sentiments

True Academy Fantasia 2026 negative feedback analysis for the topic:

`ไม่มีซิม True / ดู AF ไม่ได้ / ช่องทางรับชมเข้าถึงยาก`

## Outputs

- `outputs/true_af_nlp_dashboard.html` - interactive dashboard
- `outputs/negative_feedback_af_summary.pptx` - PowerPoint summary
- `outputs/negative_one_page_summary.png` - 16:9 one-page executive image summary
- `outputs/negative_one_page_summary.html` - HTML source for the one-page image

## Rebuild

```powershell
$env:PYTHONIOENCODING='utf-8'
python build_html_dashboard.py
python build_ppt_summary.py
python build_negative_one_page_image.py
python build_vercel_site.py
```

Raw Excel source workbooks are intentionally ignored by git.
