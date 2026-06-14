from __future__ import annotations

import shutil
from pathlib import Path


SRC_HTML = Path("outputs/true_af_nlp_dashboard.html")
SRC_PPT = Path("outputs/negative_feedback_af_summary.pptx")
SITE = Path("outputs/vercel_true_af_feedback")


def main() -> None:
    SITE.mkdir(parents=True, exist_ok=True)
    html = SRC_HTML.read_text(encoding="utf-8")
    download_css = """
    .download-deck {
      position: fixed;
      right: 22px;
      bottom: 22px;
      z-index: 30;
      background: #c2412d;
      color: #fff;
      text-decoration: none;
      padding: 11px 14px;
      border-radius: 8px;
      font-weight: 700;
      box-shadow: 0 12px 26px rgba(15, 23, 42, .22);
    }
    .download-deck:hover { background: #a93625; }
"""
    html = html.replace("</style>", download_css + "\n  </style>")
    html = html.replace(
        "</body>",
        '  <a class="download-deck" href="/negative_feedback_af_summary.pptx" download>Download PowerPoint</a>\n</body>',
    )
    (SITE / "index.html").write_text(html, encoding="utf-8")
    shutil.copy2(SRC_PPT, SITE / "negative_feedback_af_summary.pptx")
    (SITE / "vercel.json").write_text(
        '{\n  "cleanUrls": true,\n  "trailingSlash": false\n}\n',
        encoding="utf-8",
    )
    print(SITE)


if __name__ == "__main__":
    main()
