"""HTML 可视化报告：差异一律红色标注（闭环不区分预期，任何差异都是缺陷）。"""
import html
import os
from typing import List

from docxloop.compare import CompareResult


def render(results: List[CompareResult], title: str = "闭环测试报告") -> str:
    total_files = len(results)
    passed_files = sum(1 for r in results if r.passed)
    total_diffs = sum(r.total for r in results)
    pass_rate = (passed_files / total_files * 100) if total_files else 0

    cards = []
    for r in results:
        fname = html.escape(os.path.basename(r.src_path))
        status = "通过" if r.passed else "未通过"
        status_color = "#52c41a" if r.passed else "#ff4d4f"
        rows = []
        for d in r.diffs:
            rows.append(f"""
            <tr style="background:#fff1f0">
              <td>{html.escape(d.area)}</td>
              <td>{html.escape(d.location)}</td>
              <td style="font-family:monospace">{html.escape(d.src)}</td>
              <td style="font-family:monospace">{html.escape(d.out)}</td>
              <td><span style="color:#ff4d4f">差异</span></td>
            </tr>""")
        if not rows:
            rows.append('<tr><td colspan="5" style="text-align:center;color:#52c41a">完全对应，无差异</td></tr>')

        cards.append(f"""
        <div style="border:1px solid #eee;border-radius:8px;margin:16px 0;overflow:hidden">
          <div style="background:#fafafa;padding:12px 16px;border-bottom:1px solid #eee">
            <strong>{fname}</strong>
            <span style="float:right;color:{status_color}">{status}</span>
            <div style="font-size:12px;color:#888;margin-top:6px;line-height:1.6">
              原文：{html.escape(r.src_path)}<br>
              成品：{html.escape(r.out_path)}
            </div>
          </div>
          <table style="width:100%;border-collapse:collapse;font-size:13px">
            <thead>
              <tr style="background:#f5f5f5">
                <th style="padding:8px;text-align:left;width:60px">区域</th>
                <th style="padding:8px;text-align:left;width:160px">位置</th>
                <th style="padding:8px;text-align:left">原值</th>
                <th style="padding:8px;text-align:left">输出值</th>
                <th style="padding:8px;text-align:left;width:80px">判定</th>
              </tr>
            </thead>
            <tbody>{''.join(rows)}</tbody>
          </table>
        </div>""")

    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>{html.escape(title)}</title>
<style>
  body {{ font-family: -apple-system, "Segoe UI", sans-serif; margin: 24px; color: #333; }}
  h1 {{ font-size: 20px; }}
  .summary {{ display: flex; gap: 24px; margin: 16px 0; flex-wrap: wrap; }}
  .stat {{ background: #f5f5f5; border-radius: 8px; padding: 12px 20px; min-width: 120px; }}
  .stat .num {{ font-size: 24px; font-weight: bold; }}
  .stat .label {{ font-size: 12px; color: #666; }}
</style>
</head>
<body>
<h1>{html.escape(title)}</h1>
<div class="summary">
  <div class="stat"><div class="num" style="color:#1890ff">{total_files}</div><div class="label">测试文件数</div></div>
  <div class="stat"><div class="num" style="color:#52c41a">{passed_files}</div><div class="label">通过文件数</div></div>
  <div class="stat"><div class="num" style="color:#faad14">{pass_rate:.0f}%</div><div class="label">通过率</div></div>
  <div class="stat"><div class="num" style="color:#ff4d4f">{total_diffs}</div><div class="label">差异总数</div></div>
</div>
{''.join(cards)}
</body>
</html>"""
