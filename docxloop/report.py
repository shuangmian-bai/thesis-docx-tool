"""HTML 可视化报告生成。

把 CompareResult 列表渲染成自包含 HTML，人可在浏览器验收：
- 顶部总览：通过率、预期差异数、非预期差异数
- 每个文件一张卡片：差异表格，预期差异绿色标注，非预期差异红色标注
- 每个差异显示：区域、位置、原值、输出值、预期原因
"""
import html
import os
from typing import List

from docxloop.compare import CompareResult


AREA_LABELS = {
    "body": "正文",
    "cover": "封面/目录/承诺书",
    "toc": "目录",
    "promise": "承诺书",
    "header": "页眉",
    "footer": "页脚",
    "styles": "样式",
}


def render(results: List[CompareResult], title: str = "闭环测试报告") -> str:
    total_files = len(results)
    passed_files = sum(1 for r in results if r.passed)
    total_diffs = sum(r.total for r in results)
    unexpected = sum(r.unexpected for r in results)
    expected = total_diffs - unexpected
    pass_rate = (passed_files / total_files * 100) if total_files else 0

    cards = []
    for r in results:
        fname = html.escape(os.path.basename(r.src_path))
        status = "通过" if r.passed else "未通过"
        status_color = "#52c41a" if r.passed else "#ff4d4f"
        rows = []
        for d in r.diffs:
            manual = d.expected and "人工" in d.expected_reason
            if not d.expected:
                bg, tag = "#fff1f0", '<span style="color:#ff4d4f">非预期</span>'
            elif manual:
                bg, tag = "#fff7e6", '<span style="color:#fa8c16">待人工</span>'
            else:
                bg, tag = "#f6ffed", '<span style="color:#52c41a">预期</span>'
            reason = f'<div style="color:#888;font-size:12px">{html.escape(d.expected_reason)}</div>' if d.expected else ""
            rows.append(f"""
            <tr style="background:{bg}">
              <td>{html.escape(AREA_LABELS.get(d.area, d.area))}</td>
              <td>{html.escape(d.location)}</td>
              <td style="font-family:monospace">{html.escape(d.src)}</td>
              <td style="font-family:monospace">{html.escape(d.out)}</td>
              <td>{tag}{reason}</td>
            </tr>""")

        if not rows:
            rows.append('<tr><td colspan="5" style="text-align:center;color:#52c41a">无差异</td></tr>')

        cards.append(f"""
        <div style="border:1px solid #eee;border-radius:8px;margin:16px 0;overflow:hidden">
          <div style="background:#fafafa;padding:12px 16px;border-bottom:1px solid #eee">
            <strong>{fname}</strong>
            <span style="float:right;color:{status_color}">{status}</span>
            <div style="font-size:12px;color:#888;margin-top:6px;line-height:1.6">
              原文：{html.escape(r.src_path)}<br>
              成品：{html.escape(r.out_path)}<br>
              人工审核：直接打开两份 Word 对照，下表为机读差异清单
            </div>
          </div>
          <table style="width:100%;border-collapse:collapse;font-size:13px">
            <thead>
              <tr style="background:#f5f5f5">
                <th style="padding:8px;text-align:left;width:80px">区域</th>
                <th style="padding:8px;text-align:left;width:140px">位置</th>
                <th style="padding:8px;text-align:left">原值</th>
                <th style="padding:8px;text-align:left">输出值</th>
                <th style="padding:8px;text-align:left;width:120px">判定</th>
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
  <div class="stat"><div class="num" style="color:#52c41a">{expected}</div><div class="label">预期差异</div></div>
  <div class="stat"><div class="num" style="color:#ff4d4f">{unexpected}</div><div class="label">非预期差异</div></div>
</div>
{''.join(cards)}
</body>
</html>"""
