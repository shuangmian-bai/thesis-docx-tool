"""版式常量：正文区宽度与插图尺寸上限。

`CONTENT_W` 是 A4 纸减掉左右页边距后的正文区宽度（twips）。表格列宽分配
（`fragments.table_xml`）与目录的右对齐制表位（`template.toc_entry_xml`）都以它为准。
`MAX_IMG_W` / `MAX_IMG_H` 是插图换算成 EMU 后的上限，`fragments.image_xml` 按二者
等比收缩。

顺带一提，`材料/图表/figures.py` 的打印字号也是由这个宽度倒推的：
`纸面字号(pt) = 设计字号 ÷ 内容宽 × 451.3`（451.3pt = 15.92cm = `CONTENT_W`）。
"""

# 正文区可用宽度（twips）：A4 11906 − 左右页边距各 1440
CONTENT_W = 9026
EMU_PER_INCH = 914400
MAX_IMG_W = int(CONTENT_W / 1440 * EMU_PER_INCH)   # ≈ 5 730 000 EMU
MAX_IMG_H = int(8.4 * EMU_PER_INCH)                # 单张图不超过 8.4 英寸高

# 承诺书签名图的显示宽度（EMU）：3.5 cm，高度按图片比例算。
# 上限来自排版空间——「承诺人：」缩进 6096 twips 后右缘在 12.03 cm 处，
# 签名浮动锚点的横偏移是 12.25 cm（模板给的），正文右缘 15.92 cm，
# 因而可用宽度约 3.6 cm。签名是四个字横排（宽高比约 3.45），
# 若按模板示例图那 1.43 cm 的高度等比放大，宽度会到 5.7 cm 而冲出页面。
SIGN_W = 1260000

