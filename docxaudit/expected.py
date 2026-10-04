"""已知的、有意为之的偏离。

与 `README.md` 的「格式约定」收尾段保持同步：那里把预期差异归为两类，一类是内容随论文
变化（图片、core.xml、关系文件、document.xml），一类是有意引入的直接格式（TOC 大小写
标签、updateFields、jc=both、行内 Consolas、表格分页的 keepNext/cantSplit）。README 反对
「两份并存有不同步风险」，改这张表时务必回头看看那两处是否还说得通。

注意**最后一项不在这张表里**：`body.directfmt` 的统计标签不含 keepNext / cantSplit，
报不出 `WARN`，也就无从降级；这两处由 `checks/body.py` 的 `body.table.keep` 专项复核，
正常时直接报 `[通过]`。

键是 `Finding.key`，与各检查函数一一对应；改检查的 key 时务必同步这里，
否则该差异会从「有意偏离」退回「需关注」。

登记在这里只影响**报告层**的分级——`report.classify()` 把命中的 `WARN` 改判为
「有意偏离」，检查函数本身不做区分。
"""

EXPECTED_DIFFS = {
    "styles.toc": "TOC1 的 <w:caps/> 与 TOC2 的 <w:smallCaps/> 被有意移除，"
                  "否则目录里的 Android 会渲染成 ANDROID（见 README）",
    "meta.settings.updatefields": "成品多一个 <w:updateFields w:val=\"true\"/>，"
                                  "让 Word 打开时重建目录域",
    "meta.core.authorship": "core.xml 的 title / creator / lastModifiedBy 与时间戳"
                            "随论文信息变化，revision 等其余字段仍须一致",
    "parts.only_tpl.media": "模板自带的示例插图与缩略图，成品不需要",
    "parts.only_prod.media": "成品新增的论文插图",
    "parts.diff.content": "正文、著录信息、设置与关系文件随论文内容变化；"
                          "承诺书的签名图亦然——模板那张是示例签名，"
                          "成品换成论文作者的签字（image2.png，封面用图 image1 仍逐字节一致）",
    "parts.rels.root": "模板的根关系多一条指向 docProps/thumbnail.emf 的缩略图，"
                       "成品不保留缩略图（README 把关系文件归入「随内容变化」一类）",
    "body.directfmt.jc": "正文段落显式写 jc=both（模板不写）：中文论文惯例两端对齐，"
                         "OOXML 缺省是左对齐，显式写可避免导出 PDF 时右边缘参差",
    "body.directfmt.runfont": "行内 Consolas 用于 API 名、常量与路径，属技术论文正当用法",
}
