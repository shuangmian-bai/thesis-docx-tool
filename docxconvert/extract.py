"""从 docx 抽取图片到目标目录。"""
import os
import zipfile


def extract_images(docx_path, images, out_dir):
    """把 images 里列出的图片从 docx 抽到 out_dir。

    images: {zip 内路径: 建议文件名}
    返回: {zip 内路径: 最终写入的文件名}（重名时自动加后缀）
    """
    os.makedirs(out_dir, exist_ok=True)
    zin = zipfile.ZipFile(docx_path)
    written, used = {}, set()
    for zip_path, name in images.items():
        base, ext = os.path.splitext(name)
        fname = name
        i = 1
        while fname in used:
            fname = f"{base}_{i}{ext}"
            i += 1
        used.add(fname)
        with open(os.path.join(out_dir, fname), "wb") as fh:
            fh.write(zin.read(zip_path))
        written[zip_path] = fname
    zin.close()
    return written
