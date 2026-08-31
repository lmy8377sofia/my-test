#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 电子盖章工具
================

本地桌面软件：上传印章图片 → 打开 PDF → 手动 / 自动盖章 → 导出新的已盖章 PDF。
仅本地运行，不上传网络。支持 Windows 10/11（也兼容 Linux/macOS）。

依赖：PyMuPDF（import pymupdf as fitz）、Pillow；NumPy 可选（用于加速像素处理）。

命令行：
    python pdf_stamp.py                启动图形界面
    python pdf_stamp.py --selftest     无界面自检（预处理 / 自动盖章 / 导出）
    python pdf_stamp.py --demo         在 assets/demo 生成示例印章与示例 PDF
"""

from __future__ import annotations

import argparse
import io
import math
import os
import sys
import tempfile
import threading
import traceback
from dataclasses import asdict, dataclass, field

# ---------------------------------------------------------------------------
# 0. Windows 高 DPI 感知（必须在创建 Tk 窗口之前设置）
# ---------------------------------------------------------------------------
if sys.platform == "win32":
    try:
        import ctypes

        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)   # 系统 DPI 感知
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# ---------------------------------------------------------------------------
# 1. 第三方依赖导入
# ---------------------------------------------------------------------------
try:
    import pymupdf as fitz            # PyMuPDF >= 1.24（新版包名 pymupdf）
except ImportError:                   # 兼容旧版包名 fitz
    import fitz  # type: ignore

from PIL import Image, ImageDraw

try:
    import numpy as _np
    HAS_NUMPY = True
except Exception:                     # NumPy 缺失时使用纯 Python 逐像素兜底
    _np = None
    HAS_NUMPY = False

# ---------------------------------------------------------------------------
# 2. 全局常量与默认参数（规格 5.4）
# ---------------------------------------------------------------------------
APP_NAME = "PDF 电子盖章工具"
APP_VERSION = "1.0.0"

DEFAULT_SIZE_MM = 55.0        # 印章直径（mm），范围 10~80
DEFAULT_OPACITY = 0.90        # 整体透明度，范围 0.2~1.0
DEFAULT_ANGLE = 0.0           # 旋转角（度，逆时针为正）
DEFAULT_REMOVE_WHITE = True   # 去白底
DEFAULT_WHITE_THRESHOLD = 240 # 白色判定阈值
DEFAULT_PREVIEW_DPI = 120     # 屏幕预览渲染 DPI
EXPORT_DPI = 320              # 嵌入 PDF 的位图分辨率（约 300~320 DPI）
MAX_STAMP_SIDE = 1200         # 预处理时印章最长边上限（px）
PT_PER_MM = 72.0 / 25.4       # 1mm 对应的 PDF 点数

AUTO_SIZE_MM = 55.0           # 自动盖章默认直径
AUTO_OPACITY = 0.90           # 自动盖章默认透明度
KEYWORDS_COMPANY = ("华信检测", "有限公司")   # 规则 A 关键词
KEYWORD_MEMBERS = "项目成员"                  # 规则 B 关键词

# PyInstaller 打包后资源路径兼容（sys._MEIPASS）
def resource_path(rel: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


# ---------------------------------------------------------------------------
# 3. 数据结构
# ---------------------------------------------------------------------------
@dataclass
class StampPlacement:
    """一枚印章的摆放参数（规格 3.5）。"""
    page_index: int
    cx: float                       # 章心 x（pt，fitz 左上原点，Y 向下）
    cy: float                       # 章心 y（pt）
    size_mm: float = DEFAULT_SIZE_MM
    opacity: float = DEFAULT_OPACITY
    angle: float = DEFAULT_ANGLE
    stamp_path: str = ""
    remove_white: bool = DEFAULT_REMOVE_WHITE
    white_threshold: int = DEFAULT_WHITE_THRESHOLD

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "StampPlacement":
        d = dict(d)
        for k in list(d.keys()):
            if k not in cls.__dataclass_fields__:
                d.pop(k)
        return cls(**d)


# ---------------------------------------------------------------------------
# 4. 印章图像处理核心（规格 3.1 / 3.2 / 3.3）
# ---------------------------------------------------------------------------
class StampProcessor:
    """
    印章预处理管线：
      原图 → RGBA → 最长边≤1200 → 去白底 → 乘透明度 →（旋转）→ 按 mm/DPI 缩放 → PNG bytes
    两级缓存避免重复处理：
      base 缓存 key = (路径, mtime, remove_white, threshold, opacity)
      render 缓存 key = (base_key, angle, fit_px)
    """

    def __init__(self, max_side: int = MAX_STAMP_SIDE, export_dpi: int = EXPORT_DPI):
        self.max_side = max_side
        self.export_dpi = export_dpi
        self._base_cache: dict = {}    # key -> RGBA PIL Image
        self._render_cache: dict = {}  # key -> RGBA PIL Image
        self._png_cache: dict = {}     # (base_key, angle, size_mm) -> (bytes, w, h, fit_px)
        self.last_error: str | None = None

    # ------------------------------------------------------------------ 4.1
    @staticmethod
    def _base_key(stamp_path: str, remove_white: bool, white_threshold: int,
                  opacity: float) -> tuple:
        return (os.path.normcase(os.path.abspath(stamp_path)),
                round(os.path.getmtime(stamp_path), 3),
                bool(remove_white), int(white_threshold), round(float(opacity), 3))

    def prepare(self, stamp_path: str, remove_white: bool = True,
                white_threshold: int = DEFAULT_WHITE_THRESHOLD,
                opacity: float = DEFAULT_OPACITY):
        """
        加载并预处理印章。
        返回 (key, RGBA Image, 错误信息)；成功时错误信息为 None。
        """
        self.last_error = None
        try:
            if not stamp_path or not os.path.isfile(stamp_path):
                self.last_error = "印章文件不存在：%s" % stamp_path
                return None, None, self.last_error
            key = self._base_key(stamp_path, remove_white, white_threshold, opacity)
            if key in self._base_cache:
                return key, self._base_cache[key], None

            img = Image.open(stamp_path)
            img = img.convert("RGBA")

            # 3.1：最长边 > 1200px 则等比缩小，避免导出 PDF 过大
            w, h = img.size
            longest = max(w, h)
            if longest > self.max_side:
                ratio = self.max_side / float(longest)
                img = img.resize((max(1, int(w * ratio)), max(1, int(h * ratio))),
                                 Image.LANCZOS)

            # 3.1：去白底 + 边缘羽化 + 整体透明度
            if remove_white:
                img = self._remove_white(img, int(white_threshold), float(opacity))
            else:
                # 不去白底时仍然应用整体透明度
                if opacity < 0.999:
                    a = img.getchannel("A").point(lambda v: int(v * opacity))
                    img = img.copy()
                    img.putalpha(a)

            self._base_cache[key] = img
            return key, img, None
        except Exception as e:
            self.last_error = "印章处理失败：%s" % e
            return None, None, self.last_error

    # ------------------------------------------------------------------ 4.2
    def _remove_white(self, img: Image.Image, threshold: int, opacity: float) -> Image.Image:
        """去白底（规格 3.1）：R/G/B 均 ≥ threshold → Alpha=0；
        亮度 > 200 的边缘像素按比例羽化；其余像素保留原色。
        最终 Alpha *= opacity。NumPy 加速，无 NumPy 时逐像素兜底。"""
        if HAS_NUMPY:
            return self._remove_white_numpy(img, threshold, opacity)
        return self._remove_white_py(img, threshold, opacity)

    def _remove_white_numpy(self, img: Image.Image, threshold: int,
                            opacity: float) -> Image.Image:
        arr = _np.array(img, dtype=_np.int16)          # H,W,4
        r, g, b, a = arr[..., 0], arr[..., 1], arr[..., 2], arr[..., 3]
        lum = (r + g + b) / 3.0

        white = (r >= threshold) & (g >= threshold) & (b >= threshold)
        a[white] = 0                                    # 纯白 → 全透明

        fade = (lum > 200.0) & (~white)                 # 边缘羽化
        if fade.any():
            ratio = _np.clip((255.0 - lum[fade]) / 55.0, 0.0, 1.0)
            a[fade] = (a[fade].astype(_np.float32) * ratio).astype(_np.int16)

        a = _np.clip(a * opacity, 0, 255).astype(_np.uint8)
        out = _np.dstack([r.astype(_np.uint8), g.astype(_np.uint8),
                          b.astype(_np.uint8), a])
        return Image.fromarray(out, "RGBA")

    def _remove_white_py(self, img: Image.Image, threshold: int,
                         opacity: float) -> Image.Image:
        """纯 Python 逐像素兜底（无 NumPy 时）。"""
        src = img.load()
        w, h = img.size
        out = Image.new("RGBA", (w, h))
        dst = out.load()
        for y in range(h):
            for x in range(w):
                r, g, b, a = src[x, y][:4]
                lum = (r + g + b) / 3.0
                if r >= threshold and g >= threshold and b >= threshold:
                    a = 0
                elif lum > 200.0:
                    a = int(a * (255.0 - lum) / 55.0)
                a = int(a * opacity)
                dst[x, y] = (r, g, b, a)
        return out

    # ------------------------------------------------------------------ 4.3
    def render(self, base_key, angle: float, fit_px: int) -> Image.Image | None:
        """把预处理好的底图旋转并缩放到指定像素边长（较长边=fit_px），返回 RGBA。"""
        if base_key not in self._base_cache:
            return None
        rkey = (base_key, round(float(angle) % 360.0, 2), int(fit_px))
        if rkey in self._render_cache:
            return self._render_cache[rkey]

        img = self._base_cache[base_key]
        if angle % 360.0:
            img = img.rotate(float(angle), expand=True, resample=Image.BICUBIC)
        w, h = img.size
        if w >= h:
            target = (int(fit_px), max(1, int(round(h * fit_px / w))))
        else:
            target = (max(1, int(round(w * fit_px / h))), int(fit_px))
        if target != (w, h):
            img = img.resize(target, Image.LANCZOS)
        self._render_cache[rkey] = img

        # 控制缓存体积：超过 200 个条目时清理一半
        if len(self._render_cache) > 200:
            keys = list(self._render_cache.keys())[::2]
            for k in keys:
                self._render_cache.pop(k, None)
        return img

    def export_png(self, base_key, angle: float, size_mm: float,
                   dpi: int | None = None) -> tuple | None:
        """
        生成嵌入 PDF 的 PNG 字节流（规格 3.2 / 3.5）。
        返回 (png_bytes, fit_px, rot_w_px, rot_h_px)；失败返回 None。
        dpi 默认 self.export_dpi（320），px = size_mm / 25.4 * dpi。
        """
        dpi = dpi or self.export_dpi
        if base_key not in self._base_cache:
            return None
        ckey = (base_key, round(float(angle) % 360.0, 2), round(float(size_mm), 3), dpi)
        if ckey in self._png_cache:
            return self._png_cache[ckey]

        fit_px = max(1, int(round(size_mm / 25.4 * dpi)))
        img = self.render(base_key, angle, fit_px)
        if img is None:
            return None
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        result = (buf.getvalue(), fit_px, img.size[0], img.size[1])
        self._png_cache[ckey] = result
        return result

    def clear_caches(self):
        self._base_cache.clear()
        self._render_cache.clear()
        self._png_cache.clear()


# ---------------------------------------------------------------------------
# 5. 自动盖章规则（规格 四）
# ---------------------------------------------------------------------------
def _page_lines(page) -> list:
    """提取页面文本行：[(text, bbox(x0,y0,x1,y1), max_font_size)]"""
    lines = []
    try:
        d = page.get_text("dict")
    except Exception:
        return lines
    for block in d.get("blocks", []):
        for line in block.get("lines", []):
            text = "".join(s.get("text", "") for s in line.get("spans", [])).strip()
            if not text:
                continue
            bbox = tuple(line.get("bbox", (0, 0, 0, 0)))
            sizes = [s.get("size", 0.0) for s in line.get("spans", [])]
            lines.append((text, bbox, max(sizes) if sizes else 0.0))
    return lines


def find_company_line(page):
    """规则 A：第 1 页（封面）找含「华信检测」或「有限公司」的公司名行。
    打分：含「华信检测」+100，含「有限公司」+50，行越长越完整再加分，字号大加分。"""
    best, best_score = None, -1.0
    for text, bbox, fsize in _page_lines(page):
        if KEYWORDS_COMPANY[0] not in text and KEYWORDS_COMPANY[1] not in text:
            continue
        score = 0.0
        if KEYWORDS_COMPANY[0] in text:
            score += 100.0
        if KEYWORDS_COMPANY[1] in text:
            score += 50.0
        score += min(len(text), 200) * 0.5
        score += min(fsize, 60.0)
        if score > best_score:
            best, best_score = (text, bbox), score
    return best


def find_member_title_line(page):
    """规则 B：第 2 页找「项目成员」标题行。
    优先取字号最大者（标题字号通常最大），并列时取最靠上者（不是表格底部）。"""
    cands = [(t, b, fs) for t, b, fs in _page_lines(page) if KEYWORD_MEMBERS in t]
    if not cands:
        return None
    cands.sort(key=lambda c: (-c[2], c[1][1]))   # 字号降序，y0 升序
    return (cands[0][0], cands[0][1])


def compute_auto_placements(doc, stamp_path: str, size_mm: float = AUTO_SIZE_MM,
                            opacity: float = AUTO_OPACITY, angle: float = 0.0,
                            remove_white: bool = True,
                            white_threshold: int = DEFAULT_WHITE_THRESHOLD):
    """
    按默认规则计算自动盖章位置（定位跟文字走，不写死绝对坐标）。
    返回 (placements, messages)。
    """
    placements, messages = [], []
    if doc is None or doc.page_count < 1:
        return placements, ["文档为空，无法自动盖章"]

    # ---- 规则 A：第 1 页（封面），公司名行中心 ----
    page0 = doc[0]
    hit = find_company_line(page0)
    if hit:
        text, (x0, y0, x1, y1) = hit
        cx, cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
        placements.append(StampPlacement(
            page_index=0, cx=cx, cy=cy, size_mm=size_mm, opacity=opacity,
            angle=angle, stamp_path=stamp_path, remove_white=remove_white,
            white_threshold=white_threshold))
        messages.append("第 1 页：已找到公司名「%s」，印章盖在文字中心。" % text)
    else:
        messages.append("第 1 页：未找到「%s」或「%s」，请手动盖章。"
                        % KEYWORDS_COMPANY)

    # ---- 规则 B：第 2 页（项目成员），标题行中心 ----
    if doc.page_count > 1:
        page1 = doc[1]
        hit = find_member_title_line(page1)
        if hit:
            text, (x0, y0, x1, y1) = hit
            cx, cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
            placements.append(StampPlacement(
                page_index=1, cx=cx, cy=cy, size_mm=size_mm, opacity=opacity,
                angle=angle, stamp_path=stamp_path, remove_white=remove_white,
                white_threshold=white_threshold))
            messages.append("第 2 页：已找到标题「%s」，印章盖在标题中心。" % text)
        else:
            messages.append("第 2 页：未找到标题「%s」，请手动盖章。" % KEYWORD_MEMBERS)
    else:
        messages.append("第 2 页：文档只有 1 页，跳过规则 B。")

    return placements, messages


# ---------------------------------------------------------------------------
# 6. 导出（规格 3.5）
# ---------------------------------------------------------------------------
def export_pdf(src_path: str, dst_path: str, placements, proc: StampProcessor,
               progress=None) -> list:
    """
    把 placements 写入原 PDF 的副本并保存到 dst_path（绝不覆盖原文件）。
    返回被跳过的 placement 列表（页号超出文档范围等）。
    """
    skipped = []
    doc = fitz.open(src_path)
    try:
        groups: dict = {}
        for pl in placements:
            groups.setdefault(pl.page_index, []).append(pl)

        total = len(placements)
        done = 0
        for page_index in sorted(groups):
            if page_index < 0 or page_index >= doc.page_count:
                skipped.extend(groups[page_index])
                continue
            page = doc[page_index]
            pw, ph = page.rect.width, page.rect.height
            for pl in groups[page_index]:
                key, _, err = proc.prepare(pl.stamp_path, pl.remove_white,
                                           pl.white_threshold, pl.opacity)
                if key is None:
                    skipped.append(pl)
                    done += 1
                    if progress:
                        progress(done, total, "跳过（%s）" % (err or "印章无效"))
                    continue
                png = proc.export_png(key, pl.angle, pl.size_mm)
                if png is None:
                    skipped.append(pl)
                    done += 1
                    continue
                png_bytes, fit_px, rw, rh = png

                # 尺寸换算：size_pt = size_mm / 25.4 * 72（规格 3.2）
                size_pt = pl.size_mm * PT_PER_MM
                # 旋转后外接矩形：按渲染图实际像素比例换算到 pt（规格 3.3）
                width_pt = size_pt * rw / float(fit_px)
                height_pt = size_pt * rh / float(fit_px)

                cx = min(max(pl.cx, width_pt / 2), pw - width_pt / 2)
                cy = min(max(pl.cy, height_pt / 2), ph - height_pt / 2)
                rect = fitz.Rect(cx - width_pt / 2, cy - height_pt / 2,
                                 cx + width_pt / 2, cy + height_pt / 2)
                # overlay=True：盖在文字上方；keep_proportion=True：保持宽高比
                page.insert_image(rect, stream=png_bytes,
                                  keep_proportion=True, overlay=True)
                done += 1
                if progress:
                    progress(done, total, "第 %d 页 完成" % (page_index + 1))

        doc.save(dst_path, deflate=True, garbage=3)
    finally:
        doc.close()
    return skipped


# ---------------------------------------------------------------------------
# 7. 示例数据生成（--demo，供验收测试使用）
# ---------------------------------------------------------------------------
def _find_cjk_font():
    candidates = [
        # Windows
        r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\msyhbd.ttc",
        r"C:\Windows\Fonts\simsun.ttc", r"C:\Windows\Fonts\simhei.ttf",
        # Linux / macOS
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/arphic/uming.ttc",
        "/System/Library/Fonts/PingFang.ttc",
    ]
    for p in candidates:
        if os.path.isfile(p):
            return p
    return None


def generate_demo_seal(path: str, white_bg: bool = False) -> str:
    """生成一枚白底红章示例图（圆形公章样式）。"""
    import math as _m

    S = 900
    img = Image.new("RGBA", (S, S), (255, 255, 255, 255) if white_bg else (255, 255, 255, 0))
    d = ImageDraw.Draw(img)
    red = (220, 30, 30, 255)
    cx = cy = S / 2

    def ring(r, width):
        d.ellipse([cx - r, cy - r, cx + r, cy + r], outline=red, width=width)

    ring(S * 0.46, 14)
    ring(S * 0.34, 4)

    font_path = _find_cjk_font()
    font_top = font_bot = None
    try:
        from PIL import ImageFont
        if font_path:
            font_top = ImageFont.truetype(font_path, 118)
            font_bot = ImageFont.truetype(font_path, 78)
    except Exception:
        pass

    def arc_text(text, radius, size, start_deg, end_deg, font, up=True):
        """沿圆弧排布文字（公章常见样式）。"""
        if font is None:
            return
        n = len(text)
        span = (end_deg - start_deg) / max(1, n - 1) if n > 1 else 0
        for i, ch in enumerate(text):
            ang = _m.radians(start_deg + i * span)
            x = cx + radius * _m.cos(ang)
            y = cy + radius * _m.sin(ang)
            ch_img = Image.new("RGBA", (160, 160), (0, 0, 0, 0))
            cd = ImageDraw.Draw(ch_img)
            bbox = cd.textbbox((0, 0), ch, font=font)
            cd.text((80 - (bbox[2] - bbox[0]) / 2 - bbox[0],
                     80 - (bbox[3] - bbox[1]) / 2 - bbox[1]), ch,
                    font=font, fill=red)
            rot = _m.degrees(ang) - 90
            if not up:
                rot += 180
            ch_img = ch_img.rotate(rot, expand=True, resample=Image.BICUBIC)
            img.paste(ch_img, (int(x - ch_img.width / 2), int(y - ch_img.height / 2)), ch_img)

    arc_text("华信检测有限公司", S * 0.40, 118, 208, 332, font_top, up=True)
    arc_text("合同专用章", S * 0.40, 78, 28, 152, font_bot, up=False)

    # 中心五角星
    import math as _m2
    star = []
    for i in range(10):
        r = S * 0.14 if i % 2 == 0 else S * 0.062
        ang = -_m2.pi / 2 + i * _m2.pi / 5
        star.append((cx + r * _m2.cos(ang), cy + r * _m2.sin(ang)))
    d.polygon(star, fill=red)

    if white_bg:
        img = img.convert("RGB")
    img.save(path)
    return path


def generate_demo_pdf(path: str) -> str:
    """生成 2 页示例 PDF：第 1 页含公司名，第 2 页含「项目成员」标题。"""
    doc = fitz.open()
    for _ in range(2):
        doc.new_page(width=595.3, height=841.9)   # A4
    font = fitz.Font("china-s")                    # PyMuPDF 内置中文字体

    def text(page, x, y, s, size=14, color=(0, 0, 0)):
        color = tuple(c / 255.0 for c in color)
        page.insert_text((x, y), s, fontname="china-s", fontsize=size, color=color)

    # ---- 第 1 页：报价单封面 ----
    p0 = doc[0]
    text(p0, 180, 120, "报 价 单", 30, (30, 30, 120))
    text(p0, 170, 220, "致：××科技有限公司", 14)
    text(p0, 170, 260, "项目名称：检测技术服务项目", 14)
    text(p0, 170, 300, "报价日期：2026 年 8 月 31 日", 14)
    text(p0, 170, 420, "华信检测有限公司", 18, (0, 0, 0))
    text(p0, 170, 450, "地址：北京市××区××路 ×× 号", 12)
    text(p0, 170, 480, "电话：010-88888888", 12)

    # ---- 第 2 页：项目成员 ----
    p1 = doc[1]
    text(p1, 60, 90, "项目成员", 20, (30, 30, 120))
    text(p1, 60, 130, "一、项目组成员名单", 12)
    rows = [("姓名", "职务", "单位"),
            ("张三", "项目负责人", "华信检测有限公司"),
            ("李四", "技术工程师", "华信检测有限公司"),
            ("王五", "质量工程师", "华信检测有限公司"),
            ("赵六", "报告审核人", "华信检测有限公司")]
    x0, y0 = 80, 160
    col_w = (150, 120, 180)
    for i, (c1, c2, c3) in enumerate(rows):
        y = y0 + i * 30
        if i == 0:
            for k, (xx, ww) in enumerate([(x0, col_w[0]), (x0 + col_w[0], col_w[1]), (x0 + col_w[0] + col_w[1], col_w[2])]):
                p1.draw_rect(fitz.Rect(xx, y, xx + ww, y + 30), color=(0, 0, 0), width=0.8)
        for xx, ww, val in [(x0, col_w[0], c1), (x0 + col_w[0], col_w[1], c2),
                            (x0 + col_w[0] + col_w[1], col_w[2], c3)]:
            p1.draw_rect(fitz.Rect(xx, y, xx + ww, y + 30), color=(0, 0, 0), width=0.8)
            text(p1, xx + 8, y + 20, val, 11)
    text(p1, 80, y0 + len(rows) * 30 + 60, "（盖章生效）", 11, (120, 120, 120))
    doc.save(path, deflate=True, garbage=3)
    doc.close()
    return path


# ---------------------------------------------------------------------------
# 8. 自检模式（--selftest，无界面验收测试，规格 七）
# ---------------------------------------------------------------------------
def run_selftest() -> int:
    ok = True

    def check(cond: bool, name: str):
        nonlocal ok
        print(("  [通过] " if cond else "  [失败] ") + name)
        if not cond:
            ok = False

    print("== PDF 电子盖章工具 自检 ==")

    tmp = tempfile.mkdtemp(prefix="stamp_test_")
    seal_png = os.path.join(tmp, "seal.png")
    seal_jpg = os.path.join(tmp, "seal.jpg")
    doc_path = os.path.join(tmp, "demo.pdf")
    out_path = os.path.join(tmp, "demo_已盖章.pdf")

    print("[1/6] 生成测试素材...")
    generate_demo_seal(seal_png, white_bg=False)
    generate_demo_seal(seal_jpg, white_bg=True)   # 白底 JPG，测试去白底
    generate_demo_pdf(doc_path)
    check(os.path.isfile(seal_png) and os.path.isfile(seal_jpg)
          and os.path.isfile(doc_path), "测试素材生成成功")

    print("[2/6] 印章预处理（去白底 + 透明度）...")
    proc = StampProcessor()
    key, base, err = proc.prepare(seal_jpg, True, 240, 0.90)
    check(key is not None and base is not None, "白底 JPG 预处理成功")
    if base is not None:
        px = base.convert("RGBA").load()
        w, h = base.size
        corners = [px[x, y] for x, y in [(2, 2), (w - 3, 2), (2, h - 3), (w - 3, h - 3)]]
        check(all(c[3] < 16 for c in corners), "四角白色像素 Alpha≈0（白底已去除）")
        reds = [px[x, y] for x, y in [(w // 2, h // 2), (w // 2, int(h * 0.15))]]
        check(any(c[2] < 100 and c[0] > 150 and c[3] > 200 for c in reds),
              "红色像素保留且 Alpha≈0.9×255")
        check(max(w, h) <= MAX_STAMP_SIDE, "最长边 ≤ 1200px")

    print("[3/6] 无 NumPy 兜底路径...")
    global HAS_NUMPY
    saved_np = HAS_NUMPY
    try:
        import numpy as np  # noqa
        HAS_NUMPY = False
        key2, base2, err2 = proc.prepare(seal_jpg, True, 240, 0.90)
        if base is not None and base2 is not None:
            a1 = _np.array(base)[..., 3]
            a2 = _np.array(base2)[..., 3]
            diff = int(_np.abs(a1.astype(int) - a2.astype(int)).max())
            check(diff <= 2, "纯 Python 兜底与 NumPy 结果一致（最大差 %d）" % diff)
        else:
            check(False, "纯 Python 兜底路径可运行")
    finally:
        HAS_NUMPY = saved_np

    print("[4/6] 自动盖章规则（定位跟文字走）...")
    doc = fitz.open(doc_path)
    try:
        placements, msgs = compute_auto_placements(
            doc, seal_png, size_mm=55.0, opacity=0.90)
        for m in msgs:
            print("      " + m)
        check(len(placements) == 2, "规则 A+B 共生成 2 枚章")
        if len(placements) >= 2:
            p0, p1 = placements[0], placements[1]
            r0 = doc[0].rect
            check(p0.page_index == 0 and r0.width * 0.3 < p0.cx < r0.width * 0.7
                  and 250 < p0.cy < 500, "第 1 页章心落在公司名区域")
            d1 = doc[1].get_text("dict")
            title_bbox = None
            for b in d1["blocks"]:
                for ln in b.get("lines", []):
                    if "项目成员" in "".join(s["text"] for s in ln["spans"]):
                        title_bbox = ln["bbox"]
                        break
            if title_bbox:
                check(abs(p1.cx - (title_bbox[0] + title_bbox[2]) / 2) < 3
                      and abs(p1.cy - (title_bbox[1] + title_bbox[3]) / 2) < 3,
                      "第 2 页章心 = 「项目成员」标题中心（非表格底部）")
            else:
                check(False, "示例 PDF 中存在「项目成员」标题")
            check(p1.cy < 200, "第 2 页章在页面上部（标题处），不在表格底部")
    finally:
        doc.close()

    print("[5/6] 导出（尺寸换算 / 体积 / 不覆盖原文件）...")
    src_mtime = os.path.getmtime(doc_path)
    doc = fitz.open(doc_path)
    placements, _ = compute_auto_placements(doc, seal_png, 55.0, 0.90)
    doc.close()
    skipped = export_pdf(doc_path, out_path, placements, proc)
    check(skipped == [], "无跳过 placement")
    check(os.path.isfile(out_path), "导出文件已生成")
    size_kb = os.path.getsize(out_path) / 1024.0
    check(size_kb < 1024, "导出文件体积合理（%.0f KB < 1024 KB）" % size_kb)
    check(abs(os.path.getmtime(doc_path) - src_mtime) < 0.01, "原 PDF 未被修改")

    print("[6/6] 校验导出结果...")
    out = fitz.open(out_path)
    try:
        check(out.page_count == 2, "导出文档 2 页")
        imgs0 = out[0].get_images(full=True)
        imgs1 = out[1].get_images(full=True)
        xrefs = set(i[0] for i in imgs0 + imgs1)
        check(len(xrefs) == 1, "同一枚章在文档中只存一份图像数据（xref 去重）")
        r0 = out[0].get_image_rects(next(iter(xrefs)))
        r1 = out[1].get_image_rects(next(iter(xrefs)))
        check(len(r0) == 1 and len(r1) == 1, "每页实际显示 1 枚印章")
        # 检查图像 xref 尺寸与透明度
        xref = next(iter(xrefs))
        info = out.extract_image(xref)
        check(info["ext"] == "png" and info["width"] <= 700,
              "印章按 mm/DPI 缩放后嵌入（%dx%d px，非原图 900px）"
              % (info["width"], info["height"]))
        alpha_ok = info.get("smask", 0) > 0
        if not alpha_ok:
            from PIL import Image as _I
            _I.open(io.BytesIO(info["image"])).convert("RGBA").getchannel("A")
            alpha_ok = True   # PNG 内嵌 alpha 通道即视为透明
        check(alpha_ok, "印章带透明通道")
        check(r1 and r1[0].y0 < 200, "第 2 页印章位于页面上部（标题处）")
    finally:
        out.close()

    print("== 自检" + ("全部通过 ✔" if ok else "存在失败 ✘") + " ==")
    return 0 if ok else 1


# ---------------------------------------------------------------------------
# 9. 图形界面
# ---------------------------------------------------------------------------
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    TK_AVAILABLE = True
except Exception:                    # 无 Tk 环境时仅 CLI 模式（自检/示例）可用
    tk = None
    filedialog = messagebox = ttk = None
    TK_AVAILABLE = False


class StampApp:
    """Tkinter 主界面（规格 五）。"""

    def __init__(self, root: tk.Tk):
        from PIL import ImageTk          # 延迟导入：无 GUI 模式（自检）不依赖 Tk
        self.ImageTk = ImageTk
        self.root = root
        self.proc = StampProcessor()

        # 状态
        self.pdf_path: str | None = None
        self.doc = None
        self.stamp_path: str | None = None
        self.stamp_base_key = None
        self.placements: list[StampPlacement] = []
        self.selected: StampPlacement | None = None
        self.current_page = 0
        self.zoom = 1.0
        self.page_scale = 1.0          # 屏幕 px / PDF pt
        self.ox = self.oy = 20.0       # 页面图像在 canvas 上的左上角
        self.page_pil = None
        self.page_photo = None
        self._stamp_photos: list = []  # 防止 PhotoImage 被 GC
        self._disp: dict = {}          # pl -> (sx, sy, w, h) 画布显示几何
        self._drag = None
        self.dirty = False
        self._syncing = False
        self._busy = False

        # 当前参数（滑条；选中章时联动）
        self.var_size = tk.DoubleVar(value=DEFAULT_SIZE_MM)
        self.var_opacity = tk.DoubleVar(value=DEFAULT_OPACITY)
        self.var_angle = tk.DoubleVar(value=DEFAULT_ANGLE)
        self.var_remove_white = tk.BooleanVar(value=DEFAULT_REMOVE_WHITE)
        self.var_threshold = tk.IntVar(value=DEFAULT_WHITE_THRESHOLD)

        self._build_ui()
        self._bind_shortcuts()
        self._set_fonts()
        self.status.set("欢迎使用 %s v%s —— 先上传印章，再打开 PDF" % (APP_NAME, APP_VERSION))
        self._sync_param_widgets()
        self._update_title()
        self._update_all()

    # ------------------------------------------------------------------ UI
    def _set_fonts(self):
        family = "Microsoft YaHei UI" if sys.platform == "win32" else "WenQuanYi Micro Hei"
        try:
            self.root.option_add("*Font", (family, 10))
        except Exception:
            pass
        style = ttk.Style(self.root)
        try:
            style.configure("Toolbar.TButton", padding=(8, 4))
        except Exception:
            pass

    def _build_ui(self):
        self.root.title(APP_NAME)
        self.root.minsize(1120, 700)
        self.root.geometry("1280x800")

        # ---------------- 顶栏 ----------------
        top = ttk.Frame(self.root, padding=(8, 6))
        top.pack(side=tk.TOP, fill=tk.X)
        ttk.Button(top, text="① 上传印章", command=self.open_stamp).pack(side=tk.LEFT, padx=2)
        ttk.Button(top, text="② 打开 PDF", command=self.open_pdf).pack(side=tk.LEFT, padx=2)
        ttk.Separator(top, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=2)
        ttk.Button(top, text="◀ 上一页", command=self.prev_page).pack(side=tk.LEFT, padx=2)
        self.lbl_page = ttk.Label(top, text="第 0/0 页", width=10, anchor=tk.CENTER)
        self.lbl_page.pack(side=tk.LEFT, padx=2)
        ttk.Button(top, text="下一页 ▶", command=self.next_page).pack(side=tk.LEFT, padx=2)
        ttk.Separator(top, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=2)
        ttk.Button(top, text="缩小", command=lambda: self.zoom_by(1 / 1.15)).pack(side=tk.LEFT, padx=2)
        ttk.Button(top, text="放大", command=lambda: self.zoom_by(1.15)).pack(side=tk.LEFT, padx=2)
        ttk.Button(top, text="适合窗口", command=self.fit_window).pack(side=tk.LEFT, padx=2)
        ttk.Separator(top, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=2)
        ttk.Button(top, text="★ 自动盖章", command=self.auto_stamp).pack(side=tk.LEFT, padx=2)
        ttk.Button(top, text="导出 PDF", command=self.export).pack(side=tk.LEFT, padx=2)

        # ---------------- 主体 ----------------
        main = ttk.Panedwindow(self.root, orient=tk.HORIZONTAL)
        main.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=(4, 4))

        # 左侧参数面板
        left = ttk.Frame(main, width=300)
        main.add(left, weight=0)
        self._build_left_panel(left)

        # 右侧 PDF 预览
        right = ttk.Frame(main)
        main.add(right, weight=1)
        self._build_canvas(right)

        # ---------------- 底栏 ----------------
        bottom = ttk.Frame(self.root, padding=(8, 4))
        bottom.pack(side=tk.BOTTOM, fill=tk.X)
        self.status = tk.StringVar()
        self.coord = tk.StringVar(value="x=—  y=—")
        ttk.Label(bottom, textvariable=self.status).pack(side=tk.LEFT)
        ttk.Label(bottom, textvariable=self.coord).pack(side=tk.RIGHT)

    def _build_left_panel(self, parent):
        # 印章预览
        box = ttk.LabelFrame(parent, text="印章预览（处理后的效果）", padding=6)
        box.pack(fill=tk.X, padx=6, pady=4)
        self.preview_canvas = tk.Canvas(box, width=150, height=150,
                                        bg="#f2f2f2", highlightthickness=1,
                                        highlightbackground="#cccccc")
        self.preview_canvas.pack(padx=4, pady=4)
        self.lbl_stamp_name = ttk.Label(box, text="（尚未上传印章）", wraplength=270,
                                        anchor=tk.W, foreground="#666666")
        self.lbl_stamp_name.pack(fill=tk.X, padx=2, pady=2)
        self._preview_photo = None

        # 参数
        box = ttk.LabelFrame(parent, text="印章参数", padding=8)
        box.pack(fill=tk.X, padx=6, pady=4)

        self._add_slider(box, "直径（mm）", self.var_size, 10, 80, 0.5, self.on_param_size, "10~80")
        self._add_slider(box, "透明度", self.var_opacity, 0.2, 1.0, 0.01, self.on_param_opacity, "0.2~1.0")
        self._add_slider(box, "旋转角（°）", self.var_angle, -180, 180, 1, self.on_param_angle, "逆时针为正")

        row = ttk.Frame(box)
        row.pack(fill=tk.X, pady=(6, 2))
        ttk.Checkbutton(row, text="去除印章白色背景", variable=self.var_remove_white,
                        command=self.on_param_remove_white).pack(side=tk.LEFT)
        ttk.Frame(box).pack()  # 占位

        row2 = ttk.Frame(box)
        row2.pack(fill=tk.X)
        ttk.Label(row2, text="白色阈值").pack(side=tk.LEFT)
        self._add_slider(row2, "", self.var_threshold, 200, 250, 1, self.on_param_threshold, "", width=140)

        # 本页印章列表
        box = ttk.LabelFrame(parent, text="本页印章列表", padding=6)
        box.pack(fill=tk.BOTH, expand=True, padx=6, pady=4)
        listframe = ttk.Frame(box)
        listframe.pack(fill=tk.BOTH, expand=True)
        self.listbox = tk.Listbox(listframe, height=8, activestyle="dotbox",
                                  selectmode=tk.SINGLE)
        sb = ttk.Scrollbar(listframe, orient=tk.VERTICAL, command=self.listbox.yview)
        self.listbox.config(yscrollcommand=sb.set)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.bind("<<ListboxSelect>>", self.on_listbox_select)

        btns = ttk.Frame(box)
        btns.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(btns, text="删除选中", command=self.delete_selected).pack(side=tk.LEFT, padx=2)
        ttk.Button(btns, text="清空本页", command=self.clear_page).pack(side=tk.LEFT, padx=2)
        ttk.Button(btns, text="清空全部", command=self.clear_all).pack(side=tk.LEFT, padx=2)

        tip = ttk.Label(parent, text="提示：单击页面空白处盖章；\n"
                        "单击已有印章选中并拖动移动；\n"
                        "选中后调整参数实时生效；\n"
                        "Ctrl+滚轮缩放，Delete 删除。",
                        foreground="#666666", justify=tk.LEFT)
        tip.pack(fill=tk.X, padx=10, pady=6)

    def _add_slider(self, parent, label, var, from_, to, res, cmd, range_text, width=180):
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=2)
        if label:
            ttk.Label(row, text=label, width=11, anchor=tk.W).pack(side=tk.LEFT)
        tk.Scale(row, from_=from_, to=to, resolution=res, orient=tk.HORIZONTAL,
                 variable=var, command=cmd, length=width, showvalue=True,
                 highlightthickness=0).pack(side=tk.LEFT, fill=tk.X, expand=True)
        if range_text:
            ttk.Label(row, text=range_text, width=9, anchor=tk.E,
                      foreground="#999999").pack(side=tk.LEFT)

    def _build_canvas(self, parent):
        wrap = ttk.Frame(parent)
        wrap.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(wrap, bg="#3a3f44", highlightthickness=0)
        vsb = ttk.Scrollbar(wrap, orient=tk.VERTICAL, command=self.canvas.yview)
        hsb = ttk.Scrollbar(wrap, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.canvas.configure(xscrollcommand=hsb.set, yscrollcommand=vsb.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)

        self.canvas.bind("<ButtonPress-1>", self.on_canvas_press)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        self.canvas.bind("<Motion>", self.on_canvas_motion)
        self.canvas.bind("<MouseWheel>", self.on_wheel)          # Windows
        self.canvas.bind("<Button-4>", lambda e: self.on_wheel_linux(e, 1))   # Linux
        self.canvas.bind("<Button-5>", lambda e: self.on_wheel_linux(e, -1))
        self.canvas.bind("<Configure>", lambda e: None)

    def _bind_shortcuts(self):
        self.root.bind("<Control-o>", lambda e: self.open_pdf())
        self.root.bind("<Control-i>", lambda e: self.open_stamp())
        self.root.bind("<Control-s>", lambda e: self.open_stamp())
        self.root.bind("<Control-e>", lambda e: self.export())
        self.root.bind("<Control-0>", lambda e: self.fit_window())
        self.root.bind("<Control-plus>", lambda e: self.zoom_by(1.15))
        self.root.bind("<Control-minus>", lambda e: self.zoom_by(1 / 1.15))
        self.root.bind("<Delete>", lambda e: self.delete_selected())
        self.root.bind("<Left>", lambda e: self.prev_page())
        self.root.bind("<Right>", lambda e: self.next_page())
        self.root.bind("<Escape>", lambda e: self.select(None))
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ------------------------------------------------------------ 文件操作
    def open_stamp(self):
        if self._busy:
            return
        path = filedialog.askopenfilename(
            title="选择印章图片",
            filetypes=[("图片文件", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp"),
                       ("所有文件", "*.*")])
        if not path:
            return
        self.load_stamp(path)

    def load_stamp(self, path: str):
        key, img, err = self.proc.prepare(
            path, self.var_remove_white.get(), self.var_threshold.get(),
            self.var_opacity.get())
        if key is None:
            messagebox.showerror(APP_NAME, err or "印章加载失败")
            return
        self.stamp_path = path
        self.stamp_base_key = key
        self.lbl_stamp_name.config(text=os.path.basename(path),
                                   foreground="#000000")
        self.status.set("印章已加载：%s" % os.path.basename(path))
        self.dirty = True
        self._update_preview()
        self._redraw_stamps()

    def open_pdf(self):
        if self._busy:
            return
        if self.dirty and self.placements:
            r = messagebox.askyesnocancel(
                APP_NAME, "当前有未导出的修改，打开新 PDF 将丢弃这些印章。\n"
                          "是否先导出当前结果？")
            if r is None:
                return
            if r:
                if not self.export():
                    return
        path = filedialog.askopenfilename(
            title="打开 PDF 文件", filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            doc = fitz.open(path)
        except Exception as e:
            messagebox.showerror(APP_NAME, "无法打开 PDF：%s" % e)
            return
        if self.doc is not None:
            self.doc.close()
        self.doc = doc
        self.pdf_path = path
        self.placements.clear()
        self.select(None)
        self.current_page = 0
        self.zoom = 1.0
        self.dirty = False
        self.status.set("已打开：%s（共 %d 页，A4≈595×842pt）"
                        % (os.path.basename(path), doc.page_count))
        self._update_title()
        self._update_all()
        self.fit_window()

    # ------------------------------------------------------------ 页面与缩放
    def _render_page(self):
        """按当前页与 zoom 渲染页面图像（规格 3.4）。"""
        page = self.doc[self.current_page]
        scale = DEFAULT_PREVIEW_DPI / 72.0 * self.zoom
        mat = fitz.Matrix(scale, scale)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        self.page_pil = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        self.page_scale = scale

    def _update_all(self):
        """重绘页面 + 印章 + 列表 + 页码。"""
        if self.doc is None:
            self.canvas.delete("all")
            self.canvas.create_text(self.canvas.winfo_width() / 2 or 300,
                                    self.canvas.winfo_height() / 2 or 300,
                                    text="请先点击「② 打开 PDF」\n再点击页面空白处盖章",
                                    fill="#aaaaaa", font=("", 16), justify=tk.CENTER,
                                    tags="hint")
            self.lbl_page.config(text="第 0/0 页")
            self._update_listbox()
            return
        self._render_page()
        self._photo = self.ImageTk.PhotoImage(self.page_pil)   # 引用保持
        self.canvas.delete("all")
        self.canvas.create_image(self.ox, self.oy, image=self._photo,
                                 anchor=tk.NW, tags="page")
        self.canvas.config(scrollregion=(0, 0,
                                         self.ox + self.page_pil.width + 20,
                                         self.oy + self.page_pil.height + 20))
        self.lbl_page.config(text="第 %d/%d 页" % (self.current_page + 1,
                                                  self.doc.page_count))
        self._redraw_stamps()
        self._update_listbox()

    def canvas_to_pdf(self, sx: float, sy: float):
        """屏幕坐标 → PDF 坐标（规格 3.4）。页面外返回 None。"""
        if self.page_pil is None:
            return None
        x_pt = (sx - self.ox) / self.page_scale
        y_pt = (sy - self.oy) / self.page_scale
        page = self.doc[self.current_page]
        if x_pt < 0 or y_pt < 0 or x_pt > page.rect.width or y_pt > page.rect.height:
            return None
        return x_pt, y_pt

    def pdf_to_canvas(self, x_pt: float, y_pt: float):
        return (self.ox + x_pt * self.page_scale,
                self.oy + y_pt * self.page_scale)

    def _stamp_fit_px(self, pl: StampPlacement) -> int:
        size_pt = pl.size_mm * PT_PER_MM
        return max(8, int(round(size_pt * self.page_scale)))

    def set_page(self, idx: int):
        if self.doc is None:
            return
        idx = max(0, min(idx, self.doc.page_count - 1))
        if idx == self.current_page:
            return
        self.current_page = idx
        self.select(None)
        self._update_all()

    def prev_page(self):
        self.set_page(self.current_page - 1)

    def next_page(self):
        self.set_page(self.current_page + 1)

    def zoom_by(self, factor: float, anchor=None):
        if self.doc is None or self._busy:
            return
        # 以视口中心为锚点缩放
        if anchor is None:
            anchor = self.canvas_to_pdf(self.canvas.winfo_width() / 2,
                                        self.canvas.winfo_height() / 2)
        new_zoom = min(max(self.zoom * factor, 0.10), 12.0)
        if abs(new_zoom - self.zoom) < 1e-9:
            return
        self.zoom = new_zoom
        self._update_all()
        if anchor:
            sx, sy = self.pdf_to_canvas(*anchor)
            dx = self.canvas.winfo_width() / 2 - sx
            dy = self.canvas.winfo_height() / 2 - sy
            self.ox += dx
            self.oy += dy
            self._update_all()
        self.status.set("缩放：%.0f%%" % (self.zoom * 100))

    def fit_window(self):
        if self.doc is None:
            return
        page = self.doc[self.current_page]
        cw = max(200, self.canvas.winfo_width() - 60)
        ch = max(200, self.canvas.winfo_height() - 60)
        z = min(cw / page.rect.width, ch / page.rect.height)
        # 基于基准 DPI 的 zoom 倍数
        base = DEFAULT_PREVIEW_DPI / 72.0
        self.zoom = min(max(z / base, 0.05), 12.0)
        self._update_all()
        self.ox = max(20, (self.canvas.winfo_width() - self.page_pil.width) / 2)
        self.oy = max(20, (self.canvas.winfo_height() - self.page_pil.height) / 2)
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)
        self._update_all()
        self.status.set("适合窗口：%.0f%%" % (self.zoom * 100))

    def on_wheel(self, e):
        if self._busy or self.doc is None:
            return
        if e.state & 0x0004:   # Ctrl 按下 → 缩放
            self.zoom_by(1.12 if e.delta > 0 else 1 / 1.12)
        else:
            self.canvas.yview_scroll(-1 if e.delta > 0 else 1, "units")

    def on_wheel_linux(self, e, direction):
        if self._busy or self.doc is None:
            return
        if e.state & 0x0004:
            self.zoom_by(1.12 if direction > 0 else 1 / 1.12)
        else:
            self.canvas.yview_scroll(-1 if direction > 0 else 1, "units")

    # ------------------------------------------------------------ 盖章交互
    def on_canvas_press(self, e):
        if self._busy or self.doc is None or self.page_pil is None:
            return
        # 1) 命中检测：点中已有印章 → 选中并准备拖动（后盖的在上面，倒序）
        for pl in reversed([p for p in self.placements if p.page_index == self.current_page]):
            info = self._disp.get(pl)
            if not info:
                continue
            sx, sy, w, h = info
            if abs(e.x - sx) <= w / 2 + 3 and abs(e.y - sy) <= h / 2 + 3:
                self.select(pl)
                self._drag = (e.x, e.y, pl.cx, pl.cy)
                self.root.config(cursor="fleur")
                return
        # 2) 空白处单击 → 以点击处为章心放置一枚新章
        if not self.stamp_path:
            messagebox.showinfo(APP_NAME, "请先上传印章图片（点击顶栏「① 上传印章」）。")
            return
        pt = self.canvas_to_pdf(e.x, e.y)
        if pt is None:
            return
        cx, cy = pt
        pl = StampPlacement(
            page_index=self.current_page, cx=cx, cy=cy,
            size_mm=round(self.var_size.get(), 2),
            opacity=round(self.var_opacity.get(), 3),
            angle=round(self.var_angle.get(), 1),
            stamp_path=self.stamp_path,
            remove_white=self.var_remove_white.get(),
            white_threshold=int(self.var_threshold.get()))
        self.placements.append(pl)
        self.select(pl)
        self.dirty = True
        self._redraw_stamps()
        self._update_listbox()
        self.status.set("已盖章：第 %d 页 (%d, %d) pt" % (self.current_page + 1,
                                                        round(pl.cx), round(pl.cy)))

    def on_canvas_drag(self, e):
        if self._drag is None or self.selected is None or self.page_scale <= 0:
            return
        sx0, sy0, cx0, cy0 = self._drag
        pl = self.selected
        page = self.doc[self.current_page]
        size_pt = pl.size_mm * PT_PER_MM
        half = size_pt / 2
        new_cx = cx0 + (e.x - sx0) / self.page_scale
        new_cy = cy0 + (e.y - sy0) / self.page_scale
        # 限制在页面范围内（规格 5.2）
        pl.cx = min(max(new_cx, half), page.rect.width - half)
        pl.cy = min(max(new_cy, half), page.rect.height - half)
        self.dirty = True
        self._redraw_stamps()

    def on_canvas_release(self, e):
        if self._drag is not None:
            self._drag = None
            self.root.config(cursor="")
            if self.selected:
                self.status.set("已移动：第 %d 页 (%d, %d) pt"
                                % (self.current_page + 1,
                                   round(self.selected.cx), round(self.selected.cy)))

    def on_canvas_motion(self, e):
        pt = self.canvas_to_pdf(e.x, e.y) if self.doc else None
        if pt:
            self.coord.set("x=%.1f  y=%.1f pt" % pt)
        else:
            self.coord.set("x=—  y=—")

    # ------------------------------------------------------------ 选中与参数
    def select(self, pl: StampPlacement | None):
        self.selected = pl
        if pl is not None:
            self._syncing = True
            try:
                self.var_size.set(round(pl.size_mm, 2))
                self.var_opacity.set(round(pl.opacity, 3))
                self.var_angle.set(round(pl.angle, 1))
                self.var_remove_white.set(pl.remove_white)
                self.var_threshold.set(int(pl.white_threshold))
            finally:
                self._syncing = False
            # 同步列表框
            for i, p in enumerate(self.page_placements()):
                if p is pl:
                    self.listbox.selection_clear(0, tk.END)
                    self.listbox.selection_set(i)
                    self.listbox.see(i)
                    break
        else:
            self.listbox.selection_clear(0, tk.END)
        self._update_preview()
        self._draw_selection()

    def page_placements(self) -> list:
        return [p for p in self.placements if p.page_index == self.current_page]

    def on_listbox_select(self, _e):
        sel = self.listbox.curselection()
        if not sel:
            return
        items = self.page_placements()
        if sel[0] < len(items):
            self.select(items[sel[0]])

    # 参数滑条回调：选中章实时更新；未选中则作为默认参数
    def _on_param_change(self):
        if self._syncing:
            return
        pl = self.selected
        if pl is not None:
            pl.size_mm = round(self.var_size.get(), 2)
            pl.opacity = round(self.var_opacity.get(), 3)
            pl.angle = round(self.var_angle.get(), 1)
            pl.remove_white = bool(self.var_remove_white.get())
            pl.white_threshold = int(self.var_threshold.get())
            self.dirty = True
        # 预览始终显示当前参数效果
        self._update_preview()
        self._redraw_stamps()
        self._update_listbox()

    def on_param_size(self, _v=None):
        self._on_param_change()

    def on_param_opacity(self, _v=None):
        self._on_param_change()

    def on_param_angle(self, _v=None):
        self._on_param_change()

    def on_param_remove_white(self):
        self._on_param_change()

    def on_param_threshold(self, _v=None):
        self._on_param_change()

    def _sync_param_widgets(self):
        pass  # 参数已绑定变量，无需额外同步

    # ------------------------------------------------------------ 预览
    def _update_preview(self):
        """左侧印章预览：处理后的图 + 当前旋转角，固定显示高度 140px。"""
        if not self.stamp_path:
            self.preview_canvas.delete("all")
            self.preview_canvas.create_text(75, 75, text="暂无印章", fill="#999999")
            return
        key, img, err = self.proc.prepare(
            self.stamp_path, self.var_remove_white.get(), self.var_threshold.get(),
            self.var_opacity.get())
        if key is None:
            return
        if self.var_angle.get() % 360:
            img = img.rotate(float(self.var_angle.get()), expand=True,
                             resample=Image.BICUBIC)
        w, h = img.size
        scale = 140.0 / max(w, h)
        img = img.resize((max(1, int(w * scale)), max(1, int(h * scale))),
                         Image.LANCZOS)
        self._preview_photo = self.ImageTk.PhotoImage(img)
        self.preview_canvas.delete("all")
        self.preview_canvas.create_image(75, 75, image=self._preview_photo)

    # ------------------------------------------------------------ 绘制印章
    def _redraw_stamps(self):
        if self.canvas is None:
            return
        self.canvas.delete("stamp")
        self.canvas.delete("selbox")
        self._stamp_photos.clear()
        self._disp.clear()
        if self.doc is None:
            return
        for pl in self.placements:
            if pl.page_index != self.current_page:
                continue
            key, _, err = self.proc.prepare(pl.stamp_path, pl.remove_white,
                                            pl.white_threshold, pl.opacity)
            if key is None:
                continue
            fit_px = self._stamp_fit_px(pl)
            img = self.proc.render(key, pl.angle, fit_px)
            if img is None:
                continue
            sx, sy = self.pdf_to_canvas(pl.cx, pl.cy)
            photo = self.ImageTk.PhotoImage(img)
            self._stamp_photos.append(photo)
            self.canvas.create_image(sx, sy, image=photo, tags="stamp")
            self._disp[pl] = (sx, sy, img.size[0], img.size[1])
        self._draw_selection()

    def _draw_selection(self):
        self.canvas.delete("selbox")
        pl = self.selected
        if pl is None or pl.page_index != self.current_page:
            return
        info = self._disp.get(pl)
        if not info:
            return
        sx, sy, w, h = info
        self.canvas.create_rectangle(sx - w / 2 - 4, sy - h / 2 - 4,
                                     sx + w / 2 + 4, sy + h / 2 + 4,
                                     outline="#2b7de9", width=2, dash=(5, 3),
                                     tags="selbox")

    def _update_listbox(self):
        self.listbox.delete(0, tk.END)
        for i, pl in enumerate(self.page_placements()):
            self.listbox.insert(tk.END, "%d. %gmm  透明度%.2f  旋转%.0f°"
                                % (i + 1, pl.size_mm, pl.opacity, pl.angle))
        if self.selected is not None:
            for i, p in enumerate(self.page_placements()):
                if p is self.selected:
                    self.listbox.selection_set(i)
                    break

    # ------------------------------------------------------------ 删除
    def delete_selected(self):
        pl = self.selected
        if pl is None:
            messagebox.showinfo(APP_NAME, "请先在页面或列表中选中一枚印章。")
            return
        self.placements.remove(pl)
        self.select(None)
        self.dirty = True
        self._redraw_stamps()
        self._update_listbox()
        self.status.set("已删除第 %d 页的一枚印章" % (pl.page_index + 1))

    def clear_page(self):
        n = len(self.page_placements())
        if n == 0:
            return
        self.placements = [p for p in self.placements
                           if p.page_index != self.current_page]
        self.select(None)
        self.dirty = True
        self._redraw_stamps()
        self._update_listbox()
        self.status.set("已清空第 %d 页的 %d 枚印章" % (self.current_page + 1, n))

    def clear_all(self):
        if not self.placements:
            return
        self.placements.clear()
        self.select(None)
        self.dirty = True
        self._redraw_stamps()
        self._update_listbox()
        self.status.set("已清空全部印章")

    # ------------------------------------------------------------ 自动盖章
    def auto_stamp(self):
        if self._busy:
            return
        if self.doc is None:
            messagebox.showwarning(APP_NAME, "请先打开 PDF 文件。")
            return
        if not self.stamp_path:
            messagebox.showwarning(APP_NAME, "请先上传印章图片。")
            return
        placements, msgs = compute_auto_placements(
            self.doc, self.stamp_path,
            size_mm=round(self.var_size.get(), 2),
            opacity=round(self.var_opacity.get(), 3),
            angle=round(self.var_angle.get(), 1),
            remove_white=bool(self.var_remove_white.get()),
            white_threshold=int(self.var_threshold.get()))
        added = 0
        for pl in placements:
            # 去重：同一页、同一位置、同一印章不重复盖
            dup = any(p.page_index == pl.page_index and p.stamp_path == pl.stamp_path
                      and abs(p.cx - pl.cx) < 1 and abs(p.cy - pl.cy) < 1
                      and abs(p.size_mm - pl.size_mm) < 0.1
                      for p in self.placements)
            if not dup:
                self.placements.append(pl)
                added += 1
        if added:
            self.dirty = True
        self._update_all()
        miss = [m for m in msgs if "未找到" in m]
        if miss:
            head = "自动盖章完成（新增 %d 枚）。\n\n以下关键词未找到，可手动补盖：\n" % added
            messagebox.showwarning(APP_NAME, head + "\n".join("· " + m for m in miss))
        else:
            messagebox.showinfo(APP_NAME, "自动盖章完成，共新增 %d 枚印章。\n\n%s"
                                % (added, "\n".join(msgs)))

    # ------------------------------------------------------------ 导出
    def export(self, ask_path: bool = True) -> bool:
        if self._busy:
            return False
        if self.doc is None:
            messagebox.showwarning(APP_NAME, "还没有打开 PDF，无法导出。\n请先点击「② 打开 PDF」。")
            return False
        if not self.placements:
            messagebox.showwarning(APP_NAME, "当前没有印章，请先手动盖章或点「★ 自动盖章」。")
            return False
        if not self.stamp_path:
            messagebox.showwarning(APP_NAME, "请先上传印章图片。")
            return False

        dst = None
        if ask_path:
            default = os.path.splitext(self.pdf_path)[0] + "_已盖章.pdf"
            dst = filedialog.asksaveasfilename(
                title="导出已盖章 PDF", defaultextension=".pdf",
                initialfile=os.path.basename(default),
                initialdir=os.path.dirname(default) or None,
                filetypes=[("PDF 文件", "*.pdf")])
            if not dst:
                return False
        else:
            dst = os.path.splitext(self.pdf_path)[0] + "_已盖章.pdf"
            if os.path.abspath(dst) == os.path.abspath(self.pdf_path):
                dst += ".new.pdf"
        self._set_busy(True, "正在导出…")
        threading.Thread(target=self._export_worker,
                         args=(dst,), daemon=True).start()
        return True

    def _export_worker(self, dst: str):
        try:
            def progress(done, total, msg):
                self.root.after(0, lambda d=done, t=total, m=msg: self.status.set(
                    "正在导出：%d/%d  %s" % (d, t, m)))

            skipped = export_pdf(self.pdf_path, dst, list(self.placements),
                                 self.proc, progress)
            self.root.after(0, lambda: self._export_done(dst, skipped, None))
        except Exception:
            self.root.after(0, lambda: self._export_done(None, [],
                                                         traceback.format_exc()))

    def _export_done(self, dst, skipped, err):
        self._set_busy(False)
        if err:
            messagebox.showerror(APP_NAME, "导出失败：\n%s" % err)
            return
        size_kb = os.path.getsize(dst) / 1024.0
        msg = "导出成功：\n%s\n文件大小：%.0f KB\n原文件未被修改。" % (dst, size_kb)
        if skipped:
            msg += "\n\n注意：%d 枚印章因页号超出范围或印章无效被跳过。" % len(skipped)
        self.dirty = False
        self.status.set("已导出：%s（%.0f KB）" % (os.path.basename(dst), size_kb))
        if messagebox.askyesno(APP_NAME, msg + "\n\n是否打开文件所在文件夹？"):
            self._open_folder(os.path.dirname(dst))

    @staticmethod
    def _open_folder(path):
        try:
            if sys.platform == "win32":
                os.startfile(path)  # type: ignore
            elif sys.platform == "darwin":
                os.system('open "%s"' % path)
            else:
                os.system('xdg-open "%s"' % path)
        except Exception:
            pass

    # ------------------------------------------------------------ 其他
    def _set_busy(self, busy: bool, status: str | None = None):
        self._busy = busy
        self.root.config(cursor="watch" if busy else "")
        if status:
            self.status.set(status)

    def _update_title(self):
        name = os.path.basename(self.pdf_path) if self.pdf_path else "未打开 PDF"
        self.root.title("%s v%s — %s" % (APP_NAME, APP_VERSION, name))

    def on_close(self):
        if self.dirty and self.placements:
            r = messagebox.askyesnocancel(
                APP_NAME, "有未导出的修改，退出前是否先导出？")
            if r is None:
                return
            if r:
                if not self.export():
                    return
        self.root.destroy()


# ---------------------------------------------------------------------------
# 10. 入口
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description=APP_NAME)
    parser.add_argument("--selftest", action="store_true",
                        help="无界面自检（预处理/自动盖章/导出）")
    parser.add_argument("--demo", action="store_true",
                        help="在 assets/demo 生成示例印章与示例 PDF")
    args = parser.parse_args()

    if args.selftest:
        sys.exit(run_selftest())
    if args.demo:
        assets = resource_path("assets")
        demo = os.path.join(assets, "demo")
        os.makedirs(demo, exist_ok=True)
        p1 = generate_demo_seal(os.path.join(demo, "示例印章.png"), white_bg=False)
        p2 = generate_demo_seal(os.path.join(demo, "示例印章_白底.jpg"), white_bg=True)
        p3 = generate_demo_pdf(os.path.join(demo, "示例文档.pdf"))
        print("已生成：\n  %s\n  %s\n  %s" % (p1, p2, p3))
        sys.exit(0)

    if not TK_AVAILABLE:
        print("当前环境缺少 Tkinter（无图形界面支持），无法启动 GUI。\n"
              "Linux 请安装 python3-tk；Windows 请安装完整版 Python。",
              file=sys.stderr)
        sys.exit(2)
    root = tk.Tk()
    try:
        StampApp(root)
        root.mainloop()
    except tk.TclError as e:
        print("无法启动图形界面（缺少显示器或 Tk 环境）：%s" % e, file=sys.stderr)
        sys.exit(2)


if __name__ == "__main__":
    main()
