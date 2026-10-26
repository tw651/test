# -*- coding: utf-8 -*-
"""按论文《9版_投稿修订稿》内容修订开题 PPT。"""
import copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn, nsdecls
from pptx.oxml import parse_xml
from pptx.text.text import _Paragraph as Paragraph
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

DARK  = RGBColor(0x2A, 0x2D, 0x30)
BODY  = RGBColor(0x3C, 0x41, 0x3C)
GREEN = RGBColor(0x4C, 0x7A, 0x5A)
RED   = RGBColor(0x7B, 0x25, 0x32)
GOLD  = RGBColor(0xB4, 0x8C, 0x46)
GRAY  = RGBColor(0x78, 0x82, 0x78)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SERIF = 'Noto Serif SC'
SANS  = 'Noto Sans SC'
BORDER = 'D5DCD3'
HEADFILL = 'E7EDE4'

SRC = '张石林-开题报告.pptx'
DST = '张石林-开题报告_修订版.pptx'
prs = Presentation(SRC)


# ---------------------------------------------------------------- helpers
def style_run(run, size, bold, color, font, italic=False):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.name = font
    f.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    for tag in ('a:ea', 'a:cs'):
        el = rPr.find(qn(tag))
        if el is None:
            el = parse_xml('<%s %s typeface="%s"/>' % (tag, nsdecls('a'), font))
            rPr.append(el)
        else:
            el.set('typeface', font)


def set_lines(shape, lines, size=13, color=BODY, font=SERIF, bold=False,
              align=None):
    """lines: list of (str | list[(text, override_dict)])"""
    tf = shape.text_frame
    txBody = tf._txBody
    old_ps = txBody.findall(qn('a:p'))
    tmpl = copy.deepcopy(old_ps[0]) if old_ps else None
    for p in old_ps:
        txBody.remove(p)
    if tmpl is None:
        tmpl = parse_xml('<a:p %s/>' % nsdecls('a'))
    for line in lines:
        specs = line if isinstance(line, list) else [(line, {})]
        p_el = copy.deepcopy(tmpl)
        for tag in ('a:r', 'a:br', 'a:fld'):
            for el in p_el.findall(qn(tag)):
                p_el.remove(el)
        txBody.append(p_el)
        para = Paragraph(p_el, tf)
        if align is not None:
            para.alignment = align
        for text, ov in specs:
            run = para.add_run()
            run.text = text
            style_run(run,
                      ov.get('size', size),
                      ov.get('bold', bold),
                      ov.get('color', color),
                      ov.get('font', font),
                      ov.get('italic', False))


def find(slide, name):
    for sh in slide.shapes:
        if sh.name == name:
            return sh
    raise KeyError('%s not found' % name)


def cell_borders(cell, edges=('B',), color=BORDER, width=9525):
    tcPr = cell._tc.get_or_add_tcPr()
    for e in ('L', 'R', 'T', 'B'):
        tag = 'a:ln' + e
        old = tcPr.find(qn(tag))
        if old is not None:
            tcPr.remove(old)
    for i, e in enumerate(('L', 'R', 'T', 'B')):
        tag = 'a:ln' + e
        if e in edges:
            xml = ('<%%s %s w="%d" cap="flat" cmpd="sng" algn="ctr">'
                   '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
                   '<a:prstDash val="solid"/></%%s>' % (nsdecls('a'), width, color))
            xml = xml % (tag, tag)
        else:
            xml = '<%s %s w="0"><a:noFill/></%s>' % (tag, nsdecls('a'), tag)
        tcPr.insert(i, parse_xml(xml))


def fill_cell(cell, rgb):
    tcPr = cell._tc.get_or_add_tcPr()
    old = tcPr.find(qn('a:noFill'))
    if old is not None:
        tcPr.remove(old)
    for el in tcPr.findall(qn('a:solidFill')):
        tcPr.remove(el)
    sf = parse_xml('<a:solidFill %s><a:srgbClr val="%s"/></a:solidFill>'
                   % (nsdecls('a'), rgb))
    tcPr.append(sf)


def write_table(tbl, data, size=10.5, header_size=None, head_color=DARK,
                body_color=BODY, align=PP_ALIGN.CENTER, head_fill=HEADFILL,
                first_col_left=False):
    header_size = header_size or size
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            tf = cell.text_frame
            tf.word_wrap = True
            txBody = tf._txBody
            for p in txBody.findall(qn('a:p'))[1:]:
                txBody.remove(p)
            p_el = txBody.findall(qn('a:p'))[0]
            for tag in ('a:r', 'a:br', 'a:fld'):
                for el in p_el.findall(qn(tag)):
                    p_el.remove(el)
            para = Paragraph(p_el, tf)
            para.alignment = (PP_ALIGN.LEFT if (first_col_left and j == 0)
                              else align)
            run = para.add_run()
            run.text = str(val)
            style_run(run, header_size if i == 0 else size,
                      i == 0, head_color if i == 0 else body_color, SANS)
            if i == 0 and head_fill:
                fill_cell(cell, head_fill)
            edges = ('T', 'B') if i == 0 else ('B',)
            cell_borders(cell, edges)


def add_table(slide, data, left, top, col_w, row_h, size=10.5,
              first_col_left=False):
    nrows, ncols = len(data), len(data[0])
    width = Inches(sum(col_w))
    height = Inches(row_h * nrows)
    gf = slide.shapes.add_table(nrows, ncols, Inches(left), Inches(top),
                                width, height)
    tbl = gf.table
    # 去掉默认蓝色表格样式
    tblPr = tbl._tbl.tblPr
    for el in tblPr.findall(qn('a:tableStyleId')):
        tblPr.remove(el)
    tbl.first_row = False
    tbl.horz_banding = False
    for j, w in enumerate(col_w):
        tbl.columns[j].width = Inches(w)
    for i in range(nrows):
        tbl.rows[i].height = Inches(row_h)
    write_table(tbl, data, size=size, first_col_left=first_col_left)
    return gf


def add_box(slide, left, top, width, height, lines, size=12, color=BODY,
            font=SERIF, bold=False, anchor=MSO_ANCHOR.TOP, spacing=1.15):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width),
                                  Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    set_lines(tb, lines, size=size, color=color, font=font, bold=bold)
    for p in tf.paragraphs:
        p.line_spacing = spacing
        p.space_before = Pt(0)
        p.space_after = Pt(0)
    return tb


def dup_slide(src):
    new = prs.slides.add_slide(src.slide_layout)
    for sh in list(new.shapes):
        sh._element.getparent().remove(sh._element)
    for sh in src.shapes:
        new.shapes._spTree.append(copy.deepcopy(sh._element))
    bg = src._element.cSld.find(qn('p:bg'))
    if bg is not None:
        new._element.cSld.insert(0, copy.deepcopy(bg))
    return new


def drop(slide, name):
    sh = find(slide, name)
    sh._element.getparent().remove(sh._element)


# ---------------------------------------------------------------- 1 全局文字修正
BAD = [('根半径和根半径', '根半径'),
       ('和根的半径', ''),
       ('F-K滤波', 'f-k 滤波'),
       ('消除训练-推理偏差', '减小训练-推理偏差'),
       ('取得最优平衡', '取得合理平衡')]
for slide in prs.slides:
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        for para in sh.text_frame.paragraphs:
            for run in para.runs:
                t = run.text
                for a, b in BAD:
                    t = t.replace(a, b)
                if t != run.text:
                    run.text = t

# ---------------------------------------------------------------- 2 目录
s = prs.slides[1]
set_lines(find(s, 'AutoShape 20'), ['实验结果与预期成果'], size=20, bold=True,
          color=DARK, font=SERIF)
set_lines(find(s, 'AutoShape 21'), ['定量结果与成果产出'], size=12, color=GRAY,
          font=SANS)

# ---------------------------------------------------------------- 3 背景：限定对象
s = prs.slides[2]
sh = find(s, 'AutoShape 5')
tf = sh.text_frame
p = tf.add_paragraph()
r = p.add_run()
r.text = '本文聚焦粗根（直径约 24–120 mm）的深度与半径反演。'
style_run(r, 16, False, GREEN, SERIF)

# ---------------------------------------------------------------- 4 观测方法
s = prs.slides[3]
set_lines(find(s, 'AutoShape 14'), [
    '微根管、探地雷达（GPR）、电阻率成像（ERT）',
    [('无损、可重复；微根管仅能观测细根且视域小；GPR 可大范围成像，但解释受杂波干扰与弱目标识别制约',
      {'color': BODY})],
], size=12)

# ---------------------------------------------------------------- 5 GPR 原理
s = prs.slides[4]
set_lines(find(s, 'AutoShape 10'), [
    'B-scan 中粗根呈特征性双曲线回波：顶点位置对应埋深，双曲线曲率与振幅与根半径相关，是根系识别与参数反演的核心依据。'
], size=11.5, color=GREEN)

# ---------------------------------------------------------------- 6 传统方法
s = prs.slides[5]
set_lines(find(s, 'AutoShape 9'), [
    [('传统方法：', {'bold': True, 'color': RED})],
    '背景去除（MS）、f-k 滤波、小波变换、SVD/RPCA 子空间分解、稀疏表示',
    [('局限：', {'bold': True, 'color': RED})],
    '依赖窗口/秩/阈值等人工参数；土壤介电特性变化时，背景压制与弱目标保真难以兼顾',
    [('本文复现：', {'bold': True, 'color': RED})],
    'f-k 与小波的 SSIM 仅 0.025、0.035，MS/SVD 约 0.73',
], size=11.5)
find(s, 'AutoShape 9').left = Inches(9.83)
find(s, 'AutoShape 9').width = Inches(3.05)

# ---------------------------------------------------------------- 7 深度学习
s = prs.slides[6]
set_lines(find(s, 'AutoShape 6'), [
    'U-Net[4]、自编码器[23,25]、GAN[24] 及混合数据训练[26] 已用于 GPR 图像恢复与去杂波；注意力机制与多尺度融合进一步提升了弱目标特征提取能力。'
], size=12.5)
set_lines(find(s, 'AutoShape 8'), [
    [('① ', {'color': RED, 'bold': True}),
     ('参数量与计算开销随多尺度融合、注意力模块增加，边缘部署困难', {})],
    [('② ', {'color': RED, 'bold': True}),
     ('仅报告 PSNR/SSIM 不能证明去杂波图能改善下游根深度与根半径预测', {})],
    [('③ ', {'color': RED, 'bold': True}),
     ('训练用的是理想 target、推理只能用前端输出，存在训练-推理分布偏移（如文献[5]）', {})],
], size=12.5)
set_lines(find(s, 'AutoShape 10'), [
    '核心挑战：图像恢复质量、网络效率与下游预测一致性——需在同一基准下同时报告质量、FLOPs/实测延迟与下游精度'
], size=14, bold=True, color=DARK)

# ---------------------------------------------------------------- 8 科学问题与创新
s = prs.slides[7]
set_lines(find(s, 'AutoShape 6'), [
    [('Q1　', {'bold': True, 'color': GREEN}),
     ('如何在保留弱根系双曲线回波的同时，有效抑制天线耦合、直达波与土壤杂波？', {})],
    [('Q2　', {'bold': True, 'color': GREEN}),
     ('如何在统一基准（质量/FLOPs/实测延迟/显存）下取得恢复质量与计算效率的合理平衡？', {})],
    [('Q3　', {'bold': True, 'color': GREEN}),
     ('如何使前端输出与后端输入分布一致，减小训练-推理偏差？', {})],
], size=13.5)
set_lines(find(s, 'AutoShape 9'), [
    [('创新1　', {'bold': True, 'color': RED}),
     ('渐进式轻量化路线 B0→L1→L2→L3：每步仅替换一个模块，相邻变体差值即该组件贡献，并在统一基准下同时报告质量、FLOPs、实测延迟与显存', {})],
    [('创新2　', {'bold': True, 'color': RED}),
     ('raw + L3 实际预测图双通道早期融合，并与晚期/门控融合对比，减小训练-推理分布偏移', {})],
    [('创新3　', {'bold': True, 'color': RED}),
     ('同一主干联合回归根深度与根半径，并给出 MS/SVD/f-k/小波/U-Net/B0 与传统回归的完整对照', {})],
], size=12.5)

# ---------------------------------------------------------------- 9 技术路线
s = prs.slides[8]
set_lines(find(s, 'AutoShape 6'), [
    '推理链路：raw → L3(raw) → {raw, L3 预测图} → (根深度, 根半径)　｜　第二通道始终使用 L3 实际输出（非理想 target）；前端评 MSE/PSNR/SSIM 与效率，后端评 MAE/RMSE/R²'
], size=11.5)

# ---------------------------------------------------------------- 10 数据集
s = prs.slides[9]
sh = find(s, 'AutoShape 6')
sh.left = Inches(3.30)
sh.top = Inches(1.95)
sh.width = Inches(4.40)
sh.height = Inches(2.68)
set_lines(sh, [
    'gprMax（FDTD）三维薄层模型 1.0×1.0×0.001 m，网格 Δx=Δy=Δz=1 mm，PML 吸收边界',
    '400 MHz Ricker 子波；时窗 16 ns——与实测采集一致，εr≈4 时探深约 1.0 m',
    '源检距 10 mm 按零偏移（zero-offset）近似，双曲线顶点清晰；每 10 mm 一道，97 道 → 416×416',
    '根：εr 15–30，σ 0.001 S/m，埋深 470–870 mm；按“回波完整落入时窗”筛选样本',
    '均质沙土：εr 3.0–6.0，根半径 12–60 mm（2000 组）',
    '异质 Peplinski：50% 沙 + 50% 黏土、含水率 0.1%–25%、5 mm 斑块 × 50 种，εr 3.82–9.99，半径 22–36 mm（1000 组）',
    '标签：target = raw − background（同组差分，仅保留根系散射）',
], size=9.5)
for p in sh.text_frame.paragraphs:
    p.line_spacing = 1.1
    p.space_before = Pt(2.5)

set_lines(find(s, 'AutoShape 8'), ['异质土壤模型（Peplinski）'], size=10, color=GRAY,
          font=SANS)
set_lines(find(s, 'AutoShape 10'), ['标签（raw − background）'], size=10,
          color=GRAY, font=SANS)

tbl = find(s, 'Table 13').table
# 追加一行
tr = copy.deepcopy(tbl._tbl.tr_lst[-1])
tbl._tbl.append(tr)
data = [
    ['类别', '样本量', '关键参数'],
    ['均质土壤', '2000 组', '沙土 εr 3.0–6.0；根半径 12–60 mm'],
    ['异质土壤', '1000 组', 'Peplinski εr 3.82–9.99；根半径 22–36 mm；5 mm 斑块'],
    ['标签与划分', '3000 组', 'target = raw − background；测试集 600 组固定划分'],
]
write_table(tbl, data, size=10.5)
for i in range(4):
    tbl.rows[i].height = Inches(0.42)

# ---------------------------------------------------------------- 11 轻量化
s = prs.slides[10]
set_lines(find(s, 'AutoShape 5'), [
    [('B0', {'bold': True, 'color': GREEN}), (' → ', {}),
     ('L1', {'bold': True, 'color': GREEN}), ('（Ghost）→ ', {}),
     ('L2', {'bold': True, 'color': GREEN}), ('（CFFMLite）→ ', {}),
     ('L3', {'bold': True, 'color': GREEN}), ('（CALite）：逐级替换，每步只替换一个模块', {})],
], size=14)
set_lines(find(s, 'AutoShape 8'), [
    [('L1', {'bold': True, 'color': GREEN}), ('　Ghost 替换 3×3 卷积：PSNR +0.86 dB，参数 −30.7%', {})],
    [('L2', {'bold': True, 'color': GREEN}), ('　CFFMLite 替换 CFFM（4 分支→2 分支）：参数 −9.3%，延迟 −19.9%', {})],
    [('L3', {'bold': True, 'color': GREEN}), ('　CALite 替换 ESAM（坐标注意力）：参数 −44.5%，FLOPs −40.2%', {})],
    [('　', {})],
    '每步只变一个模块 → 相邻变体差值即该组件贡献（与替换顺序相关）',
    [('L3 最终：16.43 M / 149.63 G / 22.42 ms / 44.60 FPS', {'bold': True, 'color': RED})],
], size=11.5)

# ---------------------------------------------------------------- 12 双通道
s = prs.slides[11]
set_lines(find(s, 'AutoShape 6'), [
    '去杂波图增强根系弱反射，但可能削弱细节；原始图含杂波，却保留完整土壤背景与全局纹理，两者互补（文献[5]用理想 label 作第二通道）。',
    [('深度与半径均由网络从回波直接回归，避免显式几何反演所需的波速/介电常数标定。',
      {'color': GREEN})],
], size=12)
set_lines(find(s, 'AutoShape 7'), ['网络结构：ResNet18（双通道输入）'], size=14, bold=True,
          color=DARK)
tbl = find(s, 'Table 11').table
data = [
    ['输入模式', '输入', 'MAE 深度/半径'],
    ['raw-only', '原始 B-scan', '4.23 / 1.48 mm'],
    ['L3-only', 'L3 预测图', '3.46 / 1.14 mm'],
    ['dual early', 'raw + L3 预测图', '3.15 / 0.95 mm'],
]
write_table(tbl, data, size=10.5)
set_lines(find(s, 'AutoShape 12'), [
    '关键：第二通道用 L3 实际预测图（非理想 target）；早期/晚期/门控融合的半径 MAE = 0.95 / 1.06 / 2.58 mm → 取早期融合'
], size=11.5, color=RED)

# ---------------------------------------------------------------- 13 实地
s = prs.slides[12]
set_lines(find(s, 'AutoShape 4'), ['实地数据采集与实体验证'], size=24, bold=True,
          color=DARK)
set_lines(find(s, 'AutoShape 9'), ['图 5　林区环境　　图 6　埋根实验与根系剖面实测'],
          size=11, color=GRAY, font=SANS)
sh = find(s, 'AutoShape 10')
sh.left = Inches(6.53)
sh.top = Inches(4.34)
sh.width = Inches(5.83)
sh.height = Inches(2.90)
set_lines(sh, [
    [('树种：', {'bold': True}),
     ('三倍体毛白杨，株距约 3 m，胸径 15–20 cm，树高 15–20 m，根系粗壮、回波明显', {})],
    [('采集：', {'bold': True}),
     ('GSSI SIR-3000 一体机，400 MHz，时窗 16 ns，每道 512 采样点——频率/时窗/采样点与仿真正演一致', {})],
    [('对齐：', {'bold': True}),
     ('道间距 5 mm 为冗余采集，输入网络前按 10 mm 重采样；一体机收发间距大于仿真的零偏移假设，实测先做时零校正与直流去漂移，训练端以时移/增益扰动增强缩小仿真-实测差异', {})],
    [('用途：', {'bold': True}),
     ('① 实地埋根实验（已知埋深与直径）为定量检验提供真值；② 跨现场对比 MS/SVD/RPCA/U-Net/B0 与本文方法（本文给出定性结果）', {})],
], size=11)
for p in sh.text_frame.paragraphs:
    p.space_before = Pt(5)

# ---------------------------------------------------------------- 14 成果
s = prs.slides[13]
set_lines(find(s, 'AutoShape 2'), ['05　实验结果与预期成果'], size=12, bold=True,
          color=GREEN, font=SANS)
set_lines(find(s, 'AutoShape 4'), ['研究基础与预期成果'], size=24, bold=True,
          color=DARK)
set_lines(find(s, 'AutoShape 5'), ['已完成的三项方法产出'], size=15, bold=True,
          color=GREEN)
sh = find(s, 'AutoShape 6')
sh.top = Inches(2.85)
sh.height = Inches(2.10)
set_lines(sh, [
    [('1. ', {'bold': True, 'color': GREEN}),
     ('渐进式轻量化杂波抑制路线 B0→L1→L2→L3：L3 参数 16.43 M（B0 的 34.9%）、FLOPs 149.63 G（35.2%）、22.42 ms/44.60 FPS，PSNR 42.25 dB、SSIM 0.9972', {})],
    [('2. ', {'bold': True, 'color': GREEN}),
     ('raw + L3 实际预测图双通道早期融合：根深度 MAE 3.15 mm、RMSE 4.45 mm、R² 0.9992；根半径 MAE 0.95 mm、R² 0.9902（4 个随机种子）', {})],
    [('3. ', {'bold': True, 'color': GREEN}),
     ('gprMax 3000 组配对数据集（均质 2000 + 异质 1000）与评测基准：质量 + FLOPs/实测延迟 + 下游精度', {})],
], size=12.5)
for p in sh.text_frame.paragraphs:
    p.space_before = Pt(4)

set_lines(find(s, 'AutoShape 8'), ['预期学术产出与后续计划'], size=15, bold=True,
          color=GREEN)
sh = find(s, 'AutoShape 9')
sh.top = Inches(6.00)
sh.height = Inches(1.25)
set_lines(sh, [
    '· 申请发明专利 1 项　· 登记软件著作权 1–2 项　· 发表学术论文 1–2 篇　· 完成学位论文',
    '· 后续：实测定量验证与跨树种/多场址迁移；轻量化模型边缘端部署；论文修改与投稿',
    [('· 局限：仿真为薄层近似 + 零偏移假设，与一体机实测存在 sim2real 差异（以扰动增强与实测定量验证缓解）；门控融合在 3000 组规模下退化',
      {'color': GRAY})],
], size=12)
for p in sh.text_frame.paragraphs:
    p.space_before = Pt(2)

# ---------------------------------------------------------------- 新增结果页
base = prs.slides[7]

# --- 结果一：前端
r1 = dup_slide(base)
for n in ('AutoShape 5', 'AutoShape 6', 'Connector 7', 'AutoShape 8', 'AutoShape 9'):
    drop(r1, n)
set_lines(find(r1, 'AutoShape 2'), ['05　实验结果与分析'], size=12, bold=True,
          color=GREEN, font=SANS)
set_lines(find(r1, 'AutoShape 4'), ['实验结果一：杂波抑制质量与轻量化效率'], size=24,
          bold=True, color=DARK)

add_box(r1, 0.83, 1.90, 5.60, 0.35, ['杂波抑制质量（600 张测试图，seed=123）'],
        size=13, color=DARK, bold=True)
add_table(r1, [
    ['方法', 'MSE ↓', 'PSNR/dB ↑', 'SSIM ↑'],
    ['f-k', '5.57×10⁻²', '12.54', '0.025'],
    ['小波', '2.75×10⁻²', '15.61', '0.035'],
    ['MS', '1.36×10⁻²', '18.65', '0.733'],
    ['SVD', '1.35×10⁻²', '18.69', '0.735'],
    ['U-Net', '9.64×10⁻⁵', '40.16', '0.9953'],
    ['B0', '7.49×10⁻⁵', '41.25', '0.9963'],
    ['L1', '5.45×10⁻⁵', '42.64', '0.9975'],
    ['L2', '5.79×10⁻⁵', '42.37', '0.9972'],
    ['L3', '5.96×10⁻⁵', '42.25', '0.9972'],
], left=0.83, top=2.30, col_w=[1.40, 1.40, 1.40, 1.40], row_h=0.33, size=10)
add_box(r1, 0.83, 5.74, 5.60, 0.76, [
    'L1 恢复质量最高；5 个随机种子（42/123/2024/2025/777）：L1 42.48±0.53 dB、L3 41.67±0.48 dB，差异在种子波动内，故取更轻的 L3 作后端固定前端。'],
    size=10.5, color=GRAY)

add_box(r1, 6.90, 1.90, 5.60, 0.35, ['轻量化效率（V100，统一 benchmark）'], size=13,
        color=DARK, bold=True)
add_table(r1, [
    ['模型', '参数量/M', 'FLOPs/G', '延迟/ms', 'FPS', '显存/MB'],
    ['B0', '47.07', '425.5', '37.06', '26.98', '758.9'],
    ['L1', '32.64', '299.5', '36.22', '27.61', '653.5'],
    ['L2', '29.60', '250.3', '29.03', '34.45', '630.2'],
    ['L3', '16.43', '149.6', '22.42', '44.60', '528.2'],
], left=6.90, top=2.30, col_w=[0.85, 1.05, 1.00, 0.95, 0.75, 1.00], row_h=0.34,
    size=10)
add_box(r1, 6.90, 4.20, 5.60, 2.20, [
    '· L3 参数量为 B0 的 34.9%，FLOPs 为 35.2%，单图延迟 37.06→22.42 ms，FPS 26.98→44.60',
    '· 轻量化收益不能只看单次 PSNR：需同时报告恢复质量、理论计算量与目标硬件实测速度',
    '· 若只按参数量/FPS 排序，会把恢复质量不合格的 Basic U-Net 误判为最优',
    '· 效率指标仅在 V100 平台内可比',
], size=11)
for p in find(r1, r1.shapes[-1].name).text_frame.paragraphs:
    p.space_before = Pt(6)

# --- 结果二：后端
r2 = dup_slide(base)
for n in ('AutoShape 5', 'AutoShape 6', 'Connector 7', 'AutoShape 8', 'AutoShape 9'):
    drop(r2, n)
set_lines(find(r2, 'AutoShape 2'), ['05　实验结果与分析'], size=12, bold=True,
          color=GREEN, font=SANS)
set_lines(find(r2, 'AutoShape 4'), ['实验结果二：根参数预测与消融分析'], size=24,
          bold=True, color=DARK)

add_box(r2, 0.83, 1.90, 5.70, 0.35, ['根深度与根半径预测（固定 seed=123 的 L3 前端）'],
        size=13, color=DARK, bold=True)
add_table(r2, [
    ['输入/融合', '深度 MAE', '深度 RMSE', '半径 MAE', '半径 RMSE'],
    ['raw-only', '4.23', '6.40', '1.48', '1.90'],
    ['L3-only', '3.46', '4.95', '1.14', '1.52'],
    ['dual early', '3.15', '4.45', '0.95', '1.28'],
    ['dual late', '—', '—', '1.06', '1.41'],
    ['dual gated', '—', '—', '2.58', '3.20'],
], left=0.83, top=2.30, col_w=[1.30, 1.10, 1.20, 1.10, 1.20], row_h=0.34, size=10)
add_box(r2, 0.83, 4.42, 5.70, 1.90, [
    '· 单位 mm；深度为 seed=2024 单次运行，半径为 4 个种子（42/123/2024/3407）均值',
    '· R²：深度 0.9984→0.9991→0.9992；半径 0.9779→0.9861→0.9902',
    '· 相较 raw-only，dual early 的深度 MAE 降低 25.5%、RMSE 降低 30.4%；半径 MAE 降低 35.5%',
    '· 门控融合在两项任务上同时退化（深度 MAE 63.81 mm、R² 0.576），故不采用',
], size=10.5, color=BODY)

add_box(r2, 6.90, 1.90, 5.60, 0.35, ['消融归因与传统回归基线'], size=13, color=DARK,
        bold=True)
add_box(r2, 6.90, 2.30, 5.60, 4.10, [
    [('· Ghost（B0→L1）：', {'bold': True, 'color': GREEN}),
     ('PSNR +0.86 dB，参数 −30.7%，FLOPs −29.6%，是唯一质量与计算量同时获益的环节', {})],
    [('· CFFMLite（L1→L2）：', {'bold': True, 'color': GREEN}),
     ('参数 −9.3%，延迟 −19.9%（主要降低多分支结构的内存访问开销）', {})],
    [('· CALite（L2→L3）：', {'bold': True, 'color': GREEN}),
     ('参数 −44.5%，FLOPs −40.2%，延迟 −22.7%，是压缩规模的主要环节', {})],
    [('· 通道贡献：', {'bold': True, 'color': RED}),
     ('去杂波通道约占总 MAE 改善的 71.4%，原始通道约占 28.6%，同向叠加', {})],
    [('· 传统回归（同一 L3 像素特征）：', {'bold': True, 'color': RED}),
     ('最佳 Extra Trees MAE 29.2 mm，远高于双通道 ResNet18 的 3.15 mm', {})],
], size=11)
for p in find(r2, r2.shapes[-1].name).text_frame.paragraphs:
    p.space_before = Pt(8)

# ---------------------------------------------------------------- 排序：把两张结果页插到“成果”页之前
sldIdLst = prs.slides._sldIdLst
ids = list(sldIdLst)
r1_el, r2_el = ids[-2], ids[-1]
for el in (r1_el, r2_el):
    sldIdLst.remove(el)
sldIdLst.insert(13, r1_el)   # 结果一 → 第 14 页
sldIdLst.insert(14, r2_el)   # 结果二 → 第 15 页

prs.save(DST)
print('saved', DST)
for i, sl in enumerate(prs.slides, 1):
    txt = ' / '.join(sh.text_frame.text.split('\n')[0] for sh in sl.shapes
                     if sh.has_text_frame and sh.text_frame.text.strip())[:90]
    print(i, txt)
