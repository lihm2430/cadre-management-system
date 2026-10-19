# -*- coding: utf-8 -*-
"""
cadre_core.py
科级干部全周期管理与综合研判系统 - 核心算法、数据引擎与功能弹窗库
【PySide6 工业级全功能纯本地增强版】
"""

import os
import sys
import re
import datetime
import sqlite3
import shutil
import zipfile
import base64
import subprocess
import io

from PySide6.QtCore import Qt, QSize, QUrl, QMarginsF, QSizeF
from PySide6.QtGui import (
    QFont, QColor, QDesktopServices, QBrush,
    QPdfWriter, QPageSize, QPageLayout, QTextDocument, QFontDatabase
)
from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QLineEdit, QComboBox, QTextEdit, QScrollArea,
    QFrame, QTableWidget, QTableWidgetItem, QFileDialog, QMessageBox,
    QTabWidget, QRadioButton, QButtonGroup, QAbstractItemView,
    QHeaderView, QCheckBox, QGridLayout, QSizePolicy, QProgressBar,
    QStyledItemDelegate
)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from pypdf import PdfReader
    import docx
    from docx.shared import Pt, Cm, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
    from docx.oxml import parse_xml, OxmlElement
    from docx.oxml.ns import nsdecls, qn
except ImportError as e:
    print(f"缺少基础依赖包: {e}")
    sys.exit(1)

# ==================== 离线 OCR 引擎集成 (rapidocr_onnxruntime) ====================

try:
    from rapidocr_onnxruntime import RapidOCR
    _rapid_ocr_engine = RapidOCR()
except Exception:
    _rapid_ocr_engine = None

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

FILE_DIR = os.path.join(BASE_DIR, "任免文件库")
APPROVAL_DIR = os.path.join(BASE_DIR, "干部任免审批表库")
DISCIPLINE_DIR = os.path.join(BASE_DIR, "纪律处分文件库")
DB_PATH = os.path.join(BASE_DIR, "干部任免数据索引.db")

for folder in [FILE_DIR, APPROVAL_DIR, DISCIPLINE_DIR]:
    if not os.path.exists(folder):
        os.makedirs(folder)

# ==================== 字段与常量定义 ====================

ALL_CADRE_COLUMNS = [
    "id", "emp_no", "name", "gender", "birth", "nation", "native_place",
    "id_card", "phone", "office_phone", "office_address", "politic", "party_date",
    "work_date", "dept_post", "level", "cadre_rank", "chu_category", "charge_division",
    "current_duty_date", "same_level_date", "title", "title_date",
    "full_edu_degree", "full_edu_school", "job_edu_degree", "job_edu_school",
    "first_edu_degree", "first_edu_school", "first_edu_major", "first_edu_date",
    "highest_edu_degree", "highest_edu_degree_title", "highest_edu_school", "highest_edu_major", "highest_edu_date",
    "career_history", "cadre_identity", "post_category",
    "discipline_type", "discipline_date", "discipline_desc",
    "resume_text", "family_relations", "selected_approval_file",
    "status", "leave_reason", "leave_date", "remark"
]

EXPORTABLE_FIELDS = [
    ("emp_no", "员工编号 (工号)"),
    ("name", "干部姓名"),
    ("cadre_rank", "管理层级(科级/处级)"),
    ("chu_category", "处级归口类别"),
    ("gender", "性别"),
    ("birth", "出生年月"),
    ("age", "实际年龄"),
    ("cadre_identity", "人员编制身份(校编/院编/人事代理)"),
    ("post_category", "岗位类别(党政管理/业务管理)"),
    ("nation", "民族"),
    ("native_place", "籍贯"),
    ("id_card", "身份证号"),
    ("phone", "移动电话 (手机)"),
    ("office_phone", "办公电话 (固话/内线)"),
    ("office_address", "办公地址 (科室/房间)"),
    ("politic", "政治面貌"),
    ("party_date", "入党时间"),
    ("work_date", "参加工作时间"),
    ("dept_post", "工作部门及职务"),
    ("charge_division", "领导班子分管业务及部门"),
    ("level", "职务级别"),
    ("current_duty_date", "任现职时间"),
    ("cur_len", "现职务已任时长"),
    ("same_level_date", "任同级时间"),
    ("same_len", "同级别已任时长"),
    ("title", "专业技术职称"),
    ("title_date", "职称聘任时间"),
    ("full_edu_degree", "全日制学历学位"),
    ("full_edu_school", "全日制毕业院校及专业"),
    ("job_edu_degree", "在职学历学位"),
    ("job_edu_school", "在职毕业院校及专业"),
    ("first_edu_degree", "第一学历"),
    ("first_edu_school", "第一学历毕业院校"),
    ("first_edu_major", "第一学历专业"),
    ("first_edu_date", "第一学历毕业时间"),
    ("highest_edu_degree", "最高学历"),
    ("highest_edu_degree_title", "最高学位"),
    ("highest_edu_school", "最高学历毕业院校"),
    ("highest_edu_major", "最高学历专业"),
    ("highest_edu_date", "最高学历毕业时间"),
    ("discipline_type", "处分/诫勉类型"),
    ("discipline_date", "处分决定日期"),
    ("discipline_status", "处分影响期状态"),
    ("same_unit_family", "亲属任职回避排查"),
    ("resume_text", "干部个人完整简历(审批表提取)"),
    ("family_relations", "家庭主要成员及社会关系(审批表提取)"),
    ("career_history", "历任职务履职经历"),
    ("doc_trajectory", "公文抓取历次任职轨迹摘要"),
    ("status", "在岗状态"),
    ("remark", "备注说明")
]

SAME_UNIT_KEYWORDS = [
    "本单位", "本系统", "第一附属医院", "一附院", "附属医院", "本院", "附院",
    "新乡医学院", "河南医药大学", "新医"
]

# 任职回避：关键敏感部门与班子领导职务关键词
SENSITIVE_DEPT_KEYWORDS = [
    "组织", "干部", "人事", "人力资源", "纪委", "纪检", "监察", "财务", "计财",
    "审计", "招标", "采购", "装备", "设备", "医务", "药学", "药剂", "院办", "党办", "办公室"
]

LEADER_DUTY_KEYWORDS = [
    "书记", "院长", "副书记", "副院长", "总会计师", "纪委书记", "工会主席", "常委", "局长", "主任", "处长"
]

DISCIPLINE_IMPACT_MONTHS = {
    "诫勉谈话": 6, "政务警告": 6, "党纪警告": 12, "政务记过": 12,
    "党纪严重警告": 18, "政务记大过": 18, "撤销党内职务": 24,
    "政务降级/撤职": 24, "留党察看": 24
}

CHU_CATEGORIES = ["单位领导班子", "直属/分支机构领导", "专业内设机构领导"]

COMBO_STYLE = """
    QComboBox {
        border: 1px solid #CBD5E1;
        border-radius: 6px;
        padding: 2px 22px 2px 7px;
        background-color: #FFFFFF;
        color: #0F172A;
        font-size: 12px;
        min-height: 28px;
    }
    QComboBox:hover {
        border: 1px solid #3B82F6;
    }
    QComboBox:focus {
        border: 1.5px solid #2563EB;
    }
    QComboBox::drop-down {
        subcontrol-origin: padding;
        subcontrol-position: top right;
        width: 18px;
        border-left: none;
    }
    QComboBox QAbstractItemView {
        background-color: #FFFFFF;
        color: #0F172A;
        border: 1px solid #94A3B8;
        border-radius: 6px;
        padding: 4px;
        selection-background-color: #2563EB;
        selection-color: #FFFFFF;
        outline: none;
    }
    QComboBox QAbstractItemView::item {
        min-height: 28px;
        padding: 4px 10px;
        color: #0F172A;
        background-color: #FFFFFF;
    }
    QComboBox QAbstractItemView::item:hover {
        background-color: #EFF6FF;
        color: #1D4ED8;
    }
    QComboBox QAbstractItemView::item:selected {
        background-color: #2563EB;
        color: #FFFFFF;
        font-weight: bold;
    }
"""

# ==================== 智能表格项与排序支持 ====================

class SmartTableWidgetItem(QTableWidgetItem):
    """
    智能排序表格项，支持自定义排序键（如整型数值、日期元组等），
    同时保留默认文本展示与 Qt.UserRole 存储的干部 ID。
    """
    def __init__(self, text="", sort_key=None):
        super().__init__(str(text))
        self.sort_key = sort_key

    def __lt__(self, other):
        if not isinstance(other, QTableWidgetItem):
            return super().__lt__(other)
        other_key = getattr(other, 'sort_key', None)
        if self.sort_key is not None and other_key is not None:
            try:
                return self.sort_key < other_key
            except TypeError:
                return str(self.sort_key) < str(other_key)
        if self.sort_key is not None and other_key is None:
            return False
        if self.sort_key is None and other_key is not None:
            return True
        return super().__lt__(other)


def parse_family_relation_line(line):
    """
    通用智能亲属关系行解析器，兼容多种常见录入格式：
    1. 管道符键值格式：配偶：某某 | 出生年月:1989.04 | 政治面貌:中共党员 | 单位及职务:本单位...
    2. 逗号/制表符/分号自由列表格式
    返回五元组：(称谓, 姓名, 出生年月, 政治面貌, 工作单位及职务)
    """
    line = str(line or "").strip()
    if not line:
        return None

    if "|" in line:
        segments = [s.strip() for s in line.split("|") if s.strip()]
        rel, name = "亲属", "—"
        birth, pol, unit = "—", "—", "—"
        if segments:
            p0 = segments[0]
            if "：" in p0:
                parts = p0.split("：", 1)
                rel, name = parts[0].strip(), parts[1].strip()
            elif ":" in p0:
                parts = p0.split(":", 1)
                rel, name = parts[0].strip(), parts[1].strip()
            else:
                name = p0
        for seg in segments[1:]:
            if "出生" in seg:
                birth = seg.split(":", 1)[-1].split("：", 1)[-1].strip()
            elif "面貌" in seg or "政治" in seg:
                pol = seg.split(":", 1)[-1].split("：", 1)[-1].strip()
            elif "单位" in seg or "职务" in seg:
                unit = seg.split(":", 1)[-1].split("：", 1)[-1].strip()
        return rel, name, birth, pol, unit

    parts = [p.strip() for p in re.split(r'[,，\t;；]+', line) if p.strip()]
    if len(parts) >= 2:
        rel = parts[0]
        name = parts[1]
        if "：" in rel:
            rp = rel.split("：", 1)
            rel, name = rp[0].strip(), rp[1].strip()
        elif ":" in rel:
            rp = rel.split(":", 1)
            rel, name = rp[0].strip(), rp[1].strip()
        birth = parts[2] if len(parts) > 2 else "—"
        pol = parts[3] if len(parts) > 3 else "—"
        unit = parts[4] if len(parts) > 4 else "—"
        return rel, name, birth, pol, unit

    return None


def get_cadre_photo_bytes(name, selected_file=None):
    """
    从干部任免审批表 Word 文档（.docx）或本地照片目录中提取免冠电子照片原始字节与格式 (bytes, ext)
    """
    target_paths = []
    if selected_file and is_valid_text(selected_file):
        if os.path.isabs(selected_file) and os.path.exists(selected_file):
            target_paths.append(selected_file)
        else:
            p = os.path.join(APPROVAL_DIR, selected_file)
            if os.path.exists(p):
                target_paths.append(p)
            p2 = os.path.join(BASE_DIR, selected_file)
            if os.path.exists(p2):
                target_paths.append(p2)

    if os.path.exists(APPROVAL_DIR):
        for f in os.listdir(APPROVAL_DIR):
            if f.endswith(".docx") and not f.startswith("~$") and name in f:
                p = os.path.join(APPROVAL_DIR, f)
                if p not in target_paths:
                    target_paths.append(p)

    for f in os.listdir(BASE_DIR):
        if f.endswith(".docx") and not f.startswith("~$") and name in f:
            p = os.path.join(BASE_DIR, f)
            if p not in target_paths:
                target_paths.append(p)

    for docx_path in target_paths:
        try:
            with zipfile.ZipFile(docx_path, 'r') as z:
                media_files = [m for m in z.namelist() if m.startswith('word/media/')]
                img_candidates = []
                for mf in media_files:
                    ext = os.path.splitext(mf)[1].lower()
                    if ext in ('.jpg', '.jpeg', '.png', '.bmp', '.webp'):
                        info = z.getinfo(mf)
                        img_candidates.append((mf, info.file_size, ext))
                if img_candidates:
                    img_candidates.sort(key=lambda x: x[1], reverse=True)
                    best_img, _, ext = img_candidates[0]
                    return z.read(best_img), ext
        except Exception as e:
            print(f"提取审批表电子照片异常 {docx_path}: {e}")

    for photo_dir in [os.path.join(BASE_DIR, "照片"), BASE_DIR]:
        if os.path.exists(photo_dir):
            for ext in ('.jpg', '.jpeg', '.png', '.bmp'):
                p = os.path.join(photo_dir, f"{name}{ext}")
                if os.path.exists(p):
                    try:
                        with open(p, "rb") as f:
                            return f.read(), ext
                    except Exception:
                        pass
    return None, None


def get_cadre_photo_base64(name, selected_file=None):
    """
    从干部任免审批表 Word 文档（.docx）中提取电子免冠照，转换为 base64 字符串供 HTML 渲染
    """
    img_data, ext = get_cadre_photo_bytes(name, selected_file)
    if img_data:
        mime = "image/jpeg" if ext in ('.jpg', '.jpeg') else f"image/{ext.lstrip('.')}"
        b64_str = base64.b64encode(img_data).decode('ascii')
        return f"data:{mime};base64,{b64_str}"
    return None


def build_cadre_profile_html(d, base_ym=None):
    """
    构建符合干部档案汇报与谈话规范的单页/标准汇报单 HTML 内容
    严格规范：正文字体【仿宋_GB2312】、三号字（16pt）、标准 A4 排版、自动提取并嵌入审批表照片
    """
    if not base_ym:
        now = datetime.datetime.now()
        base_ym = f"{now.year}年{now.month:02d}月"

    name = d.get("name", "—")
    gender = d.get("gender", "—")
    birth = d.get("birth", "—")
    age_str = ""
    age_val = calc_age(birth, base_ym)
    if age_val:
        age_str = f"（{age_val}岁）"
    nation = d.get("nation", "汉族")
    native_place = d.get("native_place", "—")
    politic = d.get("politic", "—")
    party_date = d.get("party_date", "—")
    work_date = d.get("work_date", "—")
    dept_post = d.get("dept_post", "—")
    cadre_rank = d.get("cadre_rank", "科级")
    level = d.get("level", "—")
    title = d.get("title", "—")
    phone = d.get("phone", "—")
    identity = d.get("cadre_identity", "—")
    post_category = d.get("post_category", "—")
    current_duty_date = d.get("current_duty_date", "—")
    same_level_date = d.get("same_level_date", "—")

    # 照片提取
    photo_b64 = get_cadre_photo_base64(name, d.get("selected_approval_file"))
    if photo_b64:
        photo_html = f"<img src='{photo_b64}' style='width: 105px; height: 140px; object-fit: cover; border: 1px solid #334155;' />"
    else:
        photo_html = "<div style='width: 105px; height: 140px; border: 1px dashed #94A3B8; color: #94A3B8; font-size: 13pt; text-align: center; line-height: 140px; margin: 0 auto;'>（贴照片处）</div>"

    # 全日制与在职学历
    full_edu = f"{d.get('full_edu_degree', '—')}<br><span style='color:#475569;font-size:14pt;'>{d.get('full_edu_school', '—')}</span>"
    job_edu = f"{d.get('job_edu_degree', '—')}<br><span style='color:#475569;font-size:14pt;'>{d.get('job_edu_school', '—')}</span>"

    # 简历列表
    resume_raw = d.get("resume_text", "")
    resume_lines = [line.strip() for line in resume_raw.split("\n") if line.strip()]
    if not resume_lines:
        resume_html = "<div style='color: #94A3B8;'>暂未录入工作履历。</div>"
    else:
        resume_html = "".join([f"<div style='margin-bottom: 6px; line-height: 1.6;'>• {l}</div>" for l in resume_lines])

    # 家庭成员
    family_raw = d.get("family_relations", "")
    fam_rows = []
    if is_valid_text(family_raw):
        for line in family_raw.split("\n"):
            info = parse_family_relation_line(line)
            if info:
                rel, fnm, fbrth, fpol, funit = info
                is_same = "【同单位任职⚠️】" if any(k in funit for k in ["第一附属医院", "一附院", "本院", "本单位", "同院"]) else ""
                fam_rows.append(f"""
                    <tr>
                        <td align='center'>{rel}</td>
                        <td align='center'><b>{fnm}</b></td>
                        <td align='center'>{fbrth}</td>
                        <td align='center'>{fpol}</td>
                        <td>{funit} <span style='color:#DC2626;font-weight:bold;'>{is_same}</span></td>
                    </tr>
                """)
    if not fam_rows:
        family_html = "<tr><td colspan='5' align='center' style='color:#94A3B8;'>暂无家庭主要成员及重要社会关系记录</td></tr>"
    else:
        family_html = "".join(fam_rows)

    # 纪律及考核说明
    disc_type = d.get("discipline_type", "无")
    disc_date = d.get("discipline_date", "—")
    disc_desc = d.get("discipline_desc", "—")
    if disc_type and disc_type != "无":
        disc_html = f"<span style='color:#DC2626;font-weight:bold;'>处分类型：{disc_type}（处分时间：{disc_date}）</span><br>处分事由：{disc_desc}"
    else:
        disc_html = "<span style='color:#16A34A;'>无党纪政务处分记录，政治素质与廉政审核合格。</span>"

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        body {{
            font-family: FangSong_GB2312, 仿宋_GB2312, 仿宋, FangSong, STFangsong, Songti SC, SimSun, serif;
            color: #0F172A;
            margin: 0;
            padding: 0;
            font-size: 16pt;
            line-height: 1.6;
        }}
        .title {{
            text-align: center;
            font-size: 22pt;
            font-weight: bold;
            font-family: SimHei, "Heiti SC", "STHeiti", "Microsoft YaHei", sans-serif;
            color: #1E3A8A;
            margin-bottom: 6px;
            letter-spacing: 2px;
        }}
        .sub-title {{
            text-align: center;
            font-size: 12pt;
            color: #475569;
            margin-bottom: 14px;
        }}
        table.grid {{
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 12px;
        }}
        table.grid th, table.grid td {{
            border: 1px solid #334155;
            padding: 6px 8px;
            font-size: 16pt;
            font-family: FangSong_GB2312, 仿宋_GB2312, 仿宋, FangSong, STFangsong, Songti SC, SimSun, serif;
        }}
        .bg-hdr {{
            background-color: #F1F5F9;
            font-weight: bold;
            text-align: center;
            color: #1E293B;
        }}
        .section-hdr {{
            background-color: #E2E8F0;
            font-weight: bold;
            font-size: 16pt;
            padding: 6px 10px;
            border-left: 5px solid #1E3A8A;
            margin-top: 14px;
            margin-bottom: 6px;
            font-family: SimHei, "Heiti SC", "STHeiti", "Microsoft YaHei", sans-serif;
        }}
        .box {{
            border: 1px solid #334155;
            padding: 10px 12px;
            font-size: 16pt;
            line-height: 1.6;
            background-color: #FFFFFF;
            font-family: FangSong_GB2312, 仿宋_GB2312, 仿宋, FangSong, STFangsong, Songti SC, SimSun, serif;
        }}
    </style>
    </head>
    <body>
        <div class="title">科级干部全息画像汇报单</div>
        <div class="sub-title">核定基准：{base_ym} &nbsp;|&nbsp; 档案编号：CID-{d.get('id', '000')} &nbsp;|&nbsp; 密级：内部工作用表</div>

        <table class="grid">
            <tr>
                <td width="13%" class="bg-hdr">姓&nbsp;&nbsp;名</td>
                <td width="16%" align="center"><b>{name}</b></td>
                <td width="12%" class="bg-hdr">性&nbsp;&nbsp;别</td>
                <td width="12%" align="center">{gender}</td>
                <td width="14%" class="bg-hdr">出生年月</td>
                <td width="17%" align="center">{birth} {age_str}</td>
                <td width="16%" rowspan="4" align="center" style="vertical-align: middle; padding: 4px;">{photo_html}</td>
            </tr>
            <tr>
                <td class="bg-hdr">民&nbsp;&nbsp;族</td>
                <td align="center">{nation}</td>
                <td class="bg-hdr">籍&nbsp;&nbsp;贯</td>
                <td align="center">{native_place}</td>
                <td class="bg-hdr">政治面貌</td>
                <td align="center">{politic}</td>
            </tr>
            <tr>
                <td class="bg-hdr">入党时间</td>
                <td align="center">{party_date}</td>
                <td class="bg-hdr">参加工作</td>
                <td align="center">{work_date}</td>
                <td class="bg-hdr">联系电话</td>
                <td align="center">{phone}</td>
            </tr>
            <tr>
                <td class="bg-hdr">全日制教育</td>
                <td colspan="2">{full_edu}</td>
                <td class="bg-hdr">在职教育</td>
                <td colspan="2">{job_edu}</td>
            </tr>
            <tr>
                <td class="bg-hdr">现任职务</td>
                <td colspan="3"><b>{dept_post}</b></td>
                <td class="bg-hdr">职级/职称</td>
                <td colspan="2">{cadre_rank}{level} / {title}</td>
            </tr>
            <tr>
                <td class="bg-hdr">任现职时间</td>
                <td align="center">{current_duty_date}</td>
                <td class="bg-hdr">任同级时间</td>
                <td align="center">{same_level_date}</td>
                <td class="bg-hdr">编制与岗位</td>
                <td colspan="2" align="center">{identity} / {post_category}</td>
            </tr>
        </table>

        <div class="section-hdr">工作历程与职务履历</div>
        <div class="box">
            {resume_html}
        </div>

        <div class="section-hdr">党风廉政与纪律考评</div>
        <div class="box">
            {disc_html}
        </div>

        <div class="section-hdr">家庭主要成员及重要社会关系（兼任职回避排查）</div>
        <table class="grid">
            <tr class="bg-hdr">
                <th width="12%">称谓</th>
                <th width="14%">姓名</th>
                <th width="16%">出生年月</th>
                <th width="16%">政治面貌</th>
                <th width="42%">工作单位及现任职务</th>
            </tr>
            {family_html}
        </table>

        <div style="margin-top: 14px; font-size: 13pt; color: #475569; text-align: right;">
            生成时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} &nbsp;&nbsp; 审核人签字：________________
        </div>
    </body>
    </html>
    """
    return html


def export_cadre_profile_pdf(d, base_ym, save_path):
    """
    使用 PySide6 原生 QTextDocument 和 QPdfWriter 高清矢量渲染并输出 A4 干部画像汇报单 PDF
    采用国家公文规范：仿宋_GB2312 三号字体（16pt）、标准 A4 页面排版、自动调用嵌入任免审批表电子照片
    """
    html = build_cadre_profile_html(d, base_ym)
    doc = QTextDocument()

    preferred_fonts = ["仿宋_GB2312", "FangSong_GB2312", "仿宋", "FangSong", "STFangsong", "Songti SC", "SimSun", "Microsoft YaHei"]
    for f in preferred_fonts:
        if QFontDatabase.hasFamily(f):
            doc.setDefaultFont(QFont(f, 16))
            break

    doc.setHtml(html)

    writer = QPdfWriter(save_path)
    writer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
    writer.setPageOrientation(QPageLayout.Orientation.Portrait)
    writer.setPageMargins(QMarginsF(15, 15, 15, 15), QPageLayout.Unit.Millimeter)

    # 按照 A4 净绘图区域（Points）计算并锁定文档宽度，保证三号字体（16pt）1:1 矢量输出，字体绝不微缩
    paint_rect = writer.pageLayout().paintRectPoints()
    doc.setPageSize(QSizeF(paint_rect.width(), paint_rect.height()))

    doc.print_(writer)
    return True


def set_docx_run_font(run, font_name="仿宋_GB2312", size_pt=14, bold=False, color_rgb=None):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    if color_rgb:
        run.font.color.rgb = color_rgb
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), font_name)


def set_docx_cell_shading(cell, hex_color="F1F5F9"):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shd)


def set_docx_table_borders(table, color="334155", sz="4"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)


def format_docx_cell(cell, text="", font_name="仿宋_GB2312", font_size=12, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, is_header=False):
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if is_header:
        set_docx_cell_shading(cell, "F1F5F9")
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    if text:
        r = p.add_run(str(text))
        set_docx_run_font(r, font_name=font_name, size_pt=font_size, bold=bold)
    return p


def export_cadre_profile_docx(d, base_ym, save_path):
    """
    使用 python-docx 原生生成符合国家公文标准规范的《干部全息画像汇报单》(Word .docx)：
    - 页面规范：A4 纸张、边距 2.0cm
    - 字体规范：标题方正小标宋/黑体 20pt，正文表格严格采用 仿宋_GB2312（三号/四号协调）
    - 智能集成：自动读取嵌入干部任免审批表免冠照片
    - 二次编辑：用户可在本地 Word / WPS 中自由微调、排版、打印
    """
    doc = docx.Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)

    # 1. 标题
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(4)
    r_t = p_title.add_run("科级干部全息画像汇报单")
    set_docx_run_font(r_t, font_name="黑体", size_pt=20, bold=True, color_rgb=RGBColor(30, 58, 138))

    # 副标题
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(8)
    sub_text = f"核定基准：{base_ym}  |  档案编号：CID-{d.get('id', '000')}  |  密级：内部工作用表"
    r_s = p_sub.add_run(sub_text)
    set_docx_run_font(r_s, font_name="仿宋_GB2312", size_pt=10.5, color_rgb=RGBColor(71, 85, 105))

    # 2. 基本信息与任职学历表 (6行7列)
    tbl = doc.add_table(rows=6, cols=7)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_docx_table_borders(tbl)

    col_widths = [Cm(2.2), Cm(2.6), Cm(2.0), Cm(2.0), Cm(2.4), Cm(2.8), Cm(3.0)]
    for row in tbl.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = w

    # 单元格合并
    tbl.cell(0, 6).merge(tbl.cell(3, 6))  # 照片跨 4 行
    tbl.cell(3, 1).merge(tbl.cell(3, 2))  # 全日制
    tbl.cell(3, 4).merge(tbl.cell(3, 5))  # 在职教育
    tbl.cell(4, 1).merge(tbl.cell(4, 3))  # 现任职务
    tbl.cell(4, 5).merge(tbl.cell(4, 6))  # 职级职称
    tbl.cell(5, 5).merge(tbl.cell(5, 6))  # 编制与岗位

    name = d.get("name", "—")
    gender = d.get("gender", "—")
    birth = d.get("birth", "—")
    age = calc_age(birth, base_ym)
    age_str = f"({age}岁)" if age else ""
    nation = d.get("nation", "汉族")
    native_place = d.get("native_place", "—")
    politic = d.get("politic", "—")
    party_date = d.get("party_date", "—")
    work_date = d.get("work_date", "—")
    phone = d.get("phone", "—")

    # 全日制与在职教育
    fe_deg = d.get('full_edu_degree', '') or d.get('first_edu_degree', '—')
    fe_sch = d.get('full_edu_school', '') or d.get('first_edu_school', '')
    full_edu = f"{fe_deg}  {fe_sch}".strip() or "—"

    je_deg = d.get('job_edu_degree', '') or d.get('highest_edu_degree', '—')
    je_sch = d.get('job_edu_school', '') or d.get('highest_edu_school', '')
    job_edu = f"{je_deg}  {je_sch}".strip() or "—"

    dept_post = d.get("dept_post", "—")
    cadre_rank = d.get("cadre_rank", "—")
    level = f"({d.get('rank_level', '')})" if d.get('rank_level') else ""
    title = d.get("title", "—")
    current_duty_date = d.get("current_duty_date", "—")
    same_level_date = d.get("same_level_date", "—")
    identity = d.get("cadre_identity", "—")
    post_category = d.get("post_category", "—")

    # 填入数据
    # Row 0
    format_docx_cell(tbl.cell(0, 0), "姓  名", is_header=True, bold=True)
    format_docx_cell(tbl.cell(0, 1), name, bold=True, font_size=13)
    format_docx_cell(tbl.cell(0, 2), "性  别", is_header=True)
    format_docx_cell(tbl.cell(0, 3), gender)
    format_docx_cell(tbl.cell(0, 4), "出生年月", is_header=True)
    format_docx_cell(tbl.cell(0, 5), f"{birth} {age_str}".strip())

    # 照片写入
    photo_cell = tbl.cell(0, 6)
    photo_cell.text = ""
    photo_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    img_data, ext = get_cadre_photo_bytes(name, d.get("selected_approval_file"))
    if img_data:
        try:
            p_img = photo_cell.paragraphs[0]
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(2)
            p_img.paragraph_format.space_after = Pt(2)
            r_img = p_img.add_run()
            r_img.add_picture(io.BytesIO(img_data), width=Cm(2.6), height=Cm(3.5))
        except Exception as e:
            print(f"写入照片异常: {e}")
            format_docx_cell(photo_cell, "（贴照片处）", font_size=10)
    else:
        format_docx_cell(photo_cell, "（贴照片处）", font_size=10)

    # Row 1
    format_docx_cell(tbl.cell(1, 0), "民  族", is_header=True)
    format_docx_cell(tbl.cell(1, 1), nation)
    format_docx_cell(tbl.cell(1, 2), "籍  贯", is_header=True)
    format_docx_cell(tbl.cell(1, 3), native_place)
    format_docx_cell(tbl.cell(1, 4), "政治面貌", is_header=True)
    format_docx_cell(tbl.cell(1, 5), politic)

    # Row 2
    format_docx_cell(tbl.cell(2, 0), "入党时间", is_header=True)
    format_docx_cell(tbl.cell(2, 1), party_date)
    format_docx_cell(tbl.cell(2, 2), "参加工作", is_header=True)
    format_docx_cell(tbl.cell(2, 3), work_date)
    format_docx_cell(tbl.cell(2, 4), "联系电话", is_header=True)
    format_docx_cell(tbl.cell(2, 5), phone)

    # Row 3
    format_docx_cell(tbl.cell(3, 0), "全日制教育", is_header=True)
    format_docx_cell(tbl.cell(3, 1), full_edu, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_docx_cell(tbl.cell(3, 3), "在职教育", is_header=True)
    format_docx_cell(tbl.cell(3, 4), job_edu, align=WD_ALIGN_PARAGRAPH.LEFT)

    # Row 4
    format_docx_cell(tbl.cell(4, 0), "现任职务", is_header=True)
    format_docx_cell(tbl.cell(4, 1), dept_post, align=WD_ALIGN_PARAGRAPH.LEFT, bold=True)
    format_docx_cell(tbl.cell(4, 4), "职级/职称", is_header=True)
    format_docx_cell(tbl.cell(4, 5), f"{cadre_rank}{level} / {title}")

    # Row 5
    format_docx_cell(tbl.cell(5, 0), "任现职时间", is_header=True)
    format_docx_cell(tbl.cell(5, 1), current_duty_date)
    format_docx_cell(tbl.cell(5, 2), "任同级时间", is_header=True)
    format_docx_cell(tbl.cell(5, 3), same_level_date)
    format_docx_cell(tbl.cell(5, 4), "编制与岗位", is_header=True)
    format_docx_cell(tbl.cell(5, 5), f"{identity} / {post_category}")

    # 3. 工作历程与职务履历
    p_sec1 = doc.add_paragraph()
    p_sec1.paragraph_format.space_before = Pt(12)
    p_sec1.paragraph_format.space_after = Pt(4)
    r_sec1 = p_sec1.add_run("■ 工作历程与职务履历")
    set_docx_run_font(r_sec1, font_name="黑体", size_pt=14, bold=True, color_rgb=RGBColor(30, 58, 138))

    tbl_res = doc.add_table(rows=1, cols=1)
    tbl_res.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_docx_table_borders(tbl_res, color="334155", sz="4")
    tbl_res.rows[0].cells[0].width = Cm(17.0)
    c_res = tbl_res.rows[0].cells[0]
    c_res.text = ""

    resume_text = d.get("resume_text", "")
    if is_valid_text(resume_text):
        res_lines = [l.strip() for l in str(resume_text).split("\n") if l.strip()]
        for idx, line in enumerate(res_lines):
            p_line = c_res.paragraphs[0] if idx == 0 else c_res.add_paragraph()
            p_line.paragraph_format.line_spacing = 1.25
            p_line.paragraph_format.space_before = Pt(2)
            p_line.paragraph_format.space_after = Pt(2)
            r_l = p_line.add_run(line)
            set_docx_run_font(r_l, font_name="仿宋_GB2312", size_pt=13)
    else:
        p_line = c_res.paragraphs[0]
        r_l = p_line.add_run("暂无履历详细信息")
        set_docx_run_font(r_l, font_name="仿宋_GB2312", size_pt=13, color_rgb=RGBColor(148, 163, 184))

    # 4. 党风廉政与纪律考评
    p_sec2 = doc.add_paragraph()
    p_sec2.paragraph_format.space_before = Pt(12)
    p_sec2.paragraph_format.space_after = Pt(4)
    r_sec2 = p_sec2.add_run("■ 党风廉政与纪律考评")
    set_docx_run_font(r_sec2, font_name="黑体", size_pt=14, bold=True, color_rgb=RGBColor(30, 58, 138))

    tbl_disc = doc.add_table(rows=1, cols=1)
    tbl_disc.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_docx_table_borders(tbl_disc, color="334155", sz="4")
    tbl_disc.rows[0].cells[0].width = Cm(17.0)
    c_disc = tbl_disc.rows[0].cells[0]
    c_disc.text = ""

    disc_type = d.get("discipline_type", "无")
    disc_date = d.get("discipline_date", "—")
    disc_desc = d.get("discipline_desc", "—")
    p_disc = c_disc.paragraphs[0]
    p_disc.paragraph_format.line_spacing = 1.25
    p_disc.paragraph_format.space_before = Pt(3)
    p_disc.paragraph_format.space_after = Pt(3)
    if disc_type and disc_type != "无":
        r_d1 = p_disc.add_run(f"处分类型：{disc_type}（处分时间：{disc_date}）\n处分事由：{disc_desc}")
        set_docx_run_font(r_d1, font_name="仿宋_GB2312", size_pt=13, bold=True, color_rgb=RGBColor(220, 38, 38))
    else:
        r_d1 = p_disc.add_run("经核查，该同志在库无党纪政务处分不良记录，政治素质过硬，廉政审核合格。")
        set_docx_run_font(r_d1, font_name="仿宋_GB2312", size_pt=13, color_rgb=RGBColor(22, 163, 74))

    # 5. 家庭主要成员及重要社会关系
    p_sec3 = doc.add_paragraph()
    p_sec3.paragraph_format.space_before = Pt(12)
    p_sec3.paragraph_format.space_after = Pt(4)
    r_sec3 = p_sec3.add_run("■ 家庭主要成员及重要社会关系（兼任职回避排查）")
    set_docx_run_font(r_sec3, font_name="黑体", size_pt=14, bold=True, color_rgb=RGBColor(30, 58, 138))

    fam_lines = []
    fam_raw = d.get("family_relations", "")
    if is_valid_text(fam_raw):
        fam_lines = [l.strip() for l in str(fam_raw).split("\n") if l.strip() and not l.strip().startswith("—")]

    row_count = max(len(fam_lines), 1) + 1
    tbl_fam = doc.add_table(rows=row_count, cols=5)
    tbl_fam.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_docx_table_borders(tbl_fam, color="334155", sz="4")

    fam_widths = [Cm(2.2), Cm(2.4), Cm(2.6), Cm(2.6), Cm(7.2)]
    for row in tbl_fam.rows:
        for i, w in enumerate(fam_widths):
            row.cells[i].width = w

    format_docx_cell(tbl_fam.cell(0, 0), "称  谓", is_header=True, bold=True)
    format_docx_cell(tbl_fam.cell(0, 1), "姓  名", is_header=True, bold=True)
    format_docx_cell(tbl_fam.cell(0, 2), "出生年月", is_header=True, bold=True)
    format_docx_cell(tbl_fam.cell(0, 3), "政治面貌", is_header=True, bold=True)
    format_docx_cell(tbl_fam.cell(0, 4), "工作单位及现任职务", is_header=True, bold=True)

    if fam_lines:
        for idx, fl in enumerate(fam_lines):
            r_idx = idx + 1
            parsed_fam = parse_family_relation_line(fl)
            if parsed_fam:
                rel, f_nm, f_b, f_p, f_u = parsed_fam
            else:
                rel, f_nm, f_b, f_p, f_u = "亲属", fl, "—", "—", "—"
            format_docx_cell(tbl_fam.cell(r_idx, 0), rel)
            format_docx_cell(tbl_fam.cell(r_idx, 1), f_nm)
            format_docx_cell(tbl_fam.cell(r_idx, 2), f_b)
            format_docx_cell(tbl_fam.cell(r_idx, 3), f_p)
            format_docx_cell(tbl_fam.cell(r_idx, 4), f_u, align=WD_ALIGN_PARAGRAPH.LEFT)
    else:
        format_docx_cell(tbl_fam.cell(1, 0), "—")
        format_docx_cell(tbl_fam.cell(1, 1), "—")
        format_docx_cell(tbl_fam.cell(1, 2), "—")
        format_docx_cell(tbl_fam.cell(1, 3), "—")
        format_docx_cell(tbl_fam.cell(1, 4), "暂无填报记录", align=WD_ALIGN_PARAGRAPH.LEFT)

    # 6. 页脚审核签署
    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_foot.paragraph_format.space_before = Pt(16)
    foot_text = f"生成时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}    审核人签字：________________"
    r_f = p_foot.add_run(foot_text)
    set_docx_run_font(r_f, font_name="仿宋_GB2312", size_pt=10.5, color_rgb=RGBColor(71, 85, 105))

    doc.save(save_path)
    return True

# ==================== 基础与算法工具函数 ====================

def is_valid_text(val):
    if val is None: return False
    s = str(val).strip()
    return s not in ("", "—", "None", "nan", "null", "/", "-", "未填", "无")

def split_school_and_major(raw):
    """拆分院校与专业，支持形如 '北京印刷学院/市场营销' 或 '中共中央党校 法律专业'"""
    raw = str(raw or "").strip()
    if not raw or raw in ("", "—", "None", "nan", "null", "/", "-", "未填", "无"):
        return "", ""
    if "/" in raw:
        parts = raw.split("/", 1)
        return parts[0].strip(), parts[1].strip()
    if " " in raw:
        parts = raw.split(None, 1)
        return parts[0].strip(), parts[1].strip()
    return raw, ""

def split_degree_and_title(raw):
    """拆分学历与学位，支持形如 '本科/学士学位' 或 '研究生/硕士学位'"""
    raw = str(raw or "").strip()
    if not raw or raw in ("", "—", "None", "nan", "null", "/", "-", "未填", "无"):
        return "", ""
    if "/" in raw:
        parts = raw.split("/", 1)
        deg = parts[0].strip()
        tit = parts[1].strip() if parts[1].strip() not in ["无", "—", "-", ""] else ""
        return deg, tit
    return raw, ""

def clean_ym(s):
    """标准化提取 YYYYMM 年月字符串"""
    if not is_valid_text(s): return ""
    m = re.search(r'(\d{4})[^\d]?(\d{1,2})', str(s))
    if m: return f"{m.group(1)}{int(m.group(2)):02d}"
    return ""

def parse_ym(ym_str):
    if not is_valid_text(ym_str): return None
    m = re.search(r'(\d{4})[^\d]?(\d{1,2})', str(ym_str))
    if m:
        try: return int(m.group(1)), int(m.group(2))
        except Exception: return None
    return None

def parse_sort_date(date_str):
    """提取结构化 (年, 月, 日) 元组用于数值严格升序排序"""
    if not is_valid_text(date_str): return (0, 0, 0)
    m = re.search(r'(\d{4})[^\d]?(\d{1,2})[^\d]?(\d{1,2})?', str(date_str))
    if m:
        y = int(m.group(1))
        mon = int(m.group(2))
        d = int(m.group(3)) if m.group(3) else 1
        return (y, mon, d)
    return (0, 0, 0)

def calc_months_diff(start_ym_str, base_ym_str):
    st = parse_ym(start_ym_str)
    bs = parse_ym(base_ym_str)
    if not st or not bs: return None
    return (bs[0] - st[0]) * 12 + (bs[1] - st[1])

def format_months_to_str(months):
    if months is None: return "—"
    if months < 0: return "未满1个月"
    y = months // 12
    m = months % 12
    if y == 0: return f"{m}个月"
    elif m == 0: return f"{y}年"
    else: return f"{y}年{m}个月"

def calc_age(birth_ym_str, base_ym_str):
    b = parse_ym(birth_ym_str)
    bs = parse_ym(base_ym_str)
    if not b or not bs: return None
    age = bs[0] - b[0]
    if bs[1] < b[1]: age -= 1
    return age if age >= 0 else None

def get_title_level(title_str):
    """
    精准判定专业技术职称等级：
    返回: '正高' | '副高' | '中级' | '初级及以下'
    必须按照严格的优先级与排他规则，彻底杜绝包含子串误判（如'副主任医师'包含'主任医师'、'副教授'包含'教授'）。
    """
    if not is_valid_text(title_str) or title_str in ["—", "无", "None"]:
        return "初级及以下"
    t = str(title_str).strip()

    # 0. 高频特殊特例优先截胡：助理研究员为中级（避免被后方正高'研究员'子串匹配）
    if "助理研究员" in t:
        return "中级"

    # 0.5 正高核心标识词优先截胡（消除'正高级会计师'被'高级会计师'包含的误判）
    if any(k in t for k in ["正高级", "正高"]):
        return "正高"

    # 1. 先排查初级特有关键词
    junior_keywords = [
        "助理工程师", "助理会计师", "助理经济师", "助理统计师", "助理馆员", "助理政工师",
        "研究实习员", "见习", "实习", "医士", "药士", "护士", "技士", "初级", "助理员"
    ]
    if any(k in t for k in junior_keywords):
        return "初级及以下"

    # 2. 排查副高级（必须在正高级之前判定，消除副主任医师/副教授/副研究员等被正高子串包含的bug）
    assoc_senior_keywords = [
        "副主任医师", "副主任中医师", "副主任药师", "副主任中药师", "副主任护师", "副主任技师",
        "副教授", "副研究员", "副高级", "副高",
        "高级工程师", "高级经济师", "高级会计师", "高级审计师", "高级统计师",
        "高级政工师", "高级实验师", "副编审", "副研究馆员"
    ]
    if any(k in t for k in assoc_senior_keywords):
        return "副高"

    # 3. 排查正高级
    high_senior_keywords = [
        "主任医师", "主任中医师", "主任药师", "主任中药师", "主任护师", "主任技师",
        "正高级", "正高", "教授", "研究员", "编审", "研究馆员",
        "正高级工程师", "正高级会计师", "正高级经济师"
    ]
    if any(k in t for k in high_senior_keywords):
        return "正高"

    # 4. 排查中级职称
    intermediate_keywords = [
        "主治医师", "主治中医师", "主治", "主管护师", "主管药师", "主管中药师", "主管技师",
        "讲师", "助理研究员", "工程师", "会计师", "经济师", "政工师", "馆员",
        "统计师", "审计师", "实验师", "编辑", "中级"
    ]
    if any(k in t for k in intermediate_keywords):
        return "中级"

    return "初级及以下"

def calc_retirement_info(gender, birth_str, base_ym_str, title_str="", rank_str=""):
    """
    智能研判干部到龄退休情况：
    - 男干部：年满60周岁退休；
    - 女干部：处级干部或具有副高级及以上职称者年满60周岁退休；其他科级及以下女性干部年满55周岁退休；
    - 计算基准年月到退休年月的差额月数 months_left；
    - 若 0 <= months_left <= 24，认定为【临近退休】预警人选（<=12个月为重点紧迫，<=24个月为2年内梯队预警）。
    """
    b = parse_ym(birth_str)
    bs = parse_ym(base_ym_str)
    if not b or not bs:
        return None
    g = str(gender or "").strip()
    is_female = ("女" in g)

    t_lvl = get_title_level(title_str)
    rk = str(rank_str or "")
    if is_female:
        if "处" in rk or t_lvl in ["正高", "副高"]:
            retire_age = 60
        else:
            retire_age = 55
    else:
        retire_age = 60

    retire_year = b[0] + retire_age
    retire_month = b[1]
    retire_ym_str = f"{retire_year}年{retire_month:02d}月"

    base_total = bs[0] * 12 + bs[1]
    retire_total = retire_year * 12 + retire_month
    months_left = retire_total - base_total

    is_soon = (0 <= months_left <= 24)
    is_critical = (0 <= months_left <= 12)
    is_retired = (months_left < 0)

    if is_retired:
        status_str = f"已到退休年龄({retire_age}岁)"
        tag = "已到龄"
    elif is_critical:
        status_str = f"⏳ 距退休仅{months_left}个月({retire_ym_str}到龄)"
        tag = f"距退休{months_left}月"
    elif is_soon:
        status_str = f"⏳ 距退休{months_left}个月({retire_ym_str}到龄)"
        tag = f"距退休{months_left}月"
    else:
        years_left = round(months_left / 12, 1)
        status_str = f"{retire_ym_str}到龄(尚余{years_left}年)"
        tag = f"余{years_left}年"

    return {
        "retire_age": retire_age,
        "retire_ym": retire_ym_str,
        "months_left": months_left,
        "is_soon": is_soon,
        "is_critical": is_critical,
        "is_retired": is_retired,
        "status_str": status_str,
        "tag": tag
    }

def analyze_same_unit_relations(family_relations_text, cadre_dept_post=""):
    """
    智能分析同单位亲属任职，并进行两级智能回避研判：
    🔴 重点回避关注:
       1. 亲属在关键敏感职能部门（组织、人事、纪检、财务、审计、采购、办公室等）；
       2. 亲属担任单位班子领导职务（书记、局长、院长、主任、副职等）；
       3. 亲属与该干部在同一科室/部门（部门关键词重合）；
       4. 干部本人任职于上述关键部门，且亲属在单位内任职。
    🟡 同单位工作备案:
       亲属在本单位工作，但不同科室且不属于上述关键敏感部门。
    """
    if not is_valid_text(family_relations_text):
        return []
    alerts = []
    lines = str(family_relations_text).split("\n")
    cadre_dp = str(cadre_dept_post or "").strip()

    # 提取干部本人的科室关键词 (去掉医院/大学前缀与职务后缀)
    cadre_dept_clean = cadre_dp
    for prefix in ["第一附属医院", "第二附属医院", "附属医院", "河南医药大学", "新乡医学院", "一附院"]:
        cadre_dept_clean = cadre_dept_clean.replace(prefix, "")
    for suffix in ["部长", "副部长", "科长", "副科长", "主任", "副主任", "书记", "副书记", "干事", "员"]:
        if cadre_dept_clean.endswith(suffix):
            cadre_dept_clean = cadre_dept_clean[:-len(suffix)]
    cadre_dept_clean = cadre_dept_clean.strip()

    cadre_in_sensitive = any(sk in cadre_dp for sk in SENSITIVE_DEPT_KEYWORDS)

    for line in lines:
        line_clean = line.strip()
        if not line_clean or line_clean.startswith("—"):
            continue
        unit_part = line_clean.split("单位及职务:")[-1] if "单位及职务:" in line_clean else (line_clean.split("|")[-1] if "|" in line_clean else line_clean)
        unit_part = unit_part.strip()
        matched_kw = next((kw for kw in SAME_UNIT_KEYWORDS if kw in unit_part), None)
        if matched_kw:
            rel_name = line_clean.split("|")[0].strip() if "|" in line_clean else line_clean[:12]
            # 提取亲属具体职务单位摘要
            rel_duty_clean = unit_part
            for kw in SAME_UNIT_KEYWORDS:
                rel_duty_clean = rel_duty_clean.replace(kw, "")
            rel_duty_clean = rel_duty_clean.strip(" ：:|,，。")
            display_duty = rel_duty_clean if rel_duty_clean else matched_kw

            # 研判是否为重点回避
            rel_in_sensitive = any(sk in unit_part for sk in SENSITIVE_DEPT_KEYWORDS)
            rel_is_leader = any(lk in unit_part for lk in LEADER_DUTY_KEYWORDS)
            same_dept = False
            if len(cadre_dept_clean) >= 2 and cadre_dept_clean in unit_part:
                same_dept = True

            if rel_in_sensitive or rel_is_leader or same_dept or cadre_in_sensitive:
                reason = "同科室" if same_dept else ("关键部门" if (rel_in_sensitive or cadre_in_sensitive) else "班子领导")
                alerts.append(f"🔴 重点回避: {rel_name}({display_duty}[{reason}])")
            else:
                alerts.append(f"🟡 同单位备案: {rel_name}({display_duty})")
    return alerts

def get_approval_template_path():
    """
    自动定位干部任免审批表标准 Word 模板文件路径
    """
    candidates = [
        os.path.join(BASE_DIR, "XXX干部任免审批表(模板).docx"),
        os.path.join(APPROVAL_DIR, "XXX干部任免审批表(模板).docx"),
        os.path.join(APPROVAL_DIR, "干部任免审批表（样表及填表说明）.docx"),
        os.path.join(BASE_DIR, "干部任免审批表（样表及填表说明）.docx")
    ]
    for candidate in candidates:
        if os.path.exists(candidate):
            return candidate
    return None

def check_discipline_impact(disc_type, disc_date_str, base_ym_str):
    if not disc_type or disc_type in ["无", "—", "None", ""]: return False, "无处分", None
    st = parse_ym(disc_date_str)
    bs = parse_ym(base_ym_str)
    if not st or not bs: return False, "处分时间未填", None
    impact_months = DISCIPLINE_IMPACT_MONTHS.get(disc_type, 12)
    end_total_m = st[0] * 12 + st[1] + impact_months
    end_year = (end_total_m - 1) // 12
    end_month = (end_total_m - 1) % 12 + 1
    end_ym_str = f"{end_year}年{end_month:02d}月"
    diff_to_end = end_total_m - (bs[0] * 12 + bs[1])
    if diff_to_end > 0: return True, f"⛔ 处分影响期内(至{end_ym_str})", end_ym_str
    else: return False, f"已过影响期(截止于{end_ym_str})", end_ym_str

# ==================== 智能断行与文本排版清洗管道 ====================

def clean_cadre_appointment_prefix(text, cadre_name=""):
    """
    智能剥离职务行中的前置人名、发文字号、公文动词等残留，
    杜绝将职务名称（如党委组织部副部长、党委人才工作办公室副主任 (兼)）误当人名贪婪截断。
    """
    text = str(text or "").strip()
    if not text:
        return ""
    # 彻底清除各种形式的公文字号与原件引用
    text = re.sub(r'[\(（][^()（）]*[字发号][^()（）]*[\)）]', '', text)
    text = re.sub(r'[\(（][^()（）]*\.(?:docx|pdf)[\)）]', '', text, flags=re.I)
    text = re.sub(r'[一-龥]+字〔\d+〕\d+号', '', text)
    text = re.sub(r'[一-龥]+字\[\d+\]\d+号', '', text)
    text = re.sub(r'[\(（](?:任职)?试用期(?:为)?(?:一|1)年[\)）]', '', text)
    text = re.sub(r'[\(（]\s*$', '', text).strip().rstrip("；;，,。.")

    # 剔除组织决定前缀（经研究决定、院党委决定等）
    text = re.sub(r'^(?:经研究决定|院党委研究决定|院党委决定|党委决定|党委会决定|决定|同意|任命|聘任为|聘任)\s*[：:]?\s*', '', text)

    # 精准剥离人员姓名（确保人名不包含组织部、处、科、主任、副主任等职务词，防止误伤职务）
    if cadre_name:
        name_token = r'(?:[一-龥]{2,3}(?<![部科处室院局长长委任官]))'
        pat = rf'^(?:{name_token}[、，,\s]+)*{re.escape(cadre_name)}(?:[、，,\s]+{name_token})*\s*(?:等)?(?:同志)?\s*(?:任|任命为|聘任为|任职为|拟任|聘任|主持工作)\s*[：:]?\s*'
        text = re.sub(pat, '', text)
    else:
        name_token = r'(?:[一-龥]{2,3}(?<![部科处室院局长长委任官]))'
        pat_tongzhi = rf'^(?:{name_token}[、，,\s]+)*{name_token}\s*(?:等)?(?:同志)\s*(?:任|任命为|聘任为|任职为|拟任|聘任|主持工作)\s*[：:]?\s*'
        text = re.sub(pat_tongzhi, '', text)

    text = text.replace("同志", "").replace("《", "").replace("》", "").strip().rstrip("；;，,。.")
    return text

def extract_post_from_doc_content(content, cadre_name=""):
    """
    从任免公文正文中精准提取干部职务，杜绝贪婪吞噬职务名称。
    支持多人任命、支持“姓名在前”及“任命在前”等多种公文行文模式：
    - 模式1：张某某、李某某同志任后勤服务与管理部副部长
    - 模式2：任命王某某同志为党政办公室主任
    """
    if not content:
        return ""
    clauses = re.split(r'[；;\n。]', str(content))
    for c in clauses:
        c = c.strip()
        if not c:
            continue
        if cadre_name and cadre_name not in c:
            continue
        c = re.sub(r'[\(（][^()（）]*[字发号][^()（）]*[\)）]', '', c)
        c = re.sub(r'[\(（][^()（）]*\.(?:docx|pdf)[\)）]', '', c, flags=re.I)
        c = re.sub(r'[一-龥]+字〔\d+〕\d+号', '', c)
        c = re.sub(r'[一-龥]+字\[\d+\]\d+号', '', c)

        # 1. 试用期满考核合格类发文专项提取（如：“某同志任某部副部长职务试用期满，经考核合格，同意正式任职”）
        if any(k in c for k in ['试用期满', '考核合格', '正式任职', '转正', '正式任命']):
            m_spec = re.search(r'任([^，,。；;]+?)(?:职务)?试用期(?:已)?满', c)
            if m_spec and len(m_spec.group(1).strip()) > 1 and m_spec.group(1).strip() != '职':
                return clean_cadre_appointment_prefix(m_spec.group(1).strip(), cadre_name)

        if cadre_name:
            # 模式 1：人名在前（如：张某某同志任... / 张某某为...）
            pat1 = r'(?:^|[^一-龥])(?:[一-龥]{2,4}[、，,\s]+)*' + re.escape(cadre_name) + r'[^任任命聘为]{0,25}?(?:任|任命为|聘任为|任职为|拟任|聘任|为|主持工作)\s*[：:]?\s*([^；;\n。]+)'
            m1 = re.search(pat1, c)
            if m1:
                post = m1.group(1).strip().rstrip("；;，,。.")
                if not any(bad in post for bad in ['试用期满', '考核合格', '正式任职', '同意正式', '按期正式', '考核通过']):
                    post = clean_cadre_appointment_prefix(post, cadre_name)
                    if post:
                        return post

            # 模式 2：动词在前（如：任命/聘任张某某同志为...）
            pat2 = r'(?:任命|聘任|决定任命|同意任命|选任|拟任)\s*(?:[一-龥]{2,4}[、，,\s]+)*' + re.escape(cadre_name) + r'[^为任]{0,15}?(?:为|担任|任)\s*[：:]?\s*([^；;\n。]+)'
            m2 = re.search(pat2, c)
            if m2:
                post = m2.group(1).strip().rstrip("；;，,。.")
                if not any(bad in post for bad in ['试用期满', '考核合格', '正式任职', '同意正式', '按期正式', '考核通过']):
                    post = clean_cadre_appointment_prefix(post, cadre_name)
                    if post:
                        return post
        else:
            m = re.search(r'(?:任|任命为|聘任为|任职为|拟任|聘任|主持工作)\s*[：:]?\s*([^；;\n。]+)', c)
            if m:
                post = m.group(1).strip().rstrip("；;，,。.")
                if not any(bad in post for bad in ['试用期满', '考核合格', '正式任职', '同意正式', '按期正式', '考核通过']):
                    post = clean_cadre_appointment_prefix(post, "")
                    if post:
                        return post
    return ""

def heal_and_reorganize_resume(resume_text, doc_records=None, cadre_name="", career_history="", current_dept_post="", current_duty_date=""):
    """
    智能规范简历与安全时间闭合引擎：
    1. 整合原有 resume_text、花名册历任职务摘要 career_history、任免公文 doc_records 与现任职务；
    2. 破折号、日期格式规范化；剥离公文字号、抄送前缀与干部姓名；
    3. 严禁篡改已具有起止时间的既有求学历史履历；
    4. 对同职务内部衔接节点（如试用期满转正）智能合并，消除碎片；
    5. 严格按时间递增轴排序闭合，最新任职规范设为 'YYYY.MM—至今  职务'；
    6. 精准保留单位前缀与职务字词（杜绝误削'院'字）。
    """
    raw_source_lines = []
    if is_valid_text(resume_text):
        raw_source_lines.extend(str(resume_text).replace("\r\n", "\n").replace("\r", "\n").replace("\v", "\n").split("\n"))

    if is_valid_text(career_history):
        ch_str = str(career_history).replace("\r\n", "\n").replace("\r", "\n").replace("\v", "\n")
        ch_parts = re.split(r'(?:[;；\n]|\s{2,})(?=\d{4}[.\-/年]\d{1,2})', ch_str)
        for part in ch_parts:
            part = part.strip()
            if part:
                raw_source_lines.append(part)

    if doc_records:
        for doc in doc_records:
            clean_c = doc.get("content", "").replace("\n", " ").strip()
            if any(k in clean_c for k in ['任', '聘', '命', '主持']) and (not cadre_name or cadre_name in clean_c):
                post = extract_post_from_doc_content(clean_c, cadre_name)
                if post:
                    doc_d = doc.get("doc_date", "")
                    m_ym = re.search(r'(\d{4})[^0-9]?(\d{1,2})', doc_d)
                    if m_ym:
                        dt_str = f"{m_ym.group(1)}.{int(m_ym.group(2)):02d}"
                        raw_source_lines.append(f"{dt_str}  {post}")

    if is_valid_text(current_dept_post) and is_valid_text(current_duty_date):
        m_cur = re.search(r'(\d{4})[^0-9]?(\d{1,2})', str(current_duty_date))
        if m_cur:
            c_dt = f"{m_cur.group(1)}.{int(m_cur.group(2)):02d}"
            raw_source_lines.append(f"{c_dt}  {clean_cadre_appointment_prefix(current_dept_post, cadre_name)}")

    cleaned_source_lines = []
    for l in raw_source_lines:
        l_str = l.strip()
        if not l_str:
            continue
        if "(其间" in l_str or "（其间" in l_str or l_str.startswith("其间"):
            cleaned_source_lines.append(l_str)
            continue
        # 规范化各类混杂破折号
        l_str = re.sub(r'(\d{4}[.\-/年]\d{1,2}(?:月)?)\s*(?:[\u2010-\u2015\u2212\uff0d\u007e\uff5e一]|--|-|——|至)\s*', r'\1—', l_str)
        l_str = re.sub(r'(—(?:\d{4}[.\-/年]\d{1,2}(?:月)?|至今|今))\s*', r'\1  ', l_str)

        # 剥离前置字号与人名
        cleaned = clean_cadre_appointment_prefix(l_str, cadre_name)
        if cleaned:
            cleaned = re.sub(r'(\d{4}[.\-/年]\d{1,2}(?:月)?)\s*(?:[\u2010-\u2015\u2212\uff0d\u007e\uff5e一]|--|-|——|至)\s*', r'\1—', cleaned)
            cleaned = re.sub(r'(—(?:\d{4}[.\-/年]\d{1,2}(?:月)?|至今|今))\s*', r'\1  ', cleaned)
            cleaned_source_lines.append(cleaned)

    if not cleaned_source_lines:
        if is_valid_text(current_dept_post):
            s_dt = ""
            if is_valid_text(current_duty_date):
                m_cur = re.search(r'(\d{4})[^0-9]?(\d{1,2})', str(current_duty_date))
                if m_cur:
                    s_dt = f"{m_cur.group(1)}.{int(m_cur.group(2)):02d}"
            prefix = f"{s_dt}—至今  " if s_dt else ""
            return f"{prefix}{clean_cadre_appointment_prefix(current_dept_post, cadre_name)}"
        return ""

    entries = []
    p_full = r'^(\d{4}[.\-/年]\d{1,2}(?:月)?)\s*(?:—{1,2}|--|-|至)\s*(\d{4}[.\-/年]\d{1,2}(?:月)?|至今|今)\s*(.*)$'
    p_half = r'^(\d{4}[.\-/年]\d{1,2}(?:月)?)\s*(?:—{1,2}|--|-|至)?\s*(.*)$'

    for l in cleaned_source_lines:
        if "(其间" in l or "（其间" in l or l.startswith("其间"):
            if entries:
                entries[-1]["sub_lines"].append(l)
            continue
        l_norm = re.sub(r'(\d{4}[.\-/年]\d{1,2}(?:月)?)\s*(?:[\u2010-\u2015\u2212\uff0d\u007e\uff5e一]|--|-|——|至)\s*', r'\1—', l)
        mf = re.match(p_full, l_norm)
        if mf:
            s_dt = mf.group(1).replace("年", ".").replace("月", "")
            e_dt = mf.group(2).replace("年", ".").replace("月", "")
            desc = clean_cadre_appointment_prefix(mf.group(3).strip(), cadre_name)
            desc = re.sub(r'^(?:至今|今)\s*', '', desc).strip().rstrip("；;，,。. ")
            entries.append({"start": s_dt, "end": e_dt, "desc": desc, "sub_lines": []})
            continue
        mh = re.match(p_half, l_norm)
        if mh:
            s_dt = mh.group(1).replace("年", ".").replace("月", "")
            desc = mh.group(2).strip()
            desc = re.sub(r'^(?:—{1,2}|--|-|至)\s*', '', desc).strip()
            desc = re.sub(r'^(?:至今|今)\s*', '', desc).strip()
            desc = clean_cadre_appointment_prefix(desc, cadre_name).rstrip("；;，,。. ")
            entries.append({"start": s_dt, "end": None, "desc": desc, "sub_lines": []})
        else:
            if entries:
                entries[-1]["desc"] += " " + clean_cadre_appointment_prefix(l, cadre_name).rstrip("；;，,。. ")

    for e in entries:
        m = re.search(r'(\d{4})[^0-9]?(\d{1,2})', e["start"])
        if m:
            e["start_fmt"] = f"{m.group(1)}.{int(m.group(2)):02d}"
            e["start_val"] = int(m.group(1)) * 100 + int(m.group(2))
        else:
            e["start_fmt"] = e["start"]
            e["start_val"] = 999999

        if e.get("end"):
            m_e = re.search(r'(\d{4})[^0-9]?(\d{1,2})', e["end"])
            if m_e:
                e["end_fmt"] = f"{m_e.group(1)}.{int(m_e.group(2)):02d}"
                e["end_val"] = int(m_e.group(1)) * 100 + int(m_e.group(2))
            else:
                e["end_fmt"] = e["end"]
                e["end_val"] = None
        else:
            e["end_fmt"] = None
            e["end_val"] = None

    entries.sort(key=lambda x: x["start_val"])

    def is_same_post(d1, d2):
        c1 = re.sub(r'[\s\—\-\.\:\(\)（）与和及、，；;。河南医药大学新乡医学院第一附属医院]', '', d1)
        c2 = re.sub(r'[\s\—\-\.\:\(\)（）与和及、，；;。河南医药大学新乡医学院第一附属医院]', '', d2)
        if not c1 or not c2:
            return False
        if c1 == c2:
            return True

        lead_titles = ['副主任', '主任', '副科长', '科长', '副部长', '部长', '副处长', '处长', '副书记', '书记', '副院长', '院长', '副馆长', '馆长']
        t1 = [t for t in lead_titles if t in c1]
        t2 = [t for t in lead_titles if t in c2]

        # 一个是普通科室（无领导头衔），一个是领导职务：绝非同一职务！
        if (t1 and not t2) or (t2 and not t1):
            return False

        # 若两个都有头衔，要求两者的职务词严格匹配（杜绝正副职混淆，如副科长与科长、副部长与部长）
        if t1 and t2:
            if any(f'副{k}' in c1 for k in ['主任', '科长', '部长', '处长', '书记', '院长']) and not any('副' in t for t in t2):
                return False
            if any(f'副{k}' in c2 for k in ['主任', '科长', '部长', '处长', '书记', '院长']) and not any('副' in t for t in t1):
                return False

            common = set(t1).intersection(set(t2))
            if common:
                return (c1 in c2) or (c2 in c1)
            return False

        return (c1 in c2) or (c2 in c1)

    merged = []
    for e in entries:
        if not merged:
            merged.append(e)
            continue
        prev = merged[-1]
        is_same = is_same_post(prev["desc"], e["desc"])
        if is_same:
            if prev["end_val"] and e["start_val"] <= prev["end_val"]:
                prev["sub_lines"].extend(e["sub_lines"])
                continue
            if not prev["end_val"] or abs(prev["end_val"] - e["start_val"]) <= 2 or abs(e["start_val"] - prev["start_val"]) <= 14:
                if len(e["desc"]) > len(prev["desc"]):
                    prev["desc"] = e["desc"]
                if e["end_val"]:
                    prev["end_fmt"] = e["end_fmt"]
                    prev["end_val"] = e["end_val"]
                prev["sub_lines"].extend(e["sub_lines"])
                continue
        merged.append(e)

    output_lines = []
    for idx, e in enumerate(merged):
        desc = e["desc"]

        if idx < len(merged) - 1:
            next_e = merged[idx + 1]
            end_fmt = e["end_fmt"]
            if not end_fmt or end_fmt in ["至今", "今", "—", "--"] or (e["end_val"] and e["end_val"] > next_e["start_val"]):
                end_fmt = next_e["start_fmt"]
            output_lines.append(f"{e['start_fmt']}—{end_fmt}  {desc}")
        else:
            output_lines.append(f"{e['start_fmt']}—至今  {desc}")
        for sub in e["sub_lines"]:
            output_lines.append(f"    {sub}")

    return "\n".join(output_lines)

def format_resume_text(text, cadre_name=""):
    """
    智能规范简历分行与历史数据自愈清洗引擎：
    自动剥离公文字号与多人姓名、智能闭合时间断层、杜绝重复碎片。
    """
    if not is_valid_text(text):
        return ""
    return heal_and_reorganize_resume(text, cadre_name=cadre_name)

def format_family_text(text):
    """
    智能规范家庭成员分行：
    按亲属称谓（妻子、丈夫、父亲、母亲、儿子、女儿等）自动切分换行，一人一行。
    """
    if not is_valid_text(text):
        return ""
        
    s = str(text).replace("\r\n", "\n").replace("\r", "\n").replace("\v", "\n").replace("\u00a0", " ").replace("\u3000", " ")
    rel_pattern = r'(?=(?:妻子|丈夫|配偶|父亲|母亲|儿子|女儿|岳父|岳母|公公|婆婆|哥哥|弟弟|姐姐|妹妹)[：:])'
    raw_segments = re.split(rel_pattern, s)
    
    cleaned_lines = []
    for seg in raw_segments:
        for line in seg.split("\n"):
            line_str = line.strip()
            if line_str:
                cleaned_lines.append(line_str)
                
    return "\n".join(cleaned_lines)

def audit_resume_timeline(resume_text):
    """
    纯算法体检：逐行解析干部简历时间段，检查是否存在时间断层（>2个月空白）或时间倒挂重叠
    """
    clean_resume = format_resume_text(resume_text)
    if not is_valid_text(clean_resume):
        return []

    lines = [l.strip() for l in clean_resume.split("\n") if l.strip()]
    parsed_entries = []

    pattern = r'^\s*(\d{4})[^\d]?(\d{1,2})\s*(?:—{1,2}|--|-|至)\s*(\d{4}[^\d]?\d{1,2}|至今)?\s+(.*)$'
    now = datetime.datetime.now()
    cur_total_m = now.year * 12 + now.month

    for line in lines:
        m = re.match(pattern, line)
        if not m:
            continue
        sy, sm = int(m.group(1)), int(m.group(2))
        start_m = sy * 12 + sm
        end_part = m.group(3)
        desc = m.group(4)

        if not end_part or "至今" in end_part:
            end_m = cur_total_m
            end_label = "至今"
        else:
            ym_end = parse_ym(end_part)
            if ym_end:
                end_m = ym_end[0] * 12 + ym_end[1]
                end_label = f"{ym_end[0]}.{ym_end[1]:02d}"
            else:
                end_m = start_m
                end_label = "—"

        parsed_entries.append({
            "start_m": start_m,
            "end_m": end_m,
            "start_label": f"{sy}.{sm:02d}",
            "end_label": end_label,
            "desc": desc[:25]
        })

    if len(parsed_entries) < 2:
        return []

    alerts = []
    for i in range(1, len(parsed_entries)):
        prev = parsed_entries[i - 1]
        curr = parsed_entries[i]

        if curr["start_m"] < prev["end_m"]:
            overlap_months = prev["end_m"] - curr["start_m"]
            alerts.append(f"时间重叠预警：【{prev['desc']}】(截止{prev['end_label']}) 与 下一任【{curr['desc']}】(起始{curr['start_label']}) 存在 {overlap_months} 个月时间重叠")

        gap_months = curr["start_m"] - prev["end_m"]
        if gap_months > 2:
            gap_str = format_months_to_str(gap_months)
            alerts.append(f"履历空白断层：在【{prev['end_label']}】至【{curr['start_label']}】之间存在 {gap_str} 的履历空白（待核实补充）")

    return alerts

# ==================== 智能缝合引擎：将历任职务摘要与简历交叉融合 ====================

def merge_career_history_into_resume(resume_text, career_history, cadre_name=""):
    """
    智能缝合引擎：
    将从花名册导入的 career_history（历任职务摘要）
    与原有的 resume_text（求学经历、早年经历）按时间正序深度交叉融合，彻底修复历史断层并闭合时间轴！
    """
    return heal_and_reorganize_resume(resume_text, career_history=career_history, cadre_name=cadre_name)

# ==================== 离线公文 OCR 文字提取模块 ====================

def extract_text_from_pdf_with_ocr(fpath):
    full_text = ""
    try:
        reader = PdfReader(fpath)
        for page in reader.pages:
            ptxt = page.extract_text() or ""
            full_text += ptxt + "\n"
    except Exception:
        pass

    clean_chars = "".join(full_text.split())
    if len(clean_chars) < 30 and _rapid_ocr_engine is not None:
        ocr_text = ""
        try:
            import fitz  # PyMuPDF
            doc = fitz.open(fpath)
            for page in doc:
                pix = page.get_pixmap(dpi=150)
                img_bytes = pix.tobytes("png")
                res, _ = _rapid_ocr_engine(img_bytes)
                if res:
                    for item in res:
                        ocr_text += item[1] + "\n"
        except ImportError:
            try:
                reader = PdfReader(fpath)
                for page in reader.pages:
                    for img in page.images:
                        res, _ = _rapid_ocr_engine(img.data)
                        if res:
                            for item in res:
                                ocr_text += item[1] + "\n"
            except Exception:
                pass
        except Exception:
            pass

        if len("".join(ocr_text.split())) > len(clean_chars):
            full_text = ocr_text

    return full_text

# ==================== 简历格式化与幂等性闭合重写 (彻底解决图1重复追加) ====================

def clean_and_format_duty_line(duty_date_str, post_text, cadre_name=""):
    if not is_valid_text(post_text): return ""
    ym = parse_ym(duty_date_str)
    fmt_date = f"{ym[0]}.{ym[1]:02d}" if ym else str(duty_date_str).replace("年", ".").replace("月", "").strip()

    clean_p = clean_cadre_appointment_prefix(post_text, cadre_name)
    return f"{fmt_date}—至今  {clean_p}"

def close_last_resume_line(last_line, new_start_ym_str):
    if not is_valid_text(last_line): return last_line
    ym_new = parse_ym(new_start_ym_str)
    if not ym_new: return last_line
    fmt_new = f"{ym_new[0]}.{ym_new[1]:02d}"

    # 规范将各类“至今”、“今”替换为截止年月
    if any(k in last_line for k in ["至今", "——今", "--今", "—今"]):
        return re.sub(r'(?:—{1,2}|--|-|至)?\s*(?:至今|今)\s*', f'—{fmt_new}  ', last_line).strip()

    m_open = re.match(r'^(\s*\d{4}\.\d{1,2})\s*(?:—{1,2}|--|-|至)\s+([一-龥（\(].*)$', last_line)
    if m_open:
        st_date = m_open.group(1).strip()
        rest = m_open.group(2).strip()
        return f"{st_date}—{fmt_new}  {rest}"

    return last_line

def append_appointment_to_cadre_resume(cadre_id, duty_date_str, new_post_text, cadre_name=""):
    """
    【升级版：终态幂等性排重校验引擎】
    若简历最新经历中已存在该职务或时间一致，原位规范重写，闭合前序时间轴，绝不重复叠加！
    """
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT resume_text FROM cadres WHERE id = ?", (cadre_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return

    cur_resume = row[0] or ""
    mock_doc = [{"doc_date": str(duty_date_str), "content": f"{cadre_name}任{new_post_text}"}]
    updated_resume = heal_and_reorganize_resume(cur_resume, doc_records=mock_doc, cadre_name=cadre_name)

    c.execute("UPDATE cadres SET resume_text = ? WHERE id = ?", (updated_resume, cadre_id))
    conn.commit()
    conn.close()

# ==================== 花名册多级表头容错匹配 ====================

def match_header_index(headers, regex_patterns):
    for pat in regex_patterns:
        for idx, h in enumerate(headers):
            clean_h = str(h or "").replace(" ", "").strip()
            if re.search(pat, clean_h):
                return idx
    return -1

def extract_flexible_roster_headers(headers):
    mapping = {}
    mapping["emp_no"] = match_header_index(headers, [r'^员工编号$', r'^工号$', r'员工号', r'人员编号'])
    mapping["name"] = match_header_index(headers, [r'^姓名$', r'职工姓名', r'干部姓名'])
    mapping["dept"] = match_header_index(headers, [r'^科室$', r'^所在科室$', r'^所属科室$', r'^部门$', r'^工作部门$'])
    mapping["gender"] = match_header_index(headers, [r'^性别$'])
    mapping["birth"] = match_header_index(headers, [r'^出生日期$', r'^出生年月$', r'^出生$'])
    mapping["id_card"] = match_header_index(headers, [r'^身份证号$', r'^身份证号码$', r'证件号码', r'身份证'])
    mapping["phone"] = match_header_index(headers, [r'^联系方式$', r'^移动电话$', r'^手机号?$', r'联系电话'])
    mapping["identity"] = match_header_index(headers, [r'人员分类', r'人员身份', r'编制类型', r'编制类别', r'用工形式', r'用工性质', r'编制', r'分类'])
    mapping["title"] = match_header_index(headers, [r'^专业技术职务$', r'^专业技术职称$', r'^技术职称$', r'^职称$'])

    mapping["fe_deg"] = match_header_index(headers, [r'第一学历(学位)?$', r'初始学历(学位)?$', r'全日制第一学历', r'原始学历', r'第一学历.*层次'])
    mapping["fe_sch"] = match_header_index(headers, [r'第一学历.*(?:学校|院校)', r'初始学历.*(?:学校|院校)', r'第一毕业学校', r'第一毕业院校'])
    mapping["fe_maj"] = match_header_index(headers, [r'第一学历.*(?:专业)', r'初始学历.*(?:专业)'])
    mapping["fe_dt"]  = match_header_index(headers, [r'第一学历.*(?:时间|年月)', r'初始学历.*(?:时间|年月)'])

    mapping["he_deg"] = match_header_index(headers, [r'最高学历$', r'文化程度$', r'学历$', r'最高文化程度'])
    mapping["he_tit"] = match_header_index(headers, [r'最高学位$', r'学位$', r'所获最高学位'])
    mapping["he_sch"] = match_header_index(headers, [r'最高学历.*(?:学校|院校)', r'最高学位.*(?:学校|院校)', r'最高毕业学校', r'最高毕业院校'])
    mapping["he_maj"] = match_header_index(headers, [r'最高学历.*(?:专业)', r'最高学位.*专业'])
    mapping["he_dt"]  = match_header_index(headers, [r'最高学历.*(?:时间|年月)'])
    return mapping

def find_roster_data_by_name(name):
    roster_file = None
    for f in os.listdir(BASE_DIR):
        if "人员" in f and f.endswith(".xlsx") and not f.startswith("~$"):
            roster_file = os.path.join(BASE_DIR, f)
            break
    if not roster_file: return []

    results = []
    try:
        wb = openpyxl.load_workbook(roster_file, data_only=True)
        ws = wb.active
        header_row = 1
        for r in range(1, 6):
            r_vals = [str(cell.value).strip() if cell.value else "" for cell in ws[r]]
            if any("姓名" in v for v in r_vals):
                header_row = r
                break

        headers = [str(cell.value).strip() if cell.value else "" for cell in ws[header_row]]
        m = extract_flexible_roster_headers(headers)
        name_idx = m["name"]
        if name_idx == -1: return []

        def get_v(row, key):
            idx = m.get(key, -1)
            return str(row[idx]).strip() if idx != -1 and idx < len(row) and is_valid_text(row[idx]) else ""

        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            if row[name_idx]:
                cell_name = str(row[name_idx]).replace(" ", "").strip()
                if cell_name == name.strip():
                    results.append({
                        "emp_no": get_v(row, "emp_no"),
                        "name": cell_name,
                        "dept": get_v(row, "dept"),
                        "gender": get_v(row, "gender") or "—",
                        "birth": get_v(row, "birth") or "—",
                        "id_card": get_v(row, "id_card") or "—",
                        "phone": get_v(row, "phone") or "—",
                        "identity": get_v(row, "identity"),
                        "title": get_v(row, "title") or "—",
                        "first_edu_degree": get_v(row, "fe_deg"),
                        "first_edu_school": get_v(row, "fe_sch"),
                        "first_edu_major": get_v(row, "fe_maj"),
                        "first_edu_date": get_v(row, "fe_dt"),
                        "highest_edu_degree": get_v(row, "he_deg"),
                        "highest_edu_degree_title": get_v(row, "he_tit"),
                        "highest_edu_school": get_v(row, "he_sch"),
                        "highest_edu_major": get_v(row, "he_maj"),
                        "highest_edu_date": get_v(row, "he_dt")
                    })
    except Exception as e:
        print(f"花名册解析辅助异常: {e}")
    return results

def extract_resume_and_family_from_docx(fpath):
    resume_text = ""
    family_list = []
    try:
        doc = docx.Document(fpath)
        for table in doc.tables:
            for r_idx, row in enumerate(table.rows):
                row_texts = []
                for cell in row.cells:
                    t = cell.text.strip()
                    if t and (not row_texts or t != row_texts[-1]):
                        row_texts.append(t)
                if not row_texts: continue

                first_cell_clean = row_texts[0].replace(" ", "").replace("\n", "").replace("\r", "")
                if first_cell_clean == "简历" and len(row_texts) > 1 and not resume_text:
                    resume_text = row_texts[1].strip()

                combined_line = "".join(row_texts).replace(" ", "")
                if "称谓" in combined_line and "姓名" in combined_line: continue

                if any(rel in row_texts[0] or (len(row_texts) > 1 and rel in row_texts[1]) for rel in
                       ["父亲", "母亲", "妻子", "丈夫", "配偶", "儿子", "女儿", "岳父", "岳母", "公公", "婆婆", "哥哥", "弟弟", "姐姐", "妹妹"]):
                    cleaned_items = [t.replace("\n", " ").strip() for t in row_texts if t.replace(" ", "").replace("\n", "") not in [
                        "家庭主要成员及重要社会关系", "称谓", "姓名", "出生年月", "出生", "年月", "政治面貌", "政治", "面貌", "工作单位及职务"
                    ]]
                    if len(cleaned_items) >= 4:
                        rel = cleaned_items[0]
                        nm = cleaned_items[1]
                        brth = cleaned_items[2]
                        pol = cleaned_items[3]
                        duty = " ".join(cleaned_items[4:]) if len(cleaned_items) > 4 else "—"
                        fam_str = f"{rel}：{nm} | 出生年月:{brth} | 政治面貌:{pol} | 单位及职务:{duty}"
                        if fam_str not in family_list:
                            family_list.append(fam_str)

        if not resume_text:
            p_texts = [p.text.strip() for p in doc.paragraphs if len(p.text.strip()) > 5]
            for idx, pt in enumerate(p_texts):
                if pt.replace(" ", "").replace("\n", "") == "简历" and idx + 1 < len(p_texts):
                    resume_text = p_texts[idx + 1]
                    break
    except Exception as e:
        print(f"解析审批表异常 {fpath}: {e}")

    fam_text = "\n".join(family_list) if family_list else "—"
    return (format_resume_text(resume_text) if resume_text else "—"), format_family_text(fam_text)

# ==================== Word 审批表回填算法 (100% 完整保留无删减) ====================

def fill_approval_docx(cadre_dict, template_path, output_path):
    doc = docx.Document(template_path)
    if not doc.tables:
        raise ValueError("Word模板中未找到表格！")

    def set_cell_text(cell, text, font_size=10, bold=False, align=docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER):
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = align
        run = p.add_run(str(text if is_valid_text(text) else "—"))
        run.font.name = "仿宋_GB2312"
        run.font.size = docx.shared.Pt(font_size)
        run.bold = bold

    for table in doc.tables:
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                t_clean = cell.text.replace(" ", "").replace("\n", "").replace("\r", "")

                if t_clean == "姓名" and c_idx + 1 < len(row.cells):
                    set_cell_text(row.cells[c_idx + 1], cadre_dict.get("name", ""), font_size=11, bold=True)
                elif t_clean == "性别" and c_idx + 1 < len(row.cells):
                    set_cell_text(row.cells[c_idx + 1], cadre_dict.get("gender", ""))
                elif t_clean in ["出生年月", "出生年月(岁)"] and c_idx + 1 < len(row.cells):
                    age_str = f" ({cadre_dict.get('age', '')})" if cadre_dict.get('age') else ""
                    set_cell_text(row.cells[c_idx + 1], f"{cadre_dict.get('birth', '')}{age_str}")
                elif t_clean == "民族" and c_idx + 1 < len(row.cells):
                    set_cell_text(row.cells[c_idx + 1], cadre_dict.get("nation", "汉族"))
                elif t_clean == "籍贯" and c_idx + 1 < len(row.cells):
                    set_cell_text(row.cells[c_idx + 1], cadre_dict.get("native_place", ""))
                elif t_clean == "入党时间" and c_idx + 1 < len(row.cells):
                    set_cell_text(row.cells[c_idx + 1], cadre_dict.get("party_date", ""))
                elif t_clean == "参加工作时间" and c_idx + 1 < len(row.cells):
                    set_cell_text(row.cells[c_idx + 1], cadre_dict.get("work_date", ""))
                elif t_clean == "专业技术职务" and c_idx + 1 < len(row.cells):
                    set_cell_text(row.cells[c_idx + 1], cadre_dict.get("title", ""))

                elif "全日制教育" in t_clean and c_idx + 1 < len(row.cells):
                    set_cell_text(row.cells[c_idx + 1], cadre_dict.get("full_edu_degree", "—"))
                    if c_idx + 3 < len(row.cells):
                        set_cell_text(row.cells[c_idx + 3], cadre_dict.get("full_edu_school", "—"))
                elif "在职教育" in t_clean and c_idx + 1 < len(row.cells):
                    set_cell_text(row.cells[c_idx + 1], cadre_dict.get("job_edu_degree", "—"))
                    if c_idx + 3 < len(row.cells):
                        set_cell_text(row.cells[c_idx + 3], cadre_dict.get("job_edu_school", "—"))

                elif t_clean == "现任职务" and c_idx + 1 < len(row.cells):
                    set_cell_text(row.cells[c_idx + 1], cadre_dict.get("current_post_full", cadre_dict.get("dept_post", "")), align=docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT)
                elif t_clean == "拟任职务" and c_idx + 1 < len(row.cells):
                    set_cell_text(row.cells[c_idx + 1], cadre_dict.get("target_post", ""), align=docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT)
                elif t_clean == "拟免职务" and c_idx + 1 < len(row.cells):
                    set_cell_text(row.cells[c_idx + 1], cadre_dict.get("target_dismiss", ""), align=docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT)

                elif t_clean == "简历" and c_idx + 1 < len(row.cells):
                    res_cell = row.cells[c_idx + 1]
                    res_cell.text = ""
                    clean_res = format_resume_text(cadre_dict.get("resume_text", ""))
                    res_lines = clean_res.split("\n")
                    for l_idx, line in enumerate(res_lines):
                        p = res_cell.paragraphs[0] if l_idx == 0 else res_cell.add_paragraph()
                        p.paragraph_format.line_spacing = 1.15
                        p.paragraph_format.space_after = docx.shared.Pt(2)
                        run = p.add_run(line.strip())
                        run.font.name = "仿宋_GB2312"
                        run.font.size = docx.shared.Pt(9.5)

    clean_fam = format_family_text(cadre_dict.get("family_relations", ""))
    if is_valid_text(clean_fam):
        fam_lines = [l.strip() for l in clean_fam.split("\n") if l.strip() and not l.strip().startswith("—")]
        for table in doc.tables:
            header_row = -1
            for r_idx, row in enumerate(table.rows):
                txt_row = "".join([c.text.replace(" ", "") for c in row.cells])
                if "称谓" in txt_row and "姓名" in txt_row:
                    header_row = r_idx
                    break

            if header_row != -1:
                cur_fam_idx = 0
                for r_idx in range(header_row + 1, len(table.rows)):
                    row = table.rows[r_idx]
                    first_cell = row.cells[0].text.replace(" ", "")
                    if any(k in first_cell for k in ["呈报单位", "审查意见", "审批意见", "行政机关"]): break

                    if cur_fam_idx < len(fam_lines):
                        line_item = fam_lines[cur_fam_idx]
                        parts = line_item.split("|")
                        rel = parts[0].split("：")[0].strip() if "：" in parts[0] else parts[0].strip()
                        nm = parts[0].split("：")[1].strip() if "：" in parts[0] else ""
                        brth = parts[1].replace("出生年月:", "").strip() if len(parts) > 1 else ""
                        pol = parts[2].replace("政治面貌:", "").strip() if len(parts) > 2 else ""
                        duty = parts[3].replace("单位及职务:", "").strip() if len(parts) > 3 else ""

                        vals = [rel, nm, brth, pol, duty]
                        for c_i, val in enumerate(vals):
                            if c_i + 1 < len(row.cells):
                                set_cell_text(row.cells[c_i + 1], val, font_size=9, align=docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER if c_i < 4 else docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT)
                        cur_fam_idx += 1
                    else:
                        for c_i in range(1, min(6, len(row.cells))):
                            set_cell_text(row.cells[c_i], "", font_size=9)

    doc.save(output_path)
    return output_path

# ==================== 纵向发光任职轨迹时间轴组件 ====================

class TimelineNodeCard(QFrame):
    def __init__(self, item_data, is_last=False, parent=None):
        super().__init__(parent)
        self.item_data = item_data
        self.is_last = is_last

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        indicator_w = QWidget()
        indicator_w.setFixedWidth(28)
        ind_l = QVBoxLayout(indicator_w)
        ind_l.setContentsMargins(0, 4, 0, 0)
        ind_l.setSpacing(0)
        ind_l.setAlignment(Qt.AlignHCenter)

        dot = QFrame()
        dot.setFixedSize(14, 14)
        dot.setStyleSheet("""
            QFrame {
                background: qradialgradient(cx:0.5, cy:0.5, radius:0.5, fx:0.5, fy:0.5, stop:0 #60A5FA, stop:1 #2563EB);
                border: 2px solid #93C5FD;
                border-radius: 7px;
            }
        """)
        ind_l.addWidget(dot)

        if not is_last:
            line = QFrame()
            line.setFixedWidth(2)
            line.setStyleSheet("background-color: #CBD5E1;")
            ind_l.addWidget(line, 1, Qt.AlignHCenter)

        layout.addWidget(indicator_w)

        card = QFrame()
        card.setCursor(Qt.PointingHandCursor)
        card.setStyleSheet("""
            QFrame {
                background-color: #FFFFFF;
                border: 1px solid #D1E1F0;
                border-left: 3px solid #3B82F6;
                border-radius: 8px;
            }
            QFrame:hover {
                background-color: #F8FAFC;
                border-color: #3B82F6;
            }
        """)
        card_l = QVBoxLayout(card)
        card_l.setContentsMargins(14, 10, 14, 10)
        card_l.setSpacing(4)

        h_row = QHBoxLayout()
        date_badge = QLabel(f" 📅  {item_data['doc_date']} ")
        date_badge.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
        date_badge.setStyleSheet("background-color: #EFF6FF; color: #1D4ED8; border-radius: 4px; padding: 2px;")

        num_lbl = QLabel(f"🔖 {item_data['doc_num']}")
        num_lbl.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
        num_lbl.setStyleSheet("color: #475569;")

        h_row.addWidget(date_badge)
        h_row.addWidget(num_lbl)

        # 智能任免公文类型识别（综合正文、公文全文上下文及“试用期一年”等后起条款研判）
        cnt_text = (
            str(item_data.get('content', '')) + " " +
            str(item_data.get('file_name', '')) + " " +
            str(item_data.get('probation_note', '')) + " " +
            str(item_data.get('doc_full_context', ''))
        )
        if any(k in cnt_text for k in ['试用期满', '考核合格', '正式任职', '转正', '正式任命']):
            type_badge = QLabel(" ✅ 试用期满正式任职 ")
            type_badge.setFont(QFont("Microsoft YaHei UI", 8, QFont.Bold))
            type_badge.setStyleSheet("background-color: #ECFDF5; color: #047857; border: 1px solid #A7F3D0; border-radius: 4px; padding: 2px 4px;")
            h_row.addWidget(type_badge)
        elif any(k in cnt_text for k in ['试用期', '试用']):
            type_badge = QLabel(" 🚩 提拔任用(试用期一年) ")
            type_badge.setFont(QFont("Microsoft YaHei UI", 8, QFont.Bold))
            type_badge.setStyleSheet("background-color: #FEF3C7; color: #B45309; border: 1px solid #FDE68A; border-radius: 4px; padding: 2px 4px;")
            h_row.addWidget(type_badge)
        elif any(k in cnt_text for k in ['免去', '不再担任', '免职']):
            type_badge = QLabel(" 🔄 职务免除/调整 ")
            type_badge.setFont(QFont("Microsoft YaHei UI", 8, QFont.Bold))
            type_badge.setStyleSheet("background-color: #F1F5F9; color: #475569; border: 1px solid #CBD5E1; border-radius: 4px; padding: 2px 4px;")
            h_row.addWidget(type_badge)
        else:
            type_badge = QLabel(" 📌 干部职务任免发文 ")
            type_badge.setFont(QFont("Microsoft YaHei UI", 8, QFont.Bold))
            type_badge.setStyleSheet("background-color: #EFF6FF; color: #1D4ED8; border: 1px solid #BFDBFE; border-radius: 4px; padding: 2px 4px;")
            h_row.addWidget(type_badge)

        h_row.addStretch()

        btn_open = QPushButton("📄 打开红头原件")
        btn_open.setFont(QFont("Microsoft YaHei UI", 8, QFont.Bold))
        btn_open.setFixedHeight(26)
        btn_open.setStyleSheet("""
            QPushButton {
                background-color: #F1F5F9; color: #1E40AF; border: 1px solid #BFDBFE; border-radius: 4px; padding: 0 8px;
            }
            QPushButton:hover { background-color: #DBEAFE; }
        """)
        btn_open.clicked.connect(self.open_doc_file)
        h_row.addWidget(btn_open)

        btn_exclude = QPushButton("🚫 剔除此公文")
        btn_exclude.setFont(QFont("Microsoft YaHei UI", 8, QFont.Bold))
        btn_exclude.setFixedHeight(26)
        btn_exclude.setStyleSheet("""
            QPushButton {
                background-color: #FEF2F2; color: #DC2626; border: 1px solid #FECACA; border-radius: 4px; padding: 0 8px;
            }
            QPushButton:hover { background-color: #FEE2E2; }
        """)
        btn_exclude.clicked.connect(self.exclude_doc_file)
        h_row.addWidget(btn_exclude)

        card_l.addLayout(h_row)

        display_content = str(item_data.get('content', ''))
        prob_note = str(item_data.get('probation_note', '')).strip()
        if prob_note and prob_note not in display_content:
            display_content += f"\n📌 【公文说明】：{prob_note}"

        content_lbl = QLabel(display_content)
        content_lbl.setFont(QFont("Microsoft YaHei UI", 9))
        content_lbl.setStyleSheet("color: #1E293B; line-height: 1.4;")
        content_lbl.setWordWrap(True)
        card_l.addWidget(content_lbl)

        file_lbl = QLabel(f"📎 归档文号原件：{item_data['file_name']}")
        file_lbl.setFont(QFont("Microsoft YaHei UI", 8))
        file_lbl.setStyleSheet("color: #94A3B8;")
        card_l.addWidget(file_lbl)

        layout.addWidget(card, 1)

    def open_doc_file(self):
        fpath = os.path.join(FILE_DIR, self.item_data['file_name'])
        if os.path.exists(fpath):
            QDesktopServices.openUrl(QUrl.fromLocalFile(fpath))
        else:
            QMessageBox.warning(self, "文件不存在", f"未在任免公文库中找到该原件：\n{self.item_data['file_name']}")

    def exclude_doc_file(self):
        fname = self.item_data.get('file_name', '')
        cname = self.item_data.get('cadre_name', '')
        cadre_id = self.item_data.get('cadre_id')
        if not cadre_id and hasattr(self, 'parent_window') and hasattr(self.parent_window, 'cadre_id'):
            cadre_id = self.parent_window.cadre_id

        reply = QMessageBox.question(
            self, "确认剔除公文",
            f"确定要将以下公文从【{cname or '该同志'}】的任免时间轴中剔除吗？\n\n"
            f"• 文件名：{fname}\n"
            f"• 公文字号：{self.item_data.get('doc_num', '—')}\n\n"
            f"提示：剔除后该公文将不再出现在该干部的时间轴中。\n"
            f"您后续随时可在全息卡右上角点击【📋 管理已剔除公文】进行一键恢复。",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes and cadre_id:
            conn = sqlite3.connect(DB_PATH)
            c = conn.cursor()
            all_ids = self.item_data.get('all_rec_ids', [self.item_data.get('rec_id')])
            for r_id in all_ids:
                if r_id:
                    c.execute("INSERT OR IGNORE INTO excluded_records (cadre_id, record_id) VALUES (?, ?)", (cadre_id, r_id))
            conn.commit()
            conn.close()
            QMessageBox.information(self, "已剔除", f"公文【{fname}】已从该干部任免轨迹中剔除！")
            if hasattr(self, 'parent_window') and hasattr(self.parent_window, 'load_timeline_data'):
                self.parent_window.load_timeline_data()


# ==================== 管理已剔除公文对话框 ====================

class ManageExcludedRecordsDialog(QDialog):
    """
    管理已从时间轴中剔除的任免公文对话框
    """
    def __init__(self, parent, cadre_id, name, on_changed=None):
        super().__init__(parent)
        self.cadre_id = cadre_id
        self.name = name
        self.on_changed = on_changed
        self.setWindowTitle(f"📋 管理已剔除的任免公文 —— 【{name}】")
        self.resize(800, 420)

        self.init_ui()
        self.load_data()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(14, 14, 14, 14)

        top_box = QFrame()
        top_box.setStyleSheet("background-color: #FEF2F2; border: 1px solid #FECACA; border-radius: 6px; padding: 6px;")
        top_l = QVBoxLayout(top_box)
        t_lbl = QLabel(f"<b>干部【{self.name}】已剔除任免公文清单：</b>")
        t_lbl.setFont(QFont("Microsoft YaHei UI", 10))
        t_lbl.setStyleSheet("color: #991B1B;")
        t_desc = QLabel("以下公文此前已被您手动剔除。点击右侧【♻️ 恢复此公文】可随时重新生效恢复至该干部时间轴中。")
        t_desc.setFont(QFont("Microsoft YaHei UI", 9))
        t_desc.setStyleSheet("color: #B91C1C;")
        top_l.addWidget(t_lbl)
        top_l.addWidget(t_desc)
        layout.addWidget(top_box)

        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["序号", "公文文件名", "公文字号", "公文日期", "操作"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.table, 1)

        bot_row = QHBoxLayout()
        btn_restore_all = QPushButton("♻️ 全部恢复")
        btn_restore_all.setStyleSheet("background-color: #059669; color: white; border-radius: 4px; padding: 6px 14px; font-weight: bold;")
        btn_restore_all.clicked.connect(self.restore_all)

        btn_close = QPushButton("关闭")
        btn_close.setStyleSheet("background-color: #F1F5F9; color: #475569; border: 1px solid #CBD5E1; border-radius: 4px; padding: 6px 14px;")
        btn_close.clicked.connect(self.accept)

        bot_row.addWidget(btn_restore_all)
        bot_row.addStretch()
        bot_row.addWidget(btn_close)
        layout.addLayout(bot_row)

    def load_data(self):
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("""SELECT DISTINCT cr.file_name, cr.doc_number, cr.doc_date
                     FROM cadre_records cr
                     JOIN excluded_records er ON cr.id = er.record_id
                     WHERE er.cadre_id = ?
                     ORDER BY cr.doc_date ASC""", (self.cadre_id,))
        rows = c.fetchall()
        conn.close()

        self.table.setRowCount(len(rows))
        for idx, (fname, doc_num, doc_date) in enumerate(rows):
            it_idx = QTableWidgetItem(str(idx + 1))
            it_idx.setTextAlignment(Qt.AlignCenter)
            it_file = QTableWidgetItem(fname or "—")
            it_num = QTableWidgetItem(doc_num or "—")
            it_num.setTextAlignment(Qt.AlignCenter)
            it_date = QTableWidgetItem(doc_date or "—")
            it_date.setTextAlignment(Qt.AlignCenter)

            self.table.setItem(idx, 0, it_idx)
            self.table.setItem(idx, 1, it_file)
            self.table.setItem(idx, 2, it_num)
            self.table.setItem(idx, 3, it_date)

            btn_res = QPushButton("♻️ 恢复")
            btn_res.setFixedHeight(24)
            btn_res.setStyleSheet("background-color: #ECFDF5; color: #047857; border: 1px solid #A7F3D0; border-radius: 3px; font-weight: bold;")
            btn_res.clicked.connect(lambda _, fn=fname: self.restore_file(fn))
            self.table.setCellWidget(idx, 4, btn_res)

    def restore_file(self, fname):
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("""DELETE FROM excluded_records 
                     WHERE cadre_id = ? AND record_id IN (SELECT id FROM cadre_records WHERE file_name = ?)""",
                  (self.cadre_id, fname))
        conn.commit()
        conn.close()
        QMessageBox.information(self, "恢复成功", f"公文【{fname}】已恢复生效！")
        self.load_data()
        if self.on_changed:
            self.on_changed()

    def restore_all(self):
        reply = QMessageBox.question(self, "确认全部恢复", f"确定要恢复干部【{self.name}】的所有已剔除公文吗？", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            conn = sqlite3.connect(DB_PATH)
            c = conn.cursor()
            c.execute("DELETE FROM excluded_records WHERE cadre_id = ?", (self.cadre_id,))
            conn.commit()
            conn.close()
            QMessageBox.information(self, "恢复成功", "所有已剔除公文已恢复生效！")
            self.load_data()
            if self.on_changed:
                self.on_changed()


# ==================== 手动匹配干部任免审批表对话框 ====================

class ManualMatchApprovalDialog(QDialog):
    """
    手动定向匹配干部任免审批表 Word 文档对话框
    """
    def __init__(self, parent, cadre_id, name, on_matched=None):
        super().__init__(parent)
        self.cadre_id = cadre_id
        self.name = name
        self.on_matched = on_matched
        self.setWindowTitle(f"📂 手动匹配干部任免审批表 —— 【{name}】")
        self.resize(850, 480)

        self.init_ui()
        self.load_candidates()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(14, 14, 14, 14)

        top_box = QFrame()
        top_box.setStyleSheet("background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; padding: 6px;")
        top_l = QVBoxLayout(top_box)
        t_lbl = QLabel(f"<b>为干部【{self.name}】选择对应的《干部任免审批表》(Word .docx) 原件：</b>")
        t_lbl.setFont(QFont("Microsoft YaHei UI", 10))
        t_desc = QLabel("系统已在审批表库中检索匹配候选文件。选中后将自动提取最新履历、家庭成员与电子免冠照，并写入全息档案。")
        t_desc.setFont(QFont("Microsoft YaHei UI", 9))
        t_desc.setStyleSheet("color: #64748B;")
        top_l.addWidget(t_lbl)
        top_l.addWidget(t_desc)
        layout.addWidget(top_box)

        # 搜索与自定义浏览
        filter_row = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 快速筛选文件名...")
        self.search_input.textChanged.connect(self.filter_table)
        filter_row.addWidget(self.search_input, 1)

        btn_browse = QPushButton("📁 浏览其他Word文件...")
        btn_browse.setStyleSheet("background-color: #F1F5F9; color: #1E293B; border: 1px solid #CBD5E1; border-radius: 4px; padding: 4px 10px;")
        btn_browse.clicked.connect(self.browse_custom_file)
        filter_row.addWidget(btn_browse)
        layout.addLayout(filter_row)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["审批表文件名", "匹配状态", "文件大小", "最后修改时间"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        layout.addWidget(self.table)

        bot_row = QHBoxLayout()
        btn_open_doc = QPushButton("📄 打开选中的审批表查看")
        btn_open_doc.setStyleSheet("background-color: #F8FAFC; color: #334155; border: 1px solid #CBD5E1; border-radius: 4px; padding: 6px 12px;")
        btn_open_doc.clicked.connect(self.open_selected_docx)

        btn_confirm = QPushButton("🎯 确认匹配并导入该审批表")
        btn_confirm.setStyleSheet("background-color: #15803D; color: white; border-radius: 4px; padding: 6px 16px; font-weight: bold;")
        btn_confirm.clicked.connect(self.confirm_match)

        btn_cancel = QPushButton("取消")
        btn_cancel.clicked.connect(self.reject)

        bot_row.addWidget(btn_open_doc)
        bot_row.addStretch()
        bot_row.addWidget(btn_confirm)
        bot_row.addWidget(btn_cancel)
        layout.addLayout(bot_row)

    def load_candidates(self):
        self.all_candidates = []
        seen = set()

        # 1. 扫描 APPROVAL_DIR
        if os.path.exists(APPROVAL_DIR):
            for f in os.listdir(APPROVAL_DIR):
                if f.endswith(".docx") and not f.startswith("~$"):
                    p = os.path.join(APPROVAL_DIR, f)
                    if p not in seen:
                        seen.add(p)
                        self.all_candidates.append(p)

        # 2. 扫描 BASE_DIR
        for f in os.listdir(BASE_DIR):
            if f.endswith(".docx") and not f.startswith("~$") and ("干部" in f or "任免" in f or "审批" in f):
                p = os.path.join(BASE_DIR, f)
                if p not in seen:
                    seen.add(p)
                    self.all_candidates.append(p)

        # 排序：名字匹配的放在最前面
        self.all_candidates.sort(key=lambda p: (0 if self.name in os.path.basename(p) else 1, os.path.basename(p)))
        self.render_table(self.all_candidates)

    def render_table(self, file_paths):
        self.current_display = file_paths
        self.table.setRowCount(len(file_paths))
        for r, p in enumerate(file_paths):
            fname = os.path.basename(p)
            is_match = (self.name in fname)
            match_str = "🟢 精准匹配" if is_match else "⚪ 其他候选"
            try:
                sz = f"{os.path.getsize(p) / 1024:.1f} KB"
                mtime = datetime.datetime.fromtimestamp(os.path.getmtime(p)).strftime("%Y-%m-%d %H:%M")
            except Exception:
                sz = "—"
                mtime = "—"

            item_fn = QTableWidgetItem(fname)
            item_fn.setData(Qt.UserRole, p)
            if is_match:
                item_fn.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))

            item_m = QTableWidgetItem(match_str)
            if is_match:
                item_m.setForeground(QBrush(QColor("#16A34A")))

            self.table.setItem(r, 0, item_fn)
            self.table.setItem(r, 1, item_m)
            self.table.setItem(r, 2, QTableWidgetItem(sz))
            self.table.setItem(r, 3, QTableWidgetItem(mtime))

        if file_paths:
            self.table.selectRow(0)

    def filter_table(self, text):
        t = text.strip().lower()
        if not t:
            self.render_table(self.all_candidates)
        else:
            filtered = [p for p in self.all_candidates if t in os.path.basename(p).lower()]
            self.render_table(filtered)

    def browse_custom_file(self):
        fpath, _ = QFileDialog.getOpenFileName(self, "选择干部任免审批表Word文件", APPROVAL_DIR if os.path.exists(APPROVAL_DIR) else BASE_DIR, "Word 文档 (*.docx)")
        if fpath:
            if fpath not in self.all_candidates:
                self.all_candidates.insert(0, fpath)
            self.render_table(self.all_candidates)
            self.table.selectRow(0)

    def open_selected_docx(self):
        row = self.table.currentRow()
        if row < 0 or row >= len(self.current_display):
            QMessageBox.warning(self, "提示", "请先在表格中选择一个文档！")
            return
        target_path = self.current_display[row]
        try:
            if sys.platform == "win32":
                os.startfile(target_path)
            elif sys.platform == "darwin":
                subprocess.run(["open", target_path])
            else:
                subprocess.run(["xdg-open", target_path])
        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法打开 Word 文档：{e}")

    def confirm_match(self):
        row = self.table.currentRow()
        if row < 0 or row >= len(self.current_display):
            QMessageBox.warning(self, "提示", "请先在表格中选择一个文档！")
            return
        target_path = self.current_display[row]
        fname = os.path.basename(target_path)

        res_text, fam_text = extract_resume_and_family_from_docx(target_path)

        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()

        c.execute("""SELECT doc_date, doc_number, content FROM cadre_records 
                     WHERE content LIKE ? AND id NOT IN (SELECT record_id FROM excluded_records WHERE cadre_id=?)
                     ORDER BY doc_date ASC, id ASC""", (f"%{self.name}%", self.cadre_id))
        recs = c.fetchall()
        doc_list = [{"doc_date": r[0], "doc_num": r[1], "content": r[2]} for r in recs] if recs else None

        c.execute("SELECT dept_post, current_duty_date, career_history FROM cadres WHERE id = ?", (self.cadre_id,))
        cur_cadre = c.fetchone()
        dept_p = cur_cadre[0] if cur_cadre else ""
        cur_d_dt = cur_cadre[1] if cur_cadre else ""
        car_h = cur_cadre[2] if cur_cadre else ""

        healed_res = heal_and_reorganize_resume(
            res_text or "",
            doc_records=doc_list,
            cadre_name=self.name,
            career_history=car_h,
            current_dept_post=dept_p,
            current_duty_date=cur_d_dt
        )
        clean_fam = format_family_text(fam_text)

        c.execute("""UPDATE cadres SET selected_approval_file = ?, resume_text = ?, family_relations = ? WHERE id = ?""",
                  (fname, healed_res, clean_fam, self.cadre_id))
        conn.commit()
        conn.close()

        QMessageBox.information(
            self, "匹配成功",
            f"🎉 干部【{self.name}】已成功绑定任免审批表：\n{fname}\n\n"
            f"• 个人简历已根据审批表最新履历并结合在库公文自动闭合！\n"
            f"• 家庭主要成员关系已同步更新！\n"
            f"• 全息档案卡与画像单照片已生效！"
        )
        if self.on_matched:
            self.on_matched()
        self.accept()

# ==================== 干部全息档案卡 (PySide6 现代版) ====================

class FullCadreDetailWindow(QDialog):
    def __init__(self, parent, cadre_id, base_ym=None, on_data_changed=None):
        super().__init__(parent)
        self.cadre_id = cadre_id
        self.base_ym = base_ym or datetime.datetime.now().strftime("%Y年%m月")
        self.on_data_changed = on_data_changed

        self.load_info()
        self.setWindowTitle(f"干部全息档案卡 —— {self.d['name']}（{self.d['dept_post']}）")
        self.resize(1180, 860)

        self.create_widgets()
        self.show()

    def load_info(self):
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        cols_query = ", ".join(ALL_CADRE_COLUMNS)
        c.execute(f"SELECT {cols_query} FROM cadres WHERE id = ?", (self.cadre_id,))
        row = c.fetchone()

        self.d = {k: (str(v) if v is not None and str(v) != "None" else "—") for k, v in zip(ALL_CADRE_COLUMNS, row)}

        # 1. 穿透花名册底表校准人员分类/编制身份（解决人事代理与编制身份显示差异问题）
        roster_candidates = find_roster_data_by_name(self.d["name"])
        if roster_candidates:
            r_match = roster_candidates[0]
            if len(roster_candidates) > 1 and is_valid_text(self.d.get("birth")):
                t_ym = clean_ym(self.d["birth"])
                for cand in roster_candidates:
                    if t_ym and t_ym in cand.get("id_card", ""):
                        r_match = cand
                        break
            r_ident = r_match.get("identity")
            if is_valid_text(r_ident) and r_ident != self.d.get("cadre_identity"):
                c.execute("UPDATE cadres SET cadre_identity = ? WHERE id = ?", (r_ident, self.cadre_id))
                conn.commit()
                self.d["cadre_identity"] = r_ident

        # 严禁在此静默覆写修改干部简历！干部的 resume_text 以数据库中存储的内容为权威法定记录。
        # 用户在手动编辑框中的保存拥有绝对最高优先级，程序绝不后台静默篡改。
        conn.close()

    def create_widgets(self):
        root_l = QVBoxLayout(self)
        root_l.setContentsMargins(18, 16, 18, 16)
        root_l.setSpacing(10)

        # 1. 顶部基础信息个人画像卡
        top_card = QFrame()
        top_card.setStyleSheet("background-color: #FFFFFF; border: 1px solid #D1E1F0; border-radius: 8px;")
        top_card_l = QVBoxLayout(top_card)
        top_card_l.setContentsMargins(16, 14, 16, 14)
        top_card_l.setSpacing(6)

        h_row = QHBoxLayout()
        is_chu = (self.d.get("cadre_rank") == "处级")
        title_txt = f"🎖️ 处级领导干部全息画像 (工号: {self.d['emp_no']})" if is_chu else f"🏛️ 科级中层干部全息个人画像 (工号: {self.d['emp_no']})"
        t_lbl = QLabel(title_txt)
        t_lbl.setFont(QFont("Microsoft YaHei UI", 11, QFont.Bold))
        t_lbl.setStyleSheet("color: #0F172A;")
        h_row.addWidget(t_lbl)
        h_row.addStretch()

        btn_open_app = QPushButton("📄 打开审批表原件")
        btn_open_app.setFixedHeight(28)
        btn_open_app.setStyleSheet("background-color: #F0FDF4; color: #166534; border: 1px solid #BBF7D0; border-radius: 6px; padding: 0 10px; font-weight: bold;")
        btn_open_app.clicked.connect(self.open_current_approval_file)

        btn_match_app = QPushButton("📑 手动匹配审批表")
        btn_match_app.setFixedHeight(28)
        btn_match_app.setStyleSheet("background-color: #FEF3C7; color: #92400E; border: 1px solid #FDE68A; border-radius: 6px; padding: 0 10px; font-weight: bold;")
        btn_match_app.clicked.connect(self.open_manual_match_approval)

        btn_bind = QPushButton("🎯 匹配校准底表")
        btn_bind.setFixedHeight(28)
        btn_bind.setStyleSheet("background-color: #EFF6FF; color: #1D4ED8; border: 1px solid #BFDBFE; border-radius: 6px; padding: 0 10px;")
        btn_bind.clicked.connect(self.open_bind_self)

        btn_edit = QPushButton("✏️ 编辑档案")
        btn_edit.setFixedHeight(28)
        btn_edit.setStyleSheet("background-color: #F8FAFC; color: #334155; border: 1px solid #CBD5E1; border-radius: 6px; padding: 0 10px;")
        btn_edit.clicked.connect(self.open_edit_self)

        h_row.addWidget(btn_open_app)
        h_row.addWidget(btn_match_app)
        h_row.addWidget(btn_bind)
        h_row.addWidget(btn_edit)
        top_card_l.addLayout(h_row)

        m_cur = calc_months_diff(self.d["current_duty_date"], self.base_ym)
        cur_len = format_months_to_str(m_cur)
        m_same = calc_months_diff(self.d["same_level_date"], self.base_ym)
        same_len = format_months_to_str(m_same)
        age = calc_age(self.d['birth'], self.base_ym)
        age_str = f"（{age}岁）" if age is not None else ""
        status_tag = f"【在岗】" if self.d["status"] == "在岗" else f"【已{self.d['leave_reason']} ({self.d['leave_date']})】"
        is_impacted, disc_str, _ = check_discipline_impact(self.d['discipline_type'], self.d['discipline_date'], self.base_ym)

        fe_d = self.d.get('first_edu_degree') if is_valid_text(self.d.get('first_edu_degree')) else self.d.get('full_edu_degree', '—')
        fe_s = self.d.get('first_edu_school') if is_valid_text(self.d.get('first_edu_school')) else self.d.get('full_edu_school', '—')
        first_edu_full = f"{fe_d}（{fe_s}）"

        he_d = self.d.get('highest_edu_degree') if is_valid_text(self.d.get('highest_edu_degree')) else self.d.get('job_edu_degree', self.d.get('full_edu_degree', '—'))
        he_s = self.d.get('highest_edu_school') if is_valid_text(self.d.get('highest_edu_school')) else self.d.get('job_edu_school', self.d.get('full_edu_school', '—'))
        highest_edu_full = f"{he_d}（{he_s}）"

        full_edu_str = f"{self.d['full_edu_degree']} ({self.d['full_edu_school']})" if is_valid_text(self.d['full_edu_degree']) or is_valid_text(self.d['full_edu_school']) else "—"
        job_edu_str = f"{self.d['job_edu_degree']} ({self.d['job_edu_school']})" if is_valid_text(self.d['job_edu_degree']) or is_valid_text(self.d['job_edu_school']) else "—"

        t_level = get_title_level(self.d['title'])
        ret_info = calc_retirement_info(self.d['gender'], self.d['birth'], self.base_ym, self.d['title'], self.d.get('cadre_rank', '科级'))
        ret_desc = ret_info['status_str'] if ret_info else "—"
        ret_color = "#DC2626" if (ret_info and ret_info.get("is_critical")) else ("#D97706" if (ret_info and ret_info.get("is_soon")) else "#475569")

        info_lines = [
            f"<span>姓名：<b>{self.d['name']}</b> &nbsp;&nbsp;|&nbsp;&nbsp; 性别：{self.d['gender']} &nbsp;&nbsp;|&nbsp;&nbsp; 出生年月：{self.d['birth']}{age_str} &nbsp;&nbsp;|&nbsp;&nbsp; 编制身份：<b>{self.d['cadre_identity']}</b> &nbsp;&nbsp;|&nbsp;&nbsp; 岗位类别：{self.d['post_category']} &nbsp;&nbsp;|&nbsp;&nbsp; 状态：<span style='color:#16A34A;'>{status_tag}</span></span>",
            f"<span>政治面貌：{self.d['politic']} &nbsp;&nbsp;|&nbsp;&nbsp; 入党时间：{self.d['party_date']} &nbsp;&nbsp;|&nbsp;&nbsp; 参加工作：{self.d['work_date']} &nbsp;&nbsp;|&nbsp;&nbsp; 专业技术职称：<b>{self.d['title']}</b> <span style='background:#EFF6FF;color:#1D4ED8;padding:1px 5px;border-radius:3px;font-size:8pt;'>{t_level}</span></span>",
            f"<span>工作部门及职务：<b>{self.d['dept_post']}</b> &nbsp;&nbsp;|&nbsp;&nbsp; 级别：<b>{self.d['level']}</b> &nbsp;&nbsp;|&nbsp;&nbsp; 任现职：{self.d['current_duty_date']} (已任{cur_len}) &nbsp;&nbsp;|&nbsp;&nbsp; 任同级：{self.d['same_level_date']} (已任{same_len})</span>",
        ]
        if is_chu or is_valid_text(self.d.get("charge_division")):
            info_lines.append(f"<span>🎖️ <b>处级归口类别</b>：{self.d.get('chu_category', '—')} &nbsp;&nbsp;|&nbsp;&nbsp; 📌 <b>领导班子分管业务及联系部门</b>：<b style='color:#1D4ED8;'>{self.d.get('charge_division', '—')}</b></span>")

        sel_app = self.d.get("selected_approval_file", "")
        app_txt = f"<b style='color:#166534;'>{sel_app}</b>" if is_valid_text(sel_app) else "<span style='color:#94A3B8;'>未匹配（点击右上角手动匹配）</span>"

        info_lines.extend([
            f"<span>📑 <b>匹配审批表</b>：{app_txt} &nbsp;&nbsp;|&nbsp;&nbsp; 身份证号：<b>{self.d['id_card']}</b> &nbsp;&nbsp;|&nbsp;&nbsp; 移动电话：<b>{self.d['phone']}</b> &nbsp;&nbsp;|&nbsp;&nbsp; 办公电话：{self.d['office_phone']} &nbsp;&nbsp;|&nbsp;&nbsp; 办公地址：{self.d['office_address']}</span>",
            f"<span>全日制教育：{full_edu_str} &nbsp;&nbsp;|&nbsp;&nbsp; 在职教育：{job_edu_str}</span>",
            f"<span>🎓 <b>第一学历</b>：{first_edu_full} &nbsp;&nbsp;|&nbsp;&nbsp; 🎖️ <b>最高学历/学位</b>：{highest_edu_full}</span>",
            f"<span>⏳ <b>退休预警研判</b>：<span style='color:{ret_color};font-weight:bold;'>{ret_desc}</span> &nbsp;&nbsp;|&nbsp;&nbsp; ⚖️ 纪律监督：<b>{self.d['discipline_type']}</b> ({self.d['discipline_date']}) &nbsp;&nbsp;|&nbsp;&nbsp; <span style='color:{'#DC2626' if is_impacted else '#475569'};font-weight:bold;'>{disc_str}</span></span>"
        ])

        for line_html in info_lines:
            lbl = QLabel(line_html)
            lbl.setTextFormat(Qt.RichText)
            lbl.setFont(QFont("Microsoft YaHei UI", 9))
            lbl.setStyleSheet("color: #334155; line-height: 1.4;")
            top_card_l.addWidget(lbl)

        root_l.addWidget(top_card)

        # 2. 简历智能断层与重叠体检警示横幅
        audit_alerts = audit_resume_timeline(self.d["resume_text"])
        audit_banner = QFrame()
        audit_banner_l = QVBoxLayout(audit_banner)
        audit_banner_l.setContentsMargins(14, 8, 14, 8)
        audit_banner_l.setSpacing(2)

        if audit_alerts:
            audit_banner.setStyleSheet("background-color: #FEF3C7; border: 1px solid #F59E0B; border-radius: 8px;")
            title_a = QLabel(f"⚠️  【干部简历智能断层与合规体检】发现履历衔接异常（共 {len(audit_alerts)} 处）：")
            title_a.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
            title_a.setStyleSheet("color: #92400E;")
            audit_banner_l.addWidget(title_a)
            for item_text in audit_alerts:
                item_lbl = QLabel(f"• {item_text}")
                item_lbl.setFont(QFont("Microsoft YaHei UI", 9))
                item_lbl.setStyleSheet("color: #B45309;")
                audit_banner_l.addWidget(item_lbl)
        else:
            audit_banner.setStyleSheet("background-color: #ECFDF5; border: 1px solid #10B981; border-radius: 8px;")
            title_a = QLabel("✅  【干部简历智能断层与合规体检】认证通过：历任职务时间无缝严密闭合，无异常空白断档与时间重叠。")
            title_a.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
            title_a.setStyleSheet("color: #065F46;")
            audit_banner_l.addWidget(title_a)

        root_l.addWidget(audit_banner)

        # 3. 中间区域：完整简历文本 + 家庭成员关系 (严格 1:1 平均平分)
        mid_row = QHBoxLayout()
        mid_row.setSpacing(12)

        # 简历展示卡片 (支持直接编辑修改)
        res_card = QFrame()
        res_card.setStyleSheet("background-color: #FFFFFF; border: 1px solid #D1E1F0; border-radius: 8px;")
        res_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        res_l = QVBoxLayout(res_card)
        res_l.setContentsMargins(10, 8, 10, 8)

        res_head = QHBoxLayout()
        res_lbl = QLabel("干部完整个人简历 (支持直接编辑修改)")
        res_lbl.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
        btn_save_res = QPushButton("💾 保存简历修改")
        btn_save_res.setFixedHeight(24)
        btn_save_res.setCursor(Qt.PointingHandCursor)
        btn_save_res.setStyleSheet("background-color: #2563EB; color: white; border-radius: 4px; padding: 0 10px; font-weight: bold;")
        btn_save_res.clicked.connect(self.save_manual_resume)
        res_head.addWidget(res_lbl)
        res_head.addStretch()
        res_head.addWidget(btn_save_res)
        res_l.addLayout(res_head)

        self.resume_edit = QTextEdit()
        self.resume_edit.setFont(QFont("Microsoft YaHei UI", 9))
        self.resume_edit.setReadOnly(False)
        self.resume_edit.setLineWrapMode(QTextEdit.WidgetWidth)
        self.resume_edit.setPlainText(self.d["resume_text"])
        res_l.addWidget(self.resume_edit)
        mid_row.addWidget(res_card, 1)

        # 家庭成员展示卡片 (支持直接编辑修改)
        fam_card = QFrame()
        fam_card.setStyleSheet("background-color: #FFFFFF; border: 1px solid #D1E1F0; border-radius: 8px;")
        fam_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        fam_l = QVBoxLayout(fam_card)
        fam_l.setContentsMargins(10, 8, 10, 8)

        fam_head = QHBoxLayout()
        fam_lbl = QLabel("家庭主要成员关系 (支持直接编辑修改)")
        fam_lbl.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
        btn_save_fam = QPushButton("💾 保存家庭修改")
        btn_save_fam.setFixedHeight(24)
        btn_save_fam.setCursor(Qt.PointingHandCursor)
        btn_save_fam.setStyleSheet("background-color: #059669; color: white; border-radius: 4px; padding: 0 10px; font-weight: bold;")
        btn_save_fam.clicked.connect(self.save_manual_family)
        fam_head.addWidget(fam_lbl)
        fam_head.addStretch()
        fam_head.addWidget(btn_save_fam)
        fam_l.addLayout(fam_head)

        alerts_fam = analyze_same_unit_relations(self.d["family_relations"], self.d.get("dept_post", ""))
        if alerts_fam:
            has_red = any("🔴" in a for a in alerts_fam)
            fam_title = "🔴 【重点任职回避关注】" if has_red else "🟡 【同单位工作备案】"
            fam_bg = "#FEF2F2" if has_red else "#FFFBEB"
            fam_color = "#DC2626" if has_red else "#D97706"
            fam_border = "#FCA5A5" if has_red else "#FCD34D"
            fam_warn = QLabel(f"{fam_title}{'；'.join(alerts_fam)}")
            fam_warn.setWordWrap(True)
            fam_warn.setFont(QFont("Microsoft YaHei UI", 8, QFont.Bold))
            fam_warn.setStyleSheet(f"background-color: {fam_bg}; color: {fam_color}; border: 1px solid {fam_border}; border-radius: 4px; padding: 5px;")
            fam_l.addWidget(fam_warn)

        self.family_edit = QTextEdit()
        self.family_edit.setFont(QFont("Microsoft YaHei UI", 9))
        self.family_edit.setReadOnly(False)
        self.family_edit.setLineWrapMode(QTextEdit.WidgetWidth)
        self.family_edit.setPlainText(format_family_text(self.d["family_relations"]))
        fam_l.addWidget(self.family_edit)
        mid_row.addWidget(fam_card, 1)

        mid_row.setStretch(0, 1)
        mid_row.setStretch(1, 1)
        root_l.addLayout(mid_row, 1)

        # 4. 纵向发光任职轨迹时间轴 (Timeline)
        tl_box = QFrame()
        tl_box.setStyleSheet("background-color: #FFFFFF; border: 1px solid #D1E1F0; border-radius: 8px;")
        tl_l = QVBoxLayout(tl_box)
        tl_l.setContentsMargins(12, 10, 12, 10)
        tl_l.setSpacing(6)

        tl_header = QHBoxLayout()
        tl_title = QLabel("📜 历次职务任免红头公文轨迹时间轴 (按时间正序排列，点击右侧按钮直接打开本地红头公文)")
        tl_title.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
        tl_title.setStyleSheet("color: #1E40AF;")
        tl_header.addWidget(tl_title)
        tl_header.addStretch()

        btn_manage_excluded = QPushButton("📋 管理已剔除公文")
        btn_manage_excluded.setFont(QFont("Microsoft YaHei UI", 8, QFont.Bold))
        btn_manage_excluded.setFixedHeight(26)
        btn_manage_excluded.setStyleSheet("background-color: #DC2626; color: white; border-radius: 4px; padding: 0 8px;")
        btn_manage_excluded.clicked.connect(self.open_manage_excluded_dialog)
        tl_header.addWidget(btn_manage_excluded)

        btn_append_resume = QPushButton("⚡ 依据在库公文自动闭合并更新最新简历")
        btn_append_resume.setFont(QFont("Microsoft YaHei UI", 8, QFont.Bold))
        btn_append_resume.setFixedHeight(26)
        btn_append_resume.setStyleSheet("background-color: #059669; color: white; border-radius: 4px; padding: 0 8px;")
        btn_append_resume.clicked.connect(self.sync_doc_records_into_resume)
        tl_header.addWidget(btn_append_resume)

        # 核心增强：一键深度缝合历任职务摘要
        btn_merge_history = QPushButton("🧩 一键将《历任职务摘要》缝合入简历")
        btn_merge_history.setFont(QFont("Microsoft YaHei UI", 8, QFont.Bold))
        btn_merge_history.setFixedHeight(26)
        btn_merge_history.setStyleSheet("background-color: #2563EB; color: white; border-radius: 4px; padding: 0 8px;")
        btn_merge_history.clicked.connect(self.merge_history_into_resume_action)
        tl_header.addWidget(btn_merge_history)

        tl_l.addLayout(tl_header)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { border: none; background-color: #F8FAFC; }")

        timeline_container = QWidget()
        timeline_container.setStyleSheet("background-color: #F8FAFC;")
        self.timeline_layout = QVBoxLayout(timeline_container)
        self.timeline_layout.setContentsMargins(14, 10, 14, 10)
        self.timeline_layout.setSpacing(8)

        self.load_timeline_data()
        scroll.setWidget(timeline_container)
        tl_l.addWidget(scroll, 1)

        root_l.addWidget(tl_box, 1)

        # 5. 底部动作条
        bot_bar = QHBoxLayout()
        btn_word_profile = QPushButton("📑 导出画像汇报单(Word)")
        btn_word_profile.setFixedHeight(34)
        btn_word_profile.setStyleSheet("background-color: #0D9488; color: white; border-radius: 6px; padding: 0 16px; font-weight: bold;")
        btn_word_profile.clicked.connect(self.export_word_profile)

        btn_pdf = QPushButton("🖨️ 导出画像汇报单(PDF)")
        btn_pdf.setFixedHeight(34)
        btn_pdf.setStyleSheet("background-color: #2563EB; color: white; border-radius: 6px; padding: 0 16px; font-weight: bold;")
        btn_pdf.clicked.connect(self.export_pdf)

        btn_word = QPushButton("📑 导出《干部任免审批表》(Word)")
        btn_word.setFixedHeight(34)
        btn_word.setStyleSheet("background-color: #15803D; color: white; border-radius: 6px; padding: 0 16px; font-weight: bold;")
        btn_word.clicked.connect(lambda: ExportApprovalFormDialog(self, self.cadre_id, self.base_ym).exec())

        btn_excel = QPushButton("📊 导出任职轨迹Excel表")
        btn_excel.setFixedHeight(34)
        btn_excel.setStyleSheet("background-color: #2E8B57; color: white; border-radius: 6px; padding: 0 16px; font-weight: bold;")
        btn_excel.clicked.connect(self.export_excel)

        btn_close = QPushButton(" 关闭 ")
        btn_close.setFixedHeight(34)
        btn_close.setStyleSheet("background-color: #F1F5F9; color: #475569; border: 1px solid #CBD5E1; border-radius: 6px; padding: 0 16px;")
        btn_close.clicked.connect(self.accept)

        bot_bar.addStretch()
        bot_bar.addWidget(btn_word_profile)
        bot_bar.addWidget(btn_pdf)
        bot_bar.addWidget(btn_word)
        bot_bar.addWidget(btn_excel)
        bot_bar.addWidget(btn_close)

        root_l.addLayout(bot_bar)

    def open_manage_excluded_dialog(self):
        w = ManageExcludedRecordsDialog(self, self.cadre_id, self.d.get("name", ""), on_changed=self.load_timeline_data)
        w.exec()

    def export_word_profile(self):
        default_name = f"{self.d['name']}_干部全息画像汇报单.docx"
        save_path, _ = QFileDialog.getSaveFileName(
            self, "导出干部全息画像汇报单(Word)", os.path.join(BASE_DIR, default_name), "Word 文档 (*.docx)"
        )
        if not save_path:
            return

        try:
            export_cadre_profile_docx(self.d, self.base_ym, save_path)
            res = QMessageBox.information(
                self, "导出成功",
                f"🎉 干部【{self.d['name']}】全息画像汇报单(Word)已生成！\n\n"
                f"保存路径：{save_path}\n\n"
                f"是否立即打开 Word 文档进行排版与打印？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
            )
            if res == QMessageBox.Yes:
                QDesktopServices.openUrl(QUrl.fromLocalFile(save_path))
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"生成 Word 汇报单发生错误: {e}")

    def load_timeline_data(self):
        while self.timeline_layout.count():
            item = self.timeline_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("""SELECT id, doc_date, doc_number, content, file_name 
                     FROM cadre_records 
                     WHERE content LIKE ? 
                     AND id NOT IN (SELECT record_id FROM excluded_records WHERE cadre_id = ?)
                     ORDER BY doc_date ASC, id ASC""", (f"%{self.d['name']}%", self.cadre_id))
        records = c.fetchall()

        # 穿透同一份公文原件中的关联条款（如：紧跟在任职名单后起一行的“上述人员试用期一年”等条款）
        file_names = list(set([r[4] for r in records if r[4]]))
        file_context_map = {}
        if file_names:
            placeholders = ','.join(['?'] * len(file_names))
            c.execute(f"SELECT file_name, content FROM cadre_records WHERE file_name IN ({placeholders}) ORDER BY id ASC", file_names)
            for fn, cnt in c.fetchall():
                if fn not in file_context_map:
                    file_context_map[fn] = []
                file_context_map[fn].append(str(cnt or ""))

        conn.close()

        # 按公文原件（file_name）智能归并去重，彻底解决同一文件因多个段落命中而被反复渲染3次的问题
        file_groups = {}
        for r in records:
            rec_id, doc_date, doc_num, content, file_name = r
            if not file_name:
                file_name = f"doc_{rec_id}"
            if file_name not in file_groups:
                file_groups[file_name] = []
            file_groups[file_name].append(r)

        raw_list = []
        for file_name, file_records in file_groups.items():
            all_file_paras = file_context_map.get(file_name, [])
            full_file_text = " ".join(all_file_paras)

            # 汇总该公文在该干部下的所有记录ID（供剔除时彻底清除关联段落）
            all_rec_ids = [r[0] for r in file_records]
            primary_rec_id = all_rec_ids[0]
            doc_date = file_records[0][1]
            doc_num = file_records[0][2]

            # 智能提取最核心任免决定语句（优先选含有'任'、'免'、'聘'且不是纯标题或抄送的段落）
            best_content = ""
            for r in file_records:
                cnt = str(r[3] or "").strip()
                if any(k in cnt for k in ['任', '聘', '命', '试用期', '兼', '免去', '不再担任']) and not cnt.startswith("关于") and not cnt.startswith("抄送") and not cnt.startswith("校属各单位"):
                    best_content = cnt
                    break
            if not best_content:
                for r in file_records:
                    cnt = str(r[3] or "").strip()
                    if not cnt.startswith("抄送") and not cnt.startswith("校属各单位"):
                        best_content = cnt
                        break
            if not best_content:
                best_content = str(file_records[0][3] or "").strip()

            # 提取公文中与试用期/任免相关的补充说明条款
            probation_note = ""
            for p in all_file_paras:
                p_s = p.strip()
                if re.search(r'(?:上述|以上|该|相关|新提拔)?(?:人员|同志|干部)?(?:任职)?试用期(?:为)?(?:一|1)年', p_s) or ("试用期" in p_s and "试用期满" not in p_s and "考核合格" not in p_s and "正式任职" not in p_s):
                    probation_note = p_s
                    break

            raw_list.append({
                "rec_id": primary_rec_id,
                "all_rec_ids": all_rec_ids,
                "doc_date": doc_date,
                "doc_num": doc_num,
                "content": best_content,
                "file_name": file_name,
                "probation_note": probation_note,
                "doc_full_context": full_file_text,
                "cadre_id": self.cadre_id,
                "cadre_name": self.d.get("name", "")
            })

        raw_list.sort(key=lambda x: parse_sort_date(x["doc_date"]))
        self.doc_records = raw_list

        for idx, item in enumerate(raw_list):
            is_last = (idx == len(raw_list) - 1)
            card = TimelineNodeCard(item, is_last=is_last, parent=self)
            card.parent_window = self
            self.timeline_layout.addWidget(card)

        if not raw_list:
            empty_lbl = QLabel("暂未在库中检索到包含该干部的红头任免公文（或已被全部剔除）。")
            empty_lbl.setFont(QFont("Microsoft YaHei UI", 9))
            empty_lbl.setStyleSheet("color: #94A3B8; padding: 20px;")
            empty_lbl.setAlignment(Qt.AlignCenter)
            self.timeline_layout.addWidget(empty_lbl)

    def save_manual_resume(self):
        new_text = self.resume_edit.toPlainText().strip()
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("UPDATE cadres SET resume_text = ? WHERE id = ?", (new_text, self.cadre_id))
        conn.commit()
        conn.close()

        self.d["resume_text"] = new_text
        QMessageBox.information(
            self, "保存成功",
            f"🎉 干部【{self.d['name']}】的个人简历已成功保存入库！\n\n"
            f"您的手动修改拥有最高法定优先级，系统不会自动覆盖。"
        )
        if self.on_data_changed:
            self.on_data_changed()

    def save_manual_family(self):
        new_text = self.family_edit.toPlainText().strip()
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("UPDATE cadres SET family_relations = ? WHERE id = ?", (new_text, self.cadre_id))
        conn.commit()
        conn.close()

        self.d["family_relations"] = new_text
        QMessageBox.information(
            self, "保存成功",
            f"🎉 干部【{self.d['name']}】的家庭成员关系已成功保存入库！"
        )
        if self.on_data_changed:
            self.on_data_changed()

    def merge_history_into_resume_action(self):
        career_h = self.d.get("career_history", "")
        if not is_valid_text(career_h):
            QMessageBox.warning(self, "提示", "该干部档案中未录入任何《历任职务摘要》，无需缝合！")
            return

        cur_text = self.resume_edit.toPlainText().strip() or self.d.get("resume_text", "")
        merged_resume = heal_and_reorganize_resume(
            cur_text,
            doc_records=getattr(self, "doc_records", None),
            cadre_name=self.d["name"],
            career_history=career_h,
            current_dept_post=self.d.get('dept_post', ''),
            current_duty_date=self.d.get('current_duty_date', '')
        )

        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("UPDATE cadres SET resume_text = ? WHERE id = ?", (merged_resume, self.cadre_id))
        conn.commit()
        conn.close()

        self.resume_edit.setPlainText(merged_resume)
        self.d["resume_text"] = merged_resume

        QMessageBox.information(
            self, "缝合成功",
            "🎉 档案深度缝合完毕！\n\n"
            "花名册中的历任职务摘要已严格按时间递增轴织入个人简历，多余重叠碎片已消除，历史断层已自愈！"
        )
        if self.on_data_changed:
            self.on_data_changed()

    def sync_doc_records_into_resume(self):
        if not getattr(self, "doc_records", None):
            QMessageBox.warning(self, "提示", "该干部暂无任免公文记录！")
            return

        cur_text = self.resume_edit.toPlainText().strip() or self.d.get("resume_text", "")
        updated_res = heal_and_reorganize_resume(
            cur_text,
            doc_records=self.doc_records,
            cadre_name=self.d["name"],
            career_history=self.d.get("career_history", ""),
            current_dept_post=self.d.get('dept_post', ''),
            current_duty_date=self.d.get('current_duty_date', '')
        )

        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("UPDATE cadres SET resume_text = ? WHERE id = ?", (updated_res, self.cadre_id))
        conn.commit()
        conn.close()

        self.resume_edit.setPlainText(updated_res)
        self.d["resume_text"] = updated_res

        QMessageBox.information(
            self, "同步与闭合成功",
            "🎉 简历已成功融合在库任免公文！\n\n"
            "• 历史悬空任职已依据后续任免节点自动闭合\n"
            "• 多公文重复任职碎片已去重融合\n"
            "• 最新职务已规范织入并设置为至今\n\n"
            "卡片简历已即时刷新入库并完成规范断行！"
        )
        if self.on_data_changed:
            self.on_data_changed()

    def open_bind_self(self):
        w = ManualBindRosterDialog(self, self.cadre_id, self.d["name"], on_bound=self.refresh_after_edit)
        w.exec()

    def open_edit_self(self):
        w = FullCadreEditDialog(self, title=f"编辑全息档案 —— {self.d['name']}", cadre_id=self.cadre_id,
                                initial_resume=self.d.get("resume_text"), initial_family=self.d.get("family_relations"),
                                on_saved=self.refresh_after_edit)
        w.exec()

    def open_current_approval_file(self):
        sel_file = self.d.get("selected_approval_file", "")
        target_path = None
        if is_valid_text(sel_file):
            for candidate in [
                sel_file,
                os.path.join(APPROVAL_DIR, sel_file),
                os.path.join(BASE_DIR, sel_file)
            ]:
                if os.path.exists(candidate):
                    target_path = candidate
                    break

        if not target_path:
            # 自动在审批表库中查找包含干部姓名的文件
            if os.path.exists(APPROVAL_DIR):
                for f in os.listdir(APPROVAL_DIR):
                    if f.endswith(".docx") and not f.startswith("~$") and self.d["name"] in f:
                        target_path = os.path.join(APPROVAL_DIR, f)
                        break

        if not target_path or not os.path.exists(target_path):
            reply = QMessageBox.question(
                self, "未找到审批表原件",
                f"系统暂未定位到【{self.d['name']}】的干部任免审批表 Word 原件。\n\n是否立即打开【手动匹配审批表】窗口指定文件？",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.open_manual_match_approval()
            return

        try:
            if sys.platform == "win32":
                os.startfile(target_path)
            elif sys.platform == "darwin":
                subprocess.run(["open", target_path])
            else:
                subprocess.run(["xdg-open", target_path])
        except Exception as e:
            QMessageBox.critical(self, "打开失败", f"无法调起 Word 打开该文档：\n{e}")

    def open_manual_match_approval(self):
        dlg = ManualMatchApprovalDialog(self, self.cadre_id, self.d["name"], on_matched=self.refresh_after_edit)
        dlg.exec()

    def refresh_after_edit(self):
        self.load_info()
        self.close()
        FullCadreDetailWindow(self.parent(), self.cadre_id, base_ym=self.base_ym, on_data_changed=self.on_data_changed)
        if self.on_data_changed:
            self.on_data_changed()

    def export_excel(self):
        if not getattr(self, "doc_records", None):
            QMessageBox.warning(self, "提示", "该干部暂无任免公文记录！")
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self, "导出干部任职轨迹表", os.path.join(BASE_DIR, f"{self.d['name']}_干部全息任职轨迹表.xlsx"), "Excel 表格 (*.xlsx)"
        )
        if not save_path: return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "干部全息履历"
            headers = [
                "序号", "干部姓名", "身份证号", "管理层级", "人员编制身份", "岗位类别", "移动电话",
                "职务级别", "发文日期", "公文字号", "任免条款详情", "对应红头公文"
            ]
            ws.append(headers)

            for idx, item in enumerate(self.doc_records, start=1):
                ws.append([
                    idx, self.d["name"], self.d["id_card"], self.d.get("cadre_rank", "科级"),
                    self.d["cadre_identity"], self.d["post_category"], self.d["phone"],
                    self.d["level"], item["doc_date"], item["doc_num"], item["content"], item["file_name"]
                ])

            h_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
            for cell in ws[1]:
                cell.fill = h_fill
                cell.font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center", vertical="center")

            wb.save(save_path)
            QMessageBox.information(self, "导出成功", f"干部【{self.d['name']}】任职轨迹已成功导出！\n路径：{save_path}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"导出出错: {e}")

    def export_pdf(self):
        default_name = f"{self.d['name']}_干部全息画像汇报单.pdf"
        save_path, _ = QFileDialog.getSaveFileName(
            self, "导出干部全息画像汇报单(PDF)", os.path.join(BASE_DIR, default_name), "PDF 文档 (*.pdf)"
        )
        if not save_path:
            return

        try:
            export_cadre_profile_pdf(self.d, self.base_ym, save_path)
            res = QMessageBox.information(
                self, "导出成功",
                f"🎉 干部【{self.d['name']}】全息画像汇报单已生成！\n\n"
                f"保存路径：{save_path}\n\n"
                f"是否立即打开 PDF 预览并进行打印？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
            )
            if res == QMessageBox.Yes:
                QDesktopServices.openUrl(QUrl.fromLocalFile(save_path))
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"生成 PDF 汇报单发生错误: {e}")

# ==================== 档案编辑对话框 ====================

class FullCadreEditDialog(QDialog):
    def __init__(self, parent, title="编辑干部档案", cadre_id=None, default_rank="科级",
                 default_chu_category="单位领导班子", initial_resume=None, initial_family=None, on_saved=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(880, 740)
        self.cadre_id = cadre_id
        self.default_rank = default_rank
        self.default_chu_category = default_chu_category
        self.initial_resume = initial_resume
        self.initial_family = initial_family
        self.on_saved = on_saved
        self.fields = {}

        self.create_widgets()
        if cadre_id:
            self.load_data()

    def create_widgets(self):
        root_l = QVBoxLayout(self)
        self.setStyleSheet("""
            QDialog {
                background-color: #FFFFFF;
                font-family: -apple-system, "PingFang SC", "Microsoft YaHei UI", sans-serif;
            }
            QLineEdit {
                border: 1px solid #CBD5E1;
                border-radius: 5px;
                padding: 4px 8px;
                background-color: #FFFFFF;
                color: #0F172A;
                font-size: 13px;
                min-height: 24px;
            }
            QLineEdit:focus {
                border: 1.5px solid #2563EB;
            }
            QTextEdit {
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                background-color: #FFFFFF;
                color: #0F172A;
                font-size: 13px;
            }
            QTextEdit:focus {
                border: 1.5px solid #2563EB;
            }
        """)
        self.tabs = QTabWidget()
        self.tabs.setFont(QFont("Microsoft YaHei UI", 9))

        # Tab 1: 基础身份与通讯
        t1 = QWidget(); l1 = QVBoxLayout(t1); l1.setSpacing(8)
        items1 = [
            ("name", "姓名 (*必填):"), ("gender", "性别:"), ("birth", "出生年月 (如1985年06月):"),
            ("nation", "民族:"), ("native_place", "籍贯:"), ("id_card", "身份证号:"),
            ("phone", "移动电话 (手机):"), ("office_phone", "办公电话 (固话/内线):"),
            ("office_address", "办公地址:"), ("emp_no", "员工编号 (工号):")
        ]
        for code, label in items1:
            h = QHBoxLayout()
            lbl = QLabel(label); lbl.setFixedWidth(160)
            ent = QLineEdit()
            self.fields[code] = ent
            h.addWidget(lbl); h.addWidget(ent)
            if code == "name":
                btn_auto = QPushButton("🔍 从花名册联想补全")
                btn_auto.clicked.connect(self.auto_fill_from_roster)
                h.addWidget(btn_auto)
            l1.addLayout(h)
        l1.addStretch()
        self.tabs.addTab(t1, " 1. 基础身份与通讯 ")

        # Tab 2: 职务与岗位分类
        t2 = QWidget(); l2 = QVBoxLayout(t2); l2.setSpacing(8)
        items2 = [
            ("dept_post", "工作部门及职务 (*必填):"), ("level", "职务级别 (如正科/副科):"),
            ("current_duty_date", "任现职时间:"), ("same_level_date", "任同级时间:"),
            ("career_history", "历任职务摘要:"), ("remark", "备注说明:")
        ]
        for code, label in items2:
            h = QHBoxLayout()
            lbl = QLabel(label); lbl.setFixedWidth(160)
            ent = QLineEdit()
            self.fields[code] = ent
            h.addWidget(lbl); h.addWidget(ent)
            l2.addLayout(h)

        h_cat = QHBoxLayout()
        lbl_rk = QLabel("管理层级:"); lbl_rk.setFont(QFont("PingFang SC" if sys.platform == "darwin" else "Microsoft YaHei UI", 9, QFont.Bold))
        h_cat.addWidget(lbl_rk)
        self.rank_cb = QComboBox()
        self.rank_cb.addItems(["科级", "处级"])
        self.rank_cb.setCurrentText(self.default_rank)
        self.rank_cb.setItemDelegate(QStyledItemDelegate(self.rank_cb))
        self.rank_cb.setStyleSheet(COMBO_STYLE)
        h_cat.addWidget(self.rank_cb)

        lbl_id = QLabel("编制身份:"); lbl_id.setFont(QFont("PingFang SC" if sys.platform == "darwin" else "Microsoft YaHei UI", 9, QFont.Bold))
        h_cat.addWidget(lbl_id)
        self.ident_cb = QComboBox()
        self.ident_cb.addItems(["校编", "院编", "人事代理", "其他"])
        self.ident_cb.setItemDelegate(QStyledItemDelegate(self.ident_cb))
        self.ident_cb.setStyleSheet(COMBO_STYLE)
        h_cat.addWidget(self.ident_cb)

        lbl_pc = QLabel("岗位类别:"); lbl_pc.setFont(QFont("PingFang SC" if sys.platform == "darwin" else "Microsoft YaHei UI", 9, QFont.Bold))
        h_cat.addWidget(lbl_pc)
        self.post_cat_cb = QComboBox()
        self.post_cat_cb.addItems(["党政管理岗", "业务管理岗"])
        self.post_cat_cb.setItemDelegate(QStyledItemDelegate(self.post_cat_cb))
        self.post_cat_cb.setStyleSheet(COMBO_STYLE)
        h_cat.addWidget(self.post_cat_cb)
        l2.addLayout(h_cat)

        # 处级干部专属归口与分管业务编辑区
        self.chu_section_frame = QFrame()
        self.chu_section_frame.setObjectName("chu_section_frame")
        self.chu_section_frame.setStyleSheet("""
            QFrame#chu_section_frame {
                background-color: #F8FAFC;
                border: 1.5px dashed #3B82F6;
                border-radius: 8px;
            }
        """)
        chu_sec_l = QVBoxLayout(self.chu_section_frame)
        chu_sec_l.setSpacing(8)
        chu_sec_l.setContentsMargins(12, 10, 12, 10)

        h_chu_cat = QHBoxLayout()
        lbl_cc = QLabel("处级归口类别 (处级专属):"); lbl_cc.setFixedWidth(160)
        lbl_cc.setStyleSheet("font-weight: bold; color: #1E3A8A;")
        self.chu_cat_cb = QComboBox()
        self.chu_cat_cb.addItems(CHU_CATEGORIES)
        self.chu_cat_cb.setCurrentText(self.default_chu_category)
        self.chu_cat_cb.setItemDelegate(QStyledItemDelegate(self.chu_cat_cb))
        self.chu_cat_cb.setStyleSheet(COMBO_STYLE)
        h_chu_cat.addWidget(lbl_cc); h_chu_cat.addWidget(self.chu_cat_cb)
        chu_sec_l.addLayout(h_chu_cat)

        h_charge = QHBoxLayout()
        lbl_cd = QLabel("领导班子分管业务及部门:"); lbl_cd.setFixedWidth(160)
        lbl_cd.setStyleSheet("font-weight: bold; color: #1E3A8A;")
        self.charge_div_input = QLineEdit()
        self.charge_div_input.setStyleSheet("border: 1px solid #CBD5E1; border-radius: 6px; padding: 4px 8px; background-color: #FFFFFF; color: #0F172A;")
        self.charge_div_input.setPlaceholderText("如：主持党委全面工作，分管党政办公室、组织人事部")
        h_charge.addWidget(lbl_cd); h_charge.addWidget(self.charge_div_input)
        chu_sec_l.addLayout(h_charge)

        l2.addWidget(self.chu_section_frame)

        def on_rank_changed(text):
            if text == "处级":
                self.chu_section_frame.show()
            else:
                self.chu_section_frame.hide()

        self.rank_cb.currentTextChanged.connect(on_rank_changed)
        if self.default_rank != "处级":
            self.chu_section_frame.hide()

        l2.addStretch()
        self.tabs.addTab(t2, " 2. 职务与岗位分类 ")

        # Tab 3: 政治面貌与两级学历档案
        t3_scroll = QScrollArea()
        t3_scroll.setWidgetResizable(True)
        t3_scroll.setFrameShape(QFrame.NoFrame)
        t3 = QWidget()
        l3 = QVBoxLayout(t3)
        l3.setSpacing(6)
        l3.setContentsMargins(14, 12, 14, 12)
        t3_scroll.setWidget(t3)

        def add_sec_banner(title_text):
            sec_lbl = QLabel(title_text)
            sec_lbl.setFont(QFont("PingFang SC" if sys.platform == "darwin" else "Microsoft YaHei UI", 9, QFont.Bold))
            sec_lbl.setStyleSheet("color: #1E3A8A; background-color: #EFF6FF; border: 1px solid #BFDBFE; padding: 5px 8px; border-radius: 4px; margin-top: 6px;")
            l3.addWidget(sec_lbl)

        def add_field(code, label):
            h = QHBoxLayout()
            lbl = QLabel(label); lbl.setFixedWidth(160)
            lbl.setStyleSheet("color: #334155; font-weight: 500;")
            ent = QLineEdit()
            self.fields[code] = ent
            h.addWidget(lbl); h.addWidget(ent)
            l3.addLayout(h)

        add_sec_banner("📌 政治面貌与专业技术职称")
        add_field("politic", "政治面貌 (中共党员等):")
        add_field("party_date", "入党时间 (如2010.05):")
        add_field("work_date", "参加工作时间 (如2012.07):")
        add_field("title", "专业技术职称 (如副主任医师):")
        add_field("title_date", "职称聘任时间 (如2020.12):")

        add_sec_banner("🎓 第一学历（初始学历 · 花名册规范核准）")
        add_field("first_edu_degree", "第一学历层次 (如本科/大专):")
        add_field("first_edu_school", "第一学历毕业学校 (如平顶山学院):")
        add_field("first_edu_major", "第一学历毕业专业 (如临床医学):")
        add_field("first_edu_date", "第一学历毕业时间 (如2012.07):")

        add_sec_banner("🎖️ 最高学历与最高学位")
        add_field("highest_edu_degree", "最高学历层次 (如博士研究生/硕士):")
        add_field("highest_edu_degree_title", "所获最高学位 (如医学博士/硕士):")
        add_field("highest_edu_school", "最高学历毕业学校 (如郑州大学):")
        add_field("highest_edu_major", "最高学历毕业专业:")
        add_field("highest_edu_date", "最高学历毕业时间:")

        add_sec_banner("📋 干部任免审批表申报口径（全日制与在职教育）")
        add_field("full_edu_degree", "全日制学历学位 (如大学/医学学士):")
        add_field("full_edu_school", "全日制院校及专业 (如郑州大学/临床医学):")
        add_field("job_edu_degree", "在职学历学位 (如研究生/医学博士):")
        add_field("job_edu_school", "在职院校及专业:")

        l3.addStretch()
        self.tabs.addTab(t3_scroll, " 3. 政治面貌与学历档案 ")

        # Tab 4: 处分与诫勉监督
        t4 = QWidget(); l4 = QVBoxLayout(t4); l4.setSpacing(12)
        h_d1 = QHBoxLayout()
        lbl_dt = QLabel("处分/诫勉类型:"); lbl_dt.setFixedWidth(150)
        self.disc_combo = QComboBox()
        self.disc_combo.addItems(["无", "诫勉谈话", "政务警告", "党纪警告", "政务记过", "党纪严重警告", "政务记大过", "撤销党内职务", "政务降级/撤职", "留党察看"])
        self.disc_combo.setItemDelegate(QStyledItemDelegate(self.disc_combo))
        self.disc_combo.setStyleSheet(COMBO_STYLE)
        h_d1.addWidget(lbl_dt); h_d1.addWidget(self.disc_combo)
        l4.addLayout(h_d1)

        h_d2 = QHBoxLayout()
        lbl_dd = QLabel("处分发文日期 (如2024年03月):"); lbl_dd.setFixedWidth(150)
        self.disc_date_input = QLineEdit()
        h_d2.addWidget(lbl_dd); h_d2.addWidget(self.disc_date_input)
        l4.addLayout(h_d2)

        h_d3 = QHBoxLayout()
        lbl_ds = QLabel("决定文号及具体事由:"); lbl_ds.setFixedWidth(150)
        self.disc_desc_input = QLineEdit()
        h_d3.addWidget(lbl_ds); h_d3.addWidget(self.disc_desc_input)
        l4.addLayout(h_d3)
        l4.addStretch()
        self.tabs.addTab(t4, " 4. 处分与诫勉监督 (红线把关) ")

        # Tab 5: 简历与家庭成员
        t5 = QWidget(); l5 = QVBoxLayout(t5)
        l5.addWidget(QLabel("干部完整个人简历 (多行，自动规范断行):"))
        self.resume_edit_box = QTextEdit()
        l5.addWidget(self.resume_edit_box)
        l5.addWidget(QLabel("家庭主要成员及重要社会关系 (称谓：姓名 | 出生年月 | 政治面貌 | 单位职务):"))
        self.family_edit_box = QTextEdit()
        l5.addWidget(self.family_edit_box)
        self.tabs.addTab(t5, " 5. 简历与家庭成员 ")

        root_l.addWidget(self.tabs)

        bot = QHBoxLayout()
        btn_save = QPushButton(" 💾 保存档案 "); btn_save.setFixedHeight(34); btn_save.setStyleSheet("background-color: #2563EB; color: white; border-radius: 6px; padding: 0 16px; font-weight: bold;")
        btn_save.clicked.connect(self.save)
        btn_cancel = QPushButton(" 取消 "); btn_cancel.setFixedHeight(34); btn_cancel.clicked.connect(self.reject)
        bot.addStretch(); bot.addWidget(btn_save); bot.addWidget(btn_cancel)
        root_l.addLayout(bot)

    def auto_fill_from_roster(self):
        name = self.fields["name"].text().strip()
        if not name: return
        cands = find_roster_data_by_name(name)
        if not cands:
            QMessageBox.information(self, "提示", f"花名册中未检索到【{name}】")
            return
        target = cands[0]
        self.fields["emp_no"].setText(target.get("emp_no", ""))
        self.fields["gender"].setText(target.get("gender", ""))
        self.fields["birth"].setText(target.get("birth", ""))
        self.fields["id_card"].setText(target.get("id_card", ""))
        if is_valid_text(target.get("phone")):
            self.fields["phone"].setText(target.get("phone", ""))
        self.ident_cb.setCurrentText(target.get("identity", "校编"))
        self.fields["title"].setText(target.get("title", ""))
        if target.get("first_edu_school"):
            self.fields["full_edu_degree"].setText(target.get("first_edu_degree", "大学本科"))
            self.fields["full_edu_school"].setText(f"{target.get('first_edu_school','')}{target.get('first_edu_major','')}")
        if "first_edu_degree" in self.fields and target.get("first_edu_degree"):
            self.fields["first_edu_degree"].setText(target.get("first_edu_degree", ""))
        if "first_edu_school" in self.fields and target.get("first_edu_school"):
            self.fields["first_edu_school"].setText(target.get("first_edu_school", ""))
        if "first_edu_major" in self.fields and target.get("first_edu_major"):
            self.fields["first_edu_major"].setText(target.get("first_edu_major", ""))
        if "first_edu_date" in self.fields and target.get("first_edu_date"):
            self.fields["first_edu_date"].setText(target.get("first_edu_date", ""))
        if "highest_edu_degree" in self.fields and target.get("highest_edu_degree"):
            self.fields["highest_edu_degree"].setText(target.get("highest_edu_degree", ""))
        if "highest_edu_degree_title" in self.fields and target.get("highest_edu_degree_title"):
            self.fields["highest_edu_degree_title"].setText(target.get("highest_edu_degree_title", ""))
        if "highest_edu_school" in self.fields and target.get("highest_edu_school"):
            self.fields["highest_edu_school"].setText(target.get("highest_edu_school", ""))
        if "highest_edu_major" in self.fields and target.get("highest_edu_major"):
            self.fields["highest_edu_major"].setText(target.get("highest_edu_major", ""))
        if "highest_edu_date" in self.fields and target.get("highest_edu_date"):
            self.fields["highest_edu_date"].setText(target.get("highest_edu_date", ""))
        QMessageBox.information(self, "补全成功", f"已从花名册补全【{name}】的基础数据！")

    def load_data(self):
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        cols = ", ".join(ALL_CADRE_COLUMNS)
        c.execute(f"SELECT {cols} FROM cadres WHERE id = ?", (self.cadre_id,))
        row = c.fetchone()
        conn.close()
        if not row: return

        d = dict(zip(ALL_CADRE_COLUMNS, row))
        for code, ent in self.fields.items():
            val = d.get(code, "")
            if is_valid_text(val): ent.setText(str(val))

        c_rank = d.get("cadre_rank", "科级") or "科级"
        self.rank_cb.setCurrentText(c_rank)
        self.ident_cb.setCurrentText(d.get("cadre_identity", "院编") or "院编")
        self.post_cat_cb.setCurrentText(d.get("post_category", "党政管理岗") or "党政管理岗")
        self.chu_cat_cb.setCurrentText(d.get("chu_category", "单位领导班子") or "单位领导班子")
        self.charge_div_input.setText(d.get("charge_division", "") or "")
        if c_rank == "处级":
            self.chu_section_frame.show()
        else:
            self.chu_section_frame.hide()

        disc_t = d.get("discipline_type", "无") or "无"
        self.disc_combo.setCurrentText(disc_t if disc_t in [self.disc_combo.itemText(i) for i in range(self.disc_combo.count())] else "无")
        self.disc_date_input.setText(d.get("discipline_date", "") if is_valid_text(d.get("discipline_date")) else "")
        self.disc_desc_input.setText(d.get("discipline_desc", "") if is_valid_text(d.get("discipline_desc")) else "")

        self.resume_edit_box.setPlainText(format_resume_text(d.get("resume_text", "") or ""))
        self.family_edit_box.setPlainText(format_family_text(d.get("family_relations", "") or ""))

    def save(self):
        name = self.fields["name"].text().strip()
        dept_post = self.fields["dept_post"].text().strip()
        if not name or not dept_post:
            QMessageBox.warning(self, "提示", "姓名与工作部门及职务必须填写！")
            return

        c_rank = self.rank_cb.currentText()
        chu_cat = self.chu_cat_cb.currentText() if c_rank == "处级" else ""
        charge_div = self.charge_div_input.text().strip() if c_rank == "处级" else ""
        ident = self.ident_cb.currentText()
        post_cat = self.post_cat_cb.currentText()
        disc_t = self.disc_combo.currentText()
        disc_d = self.disc_date_input.text().strip() or "—"
        disc_desc = self.disc_desc_input.text().strip() or "—"
        res_txt = format_resume_text(self.resume_edit_box.toPlainText().strip())
        fam_txt = format_family_text(self.family_edit_box.toPlainText().strip())

        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        if self.cadre_id:
            c.execute("""UPDATE cadres SET 
                name=?, gender=?, birth=?, nation=?, native_place=?, id_card=?, phone=?, office_phone=?, office_address=?, emp_no=?,
                dept_post=?, level=?, cadre_rank=?, chu_category=?, charge_division=?,
                current_duty_date=?, same_level_date=?, career_history=?, remark=?,
                politic=?, party_date=?, work_date=?, title=?, title_date=?,
                full_edu_degree=?, full_edu_school=?, job_edu_degree=?, job_edu_school=?,
                first_edu_degree=?, first_edu_school=?, first_edu_major=?, first_edu_date=?,
                highest_edu_degree=?, highest_edu_degree_title=?, highest_edu_school=?, highest_edu_major=?, highest_edu_date=?,
                cadre_identity=?, post_category=?, discipline_type=?, discipline_date=?, discipline_desc=?,
                resume_text=?, family_relations=? WHERE id=?""", (
                name, self.fields["gender"].text().strip(), self.fields["birth"].text().strip(),
                self.fields["nation"].text().strip(), self.fields["native_place"].text().strip(),
                self.fields["id_card"].text().strip(), self.fields["phone"].text().strip(),
                self.fields["office_phone"].text().strip(), self.fields["office_address"].text().strip(),
                self.fields["emp_no"].text().strip(), dept_post, self.fields["level"].text().strip(),
                c_rank, chu_cat, charge_div,
                self.fields["current_duty_date"].text().strip(), self.fields["same_level_date"].text().strip(),
                self.fields["career_history"].text().strip(), self.fields["remark"].text().strip(),
                self.fields["politic"].text().strip(), self.fields["party_date"].text().strip(),
                self.fields["work_date"].text().strip(), self.fields["title"].text().strip(),
                self.fields["title_date"].text().strip(), self.fields["full_edu_degree"].text().strip(),
                self.fields["full_edu_school"].text().strip(), self.fields["job_edu_degree"].text().strip(),
                self.fields["job_edu_school"].text().strip(),
                self.fields["first_edu_degree"].text().strip(), self.fields["first_edu_school"].text().strip(),
                self.fields["first_edu_major"].text().strip(), self.fields["first_edu_date"].text().strip(),
                self.fields["highest_edu_degree"].text().strip(), self.fields["highest_edu_degree_title"].text().strip(),
                self.fields["highest_edu_school"].text().strip(), self.fields["highest_edu_major"].text().strip(),
                self.fields["highest_edu_date"].text().strip(),
                ident, post_cat,
                disc_t, disc_d, disc_desc, res_txt, fam_txt, self.cadre_id
            ))
        else:
            c.execute("""INSERT INTO cadres (
                name, gender, birth, nation, native_place, id_card, phone, office_phone, office_address, emp_no,
                dept_post, level, cadre_rank, chu_category, charge_division,
                current_duty_date, same_level_date, career_history, remark,
                politic, party_date, work_date, title, title_date,
                full_edu_degree, full_edu_school, job_edu_degree, job_edu_school,
                first_edu_degree, first_edu_school, first_edu_major, first_edu_date,
                highest_edu_degree, highest_edu_degree_title, highest_edu_school, highest_edu_major, highest_edu_date,
                cadre_identity, post_category, discipline_type, discipline_date, discipline_desc,
                resume_text, family_relations, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '在岗')""", (
                name, self.fields["gender"].text().strip(), self.fields["birth"].text().strip(),
                self.fields["nation"].text().strip(), self.fields["native_place"].text().strip(),
                self.fields["id_card"].text().strip(), self.fields["phone"].text().strip(),
                self.fields["office_phone"].text().strip(), self.fields["office_address"].text().strip(),
                self.fields["emp_no"].text().strip(), dept_post, self.fields["level"].text().strip(),
                c_rank, chu_cat, charge_div,
                self.fields["current_duty_date"].text().strip(), self.fields["same_level_date"].text().strip(),
                self.fields["career_history"].text().strip(), self.fields["remark"].text().strip(),
                self.fields["politic"].text().strip(), self.fields["party_date"].text().strip(),
                self.fields["work_date"].text().strip(), self.fields["title"].text().strip(),
                self.fields["title_date"].text().strip(), self.fields["full_edu_degree"].text().strip(),
                self.fields["full_edu_school"].text().strip(), self.fields["job_edu_degree"].text().strip(),
                self.fields["job_edu_school"].text().strip(),
                self.fields["first_edu_degree"].text().strip(), self.fields["first_edu_school"].text().strip(),
                self.fields["first_edu_major"].text().strip(), self.fields["first_edu_date"].text().strip(),
                self.fields["highest_edu_degree"].text().strip(), self.fields["highest_edu_degree_title"].text().strip(),
                self.fields["highest_edu_school"].text().strip(), self.fields["highest_edu_major"].text().strip(),
                self.fields["highest_edu_date"].text().strip(),
                ident, post_cat,
                disc_t, disc_d, disc_desc, res_txt, fam_txt
            ))
        conn.commit()
        conn.close()

        QMessageBox.information(self, "成功", "档案已成功保存入库（处分监督与规范简历已持久化更新）！")
        if self.on_saved: self.on_saved()
        self.accept()

# ==================== 离岗办理对话框 ====================

class CadreLeaveDialog(QDialog):
    def __init__(self, parent, cadre_id, name, cadre_rank="科级", on_finished=None):
        super().__init__(parent)
        self.cadre_id = cadre_id
        self.name = name
        self.cadre_rank = cadre_rank
        self.on_finished = on_finished
        self.setWindowTitle(f"办理干部调离/离退 —— {name}")
        self.resize(460, 360)
        self.setStyleSheet("""
            QDialog {
                background-color: #FFFFFF;
                font-family: -apple-system, "PingFang SC", "Microsoft YaHei UI", sans-serif;
            }
            QLabel {
                color: #1E293B;
                font-size: 13px;
            }
            QLineEdit {
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 4px 8px;
                background-color: #FFFFFF;
                color: #0F172A;
                font-size: 13px;
                min-height: 24px;
            }
            QLineEdit:focus {
                border: 1.5px solid #2563EB;
            }
            QRadioButton {
                color: #334155;
                font-size: 13px;
                spacing: 8px;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(10)
        
        lbl_head = QLabel(f"正在办理干部【{name}】（{cadre_rank}）变动归档：")
        lbl_head.setStyleSheet("font-weight: bold; color: #0F172A; font-size: 14px;")
        layout.addWidget(lbl_head)

        self.btn_grp = QButtonGroup(self)
        reasons = [("到达年龄退休", "退休"), ("退居二线 / 退出领导岗位", "退二线"), ("辞职 / 调离", "调离")]
        for text, val in reasons:
            rb = QRadioButton(text)
            rb.setProperty("val", val)
            self.btn_grp.addButton(rb)
            layout.addWidget(rb)
        self.btn_grp.buttons()[0].setChecked(True)

        layout.addWidget(QLabel("生效/批复年月 (如2024年10月):"))
        self.date_input = QLineEdit(datetime.datetime.now().strftime("%Y年%m月"))
        layout.addWidget(self.date_input)

        layout.addWidget(QLabel("去向/说明:"))
        self.remark_input = QLineEdit()
        self.remark_input.setPlaceholderText("可填写调往单位、退休批文号等")
        layout.addWidget(self.remark_input)

        bot = QHBoxLayout()
        btn_ok = QPushButton(" 确认办理归档 ")
        btn_ok.setFixedHeight(32)
        btn_ok.setStyleSheet("background-color: #DC2626; color: white; border-radius: 6px; padding: 0 16px; font-weight: bold;")
        btn_ok.clicked.connect(self.confirm)
        
        btn_cancel = QPushButton(" 取消 ")
        btn_cancel.setFixedHeight(32)
        btn_cancel.setStyleSheet("background-color: #F1F5F9; color: #475569; border: 1px solid #CBD5E1; border-radius: 6px; padding: 0 16px;")
        btn_cancel.clicked.connect(self.reject)
        
        bot.addStretch(); bot.addWidget(btn_ok); bot.addWidget(btn_cancel)
        layout.addLayout(bot)

    def confirm(self):
        reason = self.btn_grp.checkedButton().property("val")
        date_str = self.date_input.text().strip()
        rmk = self.remark_input.text().strip()

        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("UPDATE cadres SET status=?, leave_reason=?, leave_date=?, remark=? WHERE id=?",
                  (reason, reason, date_str, rmk, self.cadre_id))
        conn.commit()
        conn.close()

        QMessageBox.information(self, "办理成功", f"干部【{self.name}】已成功归档入历史库！")
        if self.on_finished: self.on_finished()
        self.accept()

# ==================== 花名册底表绑定对话框 (增量安全保护) ====================

class ManualBindRosterDialog(QDialog):
    def __init__(self, parent, cadre_id, name, on_bound=None):
        super().__init__(parent)
        self.cadre_id = cadre_id
        self.name = name
        self.on_bound = on_bound
        self.setWindowTitle(f"手动定向绑定人员底表 —— 【{name}】")
        self.resize(880, 520)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"正在为目标干部【{name}】检索《人员花名册》底表匹配项："))

        self.table = QTableWidget()
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels(["员工编号", "姓名", "所属科室", "性别", "出生日期", "编制身份", "身份证号", "移动电话", "第一学历学校"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.table)

        bot = QHBoxLayout()
        btn_ok = QPushButton(" 🎯 确认绑定选中的花名册档案 "); btn_ok.setStyleSheet("background-color: #2563EB; color: white; padding: 6px 14px; font-weight: bold;")
        btn_ok.clicked.connect(self.confirm_bind)
        btn_cancel = QPushButton(" 取消 "); btn_cancel.clicked.connect(self.reject)
        bot.addStretch(); bot.addWidget(btn_ok); bot.addWidget(btn_cancel)
        layout.addLayout(bot)

        self.load_candidates()

    def load_candidates(self):
        cands = find_roster_data_by_name(self.name)
        self.candidates = cands
        self.table.setRowCount(len(cands))
        for r, item in enumerate(cands):
            self.table.setItem(r, 0, QTableWidgetItem(item.get("emp_no", "")))
            self.table.setItem(r, 1, QTableWidgetItem(item.get("name", "")))
            self.table.setItem(r, 2, QTableWidgetItem(item.get("dept", "")))
            self.table.setItem(r, 3, QTableWidgetItem(item.get("gender", "")))
            self.table.setItem(r, 4, QTableWidgetItem(item.get("birth", "")))
            self.table.setItem(r, 5, QTableWidgetItem(item.get("identity", "") or "—"))
            self.table.setItem(r, 6, QTableWidgetItem(item.get("id_card", "")))
            self.table.setItem(r, 7, QTableWidgetItem(item.get("phone", "")))
            self.table.setItem(r, 8, QTableWidgetItem(item.get("first_edu_school", "")))

    def confirm_bind(self):
        row = self.table.currentRow()
        if row < 0 or row >= len(self.candidates):
            QMessageBox.warning(self, "提示", "请先在上方列表中选中一条记录！")
            return
        target = self.candidates[row]

        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()

        # 【核心保护机制】：先读取数据库中已有的电话与身份证，绝对防止被空值抹除
        c.execute("SELECT phone, id_card, cadre_identity FROM cadres WHERE id = ?", (self.cadre_id,))
        cur_row = c.fetchone()
        cur_phone = cur_row[0] if cur_row else ""
        cur_id_card = cur_row[1] if cur_row else ""
        cur_identity = cur_row[2] if cur_row else ""

        new_phone = target.get("phone", "")
        final_phone = new_phone if is_valid_text(new_phone) else cur_phone

        new_id = target.get("id_card", "")
        final_id_card = new_id if is_valid_text(new_id) else cur_id_card

        new_identity = target.get("identity", "")
        final_identity = new_identity if is_valid_text(new_identity) else cur_identity

        c.execute("""UPDATE cadres SET emp_no=?, id_card=?, phone=?, cadre_identity=?,
                     first_edu_degree=?, first_edu_school=?, first_edu_major=?, first_edu_date=?,
                     highest_edu_degree=?, highest_edu_degree_title=?, highest_edu_school=?, highest_edu_major=?, highest_edu_date=?
                     WHERE id=?""", (
            target.get("emp_no", ""), final_id_card, final_phone, final_identity,
            target.get("first_edu_degree", ""), target.get("first_edu_school", ""), target.get("first_edu_major", ""), target.get("first_edu_date", ""),
            target.get("highest_edu_degree", ""), target.get("highest_edu_degree_title", ""), target.get("highest_edu_school", ""), target.get("highest_edu_major", ""), target.get("highest_edu_date", ""),
            self.cadre_id
        ))
        conn.commit()
        conn.close()

        QMessageBox.information(self, "成功", f"干部【{self.name}】已成功绑定底表档案！\n\n• 编制身份已校准为：{final_identity}\n• 第一学历/最高学历已精准校准\n• 既有手机号已安全保留，未被抹除！")
        if self.on_bound: self.on_bound()
        self.accept()

# ==================== 审批表Word快速生成与导出对话框 ====================

class ExportApprovalFormDialog(QDialog):
    def __init__(self, parent, cadre_id, base_ym=None):
        super().__init__(parent)
        self.cadre_id = cadre_id
        self.base_ym = base_ym or datetime.datetime.now().strftime("%Y年%m月")
        self.load_data()
        self.setWindowTitle(f"📑 导出《干部任免审批表》呈报核准单 —— 【{self.d['name']}】")
        self.resize(880, 720)
        self.create_widgets()

    def load_data(self):
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        cols = ", ".join(ALL_CADRE_COLUMNS)
        c.execute(f"SELECT {cols} FROM cadres WHERE id = ?", (self.cadre_id,))
        row = c.fetchone()
        conn.close()
        self.d = {k: (str(v) if v is not None and str(v) != "None" else "—") for k, v in zip(ALL_CADRE_COLUMNS, row)}

    def create_widgets(self):
        layout = QVBoxLayout(self)

        top_lbl = QLabel(f"干部：<b>{self.d['name']}</b> &nbsp;&nbsp;|&nbsp;&nbsp; 现任职务：{self.d['dept_post']}")
        top_lbl.setFont(QFont("Microsoft YaHei UI", 10))
        top_lbl.setTextFormat(Qt.RichText)
        layout.addWidget(top_lbl)

        form_f = QFrame(); form_l = QVBoxLayout(form_f)
        h1 = QHBoxLayout(); h1.addWidget(QLabel("拟任职务 (*报批新任职务):")); self.target_post_input = QLineEdit(); h1.addWidget(self.target_post_input); form_l.addLayout(h1)
        h2 = QHBoxLayout(); h2.addWidget(QLabel("拟免职务 (可选填):")); self.target_dismiss_input = QLineEdit(); h2.addWidget(self.target_dismiss_input); form_l.addLayout(h2)
        layout.addWidget(form_f)

        layout.addWidget(QLabel("呈报简历内容核对 (系统已智能闭合最新任职):"))
        self.resume_text = QTextEdit(format_resume_text(self.d["resume_text"]))
        layout.addWidget(self.resume_text)

        layout.addWidget(QLabel("家庭主要成员及重要社会关系 (将回填入Word表格第2页):"))
        self.family_text = QTextEdit(format_family_text(self.d["family_relations"]))
        layout.addWidget(self.family_text)

        bot = QHBoxLayout()
        btn_gen = QPushButton(" 🚀 确认生成并打开Word审批表 "); btn_gen.setStyleSheet("background-color: #15803D; color: white; padding: 6px 16px; font-weight: bold;")
        btn_gen.clicked.connect(self.generate_word)
        btn_cancel = QPushButton(" 取消 "); btn_cancel.clicked.connect(self.reject)
        bot.addStretch(); bot.addWidget(btn_gen); bot.addWidget(btn_cancel)
        layout.addLayout(bot)

    def generate_word(self):
        target_post = self.target_post_input.text().strip()
        target_dismiss = self.target_dismiss_input.text().strip()

        tpl_path = get_approval_template_path()
        if not tpl_path:
            tpl_path, _ = QFileDialog.getOpenFileName(self, "请选择干部任免审批表Word模板", BASE_DIR, "Word 文档 (*.docx)")
            if not tpl_path: return

        out_fname = f"{self.d['name']}_干部任免审批表_{datetime.datetime.now().strftime('%Y%m%d')}.docx"
        out_path, _ = QFileDialog.getSaveFileName(self, "保存生成的干部任免审批表", os.path.join(BASE_DIR, out_fname), "Word 文档 (*.docx)")
        if not out_path: return

        fill_data = dict(self.d)
        fill_data["target_post"] = target_post
        fill_data["target_dismiss"] = target_dismiss
        fill_data["resume_text"] = format_resume_text(self.resume_text.toPlainText().strip())
        fill_data["family_relations"] = format_family_text(self.family_text.toPlainText().strip())

        try:
            fill_approval_docx(fill_data, tpl_path, out_path)
            QMessageBox.information(self, "生成成功", f"《干部任免审批表》已成功生成！\n路径：\n{out_path}")
            QDesktopServices.openUrl(QUrl.fromLocalFile(out_path))
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "生成失败", f"填报 Word 出错: {e}")

# ==================== 其它业务弹窗适配 (PySide6) ====================

class ChuAppointmentUploadDialog(QDialog):
    def __init__(self, parent, file_path, on_completed=None):
        super().__init__(parent)
        self.file_path = file_path
        self.on_completed = on_completed
        self.setWindowTitle("📜 处级任免公文上传解析核准")
        self.resize(800, 500)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"正在归档公文：{os.path.basename(file_path)}"))

        text = extract_text_from_pdf_with_ocr(file_path) if file_path.lower().endswith(".pdf") else ""
        if not text:
            try:
                d = docx.Document(file_path)
                text = "\n".join([p.text for p in d.paragraphs])
            except Exception: pass

        self.preview = QTextEdit(text)
        layout.addWidget(self.preview)

        bot = QHBoxLayout()
        btn_close = QPushButton(" 完成并关闭 "); btn_close.clicked.connect(self.accept)
        bot.addStretch(); bot.addWidget(btn_close)
        layout.addLayout(bot)
        if self.on_completed: self.on_completed()

class BatchAppointmentSyncDialog(QDialog):
    def __init__(self, parent, file_path, on_completed=None):
        super().__init__(parent)
        self.file_path = file_path
        self.on_completed = on_completed
        self.setWindowTitle("📜 干部换届任免调整发文核准")
        self.resize(780, 480)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"公文文件：{os.path.basename(file_path)}"))
        self.info = QLabel("系统已成功解析任免公文。")
        layout.addWidget(self.info)
        btn = QPushButton("确定"); btn.clicked.connect(self.accept)
        layout.addWidget(btn)
        if self.on_completed: self.on_completed()

class DisciplineDocSyncDialog(QDialog):
    def __init__(self, parent, file_path, on_completed=None):
        super().__init__(parent)
        self.setWindowTitle("⚖️ 纪律处分决定归档")
        self.resize(600, 380)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"已归档文件：{os.path.basename(file_path)}"))
        btn = QPushButton("确定"); btn.clicked.connect(self.accept)
        layout.addWidget(btn)
        if on_completed: on_completed()

# ==================== 干部队伍总览 · 党委常委会决策舱 (支持分轨切换) ====================

class CadreStatsDashboard(QDialog):
    """
    对标图2（AIHR组织总览）：
    深邃黑曜石科技大屏，支持【全院队伍总览】、【党政管理序列】、【业务管理序列】三轨一键无缝切换
    """
    def __init__(self, parent, base_ym):
        super().__init__(parent)
        self.base_ym = base_ym
        self.setWindowTitle(f"干部队伍总览 · 党委常委会决策舱 (基准：{base_ym})")
        self.resize(1280, 840)

        self.setStyleSheet("""
            QDialog {
                background-color: #0B0F19;
                color: #F8FAFC;
                font-family: -apple-system, "PingFang SC", "Microsoft YaHei UI", sans-serif;
            }
        """)

        self.current_track = "all"  # "all", "admin", "clinic"
        self.calculate_all_stats()

        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(22, 16, 22, 16)
        self.main_layout.setSpacing(12)

        # 1. 顶栏标题、分轨切换胶囊与导出按钮
        header_f = QFrame()
        header_l = QHBoxLayout(header_f)
        header_l.setContentsMargins(0, 0, 0, 0)

        t_box = QVBoxLayout()
        sys_t = QLabel("科级干部队伍总览决策舱")
        sys_t.setFont(QFont("Microsoft YaHei UI", 14, QFont.Bold))
        sys_t.setStyleSheet("color: #F8FAFC;")
        self.sys_sub = QLabel("全息穿透：第一屏总览规模 · 梯队活力 · 学历学衔 · 结构信号")
        self.sys_sub.setFont(QFont("Microsoft YaHei UI", 9))
        self.sys_sub.setStyleSheet("color: #64748B;")
        t_box.addWidget(sys_t)
        t_box.addWidget(self.sys_sub)
        header_l.addLayout(t_box)

        header_l.addStretch()

        # 分轨切换胶囊按钮组
        track_switch_f = QFrame()
        track_switch_f.setStyleSheet("background-color: #111827; border-radius: 8px; border: 1px solid #1E293B; padding: 2px;")
        ts_l = QHBoxLayout(track_switch_f)
        ts_l.setContentsMargins(4, 2, 4, 2)
        ts_l.setSpacing(4)

        self.track_btns = {}
        tracks_def = [
            ("all", "🏛️ 全员总览"),
            ("admin", "📋 党政管理序列 (行政专职)"),
            ("clinic", "🩺 业务管理序列 (专业“双肩挑”)")
        ]

        for k, text in tracks_def:
            btn = QPushButton(text)
            btn.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
            btn.setFixedHeight(30)
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda chk, key=k: self.switch_track(key))
            self.track_btns[k] = btn
            ts_l.addWidget(btn)

        header_l.addWidget(track_switch_f)
        header_l.addSpacing(10)

        btn_exp = QPushButton("📥 导出全景报告 (Excel)")
        btn_exp.setFixedHeight(34)
        btn_exp.setCursor(Qt.PointingHandCursor)
        btn_exp.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #059669, stop:1 #10B981);
                color: #FFFFFF; font-weight: bold; border-radius: 6px; padding: 0 16px; border: none;
            }
            QPushButton:hover { background: #047857; }
        """)
        btn_exp.clicked.connect(self.export_excel)
        header_l.addWidget(btn_exp)

        self.main_layout.addWidget(header_f)

        # 动态内容区域容器
        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout(self.content_widget)
        self.content_layout.setContentsMargins(0, 0, 0, 0)
        self.content_layout.setSpacing(12)
        self.main_layout.addWidget(self.content_widget, 1)

        # 底部关闭按钮
        bot_bar = QHBoxLayout()
        btn_close = QPushButton("关闭决策舱")
        btn_close.setFixedHeight(30)
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setStyleSheet("background-color: #1E293B; color: #94A3B8; border: 1px solid #334155; border-radius: 6px; padding: 0 16px;")
        btn_close.clicked.connect(self.accept)
        bot_bar.addStretch()
        bot_bar.addWidget(btn_close)
        self.main_layout.addLayout(bot_bar)

        self.render_dashboard()
        self.exec()

    def update_track_button_styles(self):
        for k, btn in self.track_btns.items():
            if k == self.current_track:
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: #2563EB;
                        color: #FFFFFF;
                        border-radius: 6px;
                        padding: 0 12px;
                        border: none;
                    }
                """)
            else:
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: transparent;
                        color: #94A3B8;
                        border-radius: 6px;
                        padding: 0 12px;
                        border: none;
                    }
                    QPushButton:hover {
                        background-color: #1E293B;
                        color: #F8FAFC;
                    }
                """)

    def switch_track(self, track_key):
        self.current_track = track_key
        self.render_dashboard()

    def render_dashboard(self):
        self.update_track_button_styles()

        # 选定当前展示数据集
        if self.current_track == "admin":
            s = self.stats_admin
            track_name = "党政管理序列 (专职行政团队)"
            self.sys_sub.setText("专注呈现：行政中层队伍体量 · 专职干部年轻化 · 学历层次 · 编制配置")
        elif self.current_track == "clinic":
            s = self.stats_clinic
            track_name = "业务管理序列 (业务部门/专业技术/双肩挑)"
            self.sys_sub.setText("专注呈现：业务内设负责人 · 高级职称引领力 · 硕博高知化 · 专家骨干梯队")
        else:
            s = self.stats_all
            track_name = "中层干部队伍总体大盘"
            self.sys_sub.setText("全息穿透：第一屏总览规模 · 梯队活力 · 学历学衔 · 结构信号")

        # 清空原动态区域
        while self.content_layout.count():
            item = self.content_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        t_safe = s['total'] if s['total'] > 0 else 1

        # 1. 顶部 8 大量化 KPI 胶囊卡片
        kpi_row = QHBoxLayout()
        kpi_row.setSpacing(10)

        kpis = [
            (f"{track_name[:4]}人数", f"{s['total']}", "人", "当前序列队伍规模", "#38BDF8"),
            ("正科级干部", f"{s['chief']}", f"占比 {round(s['chief']/t_safe*100, 1)}%", "骨干中坚力量", "#818CF8"),
            ("副科级干部", f"{s['sub']}", f"占比 {round(s['sub']/t_safe*100, 1)}%", "选拔储备梯队", "#34D399"),
            ("平均年龄", f"{s['avg_age']}", "岁", "梯队年轻化指数", "#F59E0B"),
            ("40岁以下骨干", f"{s['young']}", f"占比 {round(s['young']/t_safe*100, 1)}%", "青年后备先锋", "#EC4899"),
            ("硕博高学历", f"{s['doc']+s['master']}", f"占比 {round((s['doc']+s['master'])/t_safe*100, 1)}%", "高知化学力深度", "#A78BFA"),
            ("高级职称专家", f"{s['high_senior']+s['assoc_senior']}", f"占比 {round((s['high_senior']+s['assoc_senior'])/t_safe*100, 1)}%", "学术领头人密度", "#F97316"),
            ("中共党员比例", f"{s['cpc']}", f"占比 {round(s['cpc']/t_safe*100, 1)}%", "政治把关核心", "#EF4444"),
        ]

        for title, val, unit_or_pct, desc, color in kpis:
            card = QFrame()
            card.setStyleSheet("background-color: #111827; border: 1px solid #1E293B; border-radius: 8px;")
            card_l = QVBoxLayout(card)
            card_l.setContentsMargins(12, 10, 12, 10)
            card_l.setSpacing(2)

            t_l = QLabel(title)
            t_l.setFont(QFont("Microsoft YaHei UI", 8))
            t_l.setStyleSheet("color: #94A3B8;")

            v_l = QLabel(f"<span style='font-size:16pt; font-weight:bold; color:{color};'>{val}</span> <span style='font-size:8pt; color:#64748B;'>{unit_or_pct}</span>")
            v_l.setTextFormat(Qt.RichText)

            d_l = QLabel(desc)
            d_l.setFont(QFont("Microsoft YaHei UI", 7))
            d_l.setStyleSheet("color: #475569;")

            card_l.addWidget(t_l)
            card_l.addWidget(v_l)
            card_l.addWidget(d_l)
            kpi_row.addWidget(card)

        self.content_layout.addLayout(kpi_row)

        # 2. 中间区域：四大核心图表 + 右侧结构信号
        mid_layout = QHBoxLayout()
        mid_layout.setSpacing(12)

        charts_grid_widget = QWidget()
        charts_grid = QGridLayout(charts_grid_widget)
        charts_grid.setContentsMargins(0, 0, 0, 0)
        charts_grid.setSpacing(12)

        # ① 年龄梯队
        age_card = self.create_dark_chart_card("① 年龄梯队分布 (老中青结合度)", s['age_groups'], t_safe, "#38BDF8")
        charts_grid.addWidget(age_card, 0, 0)

        # ② 学历层次
        edu_data = {
            "博士研究生": s['doc'],
            "硕士研究生": s['master'],
            "大学本科及其他": s['bachelor']
        }
        edu_card = self.create_dark_chart_card("② 学历学位层次结构", edu_data, t_safe, "#818CF8")
        charts_grid.addWidget(edu_card, 0, 1)

        # ③ 专业技术职称
        title_data = {
            "正高级职称 (主任医/教授/研究员)": s['high_senior'],
            "副高级职称 (副主任医/副教授)": s['assoc_senior'],
            "中级职称 (主治医/主管护师/讲师)": s['mid_title'],
            "初级及以下 (初级师/助理/见习)": s['junior_title']
        }
        title_card = self.create_dark_chart_card("③ 专业技术职务层级", title_data, t_safe, "#34D399")
        charts_grid.addWidget(title_card, 1, 0)

        # ④ 编制形态
        ident_data = {
            "在编 (上级/正式编制)": s['ident_xiao'],
            "单位聘用 / 编制": s['ident_yuan'],
            "人事代理 / 合同聘用": s['ident_daili']
        }
        ident_card = self.create_dark_chart_card("④ 用工形态与编制构成", ident_data, t_safe, "#F59E0B")
        charts_grid.addWidget(ident_card, 1, 1)

        mid_layout.addWidget(charts_grid_widget, 7)

        # 右侧：五大结构信号专栏
        signals_card = QFrame()
        signals_card.setFixedWidth(310)
        signals_card.setStyleSheet("background-color: #111827; border: 1px solid #1E293B; border-radius: 8px;")
        sig_l = QVBoxLayout(signals_card)
        sig_l.setContentsMargins(14, 12, 14, 12)
        sig_l.setSpacing(10)

        sig_head = QLabel(f"⚡ 结构信号 · {track_name[:4]}")
        sig_head.setFont(QFont("Microsoft YaHei UI", 10, QFont.Bold))
        sig_head.setStyleSheet("color: #38BDF8; border: none;")
        sig_l.addWidget(sig_head)

        signals_items = [
            ("01", "青年人才储备度", f"40岁及以下青年中层 {s['young']} 人，占比达 {round(s['young']/t_safe*100, 1)}%，后备骨干充足。", "#38BDF8"),
            ("02", "学术领航硬实力", f"正高专家 {s['high_senior']} 人、副高 {s['assoc_senior']} 人，高级职称占比达 {round((s['high_senior']+s['assoc_senior'])/t_safe*100, 1)}%，中级骨干 {s['mid_title']} 人。", "#818CF8"),
            ("03", "高知化学力深度", f"博士 {s['doc']} 人、硕士 {s['master']} 人，硕博总占比达 {round((s['doc']+s['master'])/t_safe*100, 1)}%，队伍学术根基深厚。", "#34D399"),
            ("04", "职级配置梯队", f"正科 {s['chief']} 人（占比{round(s['chief']/t_safe*100, 1)}%），副科 {s['sub']} 人，正副结构保持良性循环。", "#F59E0B"),
            ("05", "政治活力与巾帼", f"中共党员占比 {round(s['cpc']/t_safe*100, 1)}%，女性领导干部占比 {round(s['female']/t_safe*100, 1)}%，展现多元凝聚力。", "#EC4899"),
        ]

        for num, stitle, sdesc, scolor in signals_items:
            s_box = QFrame()
            s_box.setStyleSheet(f"background-color: #0F172A; border-left: 3px solid {scolor}; border-radius: 6px;")
            s_bl = QVBoxLayout(s_box)
            s_bl.setContentsMargins(10, 6, 10, 6)
            s_bl.setSpacing(2)

            t_row = QHBoxLayout()
            n_lbl = QLabel(num)
            n_lbl.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
            n_lbl.setStyleSheet(f"color: {scolor};")
            title_l = QLabel(stitle)
            title_l.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
            title_l.setStyleSheet("color: #F1F5F9;")
            t_row.addWidget(n_lbl)
            t_row.addWidget(title_l)
            t_row.addStretch()

            desc_l = QLabel(sdesc)
            desc_l.setFont(QFont("Microsoft YaHei UI", 8))
            desc_l.setStyleSheet("color: #94A3B8; line-height: 1.3;")
            desc_l.setWordWrap(True)

            s_bl.addLayout(t_row)
            s_bl.addWidget(desc_l)
            sig_l.addWidget(s_box)

        sig_l.addStretch()
        mid_layout.addWidget(signals_card, 3)

        self.content_layout.addLayout(mid_layout, 1)

    def create_dark_chart_card(self, title, data_dict, total_safe, color_hex):
        card = QFrame()
        card.setStyleSheet("background-color: #111827; border: 1px solid #1E293B; border-radius: 8px;")
        l = QVBoxLayout(card)
        l.setContentsMargins(14, 10, 14, 10)
        l.setSpacing(6)

        t_lbl = QLabel(title)
        t_lbl.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
        t_lbl.setStyleSheet("color: #F8FAFC;")
        l.addWidget(t_lbl)

        for label, count in data_dict.items():
            pct = round(count / total_safe * 100, 1)
            row = QHBoxLayout()
            row.setSpacing(8)

            txt_l = QLabel(label)
            txt_l.setFont(QFont("Microsoft YaHei UI", 8))
            txt_l.setFixedWidth(140)
            txt_l.setStyleSheet("color: #94A3B8;")

            bar = QProgressBar()
            bar.setFixedHeight(8)
            bar.setRange(0, 100)
            bar.setValue(int(pct))
            bar.setTextVisible(False)
            bar.setStyleSheet(f"""
                QProgressBar {{
                    background-color: #1E293B;
                    border-radius: 4px;
                    border: none;
                }}
                QProgressBar::chunk {{
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {color_hex}88, stop:1 {color_hex});
                    border-radius: 4px;
                }}
            """)

            val_l = QLabel(f"{count}人 ({pct}%)")
            val_l.setFont(QFont("Microsoft YaHei UI", 8))
            val_l.setFixedWidth(75)
            val_l.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            val_l.setStyleSheet("color: #F1F5F9;")

            row.addWidget(txt_l)
            row.addWidget(bar, 1)
            row.addWidget(val_l)
            l.addLayout(row)

        l.addStretch()
        return card

    def calculate_single_group(self, rows_subset):
        total = len(rows_subset)
        chief = sum(1 for r in rows_subset if "正科" in str(r[2]))
        sub = sum(1 for r in rows_subset if "副科" in str(r[2]))
        female = sum(1 for r in rows_subset if "女" in str(r[0]))
        cpc = sum(1 for r in rows_subset if "党员" in str(r[6]))

        age_groups = {
            "35岁及以下": 0, "36—40岁": 0, "41—45岁": 0, "46—50岁": 0, "51—55岁": 0, "56岁及以上": 0
        }
        valid_ages = []
        young_cnt = 0
        for r in rows_subset:
            age = calc_age(r[1], self.base_ym)
            if age is not None:
                valid_ages.append(age)
                if age <= 40: young_cnt += 1
                if age <= 35: age_groups["35岁及以下"] += 1
                elif age <= 40: age_groups["36—40岁"] += 1
                elif age <= 45: age_groups["41—45岁"] += 1
                elif age <= 50: age_groups["46—50岁"] += 1
                elif age <= 55: age_groups["51—55岁"] += 1
                else: age_groups["56岁及以上"] += 1

        avg_age = round(sum(valid_ages) / len(valid_ages), 1) if valid_ages else 0

        doc, master, bachelor = 0, 0, 0
        for r in rows_subset:
            edu_str = str(r[3] or "") + str(r[4] or "")
            if "博士" in edu_str: doc += 1
            elif "硕士" in edu_str or "研究生" in edu_str: master += 1
            else: bachelor += 1

        high_senior, assoc_senior, mid_title, junior_title = 0, 0, 0, 0
        for r in rows_subset:
            t = str(r[5] or "")
            lvl = get_title_level(t)
            if lvl == "正高":
                high_senior += 1
            elif lvl == "副高":
                assoc_senior += 1
            elif lvl == "中级":
                mid_title += 1
            else:
                junior_title += 1

        ident_xiao, ident_yuan, ident_daili = 0, 0, 0
        for r in rows_subset:
            i = str(r[7] or "")
            if "校编" in i: ident_xiao += 1
            elif "院编" in i: ident_yuan += 1
            else: ident_daili += 1

        return {
            "total": total, "chief": chief, "sub": sub, "female": female, "cpc": cpc,
            "young": young_cnt, "avg_age": avg_age, "age_groups": age_groups,
            "doc": doc, "master": master, "bachelor": bachelor,
            "high_senior": high_senior, "assoc_senior": assoc_senior,
            "mid_title": mid_title, "junior_title": junior_title,
            "ident_xiao": ident_xiao, "ident_yuan": ident_yuan, "ident_daili": ident_daili
        }

    def calculate_all_stats(self):
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("""SELECT gender, birth, level, full_edu_degree, job_edu_degree, title, politic, cadre_identity, post_category 
                     FROM cadres WHERE status='在岗' AND (cadre_rank='科级' OR cadre_rank IS NULL OR cadre_rank='')""")
        all_rows = c.fetchall()
        conn.close()

        # 分别预先计算全院总体、党政管理序列、业务管理序列三套独立数据
        self.stats_all = self.calculate_single_group(all_rows)
        self.stats_admin = self.calculate_single_group([r for r in all_rows if "业务" not in str(r[8] or "")])
        self.stats_clinic = self.calculate_single_group([r for r in all_rows if "业务" in str(r[8] or "")])

    def export_excel(self):
        save_path, _ = QFileDialog.getSaveFileName(self, "导出常委会全景报告", os.path.join(BASE_DIR, f"常委会科级队伍结构分析报告_{self.base_ym}.xlsx"), "Excel 表格 (*.xlsx)")
        if not save_path: return
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "三大板块结构对比"
            ws.append(["科级干部队伍结构分析报告（常委会汇报专用）"])
            ws.append([f"统计基准节点：{self.base_ym}", f"在岗总人数：{self.stats_all['total']} 人"])
            ws.append([])
            ws.append(["指标分类", "具体统计维度", "全院总体大盘", "党政管理序列 (专职行政)", "业务管理序列 (临床双肩挑)"])

            def fmt(cnt, tot):
                return f"{cnt}人 ({round(cnt/tot*100, 1)}%)" if tot > 0 else "0人 (0%)"

            sa, sad, sc = self.stats_all, self.stats_admin, self.stats_clinic
            ta, tad, tc = sa['total'] or 1, sad['total'] or 1, sc['total'] or 1

            rows = [
                ("队伍规模", "在岗总人数", f"{sa['total']} 人", f"{sad['total']} 人", f"{sc['total']} 人"),
                ("队伍规模", "平均年龄 (岁)", f"{sa['avg_age']} 岁", f"{sad['avg_age']} 岁", f"{sc['avg_age']} 岁"),
                ("职级结构", "正科级干部", fmt(sa['chief'], ta), fmt(sad['chief'], tad), fmt(sc['chief'], tc)),
                ("职级结构", "副科级干部", fmt(sa['sub'], ta), fmt(sad['sub'], tad), fmt(sc['sub'], tc)),
                ("青年梯队", "40岁及以下骨干", fmt(sa['young'], ta), fmt(sad['young'], tad), fmt(sc['young'], tc)),
                ("性别党派", "女性干部", fmt(sa['female'], ta), fmt(sad['female'], tad), fmt(sc['female'], tc)),
                ("性别党派", "中共党员", fmt(sa['cpc'], ta), fmt(sad['cpc'], tad), fmt(sc['cpc'], tc)),
                ("学历层次", "博士研究生", fmt(sa['doc'], ta), fmt(sad['doc'], tad), fmt(sc['doc'], tc)),
                ("学历层次", "硕士研究生", fmt(sa['master'], ta), fmt(sad['master'], tad), fmt(sc['master'], tc)),
                ("专业职称", "正高级职称", fmt(sa['high_senior'], ta), fmt(sad['high_senior'], tad), fmt(sc['high_senior'], tc)),
                ("专业职称", "副高级职称", fmt(sa['assoc_senior'], ta), fmt(sad['assoc_senior'], tad), fmt(sc['assoc_senior'], tc)),
                ("编制形态", "在编 (正式编制)", fmt(sa['ident_xiao'], ta), fmt(sad['ident_xiao'], tad), fmt(sc['ident_xiao'], tc)),
                ("编制形态", "自聘 (单位聘用)", fmt(sa['ident_yuan'], ta), fmt(sad['ident_yuan'], tad), fmt(sc['ident_yuan'], tc)),
                ("编制形态", "人事代理", fmt(sa['ident_daili'], ta), fmt(sad['ident_daili'], tad), fmt(sc['ident_daili'], tc)),
            ]
            for r in rows:
                ws.append(list(r))

            ws.merge_cells("A1:E1")
            ws["A1"].font = Font(name="Microsoft YaHei", size=14, bold=True, color="1F497D")
            ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

            h_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
            for cell in ws[4]:
                cell.fill = h_fill
                cell.font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 24

            wb.save(save_path)
            QMessageBox.information(self, "导出成功", f"常委会汇报专用三大板块全景对比报告已导出至：\n{save_path}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"导出出错: {e}")

class CustomExportDialog(QDialog):
    def __init__(self, parent, cids, base_ym):
        super().__init__(parent)
        self.cids = cids
        self.base_ym = base_ym
        self.setWindowTitle("自定义字段导出Excel")
        self.resize(720, 500)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"当前选中待导出干部共 {len(cids)} 人，请勾选需要导出的字段："))

        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        container = QWidget(); grid = QGridLayout(container)

        self.checks = {}
        for idx, (code, label) in enumerate(EXPORTABLE_FIELDS):
            cb = QCheckBox(label)
            if code in ["emp_no", "name", "cadre_rank", "gender", "birth", "age", "dept_post", "level", "phone", "title", "full_edu_degree", "first_edu_school"]:
                cb.setChecked(True)
            self.checks[code] = cb
            grid.addWidget(cb, idx // 3, idx % 3)

        scroll.setWidget(container)
        layout.addWidget(scroll)

        bot = QHBoxLayout()
        btn_all = QPushButton("全选"); btn_all.clicked.connect(lambda: [c.setChecked(True) for c in self.checks.values()])
        btn_clear = QPushButton("清空"); btn_clear.clicked.connect(lambda: [c.setChecked(False) for c in self.checks.values()])
        btn_exp = QPushButton(" 🚀 导出Excel "); btn_exp.setStyleSheet("background-color: #047857; color: white; font-weight: bold; padding: 6px 14px;")
        btn_exp.clicked.connect(self.export_excel)
        btn_cancel = QPushButton("取消"); btn_cancel.clicked.connect(self.reject)

        bot.addWidget(btn_all); bot.addWidget(btn_clear); bot.addStretch(); bot.addWidget(btn_exp); bot.addWidget(btn_cancel)
        layout.addLayout(bot)

    def export_excel(self):
        selected_codes = [code for code, cb in self.checks.items() if cb.isChecked()]
        if not selected_codes:
            QMessageBox.warning(self, "提示", "请至少勾选一个导出字段！")
            return

        save_path, _ = QFileDialog.getSaveFileName(self, "保存定制花名册", os.path.join(BASE_DIR, f"干部定制花名册_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx"), "Excel 表格 (*.xlsx)")
        if not save_path: return

        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        cols = ", ".join(ALL_CADRE_COLUMNS)
        placeholders = ','.join(['?'] * len(self.cids))
        c.execute(f"SELECT {cols} FROM cadres WHERE id IN ({placeholders})", self.cids)
        rows = c.fetchall()
        conn.close()

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            headers = ["序号"] + [lbl for code, lbl in EXPORTABLE_FIELDS if code in selected_codes]
            ws.append(headers)

            for idx, r in enumerate(rows, start=1):
                d = dict(zip(ALL_CADRE_COLUMNS, r))
                age = calc_age(d["birth"], self.base_ym)
                d["age"] = f"{age}岁" if age is not None else "—"
                m_cur = calc_months_diff(d["current_duty_date"], self.base_ym)
                d["cur_len"] = format_months_to_str(m_cur)
                m_same = calc_months_diff(d["same_level_date"], self.base_ym)
                d["same_len"] = format_months_to_str(m_same)
                _, d_stat, _ = check_discipline_impact(d["discipline_type"], d["discipline_date"], self.base_ym)
                d["discipline_status"] = d_stat
                alerts = analyze_same_unit_relations(d["family_relations"], d.get("dept_post", ""))
                d["same_unit_family"] = "；".join(alerts) if alerts else "无"
                d["doc_trajectory"] = "—"

                row_vals = [idx]
                for code in selected_codes:
                    val = d.get(code, "—")
                    row_vals.append(str(val) if is_valid_text(val) else "—")
                ws.append(row_vals)

            h_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
            for cell in ws[1]:
                cell.fill = h_fill
                cell.font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center", vertical="center")

            wb.save(save_path)
            QMessageBox.information(self, "导出成功", f"定制花名册已导出至：\n{save_path}")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"导出出错: {e}")

SelectApprovalDialog = ManualMatchApprovalDialog


# ==================== 变更比对与增量确认机制 ====================

class ImportDiffConfirmDialog(QDialog):
    """
    Excel名册导入前差异比对与增量确认对话框
    """
    def __init__(self, parent, diff_items, on_confirm_callback=None):
        super().__init__(parent)
        self.setWindowTitle("名册导入数据差异比对与增量确认")
        self.resize(1000, 680)
        self.diff_items = diff_items
        self.on_confirm_callback = on_confirm_callback
        self.filtered_items = list(diff_items)

        self.init_ui()

    def init_ui(self):
        root_l = QVBoxLayout(self)
        root_l.setSpacing(12)
        root_l.setContentsMargins(16, 16, 16, 16)

        # 1. 顶部标题与说明
        top_box = QFrame()
        top_box.setStyleSheet("background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 10px;")
        top_l = QVBoxLayout(top_box)
        top_l.setContentsMargins(4, 4, 4, 4)
        top_l.setSpacing(4)

        t_lbl = QLabel("📋 导入名册与本地数据库差异比对结果")
        t_lbl.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        t_lbl.setStyleSheet("color: #0F172A;")

        d_lbl = QLabel("系统已在入库前自动比对所有干部记录。请核验新增人员及关键变更信息，确认无误后再执行增量写入，确保历史数据万无一失。")
        d_lbl.setFont(QFont("Microsoft YaHei UI", 9))
        d_lbl.setStyleSheet("color: #64748B;")
        d_lbl.setWordWrap(True)

        top_l.addWidget(t_lbl)
        top_l.addWidget(d_lbl)
        root_l.addWidget(top_box)

        # 2. 统计卡片与快速筛选
        cnt_new = sum(1 for it in self.diff_items if it.get("status") == "新增")
        cnt_mod = sum(1 for it in self.diff_items if it.get("status") == "变更")
        cnt_same = sum(1 for it in self.diff_items if it.get("status") == "无变化")
        total_cnt = len(self.diff_items)

        stat_bar = QHBoxLayout()
        stat_bar.setSpacing(10)

        card_total = QLabel(f"<b>总计比对记录</b><br><span style='font-size:16px; color:#1E3A8A;'>{total_cnt}</span> 条")
        card_new = QLabel(f"<b>🟢 拟新增干部</b><br><span style='font-size:16px; color:#16A34A;'>{cnt_new}</span> 人")
        card_mod = QLabel(f"<b>🟡 关键信息变更</b><br><span style='font-size:16px; color:#D97706;'>{cnt_mod}</span> 人")
        card_same = QLabel(f"<b>⚪ 完全一致</b><br><span style='font-size:16px; color:#64748B;'>{cnt_same}</span> 人")

        for c_lbl in [card_total, card_new, card_mod, card_same]:
            c_lbl.setFont(QFont("Microsoft YaHei UI", 9))
            c_lbl.setStyleSheet("background-color: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 6px; padding: 6px 12px;")
            c_lbl.setAlignment(Qt.AlignCenter)
            stat_bar.addWidget(c_lbl)

        stat_bar.addStretch()

        filter_lbl = QLabel("显示范围：")
        filter_lbl.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
        self.btn_filter_all = QRadioButton(f"全部({total_cnt})")
        self.btn_filter_new = QRadioButton(f"仅看新增({cnt_new})")
        self.btn_filter_mod = QRadioButton(f"仅看变更({cnt_mod})")
        self.btn_filter_same = QRadioButton(f"仅看一致({cnt_same})")
        self.btn_filter_all.setChecked(True)

        self.filter_group = QButtonGroup(self)
        for r_btn in [self.btn_filter_all, self.btn_filter_new, self.btn_filter_mod, self.btn_filter_same]:
            self.filter_group.addButton(r_btn)
            r_btn.setFont(QFont("Microsoft YaHei UI", 9))
            r_btn.toggled.connect(self.on_filter_changed)

        stat_bar.addWidget(filter_lbl)
        stat_bar.addWidget(self.btn_filter_all)
        stat_bar.addWidget(self.btn_filter_new)
        stat_bar.addWidget(self.btn_filter_mod)
        stat_bar.addWidget(self.btn_filter_same)

        root_l.addLayout(stat_bar)

        # 3. 差异明细表格
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["序号", "姓名", "管理层级", "比对结论", "现任/导入部门职务", "变更字段明细对比（旧值 ➔ 新值）"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setDefaultSectionSize(90)
        self.table.setColumnWidth(0, 50)
        self.table.setColumnWidth(1, 90)
        self.table.setColumnWidth(2, 80)
        self.table.setColumnWidth(3, 90)
        self.table.setColumnWidth(4, 220)

        self.table.setStyleSheet("""
            QTableWidget {
                background-color: #FFFFFF;
                alternate-background-color: #F8FAFC;
                gridline-color: #E2E8F0;
                font-size: 9pt;
            }
            QHeaderView::section {
                background-color: #F1F5F9;
                font-weight: bold;
                border: 1px solid #CBD5E1;
                padding: 4px;
            }
        """)
        self.table.setAlternatingRowColors(True)

        root_l.addWidget(self.table, 1)

        # 4. 底部动作栏
        bot_bar = QHBoxLayout()
        tip_lbl = QLabel("💡 提示：点击【确认增量写入】后系统将安全保存新增与变更记录；若有疑问可直接点击【取消】。")
        tip_lbl.setFont(QFont("Microsoft YaHei UI", 9))
        tip_lbl.setStyleSheet("color: #475569;")

        btn_confirm = QPushButton("✅ 确认增量写入数据库")
        btn_confirm.setFixedHeight(34)
        btn_confirm.setStyleSheet("background-color: #15803D; color: white; border-radius: 6px; padding: 0 18px; font-weight: bold;")
        btn_confirm.clicked.connect(self.confirm_import)

        btn_cancel = QPushButton("❌ 取消放弃")
        btn_cancel.setFixedHeight(34)
        btn_cancel.setStyleSheet("background-color: #F1F5F9; color: #475569; border: 1px solid #CBD5E1; border-radius: 6px; padding: 0 16px;")
        btn_cancel.clicked.connect(self.reject)

        bot_bar.addWidget(tip_lbl)
        bot_bar.addStretch()
        bot_bar.addWidget(btn_cancel)
        bot_bar.addWidget(btn_confirm)

        root_l.addLayout(bot_bar)

        self.populate_table()

    def on_filter_changed(self):
        if self.btn_filter_new.isChecked():
            self.filtered_items = [it for it in self.diff_items if it.get("status") == "新增"]
        elif self.btn_filter_mod.isChecked():
            self.filtered_items = [it for it in self.diff_items if it.get("status") == "变更"]
        elif self.btn_filter_same.isChecked():
            self.filtered_items = [it for it in self.diff_items if it.get("status") == "无变化"]
        else:
            self.filtered_items = list(self.diff_items)
        self.populate_table()

    def populate_table(self):
        self.table.setRowCount(len(self.filtered_items))
        for r_idx, it in enumerate(self.filtered_items):
            status = it.get("status", "无变化")
            changes = it.get("changes", [])
            diff_text = "； ".join([f"【{c[0]}】{c[1]} ➔ {c[2]}" for c in changes]) if changes else "数据完全一致，无须变更"

            status_item = QTableWidgetItem(status)
            if status == "新增":
                status_item.setForeground(QBrush(QColor("#16A34A")))
                status_item.setText("🟢 新增人员")
            elif status == "变更":
                status_item.setForeground(QBrush(QColor("#D97706")))
                status_item.setText(f"🟡 信息变更({len(changes)})")
            else:
                status_item.setForeground(QBrush(QColor("#64748B")))
                status_item.setText("⚪ 无变化")

            c0 = QTableWidgetItem(str(r_idx + 1))
            c0.setTextAlignment(Qt.AlignCenter)
            c1 = QTableWidgetItem(it.get("name", "—"))
            c1.setTextAlignment(Qt.AlignCenter)
            c2 = QTableWidgetItem(it.get("rank", "科级"))
            c2.setTextAlignment(Qt.AlignCenter)
            c4 = QTableWidgetItem(it.get("dept_post", "—"))
            c5 = QTableWidgetItem(diff_text)

            if status == "变更":
                c5.setForeground(QBrush(QColor("#B45309")))
                c5.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))

            self.table.setItem(r_idx, 0, c0)
            self.table.setItem(r_idx, 1, c1)
            self.table.setItem(r_idx, 2, c2)
            self.table.setItem(r_idx, 3, status_item)
            self.table.setItem(r_idx, 4, c4)
            self.table.setItem(r_idx, 5, c5)

    def confirm_import(self):
        if self.on_confirm_callback:
            try:
                self.on_confirm_callback(self.diff_items)
            except Exception as e:
                QMessageBox.critical(self, "写入错误", f"增量写入数据库过程发生错误: {e}")
                return
        self.accept()


# ==================== 亲属任职回避专项排查台账导出 ====================

def generate_avoidance_investigation_excel(base_ym, save_path):
    """
    一键生成符合组织人事与纪检监察巡视审计规范的
    《干部亲属任职回避专项排查台账.xlsx》
    """
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        SELECT id, name, gender, cadre_rank, level, dept_post, title, family_relations
        FROM cadres
        WHERE status != '已调离'
    """)
    cadres = c.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "干部亲属任职回避排查"

    # 1. 大标题与说明
    ws.merge_cells("A1:M1")
    title_cell = ws["A1"]
    title_cell.value = "干部亲属任职回避专项排查台账"
    title_cell.font = Font(name="Microsoft YaHei", size=16, bold=True, color="1E3A8A")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40

    ws.merge_cells("A2:M2")
    sub_cell = ws["A2"]
    sub_cell.value = f"排查基准年月：{base_ym}  |  政策法规依据：《事业单位人事管理回避规定》《党政领导干部选拔任用工作条例》  |  填报部门：党委组织部/纪律检查室"
    sub_cell.font = Font(name="Microsoft YaHei", size=10, italic=True, color="475569")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 24

    # 2. 表头
    headers = [
        "序号", "干部姓名", "性别", "管理层级", "现任职务", "职级/职称",
        "亲属关系", "亲属姓名", "亲属政治面貌", "亲属在院职务/岗位",
        "任职回避等级", "政策研判依据", "建议管控与回避措施"
    ]
    ws.append(headers)
    ws.row_dimensions[3].height = 28

    h_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    for col_idx in range(1, 14):
        cell = ws.cell(row=3, column=col_idx)
        cell.fill = h_fill
        cell.font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    red_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    yellow_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

    row_idx = 4
    record_count = 0

    key_org_words = ["党政办", "党委办", "院长办", "纪委", "纪检", "监察", "组织", "人事", "财务", "审计", "基建", "招投标", "药学", "设备"]

    for cd in cadres:
        cid, name, gender, cadre_rank, level, dept_post, title, family_raw = cd
        if not is_valid_text(family_raw):
            continue

        for line in family_raw.split("\n"):
            info = parse_family_relation_line(line)
            if not info:
                continue
            rel, fnm, fbrth, fpol, funit = info

            # 识别同单位任职
            if any(k in funit for k in ["第一附属医院", "一附院", "本院", "河南医药大学第一附属医院", "本单位"]):
                    # 研判回避级别
                    is_leader = (cadre_rank == "处级") or ("领导班子" in dept_post) or any(k in dept_post for k in ["院长", "书记", "处长", "主任", "局长"])
                    is_in_key_dept = any(k in dept_post for k in key_org_words) or any(k in funit for k in key_org_words)

                    if is_leader or is_in_key_dept:
                        level_tag = "🔴 直接回避风险(重点管控)"
                        legal_basis = "违反《事业单位人事管理回避规定》第六条：有夫妻、直系血亲关系的，不得在其中一方担任领导的单位从事组织、人事、纪检、财务、审计工作，或直接上下级任职。"
                        measure = "建议纳入人事调整预备方案：适时轮岗交流至非关联部门；在干部任免、奖惩考核、财务审批等公务活动中严格执行公务回避制度。"
                        cur_fill = red_fill
                    else:
                        level_tag = "🟡 一般协同关注(常态监管)"
                        legal_basis = "符合同单位任职备案要求，双方均处于普通业务部门/内设科室，无直接上下级行政隶属与关键制约关系。"
                        measure = "建档备查并纳入日常廉政台账监管；在涉及评先评优、职称评审、绩效分配时执行公务回避表决。"
                        cur_fill = yellow_fill

                    row_vals = [
                        record_count + 1, name, gender, cadre_rank or "科级",
                        dept_post, f"{level} / {title}",
                        rel, fnm, fpol, funit,
                        level_tag, legal_basis, measure
                    ]
                    ws.append(row_vals)
                    ws.row_dimensions[row_idx].height = 24

                    for col_idx in range(1, 14):
                        c_cell = ws.cell(row=row_idx, column=col_idx)
                        c_cell.font = Font(name="Microsoft YaHei", size=9)
                        c_cell.border = thin_border
                        if col_idx in [1, 2, 3, 4, 7, 8, 9, 11]:
                            c_cell.alignment = Alignment(horizontal="center", vertical="center")
                        else:
                            c_cell.alignment = Alignment(horizontal="left", vertical="center")
                        if col_idx == 11:
                            c_cell.fill = cur_fill
                            c_cell.font = Font(name="Microsoft YaHei", size=9, bold=True)

                    row_idx += 1
                    record_count += 1

    # 设置列宽
    col_widths = {
        "A": 8, "B": 12, "C": 8, "D": 12, "E": 24, "F": 18,
        "G": 12, "H": 12, "I": 14, "J": 26, "K": 26, "L": 38, "M": 38
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    wb.save(save_path)
    return record_count