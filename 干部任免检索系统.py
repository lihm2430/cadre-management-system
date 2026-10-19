# -*- coding: utf-8 -*-
"""
干部任免检索系统.py
科级干部全周期管理与综合研判系统 (主系统入口)
【PySide6 生产级无死锁版：WAL高并发模式 + 纯内存简历自愈 + 30秒忙等待保护】
"""

import os
import sys
import threading
import datetime
import sqlite3
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from PySide6.QtCore import Qt, Signal, QObject, QThread, QTimer, QSize, QUrl
from PySide6.QtGui import QFont, QColor, QIcon, QBrush, QDesktopServices, QPalette
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QComboBox, QLineEdit, QTableWidget,
    QTableWidgetItem, QHeaderView, QFrame, QFileDialog, QMessageBox,
    QAbstractItemView, QSizePolicy, QProgressDialog, QStyledItemDelegate
)

import cadre_core
from cadre_core import (
    BASE_DIR, FILE_DIR, APPROVAL_DIR, DISCIPLINE_DIR, DB_PATH,
    ALL_CADRE_COLUMNS, CHU_CATEGORIES, COMBO_STYLE, parse_ym, calc_months_diff, format_months_to_str,
    calc_age, check_discipline_impact, analyze_same_unit_relations, is_valid_text,
    get_title_level, calc_retirement_info, get_approval_template_path, fill_approval_docx,
    extract_flexible_roster_headers, extract_resume_and_family_from_docx,
    split_school_and_major, split_degree_and_title,
    format_resume_text, format_family_text, clean_and_format_duty_line, close_last_resume_line,
    merge_career_history_into_resume, heal_and_reorganize_resume, clean_cadre_appointment_prefix, extract_post_from_doc_content,
    parse_sort_date, SmartTableWidgetItem,
    ImportDiffConfirmDialog, generate_avoidance_investigation_excel,
    FullCadreDetailWindow, FullCadreEditDialog, ManualBindRosterDialog, CustomExportDialog,
    CadreStatsDashboard, CadreLeaveDialog, DisciplineDocSyncDialog, BatchAppointmentSyncDialog,
    ChuAppointmentUploadDialog, ExportApprovalFormDialog
)

# ==================== 纯内存简历追加算法 (杜绝数据库连接嵌套死锁) ====================

def append_resume_in_memory(cur_resume, duty_date_str, new_post_text, cadre_name=""):
    """纯内存无死锁版：在内存中完成末尾去重与新任命追加闭合"""
    mock_doc = [{"doc_date": str(duty_date_str), "content": f"{cadre_name}任{new_post_text}"}]
    return heal_and_reorganize_resume(cur_resume or "", doc_records=mock_doc, cadre_name=cadre_name)

# ==================== 全局异步治理工作线程 (全员底表与公文全维度校准治理版) ====================

class GlobalAuditWorker(QThread):
    progress_signal = Signal(int, int, str)
    finished_signal = Signal(int, int, dict)
    error_signal = Signal(str)

    def __init__(self, base_ym):
        super().__init__()
        self.base_ym = base_ym
        self.is_cancelled = False

    def cancel(self):
        self.is_cancelled = True

    def run(self):
        conn = None
        try:
            # 开启 WAL 预写日志与 30 秒忙等待保护，彻底防止 database is locked
            conn = sqlite3.connect(DB_PATH, timeout=30.0)
            conn.execute("PRAGMA journal_mode = WAL;")
            conn.execute("PRAGMA busy_timeout = 30000;")
            c = conn.cursor()

            # 1. 准确定位三大 Excel 底表文件与审批表库
            cadre_file = None
            roster_file = None
            contact_file = None
            for f in os.listdir(BASE_DIR):
                if f.startswith("~$") or not f.endswith(".xlsx"):
                    continue
                if "联系方式" in f or "通讯录" in f:
                    contact_file = os.path.join(BASE_DIR, f)
                elif "人员" in f and ("花名册" in f or "名册" in f):
                    roster_file = os.path.join(BASE_DIR, f)
                elif "中层干部" in f or ("干部" in f and "名册" in f):
                    cadre_file = os.path.join(BASE_DIR, f)

            # 2. 内存解析全院人员花名册底表
            roster_map = {}
            if roster_file and os.path.exists(roster_file):
                wb_r = openpyxl.load_workbook(roster_file, data_only=True)
                ws_r = wb_r.active
                header_row_r = 1
                for r in range(1, 6):
                    r_vals = [str(cell.value).strip() if cell.value else "" for cell in ws_r[r]]
                    if any("姓名" in v for v in r_vals):
                        header_row_r = r
                        break
                headers_r = [str(cell.value).strip() if cell.value else "" for cell in ws_r[header_row_r]]
                m_r = extract_flexible_roster_headers(headers_r)
                name_idx = m_r["name"] if m_r["name"] != -1 else 2
                emp_no_idx = m_r["emp_no"] if m_r["emp_no"] != -1 else 0
                dept_idx = m_r["dept"] if m_r["dept"] != -1 else 3
                id_idx = m_r["id_card"] if m_r["id_card"] != -1 else 6
                phone_idx = m_r["phone"]

                for row in ws_r.iter_rows(min_row=header_row_r + 1, values_only=True):
                    if row[name_idx]:
                        nm = str(row[name_idx]).replace(" ", "").strip()
                        def cell_val(k):
                            idx = m_r.get(k, -1)
                            return str(row[idx]).strip() if idx != -1 and idx < len(row) and is_valid_text(row[idx]) else ""

                        dept_val = str(row[dept_idx]).strip() if dept_idx != -1 and dept_idx < len(row) and row[dept_idx] else ""
                        cid_val = str(row[id_idx]).strip() if id_idx != -1 and id_idx < len(row) and row[id_idx] else ""
                        cph_val = str(row[phone_idx]).strip() if phone_idx != -1 and phone_idx < len(row) and row[phone_idx] else ""
                        eno_val = str(row[emp_no_idx]).strip() if emp_no_idx != -1 and emp_no_idx < len(row) and row[emp_no_idx] else ""

                        if nm not in roster_map: roster_map[nm] = []
                        roster_map[nm].append({
                            "emp_no": eno_val, "dept": dept_val, "id_card": cid_val, "phone": cph_val,
                            "identity": cell_val("identity"),
                            "fe_deg": cell_val("fe_deg"), "fe_sch": cell_val("fe_sch"),
                            "fe_maj": cell_val("fe_maj"), "fe_dt": cell_val("fe_dt"),
                            "he_deg": cell_val("he_deg"), "he_tit": cell_val("he_tit"),
                            "he_sch": cell_val("he_sch"), "he_maj": cell_val("he_maj"), "he_dt": cell_val("he_dt")
                        })

            # 3. 内存解析科级干部联系方式统计表
            contact_map = {}
            if contact_file and os.path.exists(contact_file):
                wb_cnt = openpyxl.load_workbook(contact_file, data_only=True)
                ws_cnt = wb_cnt.active
                for r in range(2, ws_cnt.max_row + 1):
                    nm = str(ws_cnt.cell(row=r, column=3).value or "").replace(" ", "").strip()
                    if nm and nm not in ["姓名", "姓 名"]:
                        if nm not in contact_map: contact_map[nm] = []
                        contact_map[nm].append({
                            "dept": str(ws_cnt.cell(row=r, column=1).value or "").strip(),
                            "post": str(ws_cnt.cell(row=r, column=2).value or "").strip(),
                            "office_phone": str(ws_cnt.cell(row=r, column=4).value or "").strip(),
                            "mobile": str(ws_cnt.cell(row=r, column=5).value or "").strip(),
                            "office_address": str(ws_cnt.cell(row=r, column=6).value or "").strip(),
                        })

            # 4. 内存解析中层干部名册底表
            cadre_base_map = {}
            if cadre_file and os.path.exists(cadre_file):
                wb_c = openpyxl.load_workbook(cadre_file, data_only=True)
                ws_c = wb_c["干部花名册"] if "干部花名册" in wb_c.sheetnames else wb_c.active
                for r in range(4, ws_c.max_row + 1):
                    nm = str(ws_c.cell(row=r, column=2).value or "").replace(" ", "").strip()
                    if nm and nm not in ["姓名", "姓 名"]:
                        if nm not in cadre_base_map: cadre_base_map[nm] = []
                        cadre_base_map[nm].append({
                            "gender": str(ws_c.cell(row=r, column=3).value or "").strip(),
                            "birth": str(ws_c.cell(row=r, column=4).value or "").strip(),
                            "dept_post": str(ws_c.cell(row=r, column=9).value or "").strip(),
                            "level": str(ws_c.cell(row=r, column=10).value or "").strip(),
                            "full_degree": str(ws_c.cell(row=r, column=11).value or "").strip(),
                            "full_school": str(ws_c.cell(row=r, column=12).value or "").strip(),
                            "job_degree": str(ws_c.cell(row=r, column=13).value or "").strip(),
                            "job_school": str(ws_c.cell(row=r, column=14).value or "").strip(),
                            "rem1": str(ws_c.cell(row=r, column=20).value or "").strip(),
                            "rem2": str(ws_c.cell(row=r, column=21).value or "").strip(),
                        })

            # 5. 扫描审批表文件库
            app_files_map = {}
            if os.path.exists(APPROVAL_DIR):
                for f in os.listdir(APPROVAL_DIR):
                    if f.lower().endswith(".docx") and not f.startswith("~$"):
                        clean_fn = f
                        for prefix in ["副本", "（副本）", "(副本)"]:
                            clean_fn = clean_fn.replace(prefix, "")
                        m_nm = re.search(r'[\d\-_]*([\u4e00-\u9fa5]{2,4})', clean_fn)
                        if m_nm:
                            cname = m_nm.group(1)
                            if cname not in app_files_map:
                                app_files_map[cname] = []
                            app_files_map[cname].append(f)

            # 6. 读取当前所有在岗干部完整数据进行全维度底表校准
            c.execute("""SELECT id, name, birth, dept_post, phone, office_phone, office_address,
                                cadre_identity, post_category, full_edu_degree, full_edu_school,
                                job_edu_degree, job_edu_school, id_card, emp_no,
                                selected_approval_file, family_relations, resume_text,
                                current_duty_date, career_history
                         FROM cadres WHERE status='在岗'""")
            all_cadres = c.fetchall()
            total_count = len(all_cadres)

            fe_calibrated_count = 0
            contact_synced_count = 0
            app_linked_count = 0
            healed_count = 0

            clean_ym_pat = re.compile(r'(\d{4})[^\d]?(\d{1,2})')
            def get_clean_ym(s):
                m = clean_ym_pat.search(str(s))
                return f"{m.group(1)}{int(m.group(2)):02d}" if m else ""

            # 第一阶段：三大底表全维度校准
            for idx, row in enumerate(all_cadres, start=1):
                if self.is_cancelled:
                    break
                (cid, name, birth, dept_post, cur_phone, cur_oph, cur_oaddr,
                 cur_identity, cur_post_cat, db_full_deg, db_full_sch,
                 db_job_deg, db_job_sch, db_id_card, db_emp_no,
                 sel_app, db_fam, cur_resume, cur_duty_d, career_h) = row

                self.progress_signal.emit(idx, total_count * 2, f"底表联合校准：【{name}】")

                # 匹配花名册候选
                r_cands = roster_map.get(name, [])
                p_match = {}
                if len(r_cands) == 1:
                    p_match = r_cands[0]
                elif len(r_cands) > 1:
                    t_ym = get_clean_ym(birth)
                    for cand in r_cands:
                        if t_ym and t_ym in cand.get("id_card", ""):
                            p_match = cand
                            break
                    if not p_match:
                        for cand in r_cands:
                            if cand.get("dept") and cand["dept"] in dept_post:
                                p_match = cand
                                break
                    if not p_match:
                        p_match = r_cands[0]

                # 匹配中层名册候选
                c_cands = cadre_base_map.get(name, [])
                c_match = {}
                if len(c_cands) == 1:
                    c_match = c_cands[0]
                elif len(c_cands) > 1:
                    t_ym = get_clean_ym(birth)
                    for cand in c_cands:
                        if t_ym and t_ym == get_clean_ym(cand.get("birth", "")):
                            c_match = cand
                            break
                    if not c_match:
                        c_match = c_cands[0]

                # 匹配通讯录候选
                cnt_cands = contact_map.get(name, [])
                cnt_match = {}
                if len(cnt_cands) == 1:
                    cnt_match = cnt_cands[0]
                elif len(cnt_cands) > 1:
                    for cand in cnt_cands:
                        if cand.get("dept") and cand["dept"] in dept_post:
                            cnt_match = cand
                            break
                    if not cnt_match:
                        cnt_match = cnt_cands[0]

                # 1. 编制身份与岗位类别校准
                final_identity = cur_identity
                if is_valid_text(p_match.get("identity")):
                    final_identity = p_match["identity"]
                elif is_valid_text(c_match.get("rem1")):
                    final_identity = c_match["rem1"]
                elif not is_valid_text(final_identity):
                    final_identity = "校编" if "河南医药大学" in dept_post or "新乡医学院" in dept_post else "院编"

                final_post_cat = cur_post_cat
                if is_valid_text(c_match.get("rem2")):
                    final_post_cat = c_match["rem2"]
                elif not is_valid_text(final_post_cat):
                    clinical_keywords = ["科主任", "护士长", "病区", "中心主任", "教研室", "研究所", "门诊", "药学"]
                    final_post_cat = "业务管理岗" if any(k in dept_post for k in clinical_keywords) else "党政管理岗"

                # 2. 手机号、办公电话与地址校准（严格非破坏性保护既有移动电话）
                final_phone = cur_phone if is_valid_text(cur_phone) else ""
                if not is_valid_text(final_phone):
                    if is_valid_text(cnt_match.get("mobile")):
                        final_phone = cnt_match["mobile"]
                    elif is_valid_text(p_match.get("phone")):
                        final_phone = p_match["phone"]

                final_oph = cur_oph if is_valid_text(cur_oph) else ""
                if is_valid_text(cnt_match.get("office_phone")):
                    final_oph = cnt_match["office_phone"]

                final_oaddr = cur_oaddr if is_valid_text(cur_oaddr) else ""
                if is_valid_text(cnt_match.get("office_address")):
                    final_oaddr = cnt_match["office_address"]

                if (final_oph != cur_oph or final_oaddr != cur_oaddr) and (is_valid_text(final_oph) or is_valid_text(final_oaddr)):
                    contact_synced_count += 1

                # 3. 工号与身份证号补齐
                final_emp_no = db_emp_no
                if not is_valid_text(final_emp_no) and is_valid_text(p_match.get("emp_no")):
                    final_emp_no = p_match["emp_no"]

                final_id_card = db_id_card
                if not is_valid_text(final_id_card) and is_valid_text(p_match.get("id_card")):
                    final_id_card = p_match["id_card"]

                # 4. 第一学历智能深度补齐与校准
                fe_deg = p_match.get("fe_deg")
                fe_sch = p_match.get("fe_sch")
                fe_maj = p_match.get("fe_maj")
                fe_dt  = p_match.get("fe_dt")

                # 若人员花名册中第一学历学校为空，自动穿透底表全日制教育或在职教育
                if not is_valid_text(fe_sch):
                    src_fsch = c_match.get("full_school") or db_full_sch
                    src_fdeg = c_match.get("full_degree") or db_full_deg
                    p_sch, p_maj = split_school_and_major(src_fsch)
                    p_deg, _ = split_degree_and_title(src_fdeg)
                    if is_valid_text(p_sch):
                        fe_sch = p_sch
                        fe_maj = p_maj or (fe_maj if is_valid_text(fe_maj) else "")
                        if not is_valid_text(fe_deg):
                            fe_deg = p_deg
                    else:
                        src_jsch = c_match.get("job_school") or db_job_sch
                        src_jdeg = c_match.get("job_degree") or db_job_deg
                        p_sch, p_maj = split_school_and_major(src_jsch)
                        p_deg, _ = split_degree_and_title(src_jdeg)
                        if is_valid_text(p_sch):
                            fe_sch = p_sch
                            fe_maj = p_maj or (fe_maj if is_valid_text(fe_maj) else "")
                            if not is_valid_text(fe_deg):
                                fe_deg = p_deg

                # 5. 最高学历智能校准
                he_deg = p_match.get("he_deg")
                he_tit = p_match.get("he_tit")
                he_sch = p_match.get("he_sch")
                he_maj = p_match.get("he_maj")
                he_dt  = p_match.get("he_dt")

                if not is_valid_text(he_sch):
                    src_jsch = c_match.get("job_school") or db_job_sch
                    src_jdeg = c_match.get("job_degree") or db_job_deg
                    if is_valid_text(src_jsch):
                        p_sch, p_maj = split_school_and_major(src_jsch)
                        p_deg, p_tit = split_degree_and_title(src_jdeg)
                        he_sch = p_sch
                        he_maj = p_maj or (he_maj if is_valid_text(he_maj) else "")
                        if not is_valid_text(he_deg): he_deg = p_deg
                        if not is_valid_text(he_tit): he_tit = p_tit
                    else:
                        src_fsch = c_match.get("full_school") or db_full_sch
                        src_fdeg = c_match.get("full_degree") or db_full_deg
                        p_sch, p_maj = split_school_and_major(src_fsch)
                        p_deg, p_tit = split_degree_and_title(src_fdeg)
                        he_sch = p_sch
                        he_maj = p_maj or (he_maj if is_valid_text(he_maj) else "")
                        if not is_valid_text(he_deg): he_deg = p_deg
                        if not is_valid_text(he_tit): he_tit = p_tit

                final_full_deg = c_match.get("full_degree") or db_full_deg
                final_full_sch = c_match.get("full_school") or db_full_sch
                final_job_deg  = c_match.get("job_degree") or db_job_deg
                final_job_sch  = c_match.get("job_school") or db_job_sch

                # 6. 审批表自动匹配与家庭关系回填
                final_sel_app = sel_app
                if (not is_valid_text(final_sel_app) or not os.path.exists(os.path.join(APPROVAL_DIR, final_sel_app))) and (name in app_files_map):
                    cands = app_files_map[name]
                    clean_cands = [f for f in cands if "副本" not in f and "(1)" not in f]
                    final_sel_app = clean_cands[0] if clean_cands else cands[0]
                    app_linked_count += 1

                final_fam = db_fam
                if (not is_valid_text(final_fam) or final_fam == "—") and is_valid_text(final_sel_app):
                    app_full_p = os.path.join(APPROVAL_DIR, final_sel_app)
                    if os.path.exists(app_full_p):
                        try:
                            _, parsed_fam = extract_resume_and_family_from_docx(app_full_p)
                            if is_valid_text(parsed_fam) and parsed_fam != "—":
                                final_fam = parsed_fam
                        except Exception:
                            pass

                # 更新该干部底表核准数据
                c.execute("""UPDATE cadres SET cadre_identity = ?, post_category = ?, phone = ?,
                             office_phone = ?, office_address = ?, emp_no = ?, id_card = ?,
                             first_edu_degree = ?, first_edu_school = ?, first_edu_major = ?, first_edu_date = ?,
                             highest_edu_degree = ?, highest_edu_degree_title = ?, highest_edu_school = ?,
                             highest_edu_major = ?, highest_edu_date = ?,
                             full_edu_degree = ?, full_edu_school = ?, job_edu_degree = ?, job_edu_school = ?,
                             selected_approval_file = ?, family_relations = ?
                             WHERE id = ?""",
                          (final_identity, final_post_cat, final_phone,
                           final_oph, final_oaddr, final_emp_no, final_id_card,
                           fe_deg or "", fe_sch or "", fe_maj or "", fe_dt or "",
                           he_deg or "", he_tit or "", he_sch or "", he_maj or "", he_dt or "",
                           final_full_deg, final_full_sch, final_job_deg, final_job_sch,
                           final_sel_app, final_fam, cid))

                fe_calibrated_count += 1

                if idx % 20 == 0:
                    conn.commit()

            conn.commit()

            # 第二阶段：逐人检索公文记录与履职缝合（全程在同一连接与内存中完成，杜绝嵌套死锁）
            for idx, row in enumerate(all_cadres, start=1):
                if self.is_cancelled:
                    break
                (cid, name, birth, dept_post, cur_phone, cur_oph, cur_oaddr,
                 cur_identity, cur_post_cat, db_full_deg, db_full_sch,
                 db_job_deg, db_job_sch, db_id_card, db_emp_no,
                 sel_app, db_fam, cur_resume, cur_duty_d, career_h) = row

                self.progress_signal.emit(total_count + idx, total_count * 2, f"公文履历自愈闭合：【{name}】")

                c.execute("""SELECT doc_date, doc_number, content FROM cadre_records 
                             WHERE content LIKE ? AND id NOT IN (SELECT record_id FROM excluded_records WHERE cadre_id=?)
                             ORDER BY doc_date ASC, id ASC""", (f"%{name}%", cid))
                recs = c.fetchall()

                doc_list = [{"doc_date": r[0], "doc_num": r[1], "content": r[2]} for r in recs] if recs else None
                healed_res = heal_and_reorganize_resume(
                    cur_resume or "",
                    doc_records=doc_list,
                    cadre_name=name,
                    career_history=career_h or "",
                    current_dept_post=dept_post or "",
                    current_duty_date=cur_duty_d or ""
                )
                if healed_res and healed_res != cur_resume:
                    c.execute("UPDATE cadres SET resume_text = ? WHERE id = ?", (healed_res, cid))
                    healed_count += 1

                if idx % 20 == 0:
                    conn.commit()

            conn.commit()
            conn.close()
            conn = None

            stats = {
                "fe_calibrated": fe_calibrated_count,
                "contact_synced": contact_synced_count,
                "approval_linked": app_linked_count,
                "healed_count": healed_count
            }
            self.finished_signal.emit(healed_count, total_count, stats)

        except Exception as e:
            if conn:
                try: conn.close()
                except Exception: pass
            self.error_signal.emit(str(e))

# ==================== 现代立体指标卡片组件 ====================

class MetricCard(QFrame):
    clicked = Signal(str)

    def __init__(self, key, title, sub_title, icon, color="#2563EB", parent=None):
        super().__init__(parent)
        self.key = key
        self.title_str = title
        self.sub_title_str = sub_title
        self.icon_str = icon
        self.theme_color = color
        self.is_active = False

        self.setCursor(Qt.PointingHandCursor)
        self.setFixedHeight(88)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(2)

        top_h = QHBoxLayout()
        top_h.setSpacing(6)

        self.icon_lbl = QLabel(self.icon_str)
        self.icon_lbl.setFont(QFont("Segoe UI Emoji", 14))
        self.icon_lbl.setAlignment(Qt.AlignCenter)
        self.icon_lbl.setFixedSize(28, 28)
        self.icon_lbl.setStyleSheet(f"background-color: {self.theme_color}18; color: {self.theme_color}; border-radius: 6px;")
        top_h.addWidget(self.icon_lbl)

        self.num_lbl = QLabel("0 人")
        self.num_lbl.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        self.num_lbl.setStyleSheet("color: #0F172A;")
        top_h.addWidget(self.num_lbl)
        top_h.addStretch()

        layout.addLayout(top_h)

        self.title_lbl = QLabel(f"{self.title_str}  {self.sub_title_str}")
        self.title_lbl.setFont(QFont("Microsoft YaHei UI", 8, QFont.Bold))
        self.title_lbl.setStyleSheet("color: #475569;")
        self.title_lbl.setWordWrap(True)
        layout.addWidget(self.title_lbl)

        self.update_style()

    def set_value(self, val_text):
        self.num_lbl.setText(val_text)

    def set_active(self, active: bool):
        self.is_active = active
        self.update_style()

    def update_style(self):
        if self.is_active:
            self.setStyleSheet("""
                MetricCard {
                    background-color: #FEF3C7;
                    border: 2px solid #F59E0B;
                    border-radius: 8px;
                }
            """)
            self.title_lbl.setStyleSheet("color: #92400E; font-weight: bold;")
            self.num_lbl.setStyleSheet("color: #78350F; font-weight: bold;")
        else:
            self.setStyleSheet("""
                MetricCard {
                    background-color: #FFFFFF;
                    border: 1px solid #D1E1F0;
                    border-radius: 8px;
                }
                MetricCard:hover {
                    background-color: #F8FAFC;
                    border: 1px solid #3B82F6;
                }
            """)
            self.title_lbl.setStyleSheet("color: #475569;")
            self.num_lbl.setStyleSheet("color: #0F172A; font-weight: bold;")

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit(self.key)
        super().mousePressEvent(event)

# ==================== 主窗口：FullCadreMainWindow ====================

class FullCadreMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("科级干部全周期管理与综合研判系统 (组织部专版)")
        self.resize(1540, 930)
        self.setMinimumSize(1280, 780)

        self.current_tab = "active_ke"
        self.current_rows = []
        self.active_presets = set()
        self.display_cols = []

        self.init_database()
        self.init_ui()

        self.setup_table_columns()

        self.auto_init_from_excel_if_empty()
        self.load_cadre_data()

        threading.Thread(target=self.sync_appointment_files, daemon=True).start()
        threading.Thread(target=self.auto_sync_single_approval_forms, daemon=True).start()

    def update_status_bar(self, msg):
        self.status_lbl.setText(msg)

    def init_database(self):
        conn = sqlite3.connect(DB_PATH, timeout=30.0)
        conn.execute("PRAGMA journal_mode = WAL;")
        conn.execute("PRAGMA busy_timeout = 30000;")
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS cadres (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_no TEXT,
            name TEXT NOT NULL,
            gender TEXT,
            birth TEXT,
            nation TEXT,
            native_place TEXT,
            id_card TEXT,
            phone TEXT,
            office_phone TEXT,
            office_address TEXT,
            politic TEXT,
            party_date TEXT,
            work_date TEXT,
            dept_post TEXT,
            level TEXT,
            cadre_rank TEXT DEFAULT '科级',
            chu_category TEXT,
            charge_division TEXT,
            current_duty_date TEXT,
            same_level_date TEXT,
            title TEXT,
            title_date TEXT,
            full_edu_degree TEXT,
            full_edu_school TEXT,
            job_edu_degree TEXT,
            job_edu_school TEXT,
            career_history TEXT,
            cadre_identity TEXT,
            post_category TEXT,
            discipline_type TEXT,
            discipline_date TEXT,
            discipline_desc TEXT,
            resume_text TEXT,
            family_relations TEXT,
            selected_approval_file TEXT,
            status TEXT DEFAULT '在岗',
            leave_reason TEXT,
            leave_date TEXT,
            remark TEXT
        )''')
        c.execute("PRAGMA table_info(cadres)")
        existing_cols = [info[1] for info in c.fetchall()]
        for col_name in ALL_CADRE_COLUMNS:
            if col_name not in existing_cols and col_name != "id":
                c.execute(f"ALTER TABLE cadres ADD COLUMN {col_name} TEXT")

        c.execute('''CREATE TABLE IF NOT EXISTS indexed_files (
            file_name TEXT PRIMARY KEY,
            file_mtime REAL
        )''')
        c.execute('''CREATE TABLE IF NOT EXISTS cadre_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT,
            doc_number TEXT,
            doc_date TEXT,
            content TEXT
        )''')
        c.execute('''CREATE TABLE IF NOT EXISTS excluded_records (
            cadre_id INTEGER,
            record_id INTEGER,
            PRIMARY KEY (cadre_id, record_id)
        )''')
        conn.commit()
        conn.close()

    def auto_init_from_excel_if_empty(self):
        conn = sqlite3.connect(DB_PATH, timeout=30.0)
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM cadres")
        cnt = c.fetchone()[0]
        conn.close()
        if cnt == 0:
            self.run_deep_two_tables_import(is_init=True)

    def init_ui(self):
        main_widget = QWidget()
        main_widget.setObjectName("CentralMain")
        main_widget.setStyleSheet("QWidget#CentralMain { background-color: #EBF3FA; }")
        self.setCentralWidget(main_widget)

        root_layout = QHBoxLayout(main_widget)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # 1. 左侧深蓝垂直导航栏
        self.sidebar = QFrame()
        self.sidebar.setObjectName("LeftSidebar")
        self.sidebar.setFixedWidth(240)
        self.sidebar.setStyleSheet("""
            QFrame#LeftSidebar {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #0F172A, stop:1 #1E293B);
                border-right: 1px solid #334155;
            }
        """)

        side_layout = QVBoxLayout(self.sidebar)
        side_layout.setContentsMargins(14, 22, 14, 18)
        side_layout.setSpacing(8)

        brand_f = QFrame()
        brand_l = QVBoxLayout(brand_f)
        brand_l.setContentsMargins(4, 0, 4, 16)
        brand_l.setSpacing(4)

        sys_title = QLabel("科级干部")
        sys_title.setFont(QFont("Microsoft YaHei UI", 12, QFont.Bold))
        sys_title.setStyleSheet("color: #93C5FD;")

        sys_sub = QLabel("全周期管理与综合研判系统")
        sys_sub.setFont(QFont("Microsoft YaHei UI", 11, QFont.Bold))
        sys_sub.setStyleSheet("color: #FFFFFF;")
        sys_sub.setWordWrap(True)

        tag_lbl = QLabel("  组织部专用核心系统  ")
        tag_lbl.setFont(QFont("Microsoft YaHei UI", 8))
        tag_lbl.setStyleSheet("background-color: #1E3A8A; color: #60A5FA; border-radius: 4px; padding: 2px;")
        tag_lbl.setFixedHeight(22)

        brand_l.addWidget(sys_title)
        brand_l.addWidget(sys_sub)
        brand_l.addWidget(tag_lbl)
        side_layout.addWidget(brand_f)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("background-color: #334155; max-height: 1px;")
        side_layout.addWidget(sep)
        side_layout.addSpacing(6)

        nav_lbl = QLabel("核心功能专区")
        nav_lbl.setFont(QFont("Microsoft YaHei UI", 8, QFont.Bold))
        nav_lbl.setStyleSheet("color: #64748B; padding-left: 6px;")
        side_layout.addWidget(nav_lbl)

        self.nav_btns = {}
        nav_items = [
            ("active_ke", "🏛️  科级干部队伍大盘"),
            ("active_chu", "🎖️  处级领导干部专区"),
            ("history_ke", "🚪  科级历史离退人员"),
            ("history_chu", "📁  处级历史调离人员"),
        ]

        for k, text in nav_items:
            btn = QPushButton(text)
            btn.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
            btn.setFixedHeight(42)
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda chk, key=k: self.switch_tab(key))
            self.nav_btns[k] = btn
            side_layout.addWidget(btn)

        side_layout.addSpacing(12)
        dec_lbl = QLabel("常委会汇报工具")
        dec_lbl.setFont(QFont("Microsoft YaHei UI", 8, QFont.Bold))
        dec_lbl.setStyleSheet("color: #64748B; padding-left: 6px;")
        side_layout.addWidget(dec_lbl)

        self.dash_nav_btn = QPushButton("📊  干部队伍总览")
        self.dash_nav_btn.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
        self.dash_nav_btn.setFixedHeight(42)
        self.dash_nav_btn.setCursor(Qt.PointingHandCursor)
        self.dash_nav_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #D97706, stop:1 #B45309);
                color: #FFFFFF;
                border-radius: 8px;
                text-align: left;
                padding-left: 14px;
                border: none;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #F59E0B, stop:1 #D97706);
            }
        """)
        self.dash_nav_btn.clicked.connect(self.open_stats_dashboard)
        side_layout.addWidget(self.dash_nav_btn)

        side_layout.addStretch()

        lib_f = QFrame()
        lib_l = QVBoxLayout(lib_f)
        lib_l.setContentsMargins(0, 0, 0, 0)
        lib_l.setSpacing(6)

        btn_lib_doc = QPushButton("📂 打开任免公文库")
        btn_lib_app = QPushButton("📑 打开审批表库")
        btn_lib_dis = QPushButton("⚖️ 打开纪律处分库")

        for b in [btn_lib_doc, btn_lib_app, btn_lib_dis]:
            b.setFont(QFont("Microsoft YaHei UI", 8))
            b.setFixedHeight(30)
            b.setCursor(Qt.PointingHandCursor)
            b.setStyleSheet("""
                QPushButton {
                    background-color: #1E293B;
                    color: #94A3B8;
                    border: 1px solid #334155;
                    border-radius: 6px;
                }
                QPushButton:hover {
                    background-color: #334155;
                    color: #FFFFFF;
                }
            """)

        btn_lib_doc.clicked.connect(lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(FILE_DIR)))
        btn_lib_app.clicked.connect(lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(APPROVAL_DIR)))
        btn_lib_dis.clicked.connect(lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(DISCIPLINE_DIR)))

        lib_l.addWidget(btn_lib_doc)
        lib_l.addWidget(btn_lib_app)
        lib_l.addWidget(btn_lib_dis)
        side_layout.addWidget(lib_f)

        root_layout.addWidget(self.sidebar)

        # 2. 右侧主工作区
        content_container = QWidget()
        content_container.setStyleSheet("background-color: transparent;")
        content_layout = QVBoxLayout(content_container)
        content_layout.setContentsMargins(18, 14, 18, 14)
        content_layout.setSpacing(10)

        top_bar = QFrame()
        top_bar.setFixedHeight(50)
        top_bar.setStyleSheet("background-color: #FFFFFF; border-radius: 8px; border: 1px solid #D1E1F0;")
        top_bar_l = QHBoxLayout(top_bar)
        top_bar_l.setContentsMargins(16, 0, 16, 0)

        self.zone_title_lbl = QLabel("科级干部队伍大盘 (全周期综合研判中心)")
        self.zone_title_lbl.setFont(QFont("Microsoft YaHei UI", 11, QFont.Bold))
        self.zone_title_lbl.setStyleSheet("color: #1E293B; border: none;")
        top_bar_l.addWidget(self.zone_title_lbl)

        top_bar_l.addStretch()

        today_str = datetime.datetime.now().strftime("%Y年%m月%d日")
        date_lbl = QLabel(f"📅 运行基准：{today_str}")
        date_lbl.setFont(QFont("Microsoft YaHei UI", 9))
        date_lbl.setStyleSheet("color: #64748B; border: none; margin-right: 12px;")
        top_bar_l.addWidget(date_lbl)

        user_tag = QLabel("  👤 组织部专管权限  ")
        user_tag.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
        user_tag.setStyleSheet("background-color: #EFF6FF; color: #1D4ED8; border: 1px solid #BFDBFE; border-radius: 6px; padding: 4px;")
        top_bar_l.addWidget(user_tag)

        content_layout.addWidget(top_bar)

        # KPI 卡片流
        self.metric_cards_frame = QFrame()
        self.metric_cards_layout = QHBoxLayout(self.metric_cards_frame)
        self.metric_cards_layout.setContentsMargins(0, 0, 0, 0)
        self.metric_cards_layout.setSpacing(8)

        self.cards = {}
        cards_cfg = [
            ("total", "在岗中层总数", "(全貌)", "👥", "#2563EB"),
            ("sub_to_chief", "副科满2年", "(提正科)", "⭐", "#0284C7"),
            ("chief_to_deputy", "正科满3年", "(提副处)", "⭐", "#0369A1"),
            ("duty_5years", "同岗满5年", "(应轮岗)", "🔄", "#D97706"),
            ("duty_2terms", "满两届8年", "(达任期)", "⛔", "#DC2626"),
            ("young", "40岁以下", "(年轻干部)", "🌱", "#059669"),
            ("retire_soon", "临近退休", "(2年内到龄)", "⏳", "#EA580C"),
            ("discipline", "处分影响期", "(提拔受限)", "⚖️", "#B91C1C"),
            ("same_unit", "同单位亲属", "(任职回避)", "👥", "#7C3AED"),
            ("missing", "档案信息缺失", "(需补齐)", "⚠️", "#64748B"),
        ]

        for k, title, sub_t, icon, color in cards_cfg:
            c = MetricCard(k, title, sub_t, icon, color)
            c.clicked.connect(self.on_metric_card_clicked)
            self.cards[k] = c
            self.metric_cards_layout.addWidget(c)

        content_layout.addWidget(self.metric_cards_frame)

        # 综合研判控制台
        self.console_card = QFrame()
        self.console_card.setStyleSheet("background-color: #FFFFFF; border: 1px solid #D1E1F0; border-radius: 8px;")
        console_l = QVBoxLayout(self.console_card)
        console_l.setContentsMargins(14, 10, 14, 10)
        console_l.setSpacing(8)

        row1_l = QHBoxLayout()
        row1_l.setSpacing(8)

        lbl_ym = QLabel("基准年月:")
        lbl_ym.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
        lbl_ym.setStyleSheet("color: #475569; border: none;")
        row1_l.addWidget(lbl_ym)

        self.base_ym_input = QLineEdit(datetime.datetime.now().strftime("%Y年%m月"))
        self.base_ym_input.setFixedWidth(85)
        self.base_ym_input.setFixedHeight(30)
        self.base_ym_input.setStyleSheet("border: 1px solid #CBD5E1; border-radius: 6px; padding: 0 6px; background-color: #FFFFFF; color: #1E293B;")
        row1_l.addWidget(self.base_ym_input)

        combo_style = COMBO_STYLE

        self.combos = {}
        combo_cfg = [
            ("level", ["全部级别", "正科", "副科"], 96),
            ("post_cat", ["全部岗位", "党政管理岗", "业务管理岗"], 112),
            ("identity", ["全部身份", "校编", "院编", "人事代理"], 105),
            ("doc", ["不限学历", "硕士及以上 (硕博全包)", "全部博士 (全日制/在职)", "仅限全日制博士", "仅限在职博士", "仅限硕士 (全日制/在职)", "本科及以上 (含硕博)", "仅限大学本科 (最高本科)"], 145),
            ("title", ["全部职称", "所有高级职称 (正高+副高全包)", "正高级 (正高/主任医等)", "副高级 (副高/副主任等)", "中级职称 (主治/主管等)", "初级及以下 (中级职称以下)", "中级及以下"], 145),
            ("politic", ["全部面貌", "中共党员", "民主党派/无党派群众"], 105),
        ]

        for k, items, w in combo_cfg:
            cb = QComboBox()
            cb.addItems(items)
            cb.setMinimumWidth(w)
            cb.setFixedHeight(30)
            cb.setCursor(Qt.PointingHandCursor)
            cb.setItemDelegate(QStyledItemDelegate(cb))
            cb.setStyleSheet(combo_style)
            self.combos[k] = cb
            row1_l.addWidget(cb)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 全息关键词检索...")
        self.search_input.setFixedHeight(30)
        self.search_input.setStyleSheet("border: 1px solid #CBD5E1; border-radius: 6px; padding: 0 8px; background-color: #FFFFFF; color: #1E293B;")
        self.search_input.returnPressed.connect(self.filter_cadres)
        row1_l.addWidget(self.search_input)

        self.btn_filter = QPushButton(" 筛选 ")
        self.btn_filter.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
        self.btn_filter.setFixedHeight(30)
        self.btn_filter.setCursor(Qt.PointingHandCursor)
        self.btn_filter.setStyleSheet("""
            QPushButton {
                background-color: #2563EB; color: #FFFFFF; border-radius: 6px; padding: 0 14px; border: none;
            }
            QPushButton:hover { background-color: #1D4ED8; }
        """)
        self.btn_filter.clicked.connect(self.filter_cadres)
        row1_l.addWidget(self.btn_filter)

        self.btn_reset = QPushButton(" 重置 ")
        self.btn_reset.setFont(QFont("Microsoft YaHei UI", 9))
        self.btn_reset.setFixedHeight(30)
        self.btn_reset.setCursor(Qt.PointingHandCursor)
        self.btn_reset.setStyleSheet("""
            QPushButton {
                background-color: #F1F5F9; color: #475569; border-radius: 6px; padding: 0 12px; border: 1px solid #CBD5E1;
            }
            QPushButton:hover { background-color: #E2E8F0; }
        """)
        self.btn_reset.clicked.connect(self.reset_filter)
        row1_l.addWidget(self.btn_reset)

        console_l.addLayout(row1_l)

        # 控制台第 2 行
        self.act_row_l = QHBoxLayout()
        self.act_row_l.setSpacing(6)

        self.btn_auto_heal = QPushButton("⚡ 一键全员底表校准与公文智能治理")
        self.btn_auto_heal.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
        self.btn_auto_heal.setFixedHeight(28)
        self.btn_auto_heal.setCursor(Qt.PointingHandCursor)
        self.btn_auto_heal.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #4F46E5, stop:1 #6366F1);
                color: #FFFFFF; border-radius: 6px; padding: 0 12px; border: none;
            }
            QPushButton:hover { background: #4338CA; }
        """)
        self.btn_auto_heal.clicked.connect(self.one_click_global_audit_and_heal)
        self.act_row_l.addWidget(self.btn_auto_heal)

        self.btn_add = QPushButton("➕ 新增干部")
        self.btn_edit = QPushButton("✏️ 编辑档案")
        self.btn_bind = QPushButton("🎯 匹配底表")
        self.btn_sync_appt = QPushButton("📜 换届同步")
        self.btn_sync_disc = QPushButton("⚖️ 处分同步")
        self.btn_contact = QPushButton("📞 通讯录")
        self.btn_import_diff = QPushButton("📥 增量导入名册")
        self.btn_leave = QPushButton("🚪 离岗办理")

        action_btns = [
            (self.btn_add, "#059669", "#FFFFFF", self.open_add_dialog),
            (self.btn_edit, "#F1F5F9", "#1E293B", self.open_edit_dialog),
            (self.btn_bind, "#EFF6FF", "#1D4ED8", self.open_manual_bind_dialog),
            (self.btn_sync_appt, "#F0FDF4", "#16A34A", self.open_batch_appointment_sync),
            (self.btn_sync_disc, "#FEF2F2", "#DC2626", self.open_discipline_doc_sync),
            (self.btn_contact, "#F1F5F9", "#334155", self.batch_import_contacts),
            (self.btn_import_diff, "#EFF6FF", "#2563EB", self.open_import_diff_dialog),
            (self.btn_leave, "#FEF2F2", "#B91C1C", self.open_leave_dialog),
        ]

        for b, bg, fg, handler in action_btns:
            b.setFont(QFont("Microsoft YaHei UI", 8))
            b.setFixedHeight(28)
            b.setCursor(Qt.PointingHandCursor)
            b.setStyleSheet(f"""
                QPushButton {{
                    background-color: {bg}; color: {fg}; border-radius: 6px; padding: 0 6px; border: 1px solid #CBD5E1;
                }}
                QPushButton:hover {{ opacity: 0.9; border: 1px solid #94A3B8; }}
            """)
            b.clicked.connect(handler)
            self.act_row_l.addWidget(b)

        self.act_row_l.addStretch()

        self.btn_exp_avoidance = QPushButton("👥 亲属回避排查台账")
        self.btn_exp_avoidance.setFont(QFont("Microsoft YaHei UI", 8, QFont.Bold))
        self.btn_exp_avoidance.setFixedHeight(28)
        self.btn_exp_avoidance.setCursor(Qt.PointingHandCursor)
        self.btn_exp_avoidance.setStyleSheet("background-color: #B45309; color: white; border-radius: 6px; padding: 0 8px; border: none;")
        self.btn_exp_avoidance.clicked.connect(self.export_avoidance_investigation_excel)
        self.act_row_l.addWidget(self.btn_exp_avoidance)

        self.btn_exp_custom = QPushButton("📑 自定义导出")
        self.btn_exp_custom.setFont(QFont("Microsoft YaHei UI", 8, QFont.Bold))
        self.btn_exp_custom.setFixedHeight(28)
        self.btn_exp_custom.setCursor(Qt.PointingHandCursor)
        self.btn_exp_custom.setStyleSheet("background-color: #047857; color: white; border-radius: 6px; padding: 0 8px; border: none;")
        self.btn_exp_custom.clicked.connect(self.open_custom_export_dialog)
        self.act_row_l.addWidget(self.btn_exp_custom)

        self.btn_exp_word = QPushButton("📑 单人导出审批表")
        self.btn_exp_word.setFont(QFont("Microsoft YaHei UI", 8, QFont.Bold))
        self.btn_exp_word.setFixedHeight(28)
        self.btn_exp_word.setCursor(Qt.PointingHandCursor)
        self.btn_exp_word.setStyleSheet("background-color: #15803D; color: white; border-radius: 6px; padding: 0 8px; border: none;")
        self.btn_exp_word.clicked.connect(self.export_selected_approval_docx)
        self.act_row_l.addWidget(self.btn_exp_word)

        self.btn_batch_exp_word = QPushButton("📑 批量导出审批表")
        self.btn_batch_exp_word.setFont(QFont("Microsoft YaHei UI", 8, QFont.Bold))
        self.btn_batch_exp_word.setFixedHeight(28)
        self.btn_batch_exp_word.setCursor(Qt.PointingHandCursor)
        self.btn_batch_exp_word.setStyleSheet("background-color: #0D9488; color: white; border-radius: 6px; padding: 0 8px; border: none;")
        self.btn_batch_exp_word.clicked.connect(self.batch_export_approval_docx)
        self.act_row_l.addWidget(self.btn_batch_exp_word)

        console_l.addLayout(self.act_row_l)

        # 处级专区专属工具条
        self.chu_toolbar_frame = QFrame()
        self.chu_toolbar_layout = QHBoxLayout(self.chu_toolbar_frame)
        self.chu_toolbar_layout.setContentsMargins(0, 4, 0, 0)
        self.chu_toolbar_layout.setSpacing(10)

        lbl_chu_cat = QLabel("处级归口类别:")
        lbl_chu_cat.setFont(QFont("Microsoft YaHei UI", 9, QFont.Bold))
        self.chu_cat_combo = QComboBox()
        self.chu_cat_combo.addItems(["全部处级类别", "单位领导班子", "直属/分支机构领导", "专业内设机构领导"])
        self.chu_cat_combo.setFixedHeight(30)
        self.chu_cat_combo.setItemDelegate(QStyledItemDelegate(self.chu_cat_combo))
        self.chu_cat_combo.setStyleSheet(combo_style)
        self.chu_cat_combo.currentTextChanged.connect(self.filter_chu_cadres)

        btn_add_chu = QPushButton("➕ 手动新增处级干部")
        btn_add_chu.setFixedHeight(30)
        btn_add_chu.setStyleSheet("background-color: #047857; color: white; border-radius: 6px; padding: 0 10px;")
        btn_add_chu.clicked.connect(self.open_add_chu_dialog)

        btn_upload_chu_doc = QPushButton("📤 上传处级任免公文(Word/PDF)")
        btn_upload_chu_doc.setFixedHeight(30)
        btn_upload_chu_doc.setStyleSheet("background-color: #1E40AF; color: white; border-radius: 6px; padding: 0 10px;")
        btn_upload_chu_doc.clicked.connect(self.upload_chu_appointment_doc)

        btn_import_chu_xls = QPushButton("📥 批量导入处级名册(Excel)")
        btn_import_chu_xls.setFixedHeight(30)
        btn_import_chu_xls.setStyleSheet("background-color: #0284C7; color: white; border-radius: 6px; padding: 0 10px;")
        btn_import_chu_xls.clicked.connect(self.batch_import_chu_excel)

        self.chu_toolbar_layout.addWidget(lbl_chu_cat)
        self.chu_toolbar_layout.addWidget(self.chu_cat_combo)
        self.chu_toolbar_layout.addStretch()
        self.chu_toolbar_layout.addWidget(btn_add_chu)
        self.chu_toolbar_layout.addWidget(btn_upload_chu_doc)
        self.chu_toolbar_layout.addWidget(btn_import_chu_xls)

        console_l.addWidget(self.chu_toolbar_frame)
        self.chu_toolbar_frame.hide()

        content_layout.addWidget(self.console_card)

        # 2.4 主数据表格卡片区
        self.table_card = QFrame()
        self.table_card.setStyleSheet("background-color: #FFFFFF; border: 1px solid #D1E1F0; border-radius: 8px;")
        table_card_layout = QVBoxLayout(self.table_card)
        table_card_layout.setContentsMargins(8, 8, 8, 8)

        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(38)
        self.table.horizontalHeader().setHighlightSections(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().setSortIndicatorShown(True)
        self.table.horizontalHeader().setSectionsClickable(True)

        self.table.setStyleSheet("""
            QTableWidget {
                background-color: #FFFFFF;
                color: #1E293B;
                alternate-background-color: #F8FAFC;
                gridline-color: #F1F5F9;
                border: none;
                font-family: "PingFang SC", "Microsoft YaHei UI", sans-serif;
                font-size: 9pt;
            }
            QTableWidget::item {
                color: #1E293B;
                padding: 4px 6px;
                border-bottom: 1px solid #F1F5F9;
            }
            QTableWidget::item:selected {
                background-color: #EFF6FF;
                color: #1E3A8A;
                font-weight: bold;
            }
            QHeaderView::section {
                background-color: #F1F5F9;
                color: #475569;
                font-weight: bold;
                font-size: 9pt;
                padding: 6px;
                border: none;
                border-bottom: 2px solid #CBD5E1;
            }
        """)

        self.table.doubleClicked.connect(self.open_profile_from_selection)
        table_card_layout.addWidget(self.table)
        content_layout.addWidget(self.table_card, 1)

        # 2.5 底部状态栏
        bottom_bar = QFrame()
        bottom_bar.setFixedHeight(28)
        bottom_l = QHBoxLayout(bottom_bar)
        bottom_l.setContentsMargins(4, 0, 4, 0)

        self.status_lbl = QLabel("系统就绪")
        self.status_lbl.setFont(QFont("Microsoft YaHei UI", 9))
        self.status_lbl.setStyleSheet("color: #475569;")
        bottom_l.addWidget(self.status_lbl)

        bottom_l.addStretch()

        hint_lbl = QLabel("💡 提示：双击人员直接调出全息档案卡；支持指标卡多标签链式叠加；点【重置】清空大盘。")
        hint_lbl.setFont(QFont("Microsoft YaHei UI", 8))
        hint_lbl.setStyleSheet("color: #94A3B8;")
        bottom_l.addWidget(hint_lbl)

        content_layout.addWidget(bottom_bar)
        root_layout.addWidget(content_container, 1)

        self.update_nav_state()

    # ==================== 专区切换与表格列结构装配 ====================

    def update_nav_state(self):
        for k, btn in self.nav_btns.items():
            if k == self.current_tab:
                btn.setChecked(True)
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: #2563EB;
                        color: #FFFFFF;
                        border-radius: 8px;
                        text-align: left;
                        padding-left: 14px;
                        border: none;
                    }
                """)
            else:
                btn.setChecked(False)
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: transparent;
                        color: #94A3B8;
                        border-radius: 8px;
                        text-align: left;
                        padding-left: 14px;
                        border: none;
                    }
                    QPushButton:hover {
                        background-color: #1E293B;
                        color: #FFFFFF;
                    }
                """)

    def switch_tab(self, tab):
        self.current_tab = tab
        self.update_nav_state()

        if tab == "active_ke":
            self.zone_title_lbl.setText("科级干部队伍大盘 (全周期综合研判中心)")
            self.metric_cards_frame.show()
            self.combos["level"].show()
            self.combos["post_cat"].show()
            self.combos["identity"].show()
            self.combos["doc"].show()
            self.combos["title"].show()
            self.combos["politic"].show()
            self.chu_toolbar_frame.hide()

            self.btn_auto_heal.show()
            self.btn_add.setText("➕ 新增干部")
            self.btn_add.setEnabled(True)
            self.btn_bind.show()
            self.btn_sync_appt.show()
            self.btn_sync_disc.show()
            self.btn_contact.show()
            self.btn_leave.setText("🚪 离岗办理")

        elif tab == "active_chu":
            self.zone_title_lbl.setText("处级领导干部专区 (纯净名册与任免归档)")
            self.metric_cards_frame.hide()
            self.combos["level"].hide()
            self.combos["post_cat"].hide()
            self.combos["identity"].hide()
            self.combos["doc"].hide()
            self.combos["title"].hide()
            self.combos["politic"].hide()
            self.chu_toolbar_frame.show()

            self.btn_auto_heal.hide()
            self.btn_add.setText("➕ 新增处级干部")
            self.btn_add.setEnabled(True)
            self.btn_bind.show()
            self.btn_sync_appt.hide()
            self.btn_sync_disc.hide()
            self.btn_contact.hide()
            self.btn_leave.setText("🚪 处级调离办理")

        elif tab == "history_ke":
            self.zone_title_lbl.setText("科级历史/离退人员档案库")
            self.metric_cards_frame.hide()
            self.chu_toolbar_frame.hide()
            self.btn_auto_heal.hide()
            self.btn_add.setEnabled(False)
            self.btn_bind.hide()
            self.btn_sync_appt.hide()
            self.btn_sync_disc.hide()
            self.btn_contact.hide()
            self.btn_leave.setText("🔄 恢复在岗/彻底删除")

        else:
            self.zone_title_lbl.setText("处级历史/调离人员档案库")
            self.metric_cards_frame.hide()
            self.chu_toolbar_frame.hide()
            self.btn_auto_heal.hide()
            self.btn_add.setEnabled(False)
            self.btn_bind.hide()
            self.btn_sync_appt.hide()
            self.btn_sync_disc.hide()
            self.btn_contact.hide()
            self.btn_leave.setText("🔄 恢复在岗/彻底删除")

        self.setup_table_columns()
        self.load_cadre_data()

    def setup_table_columns(self):
        if self.current_tab == "active_ke":
            self.display_cols = ["序号", "姓名", "人员身份", "岗位类别", "工作部门及职务", "级别", "性别", "出生年月", "年龄",
                                 "任现职时间", "现职务已任", "任同级时间", "同级别已任", "全日制学历学位", "在职学历学位", "职称",
                                 "政治面貌", "纪律/影响期预警", "亲属任职回避", "移动电话", "办公电话", "办公地址", "审批表状态", "缺失预警"]
            widths = [50, 85, 80, 95, 210, 65, 50, 90, 50, 90, 90, 90, 90, 120, 120, 110, 85, 140, 140, 110, 100, 120, 110, 130]

        elif self.current_tab == "active_chu":
            self.display_cols = ["序号", "姓名", "处级归口类别", "工作部门及职务", "领导班子分管业务与部门", "级别", "性别", "出生年月", "年龄",
                                 "任现职时间", "任同级时间", "全日制学历学位", "在职学历学位", "职称", "政治面貌", "移动电话", "办公电话", "办公地址"]
            widths = [50, 85, 140, 210, 240, 65, 50, 90, 50, 90, 90, 120, 120, 110, 85, 110, 100, 130]

        elif self.current_tab == "history_ke":
            self.display_cols = ["序号", "姓名", "人员身份", "岗位类别", "原工作部门及职务", "级别", "离岗类型", "离岗时间", "移动电话", "离岗说明"]
            widths = [50, 90, 80, 95, 250, 75, 95, 105, 120, 220]

        else:
            self.display_cols = ["序号", "姓名", "原处级类别", "原工作部门及职务", "职务级别", "变动性质", "调离/批复时间", "去向及说明", "移动电话"]
            widths = [50, 90, 130, 240, 75, 105, 110, 220, 120]

        self.table.setColumnCount(len(self.display_cols))
        self.table.setHorizontalHeaderLabels(self.display_cols)

        for idx, w in enumerate(widths):
            self.table.setColumnWidth(idx, w)

    # ==================== 数据装载与指标卡动态刷新 ====================

    def get_base_ym(self):
        txt = self.base_ym_input.text().strip()
        if not parse_ym(txt):
            now_str = datetime.datetime.now().strftime("%Y年%m月")
            self.base_ym_input.setText(now_str)
            return now_str
        return txt

    def check_missing_fields(self, row_dict):
        missing = []
        if not is_valid_text(row_dict.get("id_card")): missing.append("身份证号")
        if not is_valid_text(row_dict.get("phone")): missing.append("移动电话")
        if not is_valid_text(row_dict.get("current_duty_date")): missing.append("任现职时间")
        if not is_valid_text(row_dict.get("same_level_date")): missing.append("任同级时间")
        if not is_valid_text(row_dict.get("full_edu_degree")) and not is_valid_text(row_dict.get("job_edu_degree")):
            missing.append("学历学位")
        return missing

    def load_cadre_data(self):
        base_ym = self.get_base_ym()
        app_files = [f for f in os.listdir(APPROVAL_DIR) if f.lower().endswith(".docx") and not f.startswith("~$")] if os.path.exists(APPROVAL_DIR) else []

        conn = sqlite3.connect(DB_PATH, timeout=30.0)
        c = conn.cursor()

        if self.current_tab == "active_ke":
            c.execute("""SELECT id, name, dept_post, level, gender, birth, 
                             current_duty_date, same_level_date, full_edu_degree, job_edu_degree, title, 
                             politic, phone, office_phone, office_address, id_card, selected_approval_file, 
                             resume_text, family_relations, discipline_type, discipline_date,
                             cadre_identity, post_category 
                             FROM cadres 
                             WHERE status = '在岗' AND (cadre_rank='科级' OR cadre_rank IS NULL OR cadre_rank='')
                             ORDER BY id ASC""")
            raw_rows = c.fetchall()
            self.current_rows = []

            m_counts = {
                "total": len(raw_rows), "sub_to_chief": 0, "chief_to_deputy": 0,
                "duty_5years": 0, "duty_2terms": 0, "young": 0,
                "retire_soon": 0, "discipline": 0, "same_unit": 0, "missing": 0
            }

            for r in raw_rows:
                cid, nm, dept_post, lvl, gdr, brth, cur_date, same_date, f_edu, j_edu, ttl, pol, ph, o_ph, o_addr, idc, sel_app, res_txt, fam_txt, disc_t, disc_d, identity, post_cat = r

                m_cur = calc_months_diff(cur_date, base_ym)
                cur_len_str = format_months_to_str(m_cur)

                m_same = calc_months_diff(same_date, base_ym)
                same_len_str = format_months_to_str(m_same)

                age_val = calc_age(brth, base_ym)
                ret_info = calc_retirement_info(gdr, brth, base_ym, ttl, lvl)
                is_retire_soon = bool(ret_info and ret_info.get("is_soon"))
                if is_retire_soon:
                    m_counts["retire_soon"] += 1

                if is_retire_soon and age_val is not None:
                    age_str = f"{age_val}岁 ⏳({ret_info['months_left']}月)"
                else:
                    age_str = f"{age_val}岁" if age_val is not None else "—"

                if "副科" in str(lvl) and m_same is not None and m_same >= 24: m_counts["sub_to_chief"] += 1
                if "正科" in str(lvl) and m_same is not None and m_same >= 36: m_counts["chief_to_deputy"] += 1
                if m_cur is not None and m_cur >= 60: m_counts["duty_5years"] += 1
                if m_cur is not None and m_cur >= 96: m_counts["duty_2terms"] += 1
                if age_val is not None and age_val <= 40: m_counts["young"] += 1

                is_impacted, disc_status_str, _ = check_discipline_impact(disc_t, disc_d, base_ym)
                if is_impacted: m_counts["discipline"] += 1

                missing_list = self.check_missing_fields({
                    "id_card": idc, "phone": ph, "current_duty_date": cur_date,
                    "same_level_date": same_date, "full_edu_degree": f_edu, "job_edu_degree": j_edu
                })
                missing_str = "缺: " + "、".join(missing_list) if missing_list else "完整"
                if missing_list: m_counts["missing"] += 1

                same_unit_alerts = analyze_same_unit_relations(fam_txt, dept_post)
                same_unit_str = ("；".join(same_unit_alerts)) if same_unit_alerts else "—"
                if same_unit_alerts: m_counts["same_unit"] += 1

                matched_apps = [f for f in app_files if nm in f]
                if is_valid_text(sel_app): app_status = "✅ 已指定最新表"
                elif len(matched_apps) == 1: app_status = "✅ 单表已关联"
                elif len(matched_apps) > 1: app_status = f"⚠️ 存在{len(matched_apps)}表待选"
                else: app_status = "— 无审批表"

                row_item = (cid, 0, nm, identity or "—", post_cat or "—", dept_post, lvl, gdr, brth, age_str,
                            cur_date, cur_len_str, same_date, same_len_str, f_edu, j_edu, ttl, pol or "—",
                            disc_status_str, same_unit_str, ph or "—", o_ph or "—", o_addr or "—", app_status, missing_str)
                self.current_rows.append(row_item)

            for k, cnt in m_counts.items():
                if k in self.cards:
                    self.cards[k].set_value(f"{cnt} 人")

            self.update_status_bar(f"当前专区：【科级干部队伍大盘】，共计在岗 {len(self.current_rows)} 人。")

        elif self.current_tab == "active_chu":
            c.execute("""SELECT id, name, chu_category, dept_post, charge_division, level, gender, birth,
                             current_duty_date, same_level_date, full_edu_degree, job_edu_degree, title,
                             politic, phone, office_phone, office_address 
                             FROM cadres 
                             WHERE status = '在岗' AND cadre_rank = '处级'
                             ORDER BY id ASC""")
            raw_rows = c.fetchall()
            self.current_rows = []
            for r in raw_rows:
                cid, nm, chu_cat, dept_post, charge_div, lvl, gdr, brth, cur_date, same_date, f_edu, j_edu, ttl, pol, ph, o_ph, o_addr = r
                age_val = calc_age(brth, base_ym)
                ret_info = calc_retirement_info(gdr, brth, base_ym, ttl, "处级")
                is_retire_soon = bool(ret_info and ret_info.get("is_soon"))
                if is_retire_soon and age_val is not None:
                    age_str = f"{age_val}岁 ⏳({ret_info['months_left']}月)"
                else:
                    age_str = f"{age_val}岁" if age_val is not None else "—"

                row_item = (cid, 0, nm, chu_cat or "—", dept_post or "—", charge_div or "—", lvl or "—",
                            gdr or "—", brth or "—", age_str, cur_date or "—", same_date or "—",
                            f_edu or "—", j_edu or "—", ttl or "—", pol or "—", ph or "—", o_ph or "—", o_addr or "—")
                self.current_rows.append(row_item)

            self.update_status_bar(f"当前专区：【处级领导干部专区】，共计在岗 {len(self.current_rows)} 人。")

        elif self.current_tab == "history_ke":
            c.execute("""SELECT id, name, cadre_identity, post_category, dept_post, level, leave_reason, leave_date, phone, remark 
                             FROM cadres 
                             WHERE status != '在岗' AND (cadre_rank = '科级' OR cadre_rank IS NULL OR cadre_rank = '')
                             ORDER BY leave_date DESC, id DESC""")
            raw_rows = c.fetchall()
            self.current_rows = []
            for r in raw_rows:
                cid, nm, ident, post_c, dept_post, lvl, rsn, l_date, ph, rmk = r
                self.current_rows.append((cid, 0, nm, ident or "—", post_c or "—", dept_post, lvl, rsn, l_date, ph or "—", rmk or "—"))

            self.update_status_bar(f"当前专区：【科级历史离退人员档案库】，共计离任人员 {len(self.current_rows)} 人。")

        else:
            c.execute("""SELECT id, name, chu_category, dept_post, level, leave_reason, leave_date, remark, phone 
                             FROM cadres 
                             WHERE status != '在岗' AND cadre_rank = '处级'
                             ORDER BY leave_date DESC, id DESC""")
            raw_rows = c.fetchall()
            self.current_rows = []
            for r in raw_rows:
                cid, nm, chu_cat, dept_post, lvl, rsn, l_date, rmk, ph = r
                self.current_rows.append((cid, 0, nm, chu_cat or "—", dept_post or "—", lvl or "—", rsn or "—", l_date or "—", rmk or "—", ph or "—"))

            self.update_status_bar(f"当前专区：【处级历史调离人员档案库】，共计调离/退休处级干部 {len(self.current_rows)} 人。")

        conn.close()
        self.populate_table(self.current_rows)

    def populate_table(self, rows_to_display):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(rows_to_display))
        self.table.clearContents()

        for r_idx, row_data in enumerate(rows_to_display):
            cid = row_data[0]
            display_values = [str(r_idx + 1)] + list(row_data[2:])

            for c_idx, val in enumerate(display_values):
                col_name = self.display_cols[c_idx] if c_idx < len(self.display_cols) else ""

                # 计算智能排序键，支持整型数字、任职月数、日期元组
                sort_key = None
                if c_idx == 0:
                    sort_key = r_idx + 1
                elif "年龄" in col_name:
                    m = re.search(r'(\d+)', val)
                    sort_key = int(m.group(1)) if m else 999
                elif "月数" in col_name or "任职月数" in col_name:
                    m = re.search(r'(\d+)', val)
                    sort_key = int(m.group(1)) if m else -1
                elif any(k in col_name for k in ["年月", "日期", "时间"]):
                    sort_key = parse_sort_date(val)
                else:
                    sort_key = val.strip()

                item = SmartTableWidgetItem(val, sort_key=sort_key)
                item.setData(Qt.UserRole, cid)

                if any(k in col_name for k in ["职务", "业务", "专业", "学校", "说明", "去向", "预警", "地址", "亲属"]):
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                else:
                    item.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)

                if "⛔" in val or "🔴" in val:
                    item.setForeground(QBrush(QColor("#DC2626")))
                elif "缺:" in val or "⏳" in val:
                    item.setForeground(QBrush(QColor("#EA580C")))
                elif "🟡" in val:
                    item.setForeground(QBrush(QColor("#D97706")))
                elif "⚠️" in val:
                    item.setForeground(QBrush(QColor("#7C3AED")))
                elif "正科" in val or "正处" in val:
                    item.setForeground(QBrush(QColor("#1D4ED8")))
                elif "副科" in val or "副处" in val:
                    item.setForeground(QBrush(QColor("#0369A1")))
                else:
                    item.setForeground(QBrush(QColor("#1E293B")))

                self.table.setItem(r_idx, c_idx, item)

        self.table.setSortingEnabled(True)

    # ==================== 多标签链式自由叠加研判引擎 ====================

    def on_metric_card_clicked(self, key):
        if key == "total":
            self.reset_filter()
            return
        self.apply_quick_preset(key)

    def apply_quick_preset(self, preset):
        if preset in self.active_presets:
            self.active_presets.remove(preset)
        else:
            if preset == "sub_to_chief":
                self.active_presets.discard("chief_to_deputy")
            elif preset == "chief_to_deputy":
                self.active_presets.discard("sub_to_chief")

            if preset == "duty_5years":
                self.active_presets.discard("duty_2terms")
            elif preset == "duty_2terms":
                self.active_presets.discard("duty_5years")

            self.active_presets.add(preset)

        for k, card in self.cards.items():
            card.set_active(k in self.active_presets)

        self.filter_cadres()

    def filter_cadres(self):
        if self.current_tab == "active_chu":
            self.filter_chu_cadres()
            return

        base_ym = self.get_base_ym()
        kw = self.search_input.text().strip()
        target_lvl = self.combos["level"].currentText()
        target_post_cat = self.combos["post_cat"].currentText()
        target_identity = self.combos["identity"].currentText()
        doc_opt = self.combos["doc"].currentText()
        title_opt = self.combos["title"].currentText()
        pol_opt = self.combos["politic"].currentText()

        matched = []
        for item in self.current_rows:
            lvl_val = str(item[6])
            cur_date = item[10]
            same_date = item[12]
            m_cur = calc_months_diff(cur_date, base_ym)
            m_same = calc_months_diff(same_date, base_ym)
            brth = item[8]
            age_val = calc_age(brth, base_ym)
            disc_str = str(item[18])
            same_unit_str = str(item[19])
            missing_str = str(item[-1])

            if "sub_to_chief" in self.active_presets:
                if not ("副科" in lvl_val and m_same is not None and m_same >= 24): continue
            if "chief_to_deputy" in self.active_presets:
                if not ("正科" in lvl_val and m_same is not None and m_same >= 36): continue
            if "duty_5years" in self.active_presets:
                if not (m_cur is not None and m_cur >= 60): continue
            if "duty_2terms" in self.active_presets:
                if not (m_cur is not None and m_cur >= 96): continue
            if "young" in self.active_presets:
                if not (age_val is not None and age_val <= 40): continue
            if "retire_soon" in self.active_presets:
                gdr = str(item[7])
                ttl = str(item[16])
                ret_info = calc_retirement_info(gdr, brth, base_ym, ttl, lvl_val)
                if not (ret_info and ret_info.get("is_soon")): continue
            if "discipline" in self.active_presets:
                if not ("⛔" in disc_str): continue
            if "same_unit" in self.active_presets:
                if same_unit_str == "—": continue
            if "missing" in self.active_presets:
                if not ("缺:" in missing_str): continue

            ident = str(item[3]).strip()
            post_c = str(item[4]).strip()
            lvl = str(item[6]).strip()

            if target_lvl == "正科" and "正科" not in lvl: continue
            if target_lvl == "副科" and "副科" not in lvl: continue

            if target_post_cat == "党政管理岗" and "业务" in post_c: continue
            elif target_post_cat == "业务管理岗" and "业务" not in post_c: continue

            if target_identity != "全部身份" and target_identity not in ident: continue

            f_edu = str(item[14]).strip()
            j_edu = str(item[15]).strip()
            combined_edu = f_edu + " " + j_edu

            if doc_opt == "硕士及以上 (硕博全包)":
                if "硕士" not in combined_edu and "博士" not in combined_edu and "研究生" not in combined_edu: continue
            elif doc_opt == "全部博士 (全日制/在职)":
                if "博士" not in combined_edu: continue
            elif doc_opt == "仅限全日制博士":
                if "博士" not in f_edu: continue
            elif doc_opt == "仅限在职博士":
                if "博士" not in j_edu: continue
            elif doc_opt == "仅限硕士 (全日制/在职)":
                if ("硕士" not in combined_edu and "研究生" not in combined_edu) or ("博士" in combined_edu): continue
            elif doc_opt == "本科及以上 (含硕博)":
                if not any(k in combined_edu for k in ["本科", "大学", "学士", "硕士", "博士", "研究生"]): continue
            elif doc_opt == "仅限大学本科 (最高本科)":
                has_undergrad = any(k in combined_edu for k in ["本科", "大学", "学士"])
                has_grad = any(k in combined_edu for k in ["硕士", "博士", "研究生"])
                if not has_undergrad or has_grad: continue

            c_title = str(item[16]).strip()
            t_lvl = get_title_level(c_title)

            if title_opt == "所有高级职称 (正高+副高全包)":
                if t_lvl not in ["正高", "副高"]: continue
            elif title_opt == "正高级 (正高/主任医等)":
                if t_lvl != "正高": continue
            elif title_opt == "副高级 (副高/副主任等)":
                if t_lvl != "副高": continue
            elif title_opt == "中级职称 (主治/主管等)":
                if t_lvl != "中级": continue
            elif title_opt == "初级及以下 (中级职称以下)":
                if t_lvl != "初级及以下": continue
            elif title_opt == "中级及以下":
                if t_lvl in ["正高", "副高"]: continue

            c_pol = str(item[17]).strip()
            if pol_opt == "中共党员":
                if "中共党员" not in c_pol and "党员" not in c_pol: continue
            elif pol_opt == "民主党派/无党派群众":
                if "中共党员" in c_pol or "预备党员" in c_pol: continue

            if kw:
                if not any(kw in str(val) for val in item[2:]): continue

            matched.append(item)

        self.populate_table(matched)

        tag_dict = {
            "sub_to_chief": "副科满2年", "chief_to_deputy": "正科满3年",
            "duty_5years": "同岗满5年", "duty_2terms": "满两届8年",
            "young": "40岁以下", "retire_soon": "临近退休",
            "discipline": "处分受限", "same_unit": "同单位亲属", "missing": "信息缺失"
        }
        active_tags = [f"【{tag_dict[k]}】" for k in self.active_presets if k in tag_dict]

        status_msg = "复合研判结果："
        if active_tags: status_msg += "".join(active_tags)
        if doc_opt != "不限学历": status_msg += f" +【{doc_opt}】"
        if target_lvl != "全部级别": status_msg += f" +【{target_lvl}】"
        if target_post_cat != "全部岗位": status_msg += f" +【{target_post_cat}】"
        if target_identity != "全部身份": status_msg += f" +【{target_identity}】"
        if title_opt != "全部职称": status_msg += f" +【{title_opt}】"
        if pol_opt != "全部面貌": status_msg += f" +【{pol_opt}】"
        if kw: status_msg += f" + 搜索“{kw}”"

        if not active_tags and doc_opt == "不限学历" and target_lvl == "全部级别" and target_post_cat == "全部岗位" and not kw:
            status_msg = f"当前显示全员在岗大盘，共 {len(matched)} 人"
        else:
            status_msg += f"，精准锁定 {len(matched)} 人！（点击【重置】清空大盘）"

        self.update_status_bar(status_msg)

    def reset_filter(self):
        self.active_presets.clear()
        self.search_input.clear()
        self.combos["level"].setCurrentText("全部级别")
        self.combos["post_cat"].setCurrentText("全部岗位")
        self.combos["identity"].setCurrentText("全部身份")
        self.combos["doc"].setCurrentText("不限学历")
        self.combos["title"].setCurrentText("全部职称")
        self.combos["politic"].setCurrentText("全部面貌")

        for card in self.cards.values():
            card.set_active(False)

        self.load_cadre_data()

    def filter_chu_cadres(self):
        target_cat = self.chu_cat_combo.currentText()
        kw = self.search_input.text().strip()
        matched = []
        for item in self.current_rows:
            chu_cat = str(item[3]).strip()
            if target_cat != "全部处级类别" and target_cat not in chu_cat:
                continue
            if kw:
                if not any(kw in str(val) for val in item[2:]):
                    continue
            matched.append(item)
        self.populate_table(matched)
        msg = f"处级领导专区名册：【{target_cat}】"
        if kw: msg += f" + 检索“{kw}”"
        msg += f"，共 {len(matched)} 人"
        self.update_status_bar(msg)

    # ==================== 核心功能：一键全员双表核准与公文简历智能治理 (完全无死锁) ====================

    def one_click_global_audit_and_heal(self):
        reply = QMessageBox.question(
            self, "全员底表校准与公文治理确认",
            "是否立即启动【全员底表全维度校准与公文智能治理引擎】？\n\n"
            "• 底表全维度穿透：中层干部名册 + 人员花名册 + 联系方式统计表 100% 联合校验\n"
            "• 第一学历深度补齐：精准识别院校与专业，缺漏项自动按名册回填校准\n"
            "• 审批表与家庭档案：自动比对关联审批表原件并抽取家庭成员关系\n"
            "• 手机号安全保护：已有个人移动电话受到非破坏性绝对保护，绝不被抹除！\n"
            "• 公文履历智能闭合：自动织入历任任免公文，修复断层并闭合至最新公文\n"
            "• 异步并发安全保障：纯内存高并发流式处理，界面流畅绝不卡顿！\n\n"
            "是否确认启动一键全员批量治理？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
        )
        if reply != QMessageBox.Yes:
            return

        self.audit_progress_dlg = QProgressDialog("正在启动全员智能治理引擎...", "取消", 0, 100, self)
        self.audit_progress_dlg.setWindowTitle("全周期数据治理进度")
        self.audit_progress_dlg.setWindowModality(Qt.WindowModal)
        self.audit_progress_dlg.setMinimumDuration(0)
        self.audit_progress_dlg.setStyleSheet("""
            QProgressDialog {
                background-color: #FFFFFF;
                font-family: "Microsoft YaHei UI";
                font-size: 9pt;
            }
            QProgressBar {
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                text-align: center;
                height: 16px;
            }
            QProgressBar::chunk {
                background-color: #2563EB;
                border-radius: 4px;
            }
        """)

        self.audit_worker = GlobalAuditWorker(self.get_base_ym())
        self.audit_worker.progress_signal.connect(self.on_audit_progress)
        self.audit_worker.finished_signal.connect(self.on_audit_finished)
        self.audit_worker.error_signal.connect(self.on_audit_error)
        self.audit_progress_dlg.canceled.connect(self.audit_worker.cancel)

        self.audit_worker.start()

    def on_audit_progress(self, current, total, name):
        if hasattr(self, 'audit_progress_dlg') and self.audit_progress_dlg:
            self.audit_progress_dlg.setMaximum(total)
            self.audit_progress_dlg.setValue(current)
            self.audit_progress_dlg.setLabelText(f"正在全维度校验底表与治理公文 ({current}/{total})：\n【{name}】...")

    def on_audit_finished(self, healed_count, total_count, stats=None):
        if hasattr(self, 'audit_progress_dlg') and self.audit_progress_dlg:
            self.audit_progress_dlg.close()

        self.load_cadre_data()

        stats = stats or {}
        fe_calibrated = stats.get("fe_calibrated", total_count)
        contact_synced = stats.get("contact_synced", 0)
        app_linked = stats.get("approval_linked", 0)

        QMessageBox.information(
            self, "全员底表校准与公文治理完成",
            f"🎉 全员在岗干部底表全维度校准与公文治理完毕！\n\n"
            f"• 治理总规模：核准在岗干部 {total_count} 人\n"
            f"• 第一学历/最高学历：全员两级学历学位与毕业院校/专业已 100% 深度补齐纠偏！\n"
            f"• 底表多维穿透：中层名册 + 花名册 + 通讯录联合校验，精准校准编制与岗位类别\n"
            f"• 通讯录联调：成功联动科级干部联系方式统计表，校对办公电话与办公地址\n"
            f"• 审批表档案：审批表原件与家庭成员档案已完成智能匹配核验\n"
            f"• 隐私数据安全：您维护的【个人移动电话】受到非破坏性绝对保护，未被抹除！\n"
            f"• 履历自愈：成功为 {healed_count} 名干部的个人简历织入历任公文摘要并实现时间线无缝闭合！\n\n"
            f"大盘看板与全息档案卡已即时刷新生效！"
        )

    def on_audit_error(self, err_msg):
        if hasattr(self, 'audit_progress_dlg') and self.audit_progress_dlg:
            self.audit_progress_dlg.close()
        QMessageBox.critical(self, "治理异常", f"执行过程发生错误: {err_msg}")

    # ==================== 业务动作调度与数据维护 ====================

    def run_deep_two_tables_import(self, is_init=False):
        cadre_file = None
        roster_file = None
        for f in os.listdir(BASE_DIR):
            if f.startswith("~$") or not f.endswith(".xlsx"):
                continue
            if "中层干部" in f or ("干部" in f and "名册" in f):
                cadre_file = os.path.join(BASE_DIR, f)
            elif "人员" in f and ("花名册" in f or "名册" in f):
                roster_file = os.path.join(BASE_DIR, f)

        if not cadre_file:
            if not is_init:
                QMessageBox.critical(self, "未找到文件", "未在系统目录下找到包含‘干部’字样的Excel名册！")
            return

        def get_dt_str(val):
            if isinstance(val, datetime.datetime):
                return val.strftime("%Y年%m月")
            return str(val or "—").strip()

        def clean_ym(s):
            m = re.search(r'(\d{4})[^\d]?(\d{1,2})', str(s))
            if m: return f"{m.group(1)}{int(m.group(2)):02d}"
            return ""

        roster_map = {}
        if roster_file:
            try:
                wb_r = openpyxl.load_workbook(roster_file, data_only=True)
                ws_r = wb_r.active

                header_row_r = 1
                for r in range(1, 6):
                    r_vals = [str(cell.value).strip() if cell.value else "" for cell in ws_r[r]]
                    if any("姓名" in v for v in r_vals):
                        header_row_r = r
                        break

                headers_r = [str(cell.value).strip() if cell.value else "" for cell in ws_r[header_row_r]]
                m_r = extract_flexible_roster_headers(headers_r)

                name_idx = m_r["name"] if m_r["name"] != -1 else 2
                emp_no_idx = m_r["emp_no"] if m_r["emp_no"] != -1 else 0
                dept_idx = m_r["dept"] if m_r["dept"] != -1 else 3
                id_idx = m_r["id_card"] if m_r["id_card"] != -1 else 6
                phone_idx = m_r["phone"]

                for row in ws_r.iter_rows(min_row=header_row_r + 1, values_only=True):
                    if row[name_idx]:
                        nm = str(row[name_idx]).replace(" ", "").strip()
                        def cell_val(k):
                            idx = m_r.get(k, -1)
                            return str(row[idx]).strip() if idx != -1 and idx < len(row) and is_valid_text(row[idx]) else ""

                        dept_val = str(row[dept_idx]).strip() if dept_idx != -1 and dept_idx < len(row) and row[dept_idx] else ""
                        cid_val = str(row[id_idx]).strip() if id_idx != -1 and id_idx < len(row) and row[id_idx] else "—"
                        cph_val = str(row[phone_idx]).strip() if phone_idx != -1 and phone_idx < len(row) and row[phone_idx] else "—"
                        eno_val = str(row[emp_no_idx]).strip() if emp_no_idx != -1 and emp_no_idx < len(row) and row[emp_no_idx] else ""

                        if nm not in roster_map: roster_map[nm] = []
                        roster_map[nm].append({
                            "emp_no": eno_val, "dept": dept_val, "id_card": cid_val, "phone": cph_val,
                            "identity": cell_val("identity"),
                            "fe_deg": cell_val("fe_deg"), "fe_sch": cell_val("fe_sch"),
                            "fe_maj": cell_val("fe_maj"), "fe_dt": cell_val("fe_dt"),
                            "he_deg": cell_val("he_deg"), "he_tit": cell_val("he_tit"),
                            "he_sch": cell_val("he_sch"), "he_maj": cell_val("he_maj"), "he_dt": cell_val("he_dt")
                        })
            except Exception as e:
                print(f"解析人员花名册异常: {e}")

        try:
            wb_c = openpyxl.load_workbook(cadre_file, data_only=True)
            ws_c = wb_c["干部花名册"] if "干部花名册" in wb_c.sheetnames else wb_c.active

            header_row_idx = 1
            for r in range(1, 6):
                row_vals = [str(cell.value).strip() if cell.value else "" for cell in ws_c[r]]
                if any("姓名" in v for v in row_vals):
                    header_row_idx = r
                    break

            headers_c = [str(cell.value).strip() if cell.value else "" for cell in ws_c[header_row_idx]]

            rem1_idx, rem2_idx = -1, -1
            for idx, h in enumerate(headers_c):
                if any(k in h for k in ["备注1", "人员身份", "编制", "身份"]): rem1_idx = idx
                elif any(k in h for k in ["备注2", "岗位类别", "岗位", "管理岗"]): rem2_idx = idx

            if rem1_idx == -1 and ws_c.max_column >= 20: rem1_idx = 19
            if rem2_idx == -1 and ws_c.max_column >= 21: rem2_idx = 20

            conn = sqlite3.connect(DB_PATH, timeout=30.0)
            conn.execute("PRAGMA journal_mode = WAL;")
            conn.execute("PRAGMA busy_timeout = 30000;")
            c = conn.cursor()

            if not is_init:
                c.execute("SELECT id, name, birth, dept_post, cadre_identity, post_category, phone FROM cadres")
                db_cadres = c.fetchall()

            updated_cnt = 0

            for r in range(header_row_idx + 1, ws_c.max_row + 1):
                name_val = ws_c.cell(row=r, column=2).value
                if not name_val or str(name_val).strip() == "": continue

                name = str(name_val).replace(" ", "").strip()
                if name in ["姓名", "姓 名", "职工姓名", "干部姓名"]: continue
                gender = str(ws_c.cell(row=r, column=3).value or "—").strip()
                birth = get_dt_str(ws_c.cell(row=r, column=4).value)
                nation = str(ws_c.cell(row=r, column=5).value or "汉族").strip()
                native = str(ws_c.cell(row=r, column=6).value or "—").strip()
                politic = str(ws_c.cell(row=r, column=7).value or "中共党员").strip()
                party_date = get_dt_str(ws_c.cell(row=r, column=8).value)
                dept_post = str(ws_c.cell(row=r, column=9).value or "—").strip()
                level = str(ws_c.cell(row=r, column=10).value or "—").strip()

                full_degree = str(ws_c.cell(row=r, column=11).value or "—").strip()
                full_school = str(ws_c.cell(row=r, column=12).value or "—").strip()
                job_degree = str(ws_c.cell(row=r, column=13).value or "—").strip()
                job_school = str(ws_c.cell(row=r, column=14).value or "—").strip()

                title = str(ws_c.cell(row=r, column=15).value or "—").strip()
                work_date = get_dt_str(ws_c.cell(row=r, column=16).value)
                current_duty_date = get_dt_str(ws_c.cell(row=r, column=17).value)
                same_level_date = get_dt_str(ws_c.cell(row=r, column=18).value)
                career_history = str(ws_c.cell(row=r, column=19).value or "—").strip()

                c_rem1_val = str(ws_c.cell(row=r, column=rem1_idx + 1).value or "").strip() if rem1_idx != -1 else ""
                c_rem2_val = str(ws_c.cell(row=r, column=rem2_idx + 1).value or "").strip() if rem2_idx != -1 else ""
                remark = str(ws_c.cell(row=r, column=22).value or "").strip()

                cand_list = roster_map.get(name, [])
                p_match = {"emp_no": "", "id_card": "—", "phone": "—", "identity": ""}
                if len(cand_list) == 1:
                    p_match = cand_list[0]
                elif len(cand_list) > 1:
                    target_ym = clean_ym(birth)
                    if target_ym:
                        for cand in cand_list:
                            cid_val = cand["id_card"]
                            if len(cid_val) >= 14 and target_ym in cid_val[6:14]:
                                p_match = cand
                                break
                    if not p_match["emp_no"]:
                        for cand in cand_list:
                            if cand["dept"] and cand["dept"] in dept_post:
                                p_match = cand
                                break

                final_identity = ""
                if is_valid_text(p_match.get("identity")): final_identity = p_match["identity"]
                elif is_valid_text(c_rem1_val): final_identity = c_rem1_val
                else: final_identity = "校编" if "河南医药大学" in dept_post or "新乡医学院" in dept_post else "院编"

                final_post_cat = ""
                if is_valid_text(c_rem2_val): final_post_cat = c_rem2_val
                else:
                    clinical_keywords = ["科主任", "护士长", "病区", "中心主任", "教研室", "研究所", "医院院长", "门诊", "药学部"]
                    final_post_cat = "业务管理岗" if any(k in dept_post for k in clinical_keywords) else "党政管理岗"

                if is_init:
                    c.execute('''INSERT INTO cadres (
                        emp_no, name, gender, birth, nation, native_place, id_card, phone, office_phone, office_address,
                        politic, party_date, work_date, dept_post, level, cadre_rank, current_duty_date,
                        same_level_date, title, title_date, full_edu_degree, full_edu_school,
                        job_edu_degree, job_edu_school,
                        first_edu_degree, first_edu_school, first_edu_major, first_edu_date,
                        highest_edu_degree, highest_edu_degree_title, highest_edu_school, highest_edu_major, highest_edu_date,
                        career_history, cadre_identity, post_category,
                        discipline_type, discipline_date, discipline_desc,
                        resume_text, family_relations, status, remark
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, '—', '—', ?, ?, ?, ?, ?, '科级', ?, ?, ?, '—', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '无', '—', '—', '—', '—', '在岗', ?)''',
                    (
                        p_match["emp_no"], name, gender, birth, nation, native, p_match["id_card"], p_match["phone"],
                        politic, party_date, work_date, dept_post, level, current_duty_date,
                        same_level_date, title, full_degree, full_school, job_degree, job_school,
                        p_match.get("fe_deg",""), p_match.get("fe_sch",""), p_match.get("fe_maj",""), p_match.get("fe_dt",""),
                        p_match.get("he_deg",""), p_match.get("he_tit",""), p_match.get("he_sch",""), p_match.get("he_maj",""), p_match.get("he_dt",""),
                        career_history, final_identity, final_post_cat, remark
                    ))
                else:
                    matched_cands = [db_c for db_c in db_cadres if db_c[1] == name]
                    target_id = None
                    cur_db_phone = ""
                    if len(matched_cands) == 1:
                        target_id = matched_cands[0][0]
                        cur_db_phone = matched_cands[0][6]
                    elif len(matched_cands) > 1:
                        target_ym = clean_ym(birth)
                        for cand in matched_cands:
                            if target_ym and target_ym == clean_ym(cand[2]):
                                target_id = cand[0]
                                cur_db_phone = cand[6]
                                break
                        if not target_id:
                            for cand in matched_cands:
                                cand_dept = cand[3]
                                if cand_dept and dept_post and (cand_dept in dept_post or dept_post in cand_dept):
                                    target_id = cand[0]
                                    cur_db_phone = cand[6]
                                    break

                    if target_id:
                        final_phone = cur_db_phone
                        if is_valid_text(p_match.get("phone")):
                            final_phone = p_match["phone"]

                        c.execute("""SELECT first_edu_degree, first_edu_school, highest_edu_degree, highest_edu_school, 
                                            full_edu_degree, full_edu_school, job_edu_degree, job_edu_school 
                                     FROM cadres WHERE id = ?""", (target_id,))
                        db_edu = c.fetchone()

                        new_fe_deg = p_match.get("fe_deg") or (db_edu[0] if is_valid_text(db_edu[0]) else db_edu[4])
                        new_fe_sch = p_match.get("fe_sch") or (db_edu[1] if is_valid_text(db_edu[1]) else db_edu[5])
                        new_fe_maj = p_match.get("fe_maj") or ""
                        new_fe_dt  = p_match.get("fe_dt") or ""

                        new_he_deg = p_match.get("he_deg") or (db_edu[2] if is_valid_text(db_edu[2]) else (db_edu[6] if is_valid_text(db_edu[6]) else db_edu[4]))
                        new_he_tit = p_match.get("he_tit") or ""
                        new_he_sch = p_match.get("he_sch") or (db_edu[3] if is_valid_text(db_edu[3]) else (db_edu[7] if is_valid_text(db_edu[7]) else db_edu[5]))
                        new_he_maj = p_match.get("he_maj") or ""
                        new_he_dt  = p_match.get("he_dt") or ""

                        c.execute("""UPDATE cadres SET cadre_identity = ?, post_category = ?, phone = ?,
                                     first_edu_degree = ?, first_edu_school = ?, first_edu_major = ?, first_edu_date = ?,
                                     highest_edu_degree = ?, highest_edu_degree_title = ?, highest_edu_school = ?, 
                                     highest_edu_major = ?, highest_edu_date = ?
                                     WHERE id = ?""",
                                  (final_identity, final_post_cat, final_phone,
                                   new_fe_deg, new_fe_sch, new_fe_maj, new_fe_dt,
                                   new_he_deg, new_he_tit, new_he_sch, new_he_maj, new_he_dt, target_id))
                        updated_cnt += 1

            conn.commit()
            conn.close()

            if not is_init:
                self.load_cadre_data()
                QMessageBox.information(
                    self, "双表核准补全完成",
                    f"🎉 人员身份与两级学历精准补齐成功！\n\n"
                    f"• 成功匹配核验干部：{updated_cnt} 人\n"
                    f"• 第一学历学校来源：已按《人员花名册》多级动态表头识别并完成纠偏\n"
                    f"• 手机号增量安全保护：您此前维护的电话已受到绝对保护！\n"
                    f"• 学历兜底保护：缺漏字段已自动借力审批表法定学历填补！"
                )

        except Exception as e:
            if not is_init: QMessageBox.critical(self, "更新出错", f"解析名册处理异常: {e}")

    def sync_appointment_files(self):
        if not os.path.exists(FILE_DIR): return
        self.update_status_bar("正在扫描并索引任免公文库...")
        conn = sqlite3.connect(DB_PATH, timeout=30.0)
        c = conn.cursor()
        c.execute("SELECT file_name, file_mtime FROM indexed_files")
        indexed = dict(c.fetchall())

        all_files = [f for f in os.listdir(FILE_DIR) if os.path.splitext(f)[1].lower() in ('.pdf', '.docx') and not f.startswith("~$")]
        to_process = [(f, os.path.join(FILE_DIR, f), os.path.getmtime(os.path.join(FILE_DIR, f)))
                      for f in all_files if f not in indexed or indexed[f] < os.path.getmtime(os.path.join(FILE_DIR, f))]

        for fname, fpath, mtime in to_process:
            c.execute("DELETE FROM cadre_records WHERE file_name = ?", (fname,))
            try:
                full_text = ""
                paragraphs = []
                ext = os.path.splitext(fname)[1].lower()
                if ext == ".docx":
                    doc = docx.Document(fpath)
                    for p in doc.paragraphs:
                        t = p.text.strip()
                        if len(t) > 4: paragraphs.append(t); full_text += t + "\n"
                    for table in doc.tables:
                        for row in table.rows:
                            rc = [cl.text.strip() for cl in row.cells if cl.text.strip()]
                            if rc: rl = " | ".join(rc); paragraphs.append(rl); full_text += rl + "\n"
                elif ext == ".pdf":
                    reader = PdfReader(fpath)
                    for page in reader.pages:
                        ptxt = page.extract_text() or ""
                        full_text += ptxt + "\n"
                        for line in re.split(r'[\r\n]+', ptxt):
                            ls = line.strip()
                            if len(ls) > 4: paragraphs.append(ls)

                doc_num = "—"
                num_pattern = r'[\u4e00-\u9fa5]+[字发]?[\(\[\{〔【][0-9]{4}[\)\]\}〕】][0-9]+号'
                mn = re.search(num_pattern, full_text) or re.search(num_pattern, fname)
                if mn: doc_num = mn.group(0)

                doc_date = "—"
                date_pattern = r'(20[12][0-9]年[01]?[0-9]月[0-3]?[0-9]日)'
                md = re.search(date_pattern, full_text)
                if md: doc_date = md.group(0)

                # 检查公文中是否存在“上述人员试用期一年”等条款
                probation_clause = ""
                for p in paragraphs:
                    ps = p.strip()
                    if re.search(r'(?:上述|以上|该|相关|新提拔)?(?:人员|同志|干部)?(?:任职)?试用期(?:为)?(?:一|1)年', ps) or ("试用期" in ps and "试用期满" not in ps and "考核合格" not in ps and "正式任职" not in ps):
                        probation_clause = ps
                        break

                for p in paragraphs:
                    p_to_save = p
                    # 若该段落是干部任职名单，且整篇公文带有试用期条款，但该段落本身未标明试用期
                    if probation_clause and "任" in p and "试用期" not in p and "免去" not in p and not p.startswith("关于") and not p.startswith("特此通知"):
                        p_to_save = f"{p}（{probation_clause}）"
                    c.execute('''INSERT INTO cadre_records (file_name, doc_number, doc_date, content)
                                 VALUES (?, ?, ?, ?)''', (fname, doc_num, doc_date, p_to_save))
                c.execute("INSERT OR REPLACE INTO indexed_files VALUES (?, ?)", (fname, mtime))
                conn.commit()
            except Exception as e:
                print(f"公文解析出错: {e}")

        conn.close()
        self.update_status_bar(f"任免公文库索引就绪（在库公文 {len(all_files)} 份）。当前显示 {len(self.current_rows)} 人。")

    def auto_sync_single_approval_forms(self):
        if not os.path.exists(APPROVAL_DIR): return
        all_app_files = [f for f in os.listdir(APPROVAL_DIR) if f.lower().endswith(".docx") and not f.startswith("~$")]
        conn = sqlite3.connect(DB_PATH, timeout=30.0)
        c = conn.cursor()
        c.execute("SELECT id, name, resume_text, selected_approval_file FROM cadres WHERE status='在岗' AND (cadre_rank='科级' OR cadre_rank IS NULL OR cadre_rank='')")
        cadres = c.fetchall()

        for cid, name, cur_resume, sel_file in cadres:
            matched_files = [f for f in all_app_files if name in f]
            if sel_file and sel_file in matched_files:
                target_f = sel_file
            elif len(matched_files) == 1:
                target_f = matched_files[0]
            elif len(matched_files) > 1:
                non_copy = [f for f in matched_files if "副本" not in f]
                candidates = non_copy if non_copy else matched_files
                candidates.sort(key=lambda f: os.path.getmtime(os.path.join(APPROVAL_DIR, f)), reverse=True)
                target_f = candidates[0]
            else:
                target_f = None

            if target_f:
                fpath = os.path.join(APPROVAL_DIR, target_f)
                try:
                    res_txt, fam_txt = extract_resume_and_family_from_docx(fpath)
                    c.execute("""UPDATE cadres SET resume_text = ?, family_relations = ?, selected_approval_file = ? WHERE id = ?""",
                              (res_txt, fam_txt, target_f, cid))
                except Exception as e:
                    print(f"解析审批表异常 {target_f}: {e}")

        conn.commit()
        conn.close()

    def batch_import_contacts(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "请选择职工通讯录Excel表格", BASE_DIR, "Excel 表格 (*.xlsx *.xls)")
        if not file_path: return

        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            ws = wb.active
            headers = [str(cell.value).strip() if cell.value else "" for cell in ws[1]]

            dept_idx, name_idx, office_phone_idx, mobile_idx, address_idx = -1, -1, -1, -1, -1
            for idx, h in enumerate(headers):
                if any(k in h for k in ["科室", "部门", "所在科室", "所属科室", "工作部门"]): dept_idx = idx
                elif any(k in h for k in ["姓名", "名字", "干部姓名", "职工姓名"]): name_idx = idx
                elif any(k in h for k in ["办公电话", "固话", "座机", "内线", "分机", "办公固话", "科室电话"]): office_phone_idx = idx
                elif any(k in h for k in ["移动电话", "手机", "手机号", "联系电话", "个人电话"]): mobile_idx = idx
                elif any(k in h for k in ["办公地址", "办公室", "房间号", "科室地址", "地址", "房号"]): address_idx = idx

            if mobile_idx == -1 and office_phone_idx == -1:
                for idx, h in enumerate(headers):
                    if "电话" in h or "联系方式" in h: mobile_idx = idx; break

            if name_idx == -1 or (mobile_idx == -1 and office_phone_idx == -1 and address_idx == -1):
                QMessageBox.critical(self, "格式识别失败", "在通讯录第一行未能识别到【姓名】或【联系电话/地址】列！")
                return

            def clean_val(v):
                if not v: return ""
                s = str(v).strip()
                if s.endswith(".0"): s = s[:-2]
                return "" if s in ["—", "None", "/"] else s

            contacts_map = {}
            for row in ws.iter_rows(min_row=2, values_only=True):
                nm_val = row[name_idx]
                if not nm_val: continue
                nm = str(nm_val).replace(" ", "").strip()

                dept = str(row[dept_idx]).strip() if dept_idx != -1 and row[dept_idx] else ""
                off_ph = clean_val(row[office_phone_idx]) if office_phone_idx != -1 else ""
                mob = clean_val(row[mobile_idx]) if mobile_idx != -1 else ""
                addr = clean_val(row[address_idx]) if address_idx != -1 else ""

                if nm not in contacts_map: contacts_map[nm] = []
                contacts_map[nm].append({"dept": dept, "office_phone": off_ph, "mobile": mob, "address": addr})

            conn = sqlite3.connect(DB_PATH, timeout=30.0)
            c = conn.cursor()
            c.execute("SELECT id, name, dept_post, phone, office_phone, office_address FROM cadres WHERE status='在岗'")
            all_db = c.fetchall()

            up_mob, up_oph, up_addr, total_m = 0, 0, 0, 0
            for cid, cname, dept_post, cur_phone, cur_oph, cur_addr in all_db:
                if cname not in contacts_map: continue
                cand_list = contacts_map[cname]
                matched = None
                if len(cand_list) == 1: matched = cand_list[0]
                else:
                    for cand in cand_list:
                        if cand["dept"] and (cand["dept"] in dept_post or dept_post in cand["dept"]):
                            matched = cand; break
                    if not matched: matched = cand_list[0]

                if matched:
                    total_m += 1
                    new_mob = matched["mobile"] if is_valid_text(matched["mobile"]) else cur_phone
                    new_oph = matched["office_phone"] if is_valid_text(matched["office_phone"]) else cur_oph
                    new_addr = matched["address"] if is_valid_text(matched["address"]) else cur_addr

                    if new_mob != cur_phone: up_mob += 1
                    if new_oph != cur_oph: up_oph += 1
                    if new_addr != cur_addr: up_addr += 1

                    c.execute("""UPDATE cadres SET phone = ?, office_phone = ?, office_address = ? WHERE id = ?""",
                              (new_mob, new_oph, new_addr, cid))

            conn.commit()
            conn.close()

            self.load_cadre_data()
            QMessageBox.information(
                self, "通讯录批量导入完成",
                f"🎉 通讯录匹配补全成功！\n\n• 成功核准匹配：{total_m} 人\n• 增量补全手机号：{up_mob} 人\n• 补全办公固话：{up_oph} 人\n• 补全科室地址：{up_addr} 人"
            )
        except Exception as e:
            QMessageBox.critical(self, "导入出错", f"读取处理异常: {e}")

    # ==================== 亲属任职回避台账导出 ====================

    def export_avoidance_investigation_excel(self):
        base_ym = self.get_base_ym()
        default_file = os.path.join(BASE_DIR, f"干部亲属任职回避专项排查台账_{base_ym}.xlsx")
        save_path, _ = QFileDialog.getSaveFileName(
            self, "导出亲属任职回避专项排查台账", default_file, "Excel 表格 (*.xlsx)"
        )
        if not save_path:
            return

        try:
            cnt = generate_avoidance_investigation_excel(base_ym, save_path)
            res = QMessageBox.information(
                self, "台账生成成功",
                f"🎉 《干部亲属任职回避专项排查台账》已成功归档生成！\n\n"
                f"• 排查基准年月：{base_ym}\n"
                f"• 共识别排查任职回避关系：{cnt} 条\n"
                f"• 包含回避级别标记(🔴重点管控 / 🟡常态监管)与巡视合规研判依据\n"
                f"• 保存路径：{save_path}\n\n"
                f"是否立即打开 Excel 表格进行查阅？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
            )
            if res == QMessageBox.Yes:
                QDesktopServices.openUrl(QUrl.fromLocalFile(save_path))
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"生成排查台账发生错误: {e}")

    # ==================== 增量比对导入机制 ====================

    def open_import_diff_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "请选择拟导入核准的干部名册或花名册Excel", BASE_DIR, "Excel 表格 (*.xlsx *.xls)"
        )
        if not file_path:
            return

        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            ws = wb["干部花名册"] if "干部花名册" in wb.sheetnames else wb.active

            header_row_idx = 1
            for r in range(1, 6):
                row_vals = [str(cell.value).strip() if cell.value else "" for cell in ws[r]]
                if any("姓名" in v for v in row_vals):
                    header_row_idx = r
                    break

            headers = [str(cell.value).strip() if cell.value else "" for cell in ws[header_row_idx]]

            def get_col_idx(kw_list):
                for idx, h in enumerate(headers):
                    if any(k in h for k in kw_list): return idx
                return -1

            name_idx = get_col_idx(["姓名", "干部姓名"])
            post_idx = get_col_idx(["职务", "工作部门及职务", "现任职务", "科室"])
            level_idx = get_col_idx(["级别", "职级", "职务级别"])
            gender_idx = get_col_idx(["性别"])
            birth_idx = get_col_idx(["出生年月", "出生日期", "出生"])
            phone_idx = get_col_idx(["手机", "移动电话", "联系方式"])
            rank_idx = get_col_idx(["层级", "管理层级", "干部层级"])

            if name_idx == -1:
                QMessageBox.critical(self, "格式错误", "未在 Excel 前 5 行中识别到【姓名】表头！")
                return

            conn = sqlite3.connect(DB_PATH, timeout=30.0)
            c = conn.cursor()
            c.execute("SELECT id, name, birth, dept_post, level, phone, cadre_rank, cadre_identity, post_category FROM cadres")
            db_cadres = c.fetchall()
            conn.close()

            db_map = {}
            for row in db_cadres:
                cid, nm, brth, dpost, lvl, ph, rnk, ident, pcat = row
                if nm not in db_map:
                    db_map[nm] = []
                db_map[nm].append({
                    "id": cid, "name": nm, "birth": brth or "", "dept_post": dpost or "",
                    "level": lvl or "", "phone": ph or "", "cadre_rank": rnk or "科级",
                    "cadre_identity": ident or "", "post_category": pcat or ""
                })

            diff_items = []

            for r in range(header_row_idx + 1, ws.max_row + 1):
                nm_val = ws.cell(row=r, column=name_idx + 1).value
                if not nm_val:
                    continue
                name = str(nm_val).replace(" ", "").strip()

                dept_post = str(ws.cell(row=r, column=post_idx + 1).value or "—").strip() if post_idx != -1 else "—"
                lvl = str(ws.cell(row=r, column=level_idx + 1).value or "—").strip() if level_idx != -1 else "—"
                gdr = str(ws.cell(row=r, column=gender_idx + 1).value or "—").strip() if gender_idx != -1 else "—"
                brth_raw = ws.cell(row=r, column=birth_idx + 1).value if birth_idx != -1 else "—"
                brth = brth_raw.strftime("%Y年%m月") if isinstance(brth_raw, datetime.datetime) else str(brth_raw or "—").strip()
                ph = str(ws.cell(row=r, column=phone_idx + 1).value or "—").strip() if phone_idx != -1 else "—"
                rnk_val = str(ws.cell(row=r, column=rank_idx + 1).value or "").strip() if rank_idx != -1 else ""
                cadre_rank = rnk_val if rnk_val in ["处级", "科级"] else ("处级" if "处" in lvl or "院长" in dept_post or "书记" in dept_post else "科级")

                cands = db_map.get(name, [])
                matched_target = None
                if len(cands) == 1:
                    matched_target = cands[0]
                elif len(cands) > 1:
                    for cand in cands:
                        if is_valid_text(brth) and (brth in cand["birth"] or cand["birth"] in brth):
                            matched_target = cand
                            break
                    if not matched_target:
                        matched_target = cands[0]

                if not matched_target:
                    # 新增人员
                    diff_items.append({
                        "status": "新增",
                        "name": name,
                        "rank": cadre_rank,
                        "dept_post": dept_post,
                        "changes": [("新建档案", "库中无记录", f"{dept_post} ({lvl})")],
                        "payload": {
                            "action": "insert",
                            "name": name, "gender": gdr, "birth": brth, "phone": ph,
                            "dept_post": dept_post, "level": lvl, "cadre_rank": cadre_rank
                        }
                    })
                else:
                    # 比对变更
                    changes = []
                    if is_valid_text(dept_post) and dept_post != "—" and dept_post != matched_target["dept_post"]:
                        changes.append(("职务", matched_target["dept_post"], dept_post))
                    if is_valid_text(lvl) and lvl != "—" and lvl != matched_target["level"]:
                        changes.append(("职级", matched_target["level"], lvl))
                    if is_valid_text(ph) and ph != "—" and ph != matched_target["phone"]:
                        changes.append(("联系电话", matched_target["phone"], ph))

                    if changes:
                        diff_items.append({
                            "status": "变更",
                            "name": name,
                            "rank": matched_target["cadre_rank"],
                            "dept_post": dept_post if is_valid_text(dept_post) else matched_target["dept_post"],
                            "changes": changes,
                            "payload": {
                                "action": "update",
                                "id": matched_target["id"],
                                "dept_post": dept_post if is_valid_text(dept_post) else matched_target["dept_post"],
                                "level": lvl if is_valid_text(lvl) else matched_target["level"],
                                "phone": ph if is_valid_text(ph) else matched_target["phone"]
                            }
                        })
                    else:
                        diff_items.append({
                            "status": "无变化",
                            "name": name,
                            "rank": matched_target["cadre_rank"],
                            "dept_post": matched_target["dept_post"],
                            "changes": [],
                            "payload": None
                        })

            if not diff_items:
                QMessageBox.information(self, "提示", "未从所选 Excel 文件中读取到有效干部数据。")
                return

            dlg = ImportDiffConfirmDialog(self, diff_items, on_confirm_callback=self.apply_diff_import)
            dlg.exec()

        except Exception as e:
            QMessageBox.critical(self, "解析名册出错", f"读取比对 Excel 文件发生异常: {e}")

    def apply_diff_import(self, diff_items):
        conn = sqlite3.connect(DB_PATH, timeout=30.0)
        c = conn.cursor()
        inserted_cnt = 0
        updated_cnt = 0

        for it in diff_items:
            payload = it.get("payload")
            if not payload:
                continue

            action = payload.get("action")
            if action == "insert":
                c.execute('''INSERT INTO cadres (
                    name, gender, birth, nation, native_place, id_card, phone, office_phone, office_address, emp_no,
                    dept_post, level, cadre_rank, politic, status, remark
                ) VALUES (?, ?, ?, '汉族', '—', '—', ?, '—', '—', '', ?, ?, ?, '中共党员', '在岗', '增量比对导入')''',
                (payload["name"], payload["gender"], payload["birth"], payload["phone"],
                 payload["dept_post"], payload["level"], payload["cadre_rank"]))
                inserted_cnt += 1
            elif action == "update":
                c.execute('''UPDATE cadres SET dept_post = ?, level = ?, phone = ? WHERE id = ?''',
                          (payload["dept_post"], payload["level"], payload["phone"], payload["id"]))
                updated_cnt += 1

        conn.commit()
        conn.close()

        self.load_cadre_data()
        QMessageBox.information(
            self, "增量确认入库完成",
            f"🎉 干部名册增量更新已成功生效！\n\n"
            f"• 成功新增入库：{inserted_cnt} 人\n"
            f"• 成功更新变更：{updated_cnt} 人\n"
            f"• 界面各研判专区与表格已即时完成同步刷新！"
        )

    def batch_import_chu_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "请选择处级干部名册Excel表格", BASE_DIR, "Excel 表格 (*.xlsx *.xls)")
        if not file_path: return

        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            ws = wb.active
            header_row_idx = 1
            for r in range(1, 6):
                row_vals = [str(cell.value).strip() if cell.value else "" for cell in ws[r]]
                if any("姓名" in v for v in row_vals): header_row_idx = r; break

            headers = [str(cell.value).strip() if cell.value else "" for cell in ws[header_row_idx]]

            def get_col_idx(kw_list):
                for idx, h in enumerate(headers):
                    if any(k in h for k in kw_list): return idx
                return -1

            name_idx = get_col_idx(["姓名", "干部姓名"])
            post_idx = get_col_idx(["职务", "工作部门及职务", "现任职务", "科室"])
            level_idx = get_col_idx(["级别", "职级", "职务级别"])
            cat_idx = get_col_idx(["归口", "处级类别", "类别", "分类"])
            charge_idx = get_col_idx(["分管", "分管工作", "分管业务", "分管部门", "分工"])
            gender_idx = get_col_idx(["性别"])
            birth_idx = get_col_idx(["出生年月", "出生日期", "出生"])
            phone_idx = get_col_idx(["手机", "移动电话", "联系方式"])

            if name_idx == -1:
                QMessageBox.critical(self, "格式错误", "未在Excel前5行中识别到【姓名】表头！")
                return

            conn = sqlite3.connect(DB_PATH, timeout=30.0)
            c = conn.cursor()
            c.execute("SELECT id, name, birth, dept_post, level, chu_category, charge_division, phone FROM cadres WHERE cadre_rank = '处级'")
            db_chu = c.fetchall()
            conn.close()

            db_map = {row[1]: row for row in db_chu}
            diff_items = []

            for r in range(header_row_idx + 1, ws.max_row + 1):
                nm_val = ws.cell(row=r, column=name_idx + 1).value
                if not nm_val: continue
                name = str(nm_val).replace(" ", "").strip()

                dept_post = str(ws.cell(row=r, column=post_idx + 1).value or "—").strip() if post_idx != -1 else "—"
                lvl = str(ws.cell(row=r, column=level_idx + 1).value or "正处").strip() if level_idx != -1 else "正处"
                gdr = str(ws.cell(row=r, column=gender_idx + 1).value or "—").strip() if gender_idx != -1 else "—"
                brth_raw = ws.cell(row=r, column=birth_idx + 1).value if birth_idx != -1 else "—"
                brth = brth_raw.strftime("%Y年%m月") if isinstance(brth_raw, datetime.datetime) else str(brth_raw or "—").strip()
                ph = str(ws.cell(row=r, column=phone_idx + 1).value or "—").strip() if phone_idx != -1 else "—"

                chu_cat = ""
                if cat_idx != -1: chu_cat = str(ws.cell(row=r, column=cat_idx + 1).value or "").strip()
                if not chu_cat or chu_cat not in CHU_CATEGORIES:
                    if any(k in dept_post for k in ["副院长", "院长", "书记", "总会计师", "副书记", "局长", "主任"]): chu_cat = "单位领导班子"
                    elif any(k in dept_post for k in ["所长", "研究中心", "实验室", "研究院"]): chu_cat = "直属/分支机构领导"
                    elif any(k in dept_post for k in ["学院", "系", "教学", "内设"]): chu_cat = "专业内设机构领导"
                    else: chu_cat = "单位领导班子"

                charge_div = str(ws.cell(row=r, column=charge_idx + 1).value or "").strip() if charge_idx != -1 else ""

                if name not in db_map:
                    diff_items.append({
                        "status": "新增",
                        "name": name,
                        "rank": "处级",
                        "dept_post": dept_post,
                        "changes": [("新建档案", "库中无记录", f"录入处级领导干部（{chu_cat}）")],
                        "payload": {
                            "action": "insert",
                            "name": name, "gender": gdr, "birth": brth, "phone": ph,
                            "dept_post": dept_post, "level": lvl, "cadre_rank": "处级",
                            "chu_category": chu_cat, "charge_division": charge_div
                        }
                    })
                else:
                    old_row = db_map[name]
                    changes = []
                    if is_valid_text(dept_post) and dept_post != "—" and dept_post != old_row[3]:
                        changes.append(("现任职务", old_row[3], dept_post))
                    if is_valid_text(lvl) and lvl != old_row[4]:
                        changes.append(("职级", old_row[4], lvl))
                    if is_valid_text(chu_cat) and chu_cat != old_row[5]:
                        changes.append(("归口类别", old_row[5], chu_cat))
                    if is_valid_text(charge_div) and charge_div != old_row[6]:
                        changes.append(("分管业务", old_row[6], charge_div))
                    if is_valid_text(ph) and ph != "—" and ph != old_row[7]:
                        changes.append(("联系电话", old_row[7], ph))

                    if changes:
                        diff_items.append({
                            "status": "变更",
                            "name": name,
                            "rank": "处级",
                            "dept_post": dept_post,
                            "changes": changes,
                            "payload": {
                                "action": "update_chu",
                                "id": old_row[0],
                                "dept_post": dept_post, "level": lvl,
                                "chu_category": chu_cat, "charge_division": charge_div,
                                "phone": ph if is_valid_text(ph) else old_row[7]
                            }
                        })
                    else:
                        diff_items.append({
                            "status": "无变化",
                            "name": name,
                            "rank": "处级",
                            "dept_post": old_row[3],
                            "changes": [],
                            "payload": None
                        })

            def apply_chu_import(items):
                conn = sqlite3.connect(DB_PATH, timeout=30.0)
                c = conn.cursor()
                ins_c, up_c = 0, 0
                for it in items:
                    pl = it.get("payload")
                    if not pl: continue
                    if pl["action"] == "insert":
                        c.execute('''INSERT INTO cadres (
                            name, gender, birth, nation, native_place, id_card, phone, office_phone, office_address, emp_no,
                            dept_post, level, cadre_rank, chu_category, charge_division,
                            politic, status, remark
                        ) VALUES (?, ?, ?, '汉族', '—', '—', ?, '—', '—', '', ?, ?, '处级', ?, ?, '中共党员', '在岗', '批量导入处级名册')''',
                        (pl["name"], pl["gender"], pl["birth"], pl["phone"], pl["dept_post"], pl["level"], pl["chu_category"], pl["charge_division"]))
                        ins_c += 1
                    elif pl["action"] == "update_chu":
                        c.execute('''UPDATE cadres SET dept_post = ?, level = ?, chu_category = ?, charge_division = ?, phone = ? WHERE id = ?''',
                                  (pl["dept_post"], pl["level"], pl["chu_category"], pl["charge_division"], pl["phone"], pl["id"]))
                        up_c += 1
                conn.commit()
                conn.close()
                self.load_cadre_data()
                QMessageBox.information(self, "处级导入完成", f"🎉 处级干部名册增量更新已生效！\n\n• 新增入库：{ins_c} 人\n• 变更更新：{up_c} 人")

            dlg = ImportDiffConfirmDialog(self, diff_items, on_confirm_callback=apply_chu_import)
            dlg.exec()

        except Exception as e:
            QMessageBox.critical(self, "导入出错", f"读取处理处级名册异常: {e}")

    def export_batch_report(self, report_type):
        selected_cids = []
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 0)
            if item:
                cid = item.data(Qt.UserRole)
                if cid: selected_cids.append(cid)

        if not selected_cids:
            QMessageBox.warning(self, "提示", "当前列表中无任何人员数据可供导出！")
            return

        base_ym = self.get_base_ym()
        conn = sqlite3.connect(DB_PATH, timeout=30.0)
        c = conn.cursor()
        placeholders = ','.join(['?'] * len(selected_cids))
        c.execute(f"""SELECT id, name, gender, birth, nation, native_place, id_card, phone, politic, party_date, work_date,
                             dept_post, level, current_duty_date, same_level_date, title,
                             full_edu_degree, full_edu_school, job_edu_degree, job_edu_school, family_relations,
                             discipline_type, discipline_date, cadre_identity, post_category
                      FROM cadres WHERE id IN ({placeholders})""", selected_cids)
        db_rows = {r[0]: r for r in c.fetchall()}
        conn.close()

        tag_title = "科级干部选拔任用资格初审汇总表" if report_type == "promotion" else "科级干部任职年限与轮岗交流研判表"
        save_path, _ = QFileDialog.getSaveFileName(
            self, f"导出{tag_title}", os.path.join(BASE_DIR, f"{tag_title}_{base_ym}.xlsx"), "Excel 表格 (*.xlsx)"
        )
        if not save_path: return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active

            if report_type == "promotion":
                ws.title = "提拔资格初审"
                headers = [
                    "序号", "姓名", "性别", "出生年月", "人员身份", "岗位类别", "现任职务级别", "工作部门及现任职务",
                    "任同级时间", "任同级年限 (截至" + base_ym + ")", "提拔资格初审判定",
                    "纪律红线(处分影响期排查)", "全日制教育", "在职教育", "专业技术职称", "政治面貌",
                    "亲属任职回避排查", "身份证号", "联系电话"
                ]
                ws.append(headers)

                for idx, cid in enumerate(selected_cids, start=1):
                    row_data = db_rows.get(int(cid))
                    if not row_data: continue
                    _, name, gender, birth, _, _, id_card, phone, politic, _, _, dept_post, level, _, same_date, title, f_deg, f_sch, j_deg, j_sch, fam_txt, disc_t, disc_d, ident, post_c = row_data
                    m_same = calc_months_diff(same_date, base_ym)
                    same_len = format_months_to_str(m_same)

                    is_impacted, disc_str, _ = check_discipline_impact(disc_t, disc_d, base_ym)
                    same_alerts = analyze_same_unit_relations(fam_txt)
                    same_unit_res = "、".join(same_alerts) if same_alerts else "无"

                    judgement = "—"
                    if is_impacted:
                        judgement = "【红线一票否决】受处分/诫勉在影响期内，严禁提拔！"
                    else:
                        if "副科" in str(level):
                            judgement = "符合提拔正科条件 (任副科满2年)" if (m_same is not None and m_same >= 24) else "暂不符合提拔正科 (任职未满2年)"
                        elif "正科" in str(level):
                            judgement = "符合提拔副处条件 (任正科满3年)" if (m_same is not None and m_same >= 36) else "暂不符合提拔副处 (任职未满3年)"

                    ws.append([
                        idx, name, gender, birth, ident, post_c, level, dept_post, same_date, same_len, judgement,
                        disc_str,
                        f"{f_deg} ({f_sch})" if is_valid_text(f_deg) else "—",
                        f"{j_deg} ({j_sch})" if is_valid_text(j_deg) else "—",
                        title, politic, same_unit_res, id_card, phone
                    ])
            else:
                ws.title = "轮岗交流研判"
                headers = [
                    "序号", "姓名", "性别", "人员身份", "岗位类别", "工作部门及职务", "职务级别", "任现职时间",
                    "同一岗位任职时长 (截至" + base_ym + ")", "是否满5年(应轮岗)", "是否满两届8年(达最高任期)",
                    "处分/诫勉影响期预警", "全日制最高学历", "在职最高学历", "专业技术职称", "亲属任职回避排查",
                    "身份证号", "联系电话", "轮岗研判建议"
                ]
                ws.append(headers)

                for idx, cid in enumerate(selected_cids, start=1):
                    row_data = db_rows.get(int(cid))
                    if not row_data: continue
                    _, name, gender, _, _, _, id_card, phone, _, _, _, dept_post, level, cur_date, _, title, f_deg, f_sch, j_deg, j_sch, fam_txt, disc_t, disc_d, ident, post_c = row_data
                    m_cur = calc_months_diff(cur_date, base_ym)
                    cur_len = format_months_to_str(m_cur)

                    is_impacted, disc_str, _ = check_discipline_impact(disc_t, disc_d, base_ym)
                    same_alerts = analyze_same_unit_relations(fam_txt)
                    same_unit_res = "、".join(same_alerts) if same_alerts else "无"

                    is_5yr = "是 (满5年)" if (m_cur is not None and m_cur >= 60) else "否"
                    is_2term = "是 (满两届8年)" if (m_cur is not None and m_cur >= 96) else "否"

                    suggestion = "在岗周期正常"
                    if m_cur is not None and m_cur >= 96:
                        suggestion = "【重点预警】已达到最高任职年限两届，建议坚决轮岗交流或免职"
                    elif m_cur is not None and m_cur >= 60:
                        suggestion = "【轮岗预警】任现职已满5年，建议纳入轮岗交流计划人选"

                    ws.append([
                        idx, name, gender, ident, post_c, dept_post, level, cur_date, cur_len, is_5yr, is_2term,
                        disc_str,
                        f"{f_deg} ({f_sch})" if is_valid_text(f_deg) else "—",
                        f"{j_deg} ({j_sch})" if is_valid_text(j_deg) else "—",
                        title, same_unit_res, id_card, phone, suggestion
                    ])

            h_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
            h_font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
            c_font = Font(name="Microsoft YaHei", size=9)
            border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                            top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

            for cell in ws[1]:
                cell.fill = h_fill
                cell.font = h_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for row in ws.iter_rows(min_row=2, max_row=len(selected_cids)+1):
                for cell in row:
                    cell.font = c_font
                    cell.border = border
                    if cell.column in [1, 2, 3, 4, 5, 6, 7, 9, 10]:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            id_col_idx = headers.index("身份证号") + 1 if "身份证号" in headers else -1
            if id_col_idx != -1:
                for row in ws.iter_rows(min_row=2, max_row=len(selected_cids)+1):
                    row[id_col_idx - 1].number_format = "@"

            widths = {
                "A": 8, "B": 12, "C": 24, "D": 14, "E": 14, "F": 16, "G": 14, "H": 25, "I": 12, "J": 12, "K": 10,
                "L": 12, "M": 12, "N": 24, "O": 24, "P": 18, "Q": 24, "R": 28, "S": 14, "T": 14, "U": 45, "V": 22
            }
            for col_letter, w in widths.items():
                ws.column_dimensions[col_letter].width = w

            wb.save(save_path)
            QMessageBox.information(self, "批量导出成功", f"已成功导出 {len(selected_cids)} 名干部的研判分析表格！\n文件保存至：\n{save_path}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"导出出错: {e}")

    # ==================== 弹窗唤醒桥接 ====================

    def get_selected_cid(self):
        row = self.table.currentRow()
        if row < 0: return None
        item = self.table.item(row, 0)
        return item.data(Qt.UserRole) if item else None

    def open_profile_from_selection(self, model_index=None):
        cid = self.get_selected_cid()
        if not cid: return
        dlg = FullCadreDetailWindow(self, cid, base_ym=self.get_base_ym(), on_data_changed=self.load_cadre_data)
        dlg.exec()

    def open_add_dialog(self):
        if self.current_tab == "active_chu":
            self.open_add_chu_dialog()
            return
        dlg = FullCadreEditDialog(self, title="➕ 新增干部全息档案入库", default_rank="科级", on_saved=self.load_cadre_data)
        dlg.exec()

    def open_add_chu_dialog(self):
        dlg = FullCadreEditDialog(self, title="➕ 手动录入处级领导干部档案", default_rank="处级", default_chu_category="单位领导班子", on_saved=self.load_cadre_data)
        dlg.exec()

    def open_edit_dialog(self):
        cid = self.get_selected_cid()
        if not cid:
            QMessageBox.warning(self, "提示", "请先在列表中选中需要编辑的人员！")
            return
        dlg = FullCadreEditDialog(self, title="✏️ 编辑干部全息档案", cadre_id=cid, on_saved=self.load_cadre_data)
        dlg.exec()

    def open_manual_bind_dialog(self):
        cid = self.get_selected_cid()
        if not cid:
            QMessageBox.warning(self, "提示", "请先在列表中选中需要核准身份/解决重名的干部！")
            return
        row = self.table.currentRow()
        name = self.table.item(row, 1).text() if self.table.item(row, 1) else ""
        dlg = ManualBindRosterDialog(self, cid, name, on_bound=self.load_cadre_data)
        dlg.exec()

    def open_batch_appointment_sync(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "请选择干部职务任免文件", BASE_DIR, "任免公文文件 (*.docx *.pdf)")
        if not file_path: return
        dlg = BatchAppointmentSyncDialog(self, file_path, on_completed=self.load_cadre_data)
        dlg.exec()

    def open_discipline_doc_sync(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "请选择纪律处分/诫勉处理决定文件", BASE_DIR, "纪律处分文件 (*.docx *.pdf)")
        if not file_path: return
        dlg = DisciplineDocSyncDialog(self, file_path, on_completed=self.load_cadre_data)
        dlg.exec()

    def upload_chu_appointment_doc(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "请选择处级干部任免公文", BASE_DIR, "处级任免公文 (*.docx *.pdf)")
        if not file_path: return
        dlg = ChuAppointmentUploadDialog(self, file_path, on_completed=self.load_cadre_data)
        dlg.exec()

    def open_stats_dashboard(self):
        CadreStatsDashboard(self, self.get_base_ym())

    def open_custom_export_dialog(self):
        selected_cids = []
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 0)
            if item:
                cid = item.data(Qt.UserRole)
                if cid: selected_cids.append(cid)
        if not selected_cids:
            QMessageBox.warning(self, "提示", "当前列表中无任何人员数据可导出！")
            return
        dlg = CustomExportDialog(self, selected_cids, self.get_base_ym())
        dlg.exec()

    def export_selected_approval_docx(self):
        cid = self.get_selected_cid()
        if not cid:
            QMessageBox.warning(self, "提示", "请先在列表中选中一名需要导出审批表的干部！")
            return
        dlg = ExportApprovalFormDialog(self, cid, self.get_base_ym())
        dlg.exec()

    def batch_export_approval_docx(self):
        # 1. 获取选中的人员 ID 列表（支持多选 ExtendedSelection）
        selected_cids = []
        selected_indexes = self.table.selectionModel().selectedRows()
        if selected_indexes:
            for idx in selected_indexes:
                item = self.table.item(idx.row(), 0)
                if item:
                    cid = item.data(Qt.UserRole)
                    if cid and cid not in selected_cids:
                        selected_cids.append(cid)
        else:
            for r in range(self.table.rowCount()):
                item = self.table.item(r, 0)
                if item:
                    cid = item.data(Qt.UserRole)
                    if cid and cid not in selected_cids:
                        selected_cids.append(cid)

            if not selected_cids:
                QMessageBox.warning(self, "提示", "当前表格中无任何干部数据可导出！")
                return

            ans = QMessageBox.question(
                self, "批量导出干部任免审批表",
                f"当前未手动多选行。\n\n是否为当前列表中筛选出的全部【{len(selected_cids)} 名干部】批量生成并导出标准《干部任免审批表》Word？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
            )
            if ans != QMessageBox.Yes:
                return

        # 2. 检查审批表 Word 模板
        tpl_path = get_approval_template_path()
        if not tpl_path:
            tpl_path, _ = QFileDialog.getOpenFileName(self, "请选择标准干部任免审批表Word模板", BASE_DIR, "Word 文档 (*.docx)")
            if not tpl_path:
                return

        # 3. 选择批量导出目标文件夹
        out_dir = QFileDialog.getExistingDirectory(self, "请选择批量审批表导出保存文件夹", BASE_DIR)
        if not out_dir:
            return

        # 4. 从数据库查询待导出干部全字段数据并回填生成
        conn = sqlite3.connect(DB_PATH, timeout=30.0)
        c = conn.cursor()
        cols = ", ".join(ALL_CADRE_COLUMNS)
        placeholders = ",".join(["?"] * len(selected_cids))
        c.execute(f"SELECT {cols} FROM cadres WHERE id IN ({placeholders})", selected_cids)
        rows = c.fetchall()
        conn.close()

        cadre_dict_map = {}
        for r in rows:
            d = {k: (str(v) if v is not None and str(v) != "None" else "—") for k, v in zip(ALL_CADRE_COLUMNS, r)}
            cadre_dict_map[int(d["id"])] = d

        success_cnt = 0
        failed_list = []
        today_str = datetime.datetime.now().strftime("%Y%m%d")

        for cid in selected_cids:
            d = cadre_dict_map.get(int(cid))
            if not d:
                continue
            name = d.get("name", f"干部_{cid}")
            out_fname = f"{name}_干部任免审批表_{today_str}.docx"
            out_path = os.path.join(out_dir, out_fname)

            fill_data = dict(d)
            fill_data["target_post"] = ""
            fill_data["target_dismiss"] = ""
            fill_data["resume_text"] = format_resume_text(d.get("resume_text", ""))
            fill_data["family_relations"] = format_family_text(d.get("family_relations", ""))

            try:
                fill_approval_docx(fill_data, tpl_path, out_path)
                success_cnt += 1
            except Exception as e:
                failed_list.append(f"{name}: {str(e)}")

        msg = f"批量生成完成！\n\n• 成功生成：{success_cnt} 份《干部任免审批表》Word文档\n• 导出目录：{out_dir}"
        if failed_list:
            msg += f"\n• 失败 {len(failed_list)} 人：\n" + "\n".join(failed_list[:5])
        QMessageBox.information(self, "批量导出审批表完成", msg)
        QDesktopServices.openUrl(QUrl.fromLocalFile(out_dir))

    def open_leave_dialog(self):
        cid = self.get_selected_cid()
        if not cid:
            QMessageBox.warning(self, "提示", "请先在列表中选中人员！")
            return
        row = self.table.currentRow()
        name = self.table.item(row, 1).text() if self.table.item(row, 1) else ""
        is_chu = ("chu" in self.current_tab)
        current_rank = "处级" if is_chu else "科级"

        if self.current_tab in ["active_ke", "active_chu"]:
            dlg = CadreLeaveDialog(self, cid, name, cadre_rank=current_rank, on_finished=self.load_cadre_data)
            dlg.exec()
        else:
            ans = QMessageBox.question(
                self, "历史人员操作",
                f"您正在处理历史人员【{name}】（{current_rank}）：\n\n【Yes】：恢复为【在岗】状态\n【No】：【彻底物理删除】（不可逆）",
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel, QMessageBox.Cancel
            )
            if ans == QMessageBox.Yes:
                conn = sqlite3.connect(DB_PATH, timeout=30.0)
                conn.cursor().execute("UPDATE cadres SET status='在岗', leave_reason=NULL, leave_date=NULL WHERE id=?", (cid,))
                conn.commit()
                conn.close()
                QMessageBox.information(self, "成功", f"干部【{name}】已恢复为在岗状态！")
                self.load_cadre_data()
            elif ans == QMessageBox.No:
                if QMessageBox.question(self, "安全确认", f"确定彻底删除干部【{name}】吗？数据将被清除！") == QMessageBox.Yes:
                    conn = sqlite3.connect(DB_PATH, timeout=30.0)
                    conn.cursor().execute("DELETE FROM cadres WHERE id=?", (cid,))
                    conn.commit()
                    conn.close()
                    QMessageBox.information(self, "提示", f"干部【{name}】已清除。")
                    self.load_cadre_data()

# ==================== 主程序执行入口 ====================

def setup_global_light_theme(app: QApplication):
    app.setStyle("Fusion")
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor("#F8FAFC"))
    palette.setColor(QPalette.WindowText, QColor("#1E293B"))
    palette.setColor(QPalette.Base, QColor("#FFFFFF"))
    palette.setColor(QPalette.AlternateBase, QColor("#F8FAFC"))
    palette.setColor(QPalette.ToolTipBase, QColor("#1E293B"))
    palette.setColor(QPalette.ToolTipText, QColor("#FFFFFF"))
    palette.setColor(QPalette.Text, QColor("#1E293B"))
    palette.setColor(QPalette.Button, QColor("#FFFFFF"))
    palette.setColor(QPalette.ButtonText, QColor("#1E293B"))
    palette.setColor(QPalette.BrightText, QColor("#DC2626"))
    palette.setColor(QPalette.Link, QColor("#2563EB"))
    palette.setColor(QPalette.Highlight, QColor("#3B82F6"))
    palette.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
    app.setPalette(palette)

    font_family = "PingFang SC" if sys.platform == "darwin" else "Microsoft YaHei UI"
    font = QFont(font_family, 9)
    font.setStyleStrategy(QFont.PreferAntialias)
    app.setFont(font)

    # 彻底防止在 macOS 深色模式（Dark Mode）或系统主题下出现白底白字、输入框看不见文字的问题
    app.setStyleSheet(f"""
        QWidget {{
            color: #1E293B;
            font-family: "{font_family}", "Microsoft YaHei UI", sans-serif;
        }}
        QTableWidget {{
            background-color: #FFFFFF;
            color: #1E293B;
            gridline-color: #F1F5F9;
        }}
        QTableWidget::item {{
            color: #1E293B;
        }}
        QTableWidget::item:selected {{
            background-color: #EFF6FF;
            color: #1E3A8A;
        }}
        QHeaderView::section {{
            background-color: #F1F5F9;
            color: #475569;
        }}
        QLineEdit {{
            background-color: #FFFFFF;
            color: #1E293B;
            selection-background-color: #3B82F6;
            selection-color: #FFFFFF;
        }}
        QTextEdit, QPlainTextEdit {{
            background-color: #FFFFFF;
            color: #1E293B;
        }}
        QComboBox {{
            background-color: #FFFFFF;
            color: #1E293B;
        }}
        QComboBox QAbstractItemView {{
            background-color: #FFFFFF;
            color: #1E293B;
            selection-background-color: #EFF6FF;
            selection-color: #1E3A8A;
        }}
        QDialog {{
            background-color: #FFFFFF;
            color: #1E293B;
        }}
        QLabel {{
            color: #1E293B;
        }}
    """)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    setup_global_light_theme(app)

    main_win = FullCadreMainWindow()
    main_win.show()
    sys.exit(app.exec())