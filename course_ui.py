import argparse
import json
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Dict, List, Tuple, Set, Optional, Any

from pku_course_parser import load_courses, Course


WEEK_MIN = 1
WEEK_MAX = 16
DAYS = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]


def weekday_to_label(weekday: int) -> str:
    return DAYS[weekday - 1] if 1 <= weekday <= 7 else f"周{weekday}"


def build_occupied_cells_for_course(c: Course) -> Set[Tuple[int, int, int]]:
    occ = set()
    for week in range(WEEK_MIN, WEEK_MAX + 1):
        for m in c.meetings:
            if not m.occurs_on_week(week):
                continue
            for p in range(m.start_period, m.end_period + 1):
                occ.add((week, m.weekday, p))
    return occ


class CourseUI(tk.Tk):
    CONFIG_VERSION = 1

    def __init__(self, xlsx_path: str, sheet_name: str):
        super().__init__()
        self.title("选课课表（第二阶段 UI）")
        self.geometry("1400x820")

        # 保存源信息，写入配置文件里（用于提示，不强制校验）
        self.source_file = xlsx_path
        self.source_sheet = sheet_name

        self.res = load_courses(xlsx_path, sheet_name=sheet_name, debug=False)
        self.by_uid: Dict[Tuple[str, str], Course] = self.res.by_uid
        self.all_courses: List[Course] = list(self.by_uid.values())

        # 已选 / 冲突
        self.selected: Set[Tuple[str, str]] = set()
        self.conflicted: Set[Tuple[str, str]] = set()

        self.occ_cache: Dict[Tuple[str, str], Set[Tuple[int, int, int]]] = {
            uid: build_occupied_cells_for_course(c) for uid, c in self.by_uid.items()
        }

        self.week_var = tk.IntVar(value=1)
        self.credit_limit_var = tk.DoubleVar(value=25.0)
        self.dept_var = tk.StringVar(value="全部")

        self.unselected_iid_to_uid: Dict[str, Tuple[str, str]] = {}
        self.selected_iid_to_uid: Dict[str, Tuple[str, str]] = {}
        self.conflicted_iid_to_uid: Dict[str, Tuple[str, str]] = {}

        self._init_style()
        self._build_layout()
        self._populate_departments()
        self._refresh_lists()
        self._refresh_timetable()

    # -------------------------
    # Style
    # -------------------------

    def _init_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("Title.TLabel", font=("Arial", 13, "bold"))
        style.configure("Hint.TLabel", foreground="#333", font=("Arial", 10))
        style.configure("Toolbar.TFrame", padding=(6, 6, 6, 6))

        style.configure("Course.Treeview", rowheight=28, font=("Arial", 11))
        style.configure("Course.Treeview.Heading", font=("Arial", 11, "bold"))

    # -------------------------
    # Layout
    # -------------------------

    def _build_layout(self):
        self.columnconfigure(0, weight=3)
        self.columnconfigure(1, weight=6)
        self.columnconfigure(2, weight=3)
        self.rowconfigure(0, weight=1)

        # 左：未选
        left = ttk.Frame(self, padding=10)
        left.grid(row=0, column=0, sticky="nsew")
        left.columnconfigure(0, weight=1)
        left.rowconfigure(4, weight=1)

        ttk.Label(left, text="未选课课程目录", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )

        filter_row = ttk.Frame(left)
        filter_row.grid(row=1, column=0, sticky="ew", pady=(8, 6))
        filter_row.columnconfigure(1, weight=1)

        ttk.Label(filter_row, text="开课单位：").grid(row=0, column=0, sticky="w")
        self.dept_combo = ttk.Combobox(
            filter_row, textvariable=self.dept_var, state="readonly"
        )
        self.dept_combo.grid(row=0, column=1, sticky="ew")
        self.dept_combo.bind("<<ComboboxSelected>>", lambda e: self._refresh_lists())

        ttk.Label(left, text="可多选：Ctrl / Shift", style="Hint.TLabel").grid(
            row=2, column=0, sticky="w", pady=(0, 6)
        )

        self.unselected_tree = self._make_course_tree(parent=left)
        self.unselected_tree.grid(row=4, column=0, sticky="nsew")
        unselected_scroll = ttk.Scrollbar(
            left, orient="vertical", command=self.unselected_tree.yview
        )
        self.unselected_tree.configure(yscrollcommand=unselected_scroll.set)
        unselected_scroll.grid(row=4, column=1, sticky="ns")

        # 中：课表 + 工具栏
        mid = ttk.Frame(self, padding=10)
        mid.grid(row=0, column=1, sticky="nsew")
        mid.columnconfigure(0, weight=1)
        mid.rowconfigure(4, weight=1)

        toolbar = ttk.Frame(mid, style="Toolbar.TFrame")
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.columnconfigure(2, weight=1)

        ttk.Button(toolbar, text="<< 上一周", command=self._prev_week).grid(
            row=0, column=0, padx=(0, 6)
        )
        ttk.Button(toolbar, text="下一周 >>", command=self._next_week).grid(
            row=0, column=1, padx=(0, 12)
        )

        self.week_label = ttk.Label(toolbar, text="", style="Title.TLabel")
        self.week_label.grid(row=0, column=2, sticky="w")

        ttk.Button(
            toolbar, text="加入已选 >>", command=self._add_selected_courses
        ).grid(row=0, column=3, padx=(12, 6))
        ttk.Button(toolbar, text="<< 退选", command=self._remove_selected_courses).grid(
            row=0, column=4
        )

        # 第二行：导入/导出
        io_row = ttk.Frame(mid)
        io_row.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        io_row.columnconfigure(2, weight=1)

        ttk.Button(io_row, text="导入配置…", command=self._import_config).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Button(io_row, text="导出配置…", command=self._export_config).grid(
            row=0, column=1, sticky="w", padx=(8, 0)
        )

        # 第三行：学分上限 + 状态
        credit_row = ttk.Frame(mid)
        credit_row.grid(row=2, column=0, sticky="ew", pady=(6, 6))
        credit_row.columnconfigure(4, weight=1)

        ttk.Label(credit_row, text="学分上限：").grid(row=0, column=0, sticky="w")
        self.credit_entry = ttk.Entry(
            credit_row, width=8, textvariable=self.credit_limit_var
        )
        self.credit_entry.grid(row=0, column=1, sticky="w", padx=(4, 6))
        ttk.Button(credit_row, text="应用", command=self._apply_credit_limit).grid(
            row=0, column=2, sticky="w"
        )

        self.credit_status = ttk.Label(credit_row, text="", style="Hint.TLabel")
        self.credit_status.grid(row=0, column=4, sticky="e")

        # 课表网格
        self.table_frame = ttk.Frame(mid)
        self.table_frame.grid(row=4, column=0, sticky="nsew")
        self.table_frame.rowconfigure(0, weight=1)
        self.table_frame.columnconfigure(0, weight=1)
        self._build_timetable_grid()

        # 右：已选 + 冲突
        right = ttk.Frame(self, padding=10)
        right.grid(row=0, column=2, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)
        right.rowconfigure(7, weight=1)

        ttk.Label(right, text="已选课课程目录", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(right, text="可多选：Ctrl / Shift", style="Hint.TLabel").grid(
            row=1, column=0, sticky="w", pady=(8, 6)
        )

        self.selected_tree = self._make_course_tree(parent=right)
        self.selected_tree.grid(row=2, column=0, sticky="nsew")
        selected_scroll = ttk.Scrollbar(
            right, orient="vertical", command=self.selected_tree.yview
        )
        self.selected_tree.configure(yscrollcommand=selected_scroll.set)
        selected_scroll.grid(row=2, column=1, sticky="ns")

        ttk.Separator(right, orient="horizontal").grid(
            row=3, column=0, columnspan=2, sticky="ew", pady=(10, 8)
        )

        ttk.Label(right, text="冲突课程目录", style="Title.TLabel").grid(
            row=4, column=0, sticky="w"
        )
        ttk.Label(
            right, text="发生冲突时可暂存于此（不进课表）", style="Hint.TLabel"
        ).grid(row=5, column=0, sticky="w", pady=(6, 6))

        btn_row = ttk.Frame(right)
        btn_row.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        btn_row.columnconfigure(0, weight=1)
        ttk.Button(
            btn_row, text="移除冲突", command=self._remove_conflicted_courses
        ).grid(row=0, column=0, sticky="e")

        self.conflicted_tree = self._make_course_tree(parent=right)
        self.conflicted_tree.grid(row=7, column=0, sticky="nsew")
        conflicted_scroll = ttk.Scrollbar(
            right, orient="vertical", command=self.conflicted_tree.yview
        )
        self.conflicted_tree.configure(yscrollcommand=conflicted_scroll.set)
        conflicted_scroll.grid(row=7, column=1, sticky="ns")

    def _make_course_tree(self, parent: ttk.Frame) -> ttk.Treeview:
        tree = ttk.Treeview(
            parent,
            columns=("name", "teacher", "classno", "credits"),
            show="headings",
            selectmode="extended",
            style="Course.Treeview",
            height=22,
        )
        tree.heading("name", text="课程名")
        tree.heading("teacher", text="老师")
        tree.heading("classno", text="班号")
        tree.heading("credits", text="学分")

        tree.column("name", width=220, stretch=True, anchor="w")
        tree.column("teacher", width=140, stretch=False, anchor="w")
        tree.column("classno", width=70, stretch=False, anchor="center")
        tree.column("credits", width=60, stretch=False, anchor="center")
        return tree

    def _build_timetable_grid(self):
        header_style = dict(
            borderwidth=1,
            relief="solid",
            background="#f5f5f5",
            fg="#000000",
            font=("Arial", 11, "bold"),
            anchor="center",
        )
        cell_style = dict(
            borderwidth=1,
            relief="solid",
            background="white",
            fg="#111111",
            font=("Arial", 10),
            anchor="center",
            justify="center",
            wraplength=150,
        )

        self.cells: Dict[Tuple[int, int], tk.Label] = {}

        for r in range(0, 13):
            self.table_frame.rowconfigure(r, weight=1)
        for c in range(0, 8):
            self.table_frame.columnconfigure(c, weight=1)

        self.cells[(0, 0)] = tk.Label(
            self.table_frame, text="节次/星期", **header_style
        )
        self.cells[(0, 0)].grid(row=0, column=0, sticky="nsew")

        for i, day in enumerate(DAYS, start=1):
            lbl = tk.Label(self.table_frame, text=day, **header_style)
            lbl.grid(row=0, column=i, sticky="nsew")
            self.cells[(0, i)] = lbl

        for period in range(1, 13):
            lblp = tk.Label(self.table_frame, text=f"第{period}节", **header_style)
            lblp.grid(row=period, column=0, sticky="nsew")
            self.cells[(period, 0)] = lblp

            for daycol in range(1, 8):
                lblc = tk.Label(self.table_frame, text="", **cell_style)
                lblc.grid(row=period, column=daycol, sticky="nsew")
                self.cells[(period, daycol)] = lblc

    # -------------------------
    # Helpers
    # -------------------------

    def _populate_departments(self):
        depts = sorted(
            {
                (c.department or "").strip()
                for c in self.all_courses
                if (c.department or "").strip()
            }
        )
        values = ["全部"] + depts
        self.dept_combo["values"] = values
        if self.dept_var.get() not in values:
            self.dept_var.set("全部")

    def _filtered_unselected_uids(self) -> List[Tuple[str, str]]:
        dept = self.dept_var.get().strip()
        uids = []
        for uid, c in self.by_uid.items():
            if uid in self.selected:
                continue
            if uid in self.conflicted:
                continue
            if dept != "全部" and (c.department or "").strip() != dept:
                continue
            uids.append(uid)

        uids.sort(
            key=lambda u: (
                (self.by_uid[u].department or "").strip(),
                self.by_uid[u].course_name.strip(),
                self.by_uid[u].teacher.strip(),
                (self.by_uid[u].class_no or "").strip(),
            )
        )
        return uids

    def _selected_uids_sorted(self) -> List[Tuple[str, str]]:
        uids = list(self.selected)
        uids.sort(
            key=lambda u: (
                (self.by_uid[u].department or "").strip(),
                self.by_uid[u].course_name.strip(),
                self.by_uid[u].teacher.strip(),
                (self.by_uid[u].class_no or "").strip(),
            )
        )
        return uids

    def _conflicted_uids_sorted(self) -> List[Tuple[str, str]]:
        uids = list(self.conflicted)
        uids.sort(
            key=lambda u: (
                (self.by_uid[u].department or "").strip(),
                self.by_uid[u].course_name.strip(),
                self.by_uid[u].teacher.strip(),
                (self.by_uid[u].class_no or "").strip(),
            )
        )
        return uids

    def _current_total_credits(self) -> float:
        return sum(float(self.by_uid[uid].credits or 0.0) for uid in self.selected)

    def _update_credit_status(self):
        total = self._current_total_credits()
        limit = float(self.credit_limit_var.get() or 0.0)
        self.credit_status.config(text=f"已选总学分：{total:g} / 上限：{limit:g}")

    # -------------------------
    # Config import/export
    # -------------------------

    @staticmethod
    def _uid_to_str(uid: Tuple[str, str]) -> str:
        # 用一个不太可能出现在字段里的分隔符
        return f"{uid[0]}||{uid[1]}"

    @staticmethod
    def _str_to_uid(s: str) -> Tuple[str, str]:
        a, b = s.split("||", 1)
        return (a, b)

    def _build_config_payload(self) -> Dict[str, Any]:
        return {
            "version": self.CONFIG_VERSION,
            "source": {
                "file": os.path.abspath(self.source_file),
                "sheet": self.source_sheet,
            },
            "week": int(self.week_var.get()),
            "credit_limit": float(self.credit_limit_var.get() or 0.0),
            "selected": [self._uid_to_str(u) for u in sorted(self.selected)],
            "conflicted": [self._uid_to_str(u) for u in sorted(self.conflicted)],
        }

    def _apply_config_payload(self, payload: Dict[str, Any]):
        # 基础校验
        if not isinstance(payload, dict):
            raise ValueError("配置文件格式不正确（应为 JSON 对象）。")
        if int(payload.get("version", 0)) != self.CONFIG_VERSION:
            raise ValueError(f"配置文件版本不匹配：{payload.get('version')}")

        selected_raw = payload.get("selected", [])
        conflicted_raw = payload.get("conflicted", [])
        if not isinstance(selected_raw, list) or not isinstance(conflicted_raw, list):
            raise ValueError("配置文件格式不正确：selected/conflicted 必须为数组。")

        selected_uids = [self._str_to_uid(x) for x in selected_raw]
        conflicted_uids = [self._str_to_uid(x) for x in conflicted_raw]

        # 过滤掉当前数据源不存在的 uid
        missing_selected = [u for u in selected_uids if u not in self.by_uid]
        missing_conflicted = [u for u in conflicted_uids if u not in self.by_uid]

        selected_set = {u for u in selected_uids if u in self.by_uid}
        conflicted_set = {u for u in conflicted_uids if u in self.by_uid}

        # 保证互斥：已选优先
        conflicted_set -= selected_set

        # 应用
        self.selected = selected_set
        self.conflicted = conflicted_set

        # 恢复一些 UI 状态（安全范围内）
        try:
            credit_limit = float(
                payload.get("credit_limit", self.credit_limit_var.get())
            )
            if credit_limit > 0:
                self.credit_limit_var.set(credit_limit)
        except Exception:
            pass

        try:
            week = int(payload.get("week", self.week_var.get()))
            if WEEK_MIN <= week <= WEEK_MAX:
                self.week_var.set(week)
        except Exception:
            pass

        self._refresh_lists()
        self._refresh_timetable()

        # 提示缺失课程
        if missing_selected or missing_conflicted:
            msg_lines = ["部分课程在当前数据源中未找到，已自动忽略："]
            if missing_selected:
                msg_lines.append(f"\n- 已选缺失：{len(missing_selected)} 门")
            if missing_conflicted:
                msg_lines.append(f"\n- 冲突缺失：{len(missing_conflicted)} 门")
            messagebox.showwarning("导入完成（有缺失）", "".join(msg_lines))

    def _export_config(self):
        default_name = "course_config.json"
        path = filedialog.asksaveasfilename(
            title="导出课表配置",
            defaultextension=".json",
            initialfile=default_name,
            filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")],
        )
        if not path:
            return

        payload = self._build_config_payload()
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("导出成功", f"已保存到：\n{path}")
        except Exception as e:
            messagebox.showerror("导出失败", f"无法写入文件：\n{path}\n\n错误：{e}")

    def _import_config(self):
        path = filedialog.askopenfilename(
            title="导入课表配置",
            filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")],
        )
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            self._apply_config_payload(payload)
            messagebox.showinfo("导入成功", f"已从配置恢复课表：\n{path}")
        except Exception as e:
            messagebox.showerror("导入失败", f"配置文件无法导入：\n{path}\n\n错误：{e}")

    # -------------------------
    # Refresh
    # -------------------------

    def _refresh_lists(self):
        # 未选
        self.unselected_tree.delete(*self.unselected_tree.get_children())
        self.unselected_iid_to_uid.clear()
        for uid in self._filtered_unselected_uids():
            c = self.by_uid[uid]
            iid = self.unselected_tree.insert(
                "",
                "end",
                values=(
                    c.course_name.strip(),
                    c.teacher.strip(),
                    f"{(c.class_no or '').strip()}",
                    f"{float(c.credits or 0.0):g}",
                ),
            )
            self.unselected_iid_to_uid[iid] = uid

        # 已选
        self.selected_tree.delete(*self.selected_tree.get_children())
        self.selected_iid_to_uid.clear()
        for uid in self._selected_uids_sorted():
            c = self.by_uid[uid]
            iid = self.selected_tree.insert(
                "",
                "end",
                values=(
                    c.course_name.strip(),
                    c.teacher.strip(),
                    f"{(c.class_no or '').strip()}",
                    f"{float(c.credits or 0.0):g}",
                ),
            )
            self.selected_iid_to_uid[iid] = uid

        # 冲突
        self.conflicted_tree.delete(*self.conflicted_tree.get_children())
        self.conflicted_iid_to_uid.clear()
        for uid in self._conflicted_uids_sorted():
            c = self.by_uid[uid]
            iid = self.conflicted_tree.insert(
                "",
                "end",
                values=(
                    c.course_name.strip(),
                    c.teacher.strip(),
                    f"{(c.class_no or '').strip()}",
                    f"{float(c.credits or 0.0):g}",
                ),
            )
            self.conflicted_iid_to_uid[iid] = uid

        self._update_credit_status()

    def _refresh_timetable(self):
        week = int(self.week_var.get())
        self.week_label.config(text=f"第 {week} 周（1~16 可切换）")

        for period in range(1, 13):
            for daycol in range(1, 8):
                self.cells[(period, daycol)].config(text="", background="white")

        grid_text: Dict[Tuple[int, int], List[str]] = {}

        for uid in self.selected:
            c = self.by_uid[uid]
            for m in c.meetings:
                if not m.occurs_on_week(week):
                    continue
                daycol = m.weekday
                place = (m.room or c.key.room or "地点未知").strip()
                for p in range(m.start_period, m.end_period + 1):
                    key = (p, daycol)
                    grid_text.setdefault(key, []).append(
                        f"{c.course_name}\n{c.teacher}\n{place}"
                    )

        def color_for(name: str) -> str:
            palette = ["#e8f4ff", "#e9f8ef", "#fff2e6", "#f3e8ff", "#ffe8f0", "#f0f0ff"]
            return palette[hash(name) % len(palette)]

        for (p, d), lines in grid_text.items():
            txt = "\n---\n".join(lines)
            first_name = lines[0].split("\n")[0] if lines else ""
            self.cells[(p, d)].config(text=txt, background=color_for(first_name))

        self._update_credit_status()

    # -------------------------
    # Week navigation
    # -------------------------

    def _prev_week(self):
        w = int(self.week_var.get())
        if w > WEEK_MIN:
            self.week_var.set(w - 1)
            self._refresh_timetable()

    def _next_week(self):
        w = int(self.week_var.get())
        if w < WEEK_MAX:
            self.week_var.set(w + 1)
            self._refresh_timetable()

    # -------------------------
    # Credit limit
    # -------------------------

    def _apply_credit_limit(self):
        try:
            limit = float(self.credit_limit_var.get())
            if limit <= 0:
                raise ValueError
            self.credit_limit_var.set(limit)
            self._update_credit_status()
        except Exception:
            messagebox.showerror("学分上限错误", "请输入一个大于 0 的数字（例如 25）。")

    # -------------------------
    # Conflict detection
    # -------------------------

    def _find_conflict_between(self, uids: List[Tuple[str, str]]) -> Optional[str]:
        occupied: Dict[Tuple[int, int, int], Tuple[str, str]] = {}

        for uid in self.selected:
            for cell in self.occ_cache[uid]:
                occupied[cell] = uid

        for uid in uids:
            for cell in self.occ_cache[uid]:
                if cell in occupied:
                    other = occupied[cell]
                    week, weekday, period = cell
                    c1 = self.by_uid[other]
                    c2 = self.by_uid[uid]
                    return (
                        "检测到时间冲突：\n\n"
                        f" - {c1.course_name} | {c1.teacher} | 班{c1.class_no}\n"
                        f" - {c2.course_name} | {c2.teacher} | 班{c2.class_no}\n\n"
                        f"冲突位置：第{week}周 {weekday_to_label(weekday)} 第{period}节"
                    )
                occupied[cell] = uid
        return None

    # -------------------------
    # Add / Remove
    # -------------------------

    def _get_tree_selected_uids(
        self, tree: ttk.Treeview, mapping: Dict[str, Tuple[str, str]]
    ) -> List[Tuple[str, str]]:
        return [mapping[iid] for iid in tree.selection() if iid in mapping]

    def _add_selected_courses(self):
        uids_to_add = self._get_tree_selected_uids(
            self.unselected_tree, self.unselected_iid_to_uid
        )
        if not uids_to_add:
            return

        conflict = self._find_conflict_between(uids_to_add)
        if conflict:
            ok = messagebox.askyesno(
                "选课冲突",
                conflict + "\n\n是否将这些课程加入【冲突课程列表】？（原有课表不变）",
            )
            if ok:
                for uid in uids_to_add:
                    if uid not in self.selected:
                        self.conflicted.add(uid)
                self._refresh_lists()
            return

        current = self._current_total_credits()
        add_credits = sum(float(self.by_uid[uid].credits or 0.0) for uid in uids_to_add)
        limit = float(self.credit_limit_var.get() or 0.0)
        if current + add_credits > limit + 1e-9:
            messagebox.showwarning(
                "学分超限",
                f"加入这些课程会超出学分上限。\n\n当前：{current:g}\n新增：{add_credits:g}\n上限：{limit:g}",
            )
            return

        for uid in uids_to_add:
            self.selected.add(uid)
            self.conflicted.discard(uid)

        self._refresh_lists()
        self._refresh_timetable()

    def _remove_selected_courses(self):
        uids_to_remove = self._get_tree_selected_uids(
            self.selected_tree, self.selected_iid_to_uid
        )
        if not uids_to_remove:
            return

        for uid in uids_to_remove:
            self.selected.discard(uid)

        self._refresh_lists()
        self._refresh_timetable()

    def _remove_conflicted_courses(self):
        uids_to_remove = self._get_tree_selected_uids(
            self.conflicted_tree, self.conflicted_iid_to_uid
        )
        if not uids_to_remove:
            return

        for uid in uids_to_remove:
            self.conflicted.discard(uid)

        self._refresh_lists()


def main() -> int:
    parser = argparse.ArgumentParser(description="第二阶段：选课 UI（Tkinter）")
    parser.add_argument(
        "--file", "-f", default="pku_courses.xlsx", help="课程数据文件（.xlsx/.csv）"
    )
    parser.add_argument(
        "--sheet", "-s", default="courses", help="xlsx sheet 名（默认 courses）"
    )
    args = parser.parse_args()

    app = CourseUI(args.file, args.sheet)
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
