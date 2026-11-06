# -*- coding: utf-8 -*-
"""Xếp & In TKB Trường học - Streamlit + OR-Tools + AgGrid + ReportLab"""
import hashlib
import html
import io
import json
import os
import re
from collections import Counter
from datetime import date

import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode

APP_DIR = os.path.dirname(os.path.abspath(__file__))
RULES_FILE = os.path.join(APP_DIR, "cacrangbuocxep.json")

FIXED = ["Thứ", "Tiết", "Giờ"]
EXPECTED_ROWS = 60            # 61 dòng kể cả tiêu đề
MORNING_LAST = 5              # tiết 1-5 là buổi sáng
TOKEN_RE = re.compile(r"^[^_\s][^_]*_[A-Z]{2}$")
HELPER_COLS = ["_i", "_bg", "_ban", "_dup", "_sel", "_clk"]

DAY_RGB = {2: (255, 214, 165), 3: (255, 255, 170), 4: (190, 240, 190),
           5: (170, 225, 245), 6: (190, 200, 255), 7: (225, 190, 250)}

DEFAULT_RULES = {
    "o_cam": {"mo_ta": "Không xếp vào các ô trống cấm xếp", "bat": True, "khoa": True},
    "trung_tiet": {"mo_ta": "Không trùng tiết (GV không dạy hơn 1 lớp cùng tiết)", "bat": True},
    "tiet_trong_buoi": {"mo_ta": "Giữa tiết dạy đầu và cuối trong buổi không có hơn 1 tiết trống", "bat": True},
    "cach_2_buoi": {"mo_ta": "Dạy cả sáng và chiều thì phải có ít nhất 1 tiết trống giữa 2 buổi", "bat": True},
    "vv_tt": {"mo_ta": "Môn VV, TT: có 1 cặp tiết liền nhau và không xếp vào 2 ngày liền nhau", "bat": True},
    "gv_tranh": {"mo_ta": "Không xếp GV vào tiết GV yêu cầu tránh (file gvyctranh.xlsx)", "bat": True},
    "chao_co": {"mo_ta": "Môn CC (chào cờ) xếp vào Tiết 1 của Thứ 2", "bat": True},
    "toi_da_2": {"mo_ta": "Mỗi môn không học quá 2 tiết/lớp trong một ngày", "bat": True},
    "toi_thieu_2": {"mo_ta": "GV có tiết dạy trong buổi của ngày thì phải dạy ít nhất 2 tiết trong buổi đó", "bat": True},
}
JOINT_COLS = ["GV_Mon", "Lop", "Vi_tri"]
AVOID_COLS = ["ten_gv"] + [f"Thứ {d}" for d in range(2, 8)]


# ----------------------------------------------------------------- tiện ích
def tok(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def gv_name(t):
    return t.rsplit("_", 1)[0].strip()


def subj(t):
    return t.rsplit("_", 1)[1]


def load_rules():
    rules = json.loads(json.dumps(DEFAULT_RULES))
    if os.path.exists(RULES_FILE):
        try:
            with open(RULES_FILE, encoding="utf-8") as f:
                saved = json.load(f)
            for k, v in saved.items():
                if k in rules and not rules[k].get("khoa"):
                    rules[k]["bat"] = bool(v.get("bat", True))
        except Exception:
            pass
    else:
        save_rules(rules)
    return rules


def save_rules(rules):
    try:
        with open(RULES_FILE, "w", encoding="utf-8") as f:
            json.dump(rules, f, ensure_ascii=False, indent=2)
    except Exception:
        pass  # Streamlit Cloud có thể không cho ghi bền vững


def hex_mix(rgb, k):
    r, g, b = [int(c + (255 - c) * k) for c in rgb]
    return f"#{r:02x}{g:02x}{b:02x}"


# ------------------------------------------------------------ kiểm tra excel
def validate(df):
    """Trả về list (ok, tiêu đề, [chi tiết])."""
    rep = []
    n = len(df) + 1
    rep.append((n == EXPECTED_ROWS + 1,
                f"Số dòng kể cả dòng tiêu đề: {n} (phải là {EXPECTED_ROWS + 1})", []))

    cols = list(df.columns)
    head_ok = cols[:3] == FIXED and len(cols) > 3 and len(set(cols)) == len(cols)
    rep.append((head_ok, "Tiêu đề cột phải là: Thứ, Tiết, Giờ, rồi các lớp (không trùng tên)",
                [] if head_ok else [f"Hiện có: {cols}"]))
    if not head_ok:
        return rep
    classes = cols[3:]

    num_ok = all(pd.api.types.is_numeric_dtype(df[c]) for c in ("Thứ", "Tiết"))
    det = []
    if not num_ok:
        det.append("Cột Thứ, Tiết phải là kiểu Số (không là chuỗi).")
    else:
        bad = df[~df["Thứ"].isin(range(2, 8)) | ~df["Tiết"].isin(range(1, 11))]
        if len(bad):
            det.append(f"Thứ phải 2-7, Tiết phải 1-10. Dòng sai (số thứ tự dòng excel): "
                       f"{[int(i) + 2 for i in bad.index][:20]}")
        dupk = df[df.duplicated(["Thứ", "Tiết"], keep=False)]
        if len(dupk):
            det.append("Cặp (Thứ, Tiết) bị lặp ở dòng: "
                       f"{[int(i) + 2 for i in dupk.index][:20]}")
    rep.append((not det, "Cột Thứ, Tiết là số, đủ 6 ngày x 10 tiết, không lặp", det))

    bad_fmt = []
    for idx, c in enumerate(classes):
        letter = chr(ord("A") + idx + 3) if idx + 3 < 26 else f"cột {idx + 4}"
        looks_time = bool(re.match(r"^\d{1,2}[:h]\d{2}", c))
        vals = [tok(v) for v in df[c] if tok(v)]
        looks_num = bool(vals) and all(re.match(r"^\d+(\.\d+)?$", v) for v in vals)
        if looks_time or looks_num:
            bad_fmt.append(f"Cột {letter} có tiêu đề '{c}' không giống tên lớp (chứa giờ/số). "
                           f"Có thể file thừa cột hoặc bị lệch cột — sau 3 cột Thứ, Tiết, Giờ chỉ được là các cột lớp.")
            continue
        for i, v in df[c].items():
            t = tok(v)
            if t and not TOKEN_RE.match(t):
                bad_fmt.append(f"Dòng {int(i) + 2}, lớp {c}: '{t}'")
    rep.append((not bad_fmt,
                "Cách viết Tên GV_Môn đúng quy định (Tên + '_' + đúng 2 ký tự HOA)",
                bad_fmt[:40]))

    # Tên GV = phần trước dấu '_'. Một GV được dạy nhiều môn (VV, TT, SH...).
    # Chỉ báo lỗi khi cùng một tên nhưng viết khác nhau về hoa/thường hay khoảng trắng.
    variants = {}
    for c in classes:
        for v in df[c]:
            t = tok(v)
            if t and TOKEN_RE.match(t):
                n = gv_name(t)
                variants.setdefault(" ".join(n.split()).casefold(), set()).add(n)
    dup_names = [f"Các cách viết khác nhau của cùng một tên: {sorted(s)}"
                 for s in variants.values() if len(s) > 1]
    rep.append((not dup_names,
                f"Tên GV độc nhất, viết thống nhất (tổng {len(variants)} GV)",
                dup_names))
    return rep


def report_ok(rep):
    return all(r[0] for r in rep)


# --------------------------------------------------------------- yêu cầu tránh
PAIR_RE = re.compile(r"\(\s*(\d+)\s*,\s*(\d+)\s*\)")


def parse_avoid(adf):
    """gvyctranh.xlsx: cột 1 = ten_gv; các cột Thứ 2..Thứ 7 chứa các cặp (Thứ, tiết),
    ví dụ '(2,1), (3,1), (6,5)'. Cặp nằm ở ô nào trong dòng cũng được."""
    res = {}
    cols = list(adf.columns)
    for _, r in adf.iterrows():
        name = tok(r[cols[0]])
        if not name:
            continue
        if "_" in name and re.match(r"^[A-Z]{2}$", name.rsplit("_", 1)[1]):
            name = gv_name(name)
        for c in cols[1:]:
            for d, p in PAIR_RE.findall(tok(r[c])):
                d, p = int(d), int(p)
                if 2 <= d <= 7 and 1 <= p <= 10:
                    res.setdefault(name, set()).add((d, p))
    return res


def parse_joint(jdf):
    """Bảng tiết ghép -> list nhóm {token, classes, fixed}. Mỗi nhóm: 1 GV_Môn dạy chung nhiều lớp cùng ô."""
    groups = []
    for _, r in jdf.iterrows():
        t = tok(r["GV_Mon"])
        cl = [c.strip() for c in re.split(r"[,;]+", tok(r["Lop"])) if c.strip()]
        if not t or len(cl) < 2:
            continue
        fixed = [(int(d), int(p)) for d, p in PAIR_RE.findall(tok(r["Vi_tri"]))]
        groups.append({"token": t, "classes": cl, "fixed": fixed})
    return groups


def load_joint_df(f):
    """Đọc tietghep.xlsx: cột GV_Mon, Lop, Vi_tri (nếu tên cột khác thì lấy 3 cột đầu theo thứ tự)."""
    j = pd.read_excel(f, dtype=object)
    j.columns = [str(c).strip() for c in j.columns]
    if not all(c in j.columns for c in JOINT_COLS):
        if len(j.columns) < 2:
            raise ValueError("Cần ít nhất 2 cột: GV_Mon, Lop (và Vi_tri nếu muốn cố định vị trí).")
        j = j.iloc[:, :3].copy()
        j.columns = JOINT_COLS[:len(j.columns)]
    for c in JOINT_COLS:
        if c not in j.columns:
            j[c] = ""
    j = j[JOINT_COLS].copy()
    for c in JOINT_COLS:
        j[c] = j[c].map(tok)
    return j[j["GV_Mon"] != ""].reset_index(drop=True)


def check_joint(groups, base, classes):
    """Báo các nhóm tiết ghép không khớp với file TKB đang làm việc."""
    out = []
    for g in groups:
        t, cl = g["token"], g["classes"]
        miss = [c for c in cl if c not in classes]
        if miss:
            out.append(f"Tiết ghép {t}: lớp {', '.join(miss)} không có trong file TKB.")
            continue
        none = [c for c in cl if base[c].get(t, 0) == 0]
        if none:
            out.append(f"Tiết ghép {t}: lớp {', '.join(none)} không có tiết {t} trong file TKB.")
            continue
        if len({base[c][t] for c in cl}) > 1:
            out.append(f"Tiết ghép {t}: số tiết các lớp không bằng nhau ("
                       + ", ".join(f"{c}={base[c][t]}" for c in cl) + ").")
    return out


def joint_template():
    df = pd.DataFrame({"GV_Mon": ["Minh_TD"], "Lop": ["6A1,6A2"], "Vi_tri": ["(4,2)"]})
    b = io.BytesIO()
    df.to_excel(b, index=False)
    return b.getvalue()


def joint_map(groups):
    """'lớp|GV_Môn' -> số thứ tự nhóm ghép."""
    return {f"{c}|{g['token']}": i for i, g in enumerate(groups) for c in g["classes"]}


def cells_clash(name, parts, gmap):
    """parts: ['6A1_TD', '6A2_TD'] (lớp_môn của cùng 1 GV trong 1 ô thời gian). True nếu là trùng thật."""
    if len(parts) < 2:
        return False
    gids = set()
    for p in parts:
        c, s_ = p.rsplit("_", 1)
        gid = (gmap or {}).get(f"{c}|{name}_{s_}")
        if gid is None:
            return True
        gids.add(gid)
    return len(gids) > 1


def avoid_template():
    df = pd.DataFrame({"ten_gv": ["Anh", "Bình"],
                       "Thứ 2": ["(2,1), (3,1), (6,5)", ""],
                       "Thứ 3": ["", "(4,3)"],
                       "Thứ 4": ["", ""], "Thứ 5": ["", ""], "Thứ 6": ["", ""], "Thứ 7": ["", ""]})
    b = io.BytesIO()
    df.to_excel(b, index=False)
    return b.getvalue()


# ----------------------------------------------------------------- OR-Tools
def solve_tkb(df, ban, classes, base, rules, avoid, tlimit, groups=()):
    """Tối đa hóa số tiết xếp được, các ràng buộc bật đều là ràng buộc cứng.
    Tiết không xếp được sẽ để trống (chưa đem vào lưới)."""
    from ortools.sat.python import cp_model

    on = {k: v["bat"] for k, v in rules.items()}
    m = cp_model.CpModel()
    nrow = len(df)
    day_of = df["Thứ"].astype(int).tolist()
    per_of = df["Tiết"].astype(int).tolist()
    rowof = {(d, p): i for i, (d, p) in enumerate(zip(day_of, per_of))}
    days = sorted(set(day_of))

    x = {}
    teach = {}     # (gv, row) -> [var]
    sx = {}        # (lớp, môn, row) -> [var]

    # tiết ghép: nhiều lớp học chung 1 GV cùng ô. Lớp đầu là "đại diện" tính vào ràng buộc của GV,
    # các lớp còn lại bị buộc bằng đại diện, nên GV không bị tính trùng tiết.
    nonrep, joint_list = set(), []
    for grp in groups:
        t = grp["token"]
        cl = [c for c in grp["classes"] if c in classes]
        if len(cl) < 2 or any(base[c].get(t, 0) == 0 for c in cl):
            return False, f"Tiết ghép {t} {grp['classes']}: lớp không có trong file hoặc không có tiết {t}.", None
        if len({base[c][t] for c in cl}) > 1:
            return False, f"Tiết ghép {t}: số tiết của các lớp {cl} không bằng nhau.", None
        joint_list.append((t, cl, grp["fixed"]))
        nonrep.update((c, t) for c in cl[1:])

    for c in classes:
        cells = [i for i in range(nrow) if not ban.at[i, c]]
        counts = base[c]
        if sum(counts.values()) > len(cells):
            return False, (f"Lớp {c}: có {sum(counts.values())} tiết nhưng chỉ có "
                           f"{len(cells)} ô được phép xếp."), None
        bycell = {i: [] for i in cells}
        for t, cnt in counts.items():
            g, s = gv_name(t), subj(t)
            vs = []
            for i in cells:
                if on["chao_co"] and s == "CC" and (day_of[i], per_of[i]) != (2, 1):
                    continue
                if on["gv_tranh"] and (day_of[i], per_of[i]) in avoid.get(g, ()):
                    continue
                v = m.NewBoolVar(f"x_{c}_{t}_{i}")
                x[(c, t, i)] = v
                vs.append(v)
                bycell[i].append(v)
                if (c, t) not in nonrep:
                    teach.setdefault((g, i), []).append(v)
                sx.setdefault((c, s, i), []).append(v)
            if vs:
                m.Add(sum(vs) <= cnt)
        for vs in bycell.values():
            if vs:
                m.Add(sum(vs) <= 1)

    # ràng buộc nhóm tiết ghép: các lớp cùng ô, có thể cố định vị trí
    for t, cl, fixed in joint_list:
        rep_c = cl[0]
        for r in range(nrow):
            vr = x.get((rep_c, t, r))
            for c in cl[1:]:
                vc = x.get((c, t, r))
                if vr is None and vc is None:
                    continue
                if vr is None:
                    m.Add(vc == 0)
                elif vc is None:
                    m.Add(vr == 0)
                else:
                    m.Add(vr == vc)
        for d, p in fixed:
            r = rowof.get((d, p))
            for c in cl:
                v = x.get((c, t, r)) if r is not None else None
                if v is None:
                    return False, (f"Tiết ghép {t} cố định ở (Thứ {d}, tiết {p}): ô này bị cấm xếp "
                                   f"hoặc GV đã đăng ký tránh."), None
                m.Add(v == 1)

    # không trùng tiết (cùng Tên GV, cùng Thứ+Tiết)
    if on["trung_tiet"]:
        for vs in teach.values():
            if len(vs) > 1:
                m.Add(sum(vs) <= 1)

    # biến "GV bận" ở từng (GV, ô)
    b = {}
    if on["tiet_trong_buoi"] or on["cach_2_buoi"] or on.get("toi_thieu_2", True):
        for key, vs in teach.items():
            bv = m.NewBoolVar(f"b_{key}")
            m.AddMaxEquality(bv, vs)
            b[key] = bv
    gvs = sorted({k[0] for k in teach})

    def bval(g, d, p):
        r = rowof.get((d, p))
        return b.get((g, r)) if r is not None else None

    # giữa tiết đầu và tiết cuối trong buổi không có 2 tiết trống liên tiếp
    if on["tiet_trong_buoi"]:
        sessions = [list(range(1, MORNING_LAST + 1)), list(range(MORNING_LAST + 1, 11))]
        for g in gvs:
            for d in days:
                for ses in sessions:
                    lst = [bval(g, d, p) for p in ses]
                    for k in range(len(lst) - 1):
                        pre = [v for v in lst[:k] if v is not None]
                        post = [v for v in lst[k + 2:] if v is not None]
                        if not pre or not post:
                            continue
                        P = m.NewBoolVar("")
                        Q = m.NewBoolVar("")
                        for v in pre:
                            m.AddImplication(v, P)
                        for v in post:
                            m.AddImplication(v, Q)
                        expr = P + Q
                        for v in (lst[k], lst[k + 1]):
                            if v is not None:
                                expr = expr - v
                        m.Add(expr <= 1)

    # dạy cả 2 buổi: không được dạy liền tiết 5 (sáng) và tiết 6 (chiều)
    if on["cach_2_buoi"]:
        for g in gvs:
            for d in days:
                v1, v2 = bval(g, d, MORNING_LAST), bval(g, d, MORNING_LAST + 1)
                if v1 is not None and v2 is not None:
                    m.Add(v1 + v2 <= 1)

    # VV, TT: có 1 cặp tiết liền nhau, không xếp vào 2 ngày liền nhau
    if on["vv_tt"]:
        for c in classes:
            for s in ("VV", "TT"):
                allv = [v for i in range(nrow) for v in sx.get((c, s, i), [])]
                if not allv:
                    continue
                dv = {}
                for d in days:
                    dv[d] = m.NewBoolVar("")
                    for p in range(1, 11):
                        r = rowof.get((d, p))
                        if r is not None and (c, s, r) in sx:
                            m.Add(sum(sx[(c, s, r)]) <= dv[d])
                for d in days:
                    if d + 1 in dv:
                        m.Add(dv[d] + dv[d + 1] <= 1)
                need = sum(n for t, n in base[c].items() if subj(t) == s)
                if need < 2:
                    continue
                pairs = []
                for d in days:
                    for p in range(1, 10):
                        if p == MORNING_LAST:
                            continue          # tiết 5-6 khác buổi
                        r1, r2 = rowof.get((d, p)), rowof.get((d, p + 1))
                        if r1 is None or r2 is None:
                            continue
                        if (c, s, r1) in sx and (c, s, r2) in sx:
                            pv = m.NewBoolVar("")
                            m.Add(pv <= sum(sx[(c, s, r1)]))
                            m.Add(pv <= sum(sx[(c, s, r2)]))
                            pairs.append(pv)
                if not pairs:
                    m.Add(sum(allv) <= 1)
                else:
                    placed2 = m.NewBoolVar("")
                    m.Add(sum(allv) >= 2).OnlyEnforceIf(placed2)
                    m.Add(sum(allv) <= 1).OnlyEnforceIf(placed2.Not())
                    m.Add(sum(pairs) >= 1).OnlyEnforceIf(placed2)

    # GV có dạy trong buổi thì phải dạy ít nhất 2 tiết trong buổi đó
    if on.get("toi_thieu_2", True):
        for g in gvs:
            for d in days:
                for ses in ([1, 2, 3, 4, 5], [6, 7, 8, 9, 10]):
                    lst = [v for v in (bval(g, d, p) for p in ses) if v is not None]
                    if len(lst) == 1:
                        m.Add(lst[0] == 0)          # buổi này chỉ còn 1 ô khả dĩ: không thể đủ 2 tiết
                    elif len(lst) > 1:
                        for v in lst:
                            m.Add(sum(lst) >= 2).OnlyEnforceIf(v)

    # mỗi môn không quá 2 tiết/lớp/ngày
    if on.get("toi_da_2", True):
        for c in classes:
            for s_ in {subj(t) for t in base[c]}:
                for d in days:
                    vs = []
                    for p in range(1, 11):
                        r = rowof.get((d, p))
                        if r is not None:
                            vs += sx.get((c, s_, r), [])
                    if len(vs) > 2:
                        m.Add(sum(vs) <= 2)

    m.Maximize(sum(x.values()))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = float(tlimit)
    solver.parameters.random_seed = 7
    try:
        solver.parameters.num_workers = 4
    except AttributeError:
        solver.parameters.num_search_workers = 4
    status = solver.Solve(m)
    if status == cp_model.INFEASIBLE:
        return False, ("Các ràng buộc mâu thuẫn nhau (kiểm tra tiết ghép cố định, tránh tiết, "
                       "chào cờ...). Thử tắt bớt ràng buộc."), None
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return False, "Chưa tìm được lời giải trong thời gian cho phép. Hãy tăng thời gian.", None
    new = df.copy()
    for c in classes:
        for i in range(nrow):
            if not ban.at[i, c]:
                new.at[i, c] = ""
    for (c, t, i), v in x.items():
        if solver.BooleanValue(v):
            new.at[i, c] = t
    return True, "Thuật toán đã chạy xong.", new


# ------------------------------------------------------------- lưới / xuất
def decorate(df, ban, classes, pending, gmap=None):
    g = df.copy()
    g["_i"] = list(range(len(g)))
    bgs, dups, bans, sels = [], [], [], []
    for i in range(len(g)):
        d, p = int(g.at[i, "Thứ"]), int(g.at[i, "Tiết"])
        base = DAY_RGB.get(d, (235, 235, 235))
        bgs.append(hex_mix(base, 0.55) if p <= MORNING_LAST else hex_mix(base, 0.0))
        row_tok = {c: tok(g.at[i, c]) for c in classes if tok(g.at[i, c])}
        units = {}
        for c, t in row_tok.items():
            key = f"{c}|{t}"
            u = ("g", gmap[key]) if gmap and key in gmap else ("c", c)
            units.setdefault(gv_name(t), set()).add(u)
        dc = [c for c, t in row_tok.items() if len(units[gv_name(t)]) > 1]
        dups.append("|" + "|".join(dc) + "|" if dc else "")
        bans.append("|" + "|".join(c for c in classes if ban.at[i, c]) + "|")
        sels.append(pending[1] if pending and pending[0] == i else "")
    g["_bg"], g["_dup"], g["_ban"], g["_sel"] = bgs, dups, bans, sels
    g["_clk"] = ""
    return g, sum(1 for d in dups if d)


CELL_STYLE = JsCode("""
function(p){
  var d=p.data, f=p.colDef.field;
  if(!d) return null;
  var st={backgroundColor:d._bg, color:'#111', fontSize:'12px'};
  if(f==='Thứ'||f==='Tiết'||f==='Giờ'){
    st.backgroundColor='#1f4e9c'; st.color='#fff'; st.fontWeight='bold';
    if(d._dup && d._dup.length>0 && f!=='Giờ'){st.backgroundColor='#ff4d4d'; st.color='#fff'; st.fontWeight='bold';}
    return st;
  }
  if(d._ban.indexOf('|'+f+'|')>=0){st.backgroundColor='#bdbdbd'; return st;}
  if(d._dup.indexOf('|'+f+'|')>=0){st.backgroundColor='#ff4d4d'; st.color='#fff'; st.fontWeight='bold';}
  if(d._sel===f){st.border='3px solid #0057ff';}
  return st;
}
""")

EDITABLE_JS = JsCode("""
function(p){
  var f=p.colDef.field;
  if(f==='Thứ'||f==='Tiết'||f==='Giờ') return false;
  return p.data._ban.indexOf('|'+f+'|')<0;
}
""")


def count_single(df, classes):
    """Số (GV, ngày, buổi) mà GV chỉ dạy đúng 1 tiết trong buổi."""
    per = {}
    for i in range(len(df)):
        d, p = int(df.at[i, "Thứ"]), int(df.at[i, "Tiết"])
        ses = 0 if p <= MORNING_LAST else 1
        for g in {gv_name(tok(df.at[i, c])) for c in classes if tok(df.at[i, c])}:
            per.setdefault((g, d, ses), set()).add(p)
    return sum(1 for v in per.values() if len(v) == 1)


def fmt_date(d):
    return d.strftime("%d/%m/%Y").lstrip("0").replace("/0", "/")


def week_table(kind, df, classes, name):
    """kind='gv' hoặc 'lop'. Trả về (danh sách thứ, các hàng [tiết, giờ, [ô thứ 2], [ô thứ 3]...])."""
    days = sorted(set(df["Thứ"].astype(int)))
    times = {}
    for p, h in zip(df["Tiết"], df["Giờ"]):
        times.setdefault(int(p), h)
    rows = []
    for p in range(1, 11):
        row = [p, times.get(p, "")]
        for d in days:
            r = df[(df["Thứ"] == d) & (df["Tiết"] == p)].index
            cells = []
            if len(r):
                r = r[0]
                if kind == "gv":
                    cells = [f"{c}_{subj(tok(df.at[r, c]))}" for c in classes
                             if tok(df.at[r, c]) and gv_name(tok(df.at[r, c])) == name]
                else:
                    t = tok(df.at[r, name])
                    cells = [t] if t else []
            row.append(cells)
        rows.append(row)
    return days, rows


def week_html(kind, df, classes, name, school, apply_from, gmap=None):
    """TKB tuần của 1 GV / 1 lớp, trình bày giống bản in PDF."""
    days, rows = week_table(kind, df, classes, name)
    title = ("TKB của GV: " if kind == "gv" else "TKB của Lớp: ") + name
    b = "border:1px solid #000;padding:4px 8px;text-align:center;color:#000;"
    out = ["<div style='background:#fff;color:#000;padding:10px 14px;border-radius:6px;overflow-x:auto'>",
           "<div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:8px'>"
           f"<div>Trường : {html.escape(school)}</div>"
           f"<div style='font-weight:bold;font-size:1.15em'>{html.escape(title)}</div>"
           f"<div>Áp dụng từ: {fmt_date(apply_from)}</div></div>",
           "<table style='border-collapse:collapse;width:100%'><tr>"
           + f"<th style='{b}background:#d9d9d9'>Tiết</th><th style='{b}background:#d9d9d9'>Giờ</th>"
           + "".join(f"<th style='{b}background:#d9d9d9'>Thứ {d}</th>" for d in days) + "</tr>"]
    total = 0
    for p, gio, *cells in rows:
        bg = "#e6e6e6" if p > MORNING_LAST else "#ffffff"
        tds = [f"<td style='{b}background:{bg}'>{p}</td>", f"<td style='{b}background:{bg}'>{html.escape(str(gio))}</td>"]
        for cl in cells:
            total += len(cl)
            cbg = "#ffb3b3" if (kind == "gv" and cells_clash(name, cl, gmap)) else bg
            tds.append(f"<td style='{b}background:{cbg}'>{'<br>'.join(html.escape(x) for x in cl)}</td>")
        out.append("<tr>" + "".join(tds) + "</tr>")
    out.append("</table></div>")
    return "".join(out), total


def load_avoid_df(f):
    a = pd.read_excel(f, dtype=object)
    a.columns = [str(c).strip() for c in a.columns]
    a = a.rename(columns={a.columns[0]: "ten_gv"})
    for c in AVOID_COLS:
        if c not in a.columns:
            a[c] = ""
    a = a[AVOID_COLS].copy()
    for c in AVOID_COLS:
        a[c] = a[c].map(tok)
    return a[a["ten_gv"] != ""].reset_index(drop=True)


def avoid_set_row(adf, name, pairs):
    """Cập nhật các cột Thứ 2..Thứ 7 của GV (nối thêm dòng nếu chưa có tên). Mỗi cột Thứ d ghi các ô (d,tiết)."""
    adf = adf.copy()
    vals = {f"Thứ {d}": ", ".join(f"({d},{p})" for p in sorted(pp for dd, pp in pairs if dd == d))
            for d in range(2, 8)}
    mask = adf["ten_gv"].map(lambda v: tok(v).casefold()) == name.casefold()
    if mask.any():
        for col, v in vals.items():
            adf.loc[mask, col] = v
    elif pairs:
        adf = pd.concat([adf, pd.DataFrame([{"ten_gv": name, **vals}])], ignore_index=True)
    return adf


def avoid_bytes(adf):
    out = adf.copy()
    for c in out.columns:
        out[c] = out[c].map(lambda v: tok(v) or None)
    b = io.BytesIO()
    out.to_excel(b, index=False)
    return b.getvalue()


def count_over2(df, classes):
    """Số trường hợp (lớp, môn, ngày) có hơn 2 tiết trong lưới hiện tại."""
    n = 0
    for c in classes:
        for _, grp in df.groupby("Thứ"):
            cnt = Counter(subj(tok(v)) for v in grp[c] if tok(v))
            n += sum(1 for k in cnt.values() if k > 2)
    return n


def remaining(df, ban, classes, base):
    """Các tiết (theo lớp) còn chưa có trong lưới."""
    out = {}
    for c in classes:
        placed = Counter(tok(df.at[i, c]) for i in range(len(df))
                         if not ban.at[i, c] and tok(df.at[i, c]))
        out[c] = Counter(base[c]) - placed
    return out


def recalc_js(gmap):
    """JS tính lại cột ẩn _dup cho 1 dòng (GV trùng tiết), bỏ qua các nhóm tiết ghép."""
    return "var JG=" + json.dumps(gmap or {}, ensure_ascii=False) + ";\n" + """
function recalc(node){
  var d=node.data, units={}, items=[];
  for(var k in d){
    if(k==='Thứ'||k==='Tiết'||k==='Giờ'||k.charAt(0)==='_') continue;
    var v=d[k]; if(!v) continue;
    var t=String(v).trim(), i=t.lastIndexOf('_');
    var n=(i>=0?t.substring(0,i):t).trim();
    var key=k+'|'+t;
    var u=Object.prototype.hasOwnProperty.call(JG,key)?('g'+JG[key]):('c'+k);
    if(!units[n]) units[n]={};
    units[n][u]=1; items.push([k,n]);
  }
  var dc=[];
  items.forEach(function(e){ if(Object.keys(units[e[1]]).length>1) dc.push(e[0]); });
  d._dup = dc.length ? '|'+dc.join('|')+'|' : '';
}
"""


def click_js(mode, gmap=None):
    """swap_1: hoán vị ngay trong trình duyệt (êm, không tải lại lưới), rồi báo Python 1 lần.
    swap_2: ghi '<lớp>|<mốc giờ>' vào cột ẩn _clk để Python mở popup TKB giáo viên."""
    if mode == "swap1":
        return JsCode("""
function(p){
  """ + recalc_js(gmap) + """
  var f=p.column.getColId();
  if(f==='Thứ'||f==='Tiết'||f==='Giờ'||f.charAt(0)==='_') return;
  var d=p.node.data, api=p.api;
  if(d._ban.indexOf('|'+f+'|')>=0) return;
  var w=window.__tkb||(window.__tkb={});
  if(!w.pend){
    w.pend={node:p.node,f:f}; d._sel=f;
    api.refreshCells({rowNodes:[p.node],force:true});
    return;
  }
  var a=w.pend; w.pend=null; a.node.data._sel='';
  if(a.node===p.node && a.f===f){ api.refreshCells({rowNodes:[p.node],force:true}); return; }
  var t=a.node.data[a.f]; a.node.data[a.f]=d[f]; d[f]=t;
  recalc(a.node); recalc(p.node);
  api.refreshCells({rowNodes:[a.node,p.node],force:true});
  p.node.setDataValue('_clk','sync|'+Date.now());
}
""")
    if mode == "swap2":
        return JsCode("""
function(p){
  var f=p.column.getColId();
  if(f==='Thứ'||f==='Tiết'||f==='Giờ'||f.charAt(0)==='_') return;
  p.node.setDataValue('_clk', f+'|'+Date.now());
}
""")
    return None


def edit_js(gmap=None):
    """Chế độ tự do: sau khi sửa ô, kiểm tra dạng Tên_Môn và cập nhật màu đỏ ngay trong trình duyệt."""
    return JsCode("""
function(p){
  """ + recalc_js(gmap) + """
  var f=p.colDef.field;
  if(f==='Thứ'||f==='Tiết'||f==='Giờ'||f.charAt(0)==='_') return;
  var v=(p.newValue==null)?'':String(p.newValue).trim();
  if(v!=='' && !/^[^_\\s][^_]*_[A-Z]{2}$/.test(v)){
    p.node.data[f]=(p.oldValue==null)?'':p.oldValue;
  } else { p.node.data[f]=v; }
  recalc(p.node);
  p.api.refreshCells({rowNodes:[p.node],force:true});
}
""")


def teacher_table(df, classes, name):
    days = sorted(set(df["Thứ"].astype(int)))
    times = {int(p): h for p, h in zip(df["Tiết"], df["Giờ"])}
    rows = []
    for p in range(1, 11):
        row = {"Tiết": p, "Giờ": times.get(p, "")}
        for d in days:
            sel = df[(df["Thứ"] == d) & (df["Tiết"] == p)]
            cells = []
            if len(sel):
                r = sel.index[0]
                cells = [f"{c}_{subj(tok(df.at[r, c]))}" for c in classes
                         if tok(df.at[r, c]) and gv_name(tok(df.at[r, c])) == name]
            row[f"Thứ {d}"] = " + ".join(cells)
        rows.append(row)
    return pd.DataFrame(rows)


def to_excel_bytes(df, classes):
    out = df.copy()
    for c in classes:
        out[c] = out[c].map(lambda v: tok(v) or None)
    b = io.BytesIO()
    with pd.ExcelWriter(b, engine="openpyxl") as w:
        out.to_excel(w, index=False, sheet_name="TKB")
    return b.getvalue()


# ---------------------------------------------------------------------- PDF
def register_fonts():
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    cands = [
        (os.path.join(APP_DIR, "fonts", "DejaVuSans.ttf"), os.path.join(APP_DIR, "fonts", "DejaVuSans-Bold.ttf")),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/dejavu/DejaVuSans.ttf", "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf"),
        ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
        ("/Library/Fonts/Arial.ttf", "/Library/Fonts/Arial Bold.ttf"),
    ]
    for reg, bold in cands:
        if os.path.exists(reg):
            pdfmetrics.registerFont(TTFont("VN", reg))
            pdfmetrics.registerFont(TTFont("VN-B", bold if os.path.exists(bold) else reg))
            return
    raise RuntimeError("Không tìm thấy font Unicode. Thêm 'fonts-dejavu-core' vào packages.txt "
                       "hoặc đặt DejaVuSans.ttf vào thư mục fonts/.")


def build_pdf(kind, df, classes, school, apply_from, only=None):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A3, A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    register_fonts()
    days = sorted(set(df["Thứ"].astype(int)))
    times = {}
    for p, h in zip(df["Tiết"], df["Giờ"]):
        times.setdefault(int(p), h)
    date_txt = fmt_date(apply_from)

    big = kind == "tkb_all" and len(classes) > 22
    page = landscape(A3 if big else A4)
    margin = 24
    avail = page[0] - 2 * margin
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=page, leftMargin=margin, rightMargin=margin,
                            topMargin=margin, bottomMargin=margin)
    st_n = ParagraphStyle("n", fontName="VN", fontSize=10, leading=13)
    st_b = ParagraphStyle("b", fontName="VN-B", fontSize=13, leading=16, alignment=1)
    st_r = ParagraphStyle("r", fontName="VN", fontSize=10, leading=13, alignment=2)

    def header(title):
        t = Table([[Paragraph(f"Trường : {school}", st_n), "", ""],
                   ["", Paragraph(title, st_b), Paragraph(f"Áp dụng từ: {date_txt}", st_r)]],
                  colWidths=[avail * 0.3, avail * 0.4, avail * 0.3])
        t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
        return [t, Spacer(1, 8)]

    def grid(data, widths, fs, heights=None, shade=()):
        t = Table(data, colWidths=widths, rowHeights=heights, repeatRows=1)
        style = [
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("FONTNAME", (0, 0), (-1, -1), "VN"),
            ("FONTNAME", (0, 0), (-1, 0), "VN-B"),
            ("FONTSIZE", (0, 0), (-1, -1), fs),
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]
        for r in shade:      # buổi chiều: xám nhạt
            style.append(("BACKGROUND", (0, r), (-1, r), colors.Color(0.9, 0.9, 0.9)))
        t.setStyle(TableStyle(style))
        return t

    el = []
    if kind == "tkb_all":
        n = len(classes)
        fs = 9 if n <= 12 else 8 if n <= 20 else 7 if n <= 30 else 6
        fixed_w = [34, 34, 76]
        cw = max(30, (avail - sum(fixed_w)) / max(n, 1))
        for k, d in enumerate(days):
            el += header("TKB Toàn Trường")
            data = [FIXED + classes]
            sub = df[df["Thứ"] == d]
            for r in sub.index:
                data.append([int(df.at[r, "Thứ"]), int(df.at[r, "Tiết"]), df.at[r, "Giờ"]] +
                            [tok(df.at[r, c]) for c in classes])
            shade = [k2 + 1 for k2, r in enumerate(sub.index) if int(df.at[r, "Tiết"]) > MORNING_LAST]
            el.append(grid(data, fixed_w + [cw] * n, fs, None, shade))
            if k < len(days) - 1:
                el.append(PageBreak())
    else:
        if kind == "tkb_moi_gv":
            names = sorted({gv_name(tok(v)) for c in classes for v in df[c] if tok(v)})
            title_fmt = "TKB của GV: {}"
        else:
            names = list(classes)
            title_fmt = "TKB của Lớp: {}"
        if only:
            names = [only]
        cw = [40, 80] + [(avail - 120) / len(days)] * len(days)
        for k, nm in enumerate(names):
            el += header(title_fmt.format(nm))
            data = [["Tiết", "Giờ"] + [f"Thứ {d}" for d in days]]
            for p in range(1, 11):
                row = [p, times.get(p, "")]
                for d in days:
                    r = df[(df["Thứ"] == d) & (df["Tiết"] == p)].index
                    txt = ""
                    if len(r):
                        r = r[0]
                        if kind == "tkb_moi_gv":
                            txt = "\n".join(f"{c}_{subj(tok(df.at[r, c]))}" for c in classes
                                            if tok(df.at[r, c]) and gv_name(tok(df.at[r, c])) == nm)
                        else:
                            txt = tok(df.at[r, nm])
                    row.append(txt)
                data.append(row)
            el.append(grid(data, cw, 10, [24] + [34] * 10, range(MORNING_LAST + 1, 11)))
            if k < len(names) - 1:
                el.append(PageBreak())
    doc.build(el)
    return buf.getvalue()


# =============================================================== GIAO DIỆN
st.set_page_config(page_title="Xếp & In TKB Trường học", layout="wide")
st.header("⚙️ Xếp & In TKB Trường học với tiensg89@")

S = st.session_state
if "joint" not in st.session_state:
    st.session_state.joint = pd.DataFrame(columns=JOINT_COLS)
if "avoid_df" not in st.session_state:
    st.session_state.avoid_df = pd.DataFrame(columns=AVOID_COLS)
for k, v in {"avoid_sig": None, "avoid_ver": 0, "joint_sig": None, "auth": False, "df": None, "ban": None, "classes": [], "pending": None,
             "gk": 0, "report": None, "sig": None, "avoid": {}, "view_gv": None,
             "flash": None, "swap1": False, "swap2": False, "base": {}, "last_clk": ""}.items():
    S.setdefault(k, v)
if "rules" not in S:
    S.rules = load_rules()
for _k, _v in DEFAULT_RULES.items():
    S.rules.setdefault(_k, json.loads(json.dumps(_v)))

# ---- 2. mật mã
if not S.auth:
    pw = st.text_input("Nhập mật mã để làm việc tiếp", type="password")
    real = st.secrets.get("APP_PASSWORD", "tkb@2026") if hasattr(st, "secrets") else "tkb@2026"
    if pw:
        if pw == real:
            S.auth = True
            st.rerun()
        else:
            st.error("Mật mã không đúng.")
    st.stop()


# ---- dialogs
@st.dialog("Kết quả kiểm tra file Excel", width="large")
def show_report(rep):
    for ok, title, det in rep:
        st.markdown(("✅ " if ok else "❌ ") + title)
        for d in det:
            st.caption("• " + d)
    if report_ok(rep):
        st.success("Cấu trúc tệp ĐÚNG. OK!")
    else:
        st.error("Cấu trúc tệp SAI → từ chối xử lí. Vui lòng sửa tệp và upload lại.")


@st.dialog("Chọn ràng buộc xếp TKB", width="large")
def constraint_dialog():
    with st.expander("Nội dung file cacrangbuocxep.json", expanded=True):
        st.json(S.rules)
    new = json.loads(json.dumps(S.rules))
    for k, v in S.rules.items():
        new[k]["bat"] = st.checkbox(v["mo_ta"], value=v["bat"], disabled=v.get("khoa", False),
                                    key=f"rule_{k}")
    if st.button("Lưu lựa chọn", type="primary"):
        S.rules = new
        save_rules(new)
        st.rerun()


def on_swap1():
    if S.swap1:
        S.swap2 = False
    S.pending = None
    S.gk += 1


def on_swap2():
    if S.swap2:
        S.swap1 = False
    S.pending = None
    S.gk += 1


@st.dialog("TKB của giáo viên", width="large")
def show_teacher(name, df, classes, hit):
    tt = teacher_table(df, classes, name)
    n_les = int(sum(1 for c in tt.columns if c.startswith("Thứ") for v in tt[c] if v))
    st.markdown(f"### GV: {name}  —  {n_les} tiết/tuần")

    def sty(t):
        css = pd.DataFrame("", index=t.index, columns=t.columns)
        for r in t.index:
            p = int(t.at[r, "Tiết"])
            for c in t.columns:
                if p > MORNING_LAST:
                    css.at[r, c] = "background-color:#eeeeee"
                if c.startswith("Thứ"):
                    d = int(c.split()[1])
                    if (d, p) == hit:
                        css.at[r, c] = "background-color:#fff176;font-weight:bold"
                    if cells_clash(name, [x_ for x_ in str(t.at[r, c]).split(" + ") if x_], S.gmap):
                        css.at[r, c] = "background-color:#ff4d4d;color:white;font-weight:bold"
        return css

    st.dataframe(tt.style.apply(sty, axis=None), use_container_width=True, hide_index=True)
    st.caption("Ô vàng: ô vừa bấm. Ô đỏ: GV bị trùng tiết. Hàng xám: buổi chiều.")


# ---- 3. upload
st.subheader("1️⃣Lấy dữ liệu")
up = st.file_uploader("Upload tệp Excel TKB (.xlsx)", type=["xlsx"])
if up is not None:
    sig = (up.name, up.size)
    if S.sig != sig:
        S.sig = sig
        try:
            raw = pd.read_excel(up)
            raw.columns = [str(c).strip() for c in raw.columns]
            raw = raw.drop(columns=[c for c in raw.columns
                                    if c.startswith("Unnamed") and raw[c].isna().all()])
            rep = validate(raw)
        except Exception as e:
            raw, rep = None, [(False, f"Không đọc được tệp: {e}", [])]
        S.report = rep
        S.pending, S.view_gv = None, None
        S.gk += 1
        if report_ok(rep):
            classes = list(raw.columns[3:])
            df = raw.copy()
            df["Thứ"] = df["Thứ"].astype(int)
            df["Tiết"] = df["Tiết"].astype(int)
            df["Giờ"] = df["Giờ"].astype(str)
            for c in classes:
                df[c] = df[c].map(tok)
            df = df.reset_index(drop=True)
            S.ban = pd.DataFrame({c: (df[c] == "") for c in classes})
            S.df, S.classes = df, classes
            S.base = {c: Counter(t for t in df[c] if t) for c in classes}
        else:
            S.df = None
else:
    S.sig = None
    S.df = None

if S.report is not None:
    c1, c2 = st.columns([1, 3])
    if c1.button("🔍 Kiểm tra file excel"):
        show_report(S.report)
    if report_ok(S.report):
        c2.success("Tệp hợp lệ.")
    else:
        c2.error("Tệp không đúng cấu trúc — từ chối xử lí. Bấm 'Kiểm tra file excel' để xem lỗi, rồi làm lại.")

if S.df is None:
    st.info("Hãy upload tệp Excel TKB hợp lệ để bắt đầu.")
    st.stop()

df, ban, classes = S.df, S.ban, S.classes
S.groups = parse_joint(S.joint)
S.gmap = joint_map(S.groups)

# ---- 6, 7. xếp TKB
st.subheader("2️⃣Xếp TKB tự động (Google OR-Tools)")
a1, a2, a3 = st.columns([2, 2, 2])
with a1:
    avoid_file = st.file_uploader("Lấy tệp gvyctranh.xlsx (cột ten_gv, Thứ 2..Thứ 7; ô ghi các cặp (Thứ,tiết))", type=["xlsx"], key="avoid_up")
    st.download_button("Tải mẫu gvyctranh.xlsx", avoid_template(), "gvyctranh_mau.xlsx")
    if avoid_file is None:
        S.avoid_sig = None       # bỏ file đi thì lần upload sau sẽ được nạp lại
    else:
        _sig = (avoid_file.name, hashlib.md5(avoid_file.getvalue()).hexdigest())
        if S.avoid_sig != _sig:
            try:
                S.avoid_df = load_avoid_df(avoid_file)
                S.avoid_sig = _sig
                S.avoid_ver += 1
            except Exception as e:
                st.error(f"Lỗi đọc gvyctranh.xlsx: {e}")
    S.avoid = parse_avoid(S.avoid_df)
    st.caption(f"Đang có yêu cầu tránh tiết của {len(S.avoid)} GV (nhập thêm bằng nút bên dưới).")
    _known = {gv_name(t).casefold() for c in classes for t in S.base[c]}
    _unknown = [n for n in S.avoid if n.casefold() not in _known]
    if _unknown:
        st.warning("GV trong gvyctranh.xlsx không có trong file TKB: " + ", ".join(_unknown))

    st.markdown("---")
    joint_file = st.file_uploader("tietghep.xlsx (cột GV_Mon, Lop, Vi_tri)", type=["xlsx"], key="joint_up")
    st.download_button("Tải mẫu tietghep.xlsx", joint_template(), "tietghep_mau.xlsx", key="dl_joint_tpl")
    if joint_file is None:
        S.joint_sig = None
    else:
        _jsig = hashlib.md5(joint_file.getvalue()).hexdigest()
        if S.joint_sig != _jsig:
            try:
                S.joint = load_joint_df(joint_file)
                S.joint_sig = _jsig
                S.gk += 1          # dựng lại lưới để cập nhật màu theo nhóm ghép
            except Exception as e:
                st.error(f"Lỗi đọc tietghep.xlsx: {e}")
    S.groups = parse_joint(S.joint)
    S.gmap = joint_map(S.groups)
    st.caption(f"Đang có {len(S.groups)} nhóm tiết ghép (quản lý bằng nút Tiết ghép bên dưới).")
    if len(S.joint) > len(S.groups):
        st.warning(f"Có {len(S.joint) - len(S.groups)} dòng tiết ghép bị bỏ qua (thiếu GV_Mon hoặc dưới 2 lớp).")
    for _p in check_joint(S.groups, S.base, classes):
        st.warning(_p)
with a2:
    tlimit = st.select_slider("Thời gian chạy thuật toán (giây)", options=[30, 60, 120, 180, 240, 300], value=60)
    if st.button("⚙️ Chọn lại ràng buộc"):
        constraint_dialog()
with a3:
    st.write("Ràng buộc đang bật:")
    for k, v in S.rules.items():
        st.caption(("✅ " if v["bat"] else "⬜ ") + v["mo_ta"])

# ---- nhập yêu cầu tránh tiết cho GV (dialog, đặt ngay sau phần gvyctranh.xlsx, trước khi xếp TKB)
@st.dialog("Nhập yêu cầu tránh tiết cho GV", width="large")
def dlg_avoid():
    names = sorted({gv_name(t) for c in S.classes for t in S.base[c]}
                   | {gv_name(tok(v)) for c in S.classes for v in S.df[c] if tok(v)})
    sel = st.selectbox("Chọn giáo viên", names, index=None, placeholder="Chọn giáo viên...",
                       key="menu_av")
    if not sel:
        return
    days_ = sorted(set(S.df["Thứ"].astype(int)))
    cur = next((v for k, v in S.avoid.items() if k.casefold() == sel.casefold()), set())
    st.markdown(f"**Yêu cầu xếp tránh tiết của GV: {sel}**")
    st.caption("Nhập ký tự x vào ô phù hợp với các tiết cần tránh xếp trong ngày.")
    tbl = pd.DataFrame({"Tiết": list(range(1, 11))})
    for d in days_:
        tbl[f"Thứ {d}"] = ["x" if (d, p) in cur else "" for p in range(1, 11)]
    cfg = {"Tiết": st.column_config.NumberColumn("Tiết", width="small")}
    for d in days_:
        cfg[f"Thứ {d}"] = st.column_config.TextColumn(f"Thứ {d}", max_chars=1, width="small")
    edited = st.data_editor(tbl, hide_index=True, disabled=["Tiết"], column_config=cfg,
                            key=f"avoid_ed_{sel}_{S.avoid_ver}")
    if st.button("💾 Lưu dữ liệu", key=f"save_av_{sel}"):
        pairs = {(d, int(p)) for d in days_
                 for p, v in zip(edited["Tiết"], edited[f"Thứ {d}"]) if tok(v)}
        S.avoid_df = avoid_set_row(S.avoid_df, sel, pairs)
        S.avoid = parse_avoid(S.avoid_df)
        S.avoid_ver += 1
        S.avoid_msg = (f"Đã lưu {len(pairs)} tiết tránh của GV {sel}. "
                       "Bấm 'Tải về gvyctranh.xlsx' để lấy file đã cập nhật.")
        st.rerun()   # đóng dialog và chạy lại để nút tải nhận dữ liệu mới


@st.dialog("Tiết ghép: 1 GV dạy nhiều lớp cùng lúc", width="large")
def dlg_joint():
    st.caption("Dùng cho môn như Thể dục: 1 GV dạy nhiều lớp chung 1 tiết ngoài sân. "
               "Các lớp trong nhóm luôn học cùng ô và không bị tính là GV trùng tiết.")
    toks = sorted({t for c in S.classes for t in S.base[c]})
    tk = st.selectbox("GV_Môn", toks, index=None, placeholder="Chọn GV_Môn (ví dụ Minh_TD)...", key="jn_tok")
    opts = [c for c in S.classes if tk and S.base[c].get(tk, 0) > 0]
    cl = st.multiselect("Các lớp học chung", opts, key="jn_cls")
    fx = st.text_input("Vị trí cố định (Thứ,Tiết) - để trống nếu để máy tự chọn. Ví dụ: (4,2)", key="jn_fix")
    if st.button("➕ Thêm nhóm ghép"):
        pairs = PAIR_RE.findall(fx)
        if not tk or len(cl) < 2:
            st.error("Hãy chọn GV_Môn và ít nhất 2 lớp.")
        elif fx.strip() and not pairs:
            st.error("Vị trí cố định phải có dạng (Thứ,Tiết), ví dụ (4,2).")
        else:
            row = {"GV_Mon": tk, "Lop": ",".join(cl),
                   "Vi_tri": ", ".join(f"({d},{p})" for d, p in pairs)}
            S.joint = pd.concat([S.joint, pd.DataFrame([row])], ignore_index=True)
            S.gk += 1
            st.success(f"Đã thêm nhóm ghép {tk}: {', '.join(cl)}.")
    if len(S.joint):
        st.markdown("**Các nhóm tiết ghép hiện có**")
        show = S.joint.copy()
        show.insert(0, "STT", range(1, len(show) + 1))
        st.dataframe(show, hide_index=True, use_container_width=True)
        dsel = st.selectbox("Xóa nhóm số", list(range(1, len(S.joint) + 1)), index=None,
                            placeholder="Chọn STT...", key="jn_del")
        if st.button("🗑 Xóa nhóm đã chọn") and dsel:
            S.joint = S.joint.drop(index=dsel - 1).reset_index(drop=True)
            S.gk += 1
            st.rerun()
    st.caption("Muốn nạp từ file: upload tietghep.xlsx ở mục xếp TKB (cạnh gvyctranh.xlsx).")
    st.download_button("⬇️ Tải về tietghep.xlsx", avoid_bytes(S.joint.rename(columns={}).copy()),
                       "tietghep.xlsx", key="dl_joint",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if st.button("✔ Xong, đóng"):
        st.rerun()


d1, d2, d3 = st.columns(3)
if d3.button("🤝 Tiết ghép (1 GV nhiều lớp)"):
    dlg_joint()
d1.download_button("⬇️ Tải về gvyctranh.xlsx", avoid_bytes(S.avoid_df), "gvyctranh.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   key="dl_avoid")
if d2.button("✍️ Nhập / xem tránh tiết cho GV"):
    S.menu_av = None
    dlg_avoid()
if S.get("avoid_msg"):
    st.success(S.avoid_msg)
    S.avoid_msg = None
st.caption("Yêu cầu tránh tiết được giữ trong phiên làm việc này. Hãy tải gvyctranh.xlsx về, "
           "lần sau upload lại ở trên.")
if len(S.avoid_df):
    with st.expander("Xem nội dung gvyctranh.xlsx hiện tại"):
        st.dataframe(S.avoid_df, hide_index=True, use_container_width=True)

run = st.button("▶️ Chạy xếp TKB", type="primary")

if run:
    with st.spinner(f"Đang xếp TKB (tối đa {tlimit} giây)...", show_time=True):
        ok, msg, new = solve_tkb(df, ban, classes, S.base, S.rules, S.avoid, tlimit, S.groups)
    if ok:
        S.df = new
        S.pending = None
        S.gk += 1
        S.flash = ("success", msg)
        st.rerun()
    else:
        st.error(msg)

if S.flash:
    getattr(st, S.flash[0])(S.flash[1])
    S.flash = None

# ---- thống kê tiết đã / chưa xếp
rem = remaining(df, ban, classes, S.base)
n_rem = sum(sum(v.values()) for v in rem.values())
n_total = sum(sum(v.values()) for v in S.base.values())
pct = round(100 * (n_total - n_rem) / n_total) if n_total else 100
st.markdown(f"**Tỉ lệ % đã xếp được là {pct}%**  \n"
            f"Còn **{n_rem}** tiết chưa xếp (chưa đem vào lưới và file kết quả).")
if n_rem:
    st.write("❓ Có cần đem hết vào lưới để tự xếp không (Y/N)? Chọn Y thì bấm nút bên dưới.")
    with st.expander("Xem các tiết chưa xếp"):
        st.dataframe(pd.DataFrame([{"Lớp": c, "Số tiết": sum(v.values()),
                                    "Các tiết": ", ".join(f"{t} x{n}" for t, n in sorted(v.items()))}
                                   for c, v in rem.items() if v]),
                     hide_index=True, use_container_width=True)
if st.button("📥 Đưa các tiết chưa xếp vào lưới", disabled=(n_rem == 0)):
    newdf = df.copy()
    for c in classes:
        free_cells = [i for i in range(len(df)) if not ban.at[i, c] and df.at[i, c] == ""]
        for i, t in zip(free_cells, list(rem[c].elements())):
            newdf.at[i, c] = t
    S.df = newdf
    S.pending = None
    S.gk += 1
    S.flash = ("info", "Đã đưa các tiết chưa xếp vào các ô trống của lưới. Hãy tinh chỉnh bằng tay (ô đỏ là chỗ GV trùng tiết).")
    st.rerun()

# ---- 5, 8. lưới
st.subheader("3️⃣Tinh chỉnh TKB)")
t1, t2, t3 = st.columns([1, 1, 3])
t1.toggle("Bật/Tắt: Hoán vị ô", key="swap1", on_change=on_swap1)
t2.toggle("Bật/Tắt: Xem TKB của GV", key="swap2", on_change=on_swap2)
if S.swap1:
    t3.info("Hoán vị: bấm ô thứ nhất (viền xanh) rồi bấm ô thứ hai để đổi chỗ ngay trên lưới. Ô xám cấm xếp không đổi được.")
elif S.swap2:
    t3.info("Bấm vào ô có tên GV để mở cửa sổ TKB của GV đó.")
else:
    t3.info("Chế độ tự do: bấm đúp vào ô để sửa trực tiếp (dạng Tên_Môn, môn 2 ký tự HOA).")

# Lưới được dựng 1 lần cho mỗi "phiên bản" gk. Hoán vị/sửa ô diễn ra ngay trong trình duyệt
# nên lưới không bị tải lại (không chớp, không mất vị trí con trỏ).
if S.get("grid_df") is None or S.get("grid_gk") != S.gk:
    S.grid_df, _ = decorate(df, ban, classes, None, S.gmap)
    S.grid_gk = S.gk
gdf = S.grid_df
_, ndup = decorate(df, ban, classes, None, S.gmap)
if ndup:
    st.warning(f"Có {ndup} dòng có GV trùng tiết (tô đỏ).")
else:
    st.success("Không có GV nào trùng tiết.")
if S.rules.get("toi_da_2", {}).get("bat"):
    over = count_over2(df, classes)
    if over:
        st.warning(f"Có {over} trường hợp một môn học quá 2 tiết/lớp trong một ngày.")
if S.rules.get("toi_thieu_2", {}).get("bat"):
    single = count_single(df, classes)
    if single:
        st.warning(f"Có {single} buổi mà GV chỉ dạy đúng 1 tiết (nên có ít nhất 2 tiết/buổi).")

mode = "swap1" if S.swap1 else "swap2" if S.swap2 else "free"
gb = GridOptionsBuilder.from_dataframe(gdf)
gb.configure_default_column(editable=EDITABLE_JS if mode == "free" else False, sortable=False,
                            filter=False, resizable=True, suppressMovable=True, minWidth=80,
                            cellStyle=CELL_STYLE)
for c in HELPER_COLS:
    gb.configure_column(c, hide=True)
gb.configure_column("Thứ", pinned="left", width=70, editable=False)
gb.configure_column("Tiết", pinned="left", width=70, editable=False)
gb.configure_column("Giờ", hide=True)
cj = click_js(mode, S.gmap)
if cj is not None:
    gb.configure_grid_options(onCellClicked=cj)
if mode == "free":
    gb.configure_grid_options(onCellValueChanged=edit_js(S.gmap))
resp = AgGrid(gdf, gridOptions=gb.build(), update_on=["cellValueChanged"],
              allow_unsafe_jscode=True, key=f"grid_{S.gk}", height=620)

nd = resp.data if hasattr(resp, "data") else None
if nd is not None and len(nd):
    nd = pd.DataFrame(nd)
    if "_i" in nd.columns:
        if mode in ("free", "swap1"):
            # đồng bộ dữ liệu từ lưới về Python (không tải lại lưới)
            nd = nd.set_index("_i")
            changed, bad = False, False
            newdf = df.copy()
            for i in range(len(df)):
                if i not in nd.index:
                    continue
                for c in classes:
                    if ban.at[i, c]:
                        continue
                    v = tok(nd.at[i, c])
                    if v != df.at[i, c]:
                        if v and not TOKEN_RE.match(v):
                            bad = True
                        else:
                            newdf.at[i, c] = v
                            changed = True
            if changed or bad:
                if bad:
                    st.toast("Giá trị không đúng dạng Tên_Môn (2 ký tự HOA) nên bị bỏ qua.")
                    S.gk += 1          # chỉ khi sai mới phải dựng lại lưới
                S.df = newdf
                st.rerun()
        elif "_clk" in nd.columns:
            # swap_2: bấm ô -> popup TKB giáo viên
            clk = nd[nd["_clk"].astype(str).str.contains("|", regex=False)].copy()
            if len(clk):
                clk["_ts"] = clk["_clk"].astype(str).str.rsplit("|", n=1).str[-1].astype(float)
                last = clk.sort_values("_ts").iloc[-1]
                val = str(last["_clk"])
                if val != S.last_clk:
                    S.last_clk = val
                    col = val.rsplit("|", 1)[0]
                    i = int(last["_i"])
                    if col in classes:
                        t = tok(df.at[i, col])
                        if t:
                            show_teacher(gv_name(t), df, classes,
                                         (int(df.at[i, "Thứ"]), int(df.at[i, "Tiết"])))
                        else:
                            st.toast("Ô này chưa có GV.")

# ---- Xem/In rời rạc & nhập tránh tiết
st.subheader("4️⃣Chỉnh tiêu đề bản in")
q1, q2 = st.columns(2)
school = q1.text_input("Trường", "ABC")
apply_from = q2.date_input("Áp dụng từ", value=date(2026, 9, 5), format="DD/MM/YYYY")

st.subheader("5️⃣Xem/Tải TKB Giáo viên/Lớp")
gv_list = sorted({gv_name(t) for c in classes for t in S.base[c]}
                 | {gv_name(tok(v)) for c in classes for v in df[c] if tok(v)})
def _pick(kind, key):
    """Chọn trong selectbox -> yêu cầu mở dialog, rồi trả selectbox về trạng thái chưa chọn
    (để lần sau chọn lại đúng GV/lớp đó vẫn mở được dialog)."""
    v = S.get(key)
    if v:
        S.open_dlg = (kind, v)
        S[key] = None


m1, m2 = st.columns(2)
m1.selectbox("a - Xem/In TKB GV", gv_list, index=None, placeholder="Chọn giáo viên...",
             key="menu_gv", on_change=_pick, args=("gv", "menu_gv"))
m2.selectbox("b - Xem/In TKB Lớp", classes, index=None, placeholder="Chọn lớp...",
             key="menu_lop", on_change=_pick, args=("lop", "menu_lop"))


@st.dialog("Xem / Tải TKB", width="large")
def dlg_week(kind, name):
    cur_df, cur_classes = S.df, S.classes
    page, total = week_html(kind, cur_df, cur_classes, name, school, apply_from, S.gmap)
    st.markdown(page, unsafe_allow_html=True)
    st.caption(f"Tổng cộng {total} tiết/tuần.")
    try:
        data = build_pdf("tkb_moi_gv" if kind == "gv" else "tkb_moi_lop", cur_df, cur_classes,
                         school, apply_from, only=name)
        st.download_button(f"⬇️ Tải PDF {name}", data, f"tkb_{name}.pdf",
                           mime="application/pdf", key=f"dl_{kind}_{name}")
    except Exception as e:
        st.caption(f"Không tạo được PDF: {e}")


if S.get("open_dlg"):
    _kind, _name = S.open_dlg
    S.open_dlg = None
    dlg_week(_kind, _name)

# ---- 9. download excel
st.subheader("6️⃣Tải về tệp TKB excel chung")
st.download_button("⬇️ Tải file TKB excel đã xếp", to_excel_bytes(S.df, classes),
                   "tkb_da_xep.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---- 10. PDF
st.subheader("7️⃣Tải các Tkb.pdf : Trường, Giáo viên, Lớp")
kind = st.selectbox("Chọn loại PDF", ["tkb_all", "tkb_moi_gv", "tkb_moi_lop"],
                    format_func=lambda k: {"tkb_all": "tkb_all.pdf - TKB toàn trường",
                                           "tkb_moi_gv": "tkb_moi_gv.pdf - TKB từng giáo viên",
                                           "tkb_moi_lop": "tkb_moi_lop.pdf - TKB từng lớp"}[k])
if st.button("Tạo PDF"):
    try:
        with st.spinner("Đang tạo PDF..."):
            S.pdf = (kind, build_pdf(kind, S.df, classes, school, apply_from))
    except Exception as e:
        st.error(f"Lỗi tạo PDF: {e}")
if S.get("pdf"):
    k, data = S.pdf
    st.download_button(f"⬇️ Tải {k}.pdf", data, f"{k}.pdf", mime="application/pdf")
