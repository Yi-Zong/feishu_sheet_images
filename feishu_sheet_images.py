# -*- coding: utf-8 -*-
"""
飞书电子表格 图片工具
  [1] 批量导出表格内图片
  [2] 把本地处理好的图片写回原单元格

Author: YX&Grok
依赖: pip install requests
运行: py feishu_sheet_images.py
"""
import os
import re
import json
import time
import requests

AUTHOR = "YX&Grok"
VERSION = "1.0"
BASE = "https://open.feishu.cn/open-apis"
TOKEN_HELP = "https://open.feishu.cn/api-explorer"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
VIEW_SUFFIXES = ("正", "侧", "背", "俯", "仰", "低")

try:
    import msvcrt
except ImportError:
    msvcrt = None

PAUSED = False
STOP = False


def banner(subtitle=""):
    print("=" * 60)
    print("  飞书电子表格  图片导出 / 回传")
    print("  Author: %s    v%s" % (AUTHOR, VERSION))
    if subtitle:
        print("  " + subtitle)
    print("=" * 60)


def hr():
    print("-" * 60)


def pause_exit():
    print()
    print("Author:", AUTHOR)
    try:
        input("按回车退出...")
    except Exception:
        pass


def token_help():
    print()
    print("如何获取 Token：")
    print("  1. 打开  " + TOKEN_HELP)
    print("  2. 右上角登录你的飞书账号")
    print("  3. Authorization 选 user_access_token（不要选 tenant）")
    print("  4. 点「获取 Token」，复制 u- 开头的那串")
    print()
    print("注意：")
    print("  · t- 开头是应用 Token，只能读、不能写回图片")
    print("  · u- 开头是你本人 Token，下载和回传都用这个")
    print("  · Token 大约 2 小时过期，过期重新获取即可")
    print("  · 应用需开通：查看/编辑电子表格，并发布版本")
    hr()


def ask_token():
    token_help()
    token = input("请粘贴 USER_ACCESS_TOKEN: ").strip().strip('"').strip("'")
    if not token:
        raise SystemExit("未输入 token")
    if token.startswith("t-"):
        print()
        print("警告：这是 tenant_access_token（t-）。")
        print("导出或许能用，回传一般会报 99991672 No permission。")
        print("建议改用 u- 开头的 user_access_token。")
        go = input("仍要继续？(Y/N，回车=N): ").strip().lower()
        if go not in ("y", "yes", "是"):
            raise SystemExit("已取消，请改用 u- Token 后重跑")
    elif not token.startswith("u-"):
        print("提示：常见 Token 以 u- 或 t- 开头，请确认复制完整。")
    return token


def ask_spreadsheet_token():
    while True:
        link = input("请粘贴飞书表格链接（或直接贴 token）: ").strip()
        tok = parse_spreadsheet_token(link)
        if tok:
            print("已解析表格 token:", tok)
            return tok
        print("无法识别，示例：https://xxx.feishu.cn/sheets/xxxxxxxx")


def ask_folder(prompt, default_name="feishu_images", must_exist=False):
    print()
    print(prompt)
    print("  请填完整路径，例如  D:\\user\\下载\\人设图")
    print("  当前脚本目录：", SCRIPT_DIR)
    print("  直接回车 = 脚本同目录下的 %s" % default_name)
    raw = input("路径: ").strip().strip('"').strip("'")
    if not raw:
        path = os.path.join(SCRIPT_DIR, default_name)
    else:
        path = os.path.abspath(os.path.expanduser(raw))
    if must_exist:
        if not os.path.isdir(path):
            raise SystemExit("目录不存在: " + path)
    else:
        os.makedirs(path, exist_ok=True)
    print("实际路径：", path)
    return path


def col_letter(n):
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def letter_to_col(s):
    s = s.strip().upper()
    n = 0
    for ch in s:
        if not ("A" <= ch <= "Z"):
            return None
        n = n * 26 + (ord(ch) - 64)
    return n


def safe_name(s, max_len=60):
    s = re.sub(r"[\r\n\t]+", " ", str(s))
    s = s.replace("|", " ")
    s = re.sub(r'[\\/:*?"<>]', "_", s)
    s = re.sub(r"\s+", " ", s).strip(" ._")
    s = s.replace(" ", "_")
    s = re.sub(r"_+", "_", s).strip("_")
    return s[:max_len]


def clip_note(text):
    text = (text or "").strip()
    if "，" in text:
        text = text.split("，", 1)[0].strip()
    return text


def cell_text(cell):
    if cell is None:
        return ""
    if isinstance(cell, (int, float)):
        return str(cell)
    if isinstance(cell, str):
        return cell.strip()
    if isinstance(cell, list):
        return "".join(cell_text(x) for x in cell).strip()
    if isinstance(cell, dict):
        for k in ("text", "name", "value"):
            if cell.get(k):
                return str(cell[k]).strip()
        if cell.get("type") == "embed-image":
            return ""
        return "".join(cell_text(v) for v in cell.values()).strip()
    return str(cell).strip()


def parse_choice(raw, n):
    raw = (raw or "").strip()
    if raw == "" or raw.lower() in ("all", "a", "全部"):
        return list(range(1, n + 1))
    out = []
    for part in re.split(r"[,\s，、]+", raw):
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            out.extend(range(int(a), int(b) + 1))
        else:
            out.append(int(part))
    out = [i for i in out if 1 <= i <= n]
    return out or list(range(1, n + 1))


def parse_cols(raw, max_col, default_all=True):
    raw = (raw or "").strip()
    if raw == "" or raw.lower() in ("all", "a", "全部"):
        if raw == "" and not default_all:
            return []
        return list(range(1, max_col + 1))
    out = []
    for part in re.split(r"[,\s，、]+", raw):
        if not part:
            continue
        if part.isdigit():
            out.append(int(part))
        else:
            c = letter_to_col(part)
            if c:
                out.append(c)
    out = [c for c in out if 1 <= c <= max_col]
    return sorted(set(out))


def parse_spreadsheet_token(text):
    text = (text or "").strip().strip('"').strip("'")
    m = re.search(r"/sheets/([A-Za-z0-9]+)", text)
    if m:
        return m.group(1)
    m = re.search(r"/spreadsheets/([A-Za-z0-9]+)", text)
    if m:
        return m.group(1)
    m = re.search(r"[?&](?:spreadsheetToken|spreadsheet_token)=([A-Za-z0-9]+)", text)
    if m:
        return m.group(1)
    if re.fullmatch(r"[A-Za-z0-9]{10,}", text):
        return text
    return None


def header_key(text):
    return re.sub(r"\s+", "", text or "")


def find_header_col(headers, keywords):
    for c in sorted(headers):
        t = header_key(headers.get(c))
        if any(k in t for k in keywords):
            return c
    return None


def note_cols_of(headers, name_cols):
    return [c for c in (name_cols or []) if "备注" in header_key(headers.get(c))]


def default_suffix_for_header(header):
    h = header_key(header)
    if not h:
        return ""
    if h in ("图", "图片", "配图", "截图", "参考图", "形象图"):
        return ""
    if h in VIEW_SUFFIXES:
        return h
    rules = (
        ("正面", "正"),
        ("侧面", "侧"),
        ("背面", "背"),
        ("反面", "背"),
        ("俯视", "俯"),
        ("全景", "俯"),
        ("仰视", "仰"),
        ("低角度", "低"),
        ("低角", "低"),
    )
    for key, short in rules:
        if key in h:
            return short
    return ""


def row_name_text(grid, r, name_cols, note_cols):
    if not name_cols:
        return ""
    note_cols = set(note_cols or [])
    parts = []
    for c in name_cols:
        t = cell_text(grid.get((r, c)))
        if c in note_cols:
            t = clip_note(t)
        t = safe_name(t)
        if t:
            parts.append(t)
    return "_".join(parts)


def make_filename(r, name_txt, col, suffix, extra_index):
    if name_txt:
        fn = f"{r}_{name_txt}"
    else:
        fn = f"{r}_{col}"
    if suffix:
        fn = f"{fn}_{suffix}"
    if extra_index > 1:
        fn = f"{fn}_{extra_index}"
    return fn


def check_pause():
    global PAUSED, STOP
    if not msvcrt:
        return

    def read_keys():
        global PAUSED, STOP
        while msvcrt.kbhit():
            ch = msvcrt.getwch()
            if ch in ("q", "Q"):
                STOP = True
                PAUSED = False
            elif ch in ("p", "P", " "):
                PAUSED = not PAUSED
                if PAUSED:
                    print("\n已暂停。按 P 继续，按 Q 结束")
                else:
                    print("继续下载...")

    read_keys()
    while PAUSED and not STOP:
        time.sleep(0.15)
        read_keys()


def api_get(token, path, **kwargs):
    r = requests.get(
        f"{BASE}{path}",
        headers={"Authorization": "Bearer " + token},
        timeout=120,
        **kwargs,
    )
    r.raise_for_status()
    return r


def api_get_json(token, path):
    r = requests.get(
        f"{BASE}{path}",
        headers={"Authorization": "Bearer " + token},
        timeout=120,
    )
    try:
        return r.status_code, r.json()
    except Exception:
        return r.status_code, {"raw": r.text[:800]}


def list_sheets(token, ss_token):
    code, meta = api_get_json(token, "/sheets/v3/spreadsheets/%s/sheets/query" % ss_token)
    if not isinstance(meta, dict) or meta.get("code") != 0:
        print("读取表格失败：", code, meta)
        raise SystemExit(1)
    return meta["data"]["sheets"]


def print_sheets(sheets, with_id=False):
    print()
    print("工作表：")
    for i, sh in enumerate(sheets, 1):
        gp = sh.get("grid_properties") or {}
        extra = "  id=%s" % sh.get("sheet_id") if with_id else ""
        print("  [%d] %s  %s行 x %s列%s" % (
            i,
            sh.get("title"),
            gp.get("row_count"),
            gp.get("column_count"),
            extra,
        ))


def open_folder(path):
    try:
        os.startfile(path)
    except Exception:
        pass


def ask_name_cols(headers, cols, hint):
    raw = input(hint).strip()
    name_cols = parse_cols(raw, cols, default_all=False)
    if not name_cols:
        print("未指定列，默认用 A 列")
        name_cols = [1]
    labels = " + ".join(
        "%s:%s" % (col_letter(c), headers.get(c) or "(空)") for c in name_cols
    )
    print("文件名主体 = 行号_%s" % labels)
    return name_cols


def ask_suffixes(img_cols, headers, confirm_first=True):
    suffixes = {}
    print()
    print("图片列后缀（只认：图=不加，以及 正/侧/背/俯/仰/低）：")
    for c in img_cols:
        h = headers.get(c) or ""
        suffixes[c] = default_suffix_for_header(h)
        print("  %s [%s]  →  %s" % (
            col_letter(c), h or "(空)", suffixes[c] or "(不加)",
        ))
    if confirm_first:
        raw = input("回车确认，输入 N 手动改: ").strip().lower()
        if raw not in ("n", "no", "改", "手动"):
            return suffixes
    print("  回车=推荐值，输入 - =不加，也可直接填 正/侧/背/俯/仰/低")
    for c in img_cols:
        h = headers.get(c) or ""
        rec = suffixes[c]
        rec_show = rec if rec else "(不加)"
        raw = input("  列 %s [%s] 推荐 %s: " % (
            col_letter(c), h or "(空)", rec_show,
        )).strip()
        if raw == "":
            continue
        if raw in ("-", "无", "n", "none", "不加", "图"):
            suffixes[c] = ""
        else:
            suffixes[c] = raw if raw in VIEW_SUFFIXES else safe_name(raw, 8)
    return suffixes


def print_name_preview(name_cols, suffixes, headers):
    body = "_".join(
        "{%s}" % (headers.get(c) or col_letter(c)) for c in (name_cols or [])
    )
    if body:
        print("文件名格式：行号_%s[_正/_侧/_背/_俯/_仰/_低]" % body)
    else:
        print("文件名格式：行号_列号")
    if suffixes:
        print("  各图片列示例：")
        for c, sfx in suffixes.items():
            h = headers.get(c) or col_letter(c)
            if body and sfx:
                print("    %s %s  →  2_%s_%s.png" % (col_letter(c), h, body, sfx))
            elif body:
                print("    %s %s  →  2_%s.png" % (col_letter(c), h, body))
            else:
                print("    %s %s  →  2_%s.png" % (col_letter(c), h, c))


def collect_images(obj, acc):
    if isinstance(obj, dict):
        if obj.get("type") == "embed-image" and obj.get("fileToken"):
            acc.append(obj["fileToken"])
        else:
            for v in obj.values():
                collect_images(v, acc)
    elif isinstance(obj, list):
        for v in obj:
            collect_images(v, acc)


def unique_path(folder, filename, used_names):
    path = os.path.join(folder, filename)
    base, ext = os.path.splitext(filename)
    i = 2
    while path in used_names or os.path.exists(path):
        path = os.path.join(folder, "%s_%d%s" % (base, i, ext))
        i += 1
    used_names.add(path)
    return path


def run_export():
    global PAUSED, STOP, n_saved
    PAUSED = False
    STOP = False
    n_saved = 0

    banner("功能：从飞书表格批量下载图片")
    print("导出说明：")
    print("  · 先选工作表，再选图片所在列")
    print("  · 命名可选：行_列 / 自定义列 / 人物表 / 场景表")
    print("  · 人物：行号_角色_备注（备注只取第一个「，」前）")
    print("  · 场景：行号_场地名称 或 行号_场地名称_正/侧/背…")
    print("  · 下载中按 P 暂停/继续，按 Q 结束")
    print("  · 若准备回传，建议命名选 [1] 行_列，位置最准")
    hr()

    token = ask_token()
    ss_token = ask_spreadsheet_token()
    out_dir = ask_folder("下载保存目录：")

    def read_range(sid, a1):
        return api_get(
            token,
            "/sheets/v2/spreadsheets/%s/values/%s!%s" % (ss_token, sid, a1),
        ).json()

    def pull_values(sid, cols, r1, r2):
        if not cols:
            return {}
        cols = sorted(set(cols))
        ranges = []
        start = prev = cols[0]
        for c in cols[1:]:
            if c == prev + 1:
                prev = c
            else:
                ranges.append((start, prev))
                start = prev = c
        ranges.append((start, prev))
        grid = {}

        def load(c1, c2, a, b):
            a1 = "%s%d:%s%d" % (col_letter(c1), a, col_letter(c2), b)
            data = read_range(sid, a1)
            if data.get("code") == 90221 and (b > a or c2 > c1):
                if b > a:
                    mid = (a + b) // 2
                    load(c1, c2, a, mid)
                    load(c1, c2, mid + 1, b)
                    return
                if c2 > c1:
                    mid = (c1 + c2) // 2
                    load(c1, mid, a, b)
                    load(mid + 1, c2, a, b)
                    return
                print("  skip too large", a1)
                return
            if data.get("code") != 0:
                print("  read fail", a1, data.get("msg") or data)
                return
            values = (data.get("data") or {}).get("valueRange", {}).get("values") or []
            for i, row in enumerate(values):
                for j, cell in enumerate(row or []):
                    grid[(a + i, c1 + j)] = cell

        for c1, c2 in ranges:
            load(c1, c2, r1, r2)
        return grid

    def download_token(ft, path, used_names):
        global n_saved
        resp = api_get(token, "/drive/v1/medias/%s/download" % ft)
        ctype = resp.headers.get("Content-Type", "")
        if not path.lower().endswith((".png", ".jpg", ".jpeg", ".gif", ".webp")):
            ext = ".jpg" if ("jpeg" in ctype or "jpg" in ctype) else (
                ".gif" if "gif" in ctype else (".webp" if "webp" in ctype else ".png")
            )
            folder, name = os.path.split(path + ext)
            path = unique_path(folder, name, used_names)
        with open(path, "wb") as f:
            f.write(resp.content)
        n_saved += 1
        print("  saved", path, len(resp.content))
        time.sleep(0.25)

    sheets = list_sheets(token, ss_token)
    print_sheets(sheets)
    sel = parse_choice(
        input("\n选择要导出的表（如 1,3 或 全部，回车=全部）: "),
        len(sheets),
    )
    chosen = [sheets[i - 1] for i in sel]

    jobs = []
    for sh in chosen:
        sid = sh["sheet_id"]
        title = sh.get("title") or sid
        gp = sh.get("grid_properties") or {}
        rows = int(gp.get("row_count") or 100)
        cols = int(gp.get("column_count") or 10)
        print()
        print("=" * 60)
        print("  %s    %d行 x %d列" % (title, rows, cols))
        print("=" * 60)
        print("读取表头...")
        header_grid = pull_values(sid, list(range(1, cols + 1)), 1, 1)
        headers = {}
        for c in range(1, cols + 1):
            t = cell_text(header_grid.get((1, c)))
            headers[c] = t
            print("  [%s=%d] %s" % (col_letter(c), c, t if t else "(空/图片)"))

        img_cols = parse_cols(
            input("选择图片列（如 H,I 或 I,J,L，回车=全部列）: "),
            cols,
        )
        if not img_cols:
            print("未选择图片列，跳过")
            continue
        print("已选图片列：", ", ".join(
            "%s:%s" % (col_letter(c), headers.get(c) or "(空)") for c in img_cols
        ))

        print()
        print("命名方式：")
        print("  [1] 默认：行_列号            例  3_8.png   ← 回传请选这个")
        print("  [2] 自定义：行号_所选列文字[_正/_侧/_背…]")
        print("  [3] 人物表：行号_角色_备注（备注只取第一个「，」前；无备注不加_）")
        print("  [4] 场景表：行号_场地名称 或 行号_场地名称_正/侧/背/俯/仰/低")
        mode = (input("选择 1 / 2 / 3 / 4（回车=1）: ").strip() or "1").lower()
        mode = {"人物": "3", "角色": "3", "场景": "4", "场地": "4", "自定义": "2"}.get(mode, mode)

        name_cols = []
        suffixes = {c: "" for c in img_cols}

        if mode == "3":
            role_c = find_header_col(headers, ("角色", "人物", "姓名", "名字"))
            note_c = find_header_col(headers, ("备注",))
            if not role_c:
                name_cols = ask_name_cols(
                    headers, cols,
                    "未找到「角色」列，请指定命名列（可多列，如 B,C）: ",
                )
            else:
                name_cols = [role_c]
                if note_c and note_c != role_c:
                    name_cols.append(note_c)
                print("人物表命名列：", " + ".join(
                    "%s:%s" % (col_letter(c), headers.get(c)) for c in name_cols
                ))
            suffixes = ask_suffixes(img_cols, headers)
        elif mode == "4":
            place_c = find_header_col(headers, ("场地名称", "场地名", "场景名称", "场景名", "场地"))
            if not place_c:
                name_cols = ask_name_cols(
                    headers, cols,
                    "未找到「场地名称」列，请指定命名列（如 C）: ",
                )
            else:
                name_cols = [place_c]
                print("场景表命名列：%s:%s" % (col_letter(place_c), headers.get(place_c)))
            suffixes = ask_suffixes(img_cols, headers)
        elif mode == "2":
            name_cols = ask_name_cols(
                headers, cols,
                "用作文件名的列（可多列，人物如 B,C，场景如 C）: ",
            )
            suffixes = ask_suffixes(img_cols, headers)
        else:
            name_cols = []
            suffixes = {c: "" for c in img_cols}
            print("文件名 = 行号_列号")

        note_cols = note_cols_of(headers, name_cols)
        print_name_preview(name_cols, suffixes, headers)
        jobs.append({
            "sid": sid, "title": title, "rows": rows,
            "img_cols": img_cols, "name_cols": name_cols,
            "note_cols": note_cols, "suffixes": suffixes,
        })

    seen_tokens = set()
    used_names = set()

    print()
    hr()
    print("开始下载")
    print("保存到：", out_dir)
    print("下载中可按 P 暂停/继续，按 Q 结束")
    hr()

    for job in jobs:
        if STOP:
            break
        sid, title = job["sid"], job["title"]
        rows, img_cols = job["rows"], job["img_cols"]
        name_cols, note_cols = job["name_cols"], job["note_cols"]
        suffixes = job["suffixes"]
        need_cols = sorted(set(img_cols + list(name_cols or [])))
        folder = os.path.join(out_dir, safe_name(title) or "sheet")
        os.makedirs(folder, exist_ok=True)
        print()
        print(">>", title)
        print("  子文件夹：", folder)
        step = 10
        for r1 in range(1, rows + 1, step):
            if STOP:
                break
            r2 = min(rows, r1 + step - 1)
            print("  rows %d-%d" % (r1, r2))
            grid = pull_values(sid, need_cols, r1, r2)
            for r in range(r1, r2 + 1):
                if STOP:
                    break
                name_txt = row_name_text(grid, r, name_cols, note_cols)
                for c in img_cols:
                    if STOP:
                        break
                    tokens = []
                    collect_images(grid.get((r, c)), tokens)
                    for k, ft in enumerate(tokens, 1):
                        check_pause()
                        if STOP:
                            break
                        if ft in seen_tokens:
                            continue
                        seen_tokens.add(ft)
                        fn = make_filename(r, name_txt, c, suffixes.get(c, ""), k)
                        path = unique_path(folder, fn, used_names)
                        try:
                            download_token(ft, path, used_names)
                        except Exception as e:
                            print("  fail", ft, e)

    abs_out = os.path.abspath(out_dir)
    print()
    print("%s，共 %d 张" % ("已中止" if STOP else "完成", n_saved))
    print("文件夹：", abs_out)
    open_folder(abs_out)
    print("已尝试打开文件夹")
    pause_exit()


def run_upload():
    banner("功能：把本地图片按文件名写回飞书表格")
    print("回传说明：")
    print("  · 必须用 u- 开头的 user_access_token，t- 会报 No permission")
    print("  · 推荐导出时用「行_列」命名，例如 3_8.png → 第3行 H 列")
    print("  · 会覆盖目标单元格原图，文件名不要改")
    print("  · 接口要 PNG/JPG/GIF 等，不支持 webp")
    print("  · 单张过大（十多 MB）可能失败，可先转高质量 JPG，像素仍是 4K")
    hr()

    token = ask_token()
    ss_token = ask_spreadsheet_token()
    img_dir = ask_folder("本地图片文件夹（完整路径）：", must_exist=True)

    headers = {
        "Authorization": "Bearer " + token,
        "Content-Type": "application/json; charset=utf-8",
    }

    sheets = list_sheets(token, ss_token)
    print_sheets(sheets, with_id=True)
    idx = int(input("写回哪个表（编号）: ").strip())
    sheet = sheets[idx - 1]
    sid = sheet["sheet_id"]
    print("目标表：", sheet.get("title"), sid)

    print()
    print("文件名如何对应单元格：")
    print("  [1] 文件名是 行_列   例 3_8.png → H3")
    print("  [2] 文件名只有行号，全部写入同一列  例 3_林悦.png → 指定列第3行")
    mode = input("选择 1 或 2（回车=1）: ").strip() or "1"

    fixed_col = None
    if mode == "2":
        c = input("写回哪一列（如 H 或 8）: ").strip()
        fixed_col = int(c) if c.isdigit() else letter_to_col(c)
        if not fixed_col:
            raise SystemExit("列无效")

    def parse_name(name):
        base = os.path.splitext(name)[0]
        m = re.match(r"^(\d+)_(\d+)(?:_\d+)?$", base)
        if m and mode == "1":
            return int(m.group(1)), int(m.group(2))
        m = re.match(r"^(\d+)_.+", base)
        if m and mode == "2":
            return int(m.group(1)), fixed_col
        return None, None

    exts = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp"}
    files = [f for f in os.listdir(img_dir) if os.path.splitext(f)[1].lower() in exts]
    files.sort()
    print("找到 %d 张图" % len(files))

    ok = fail = 0
    for name in files:
        row, col = parse_name(name)
        if not row or not col:
            print("跳过（文件名无法解析）:", name)
            fail += 1
            continue

        path = os.path.join(img_dir, name)
        raw = open(path, "rb").read()
        size_mb = len(raw) / 1024 / 1024
        ext = os.path.splitext(name)[1].lower()
        api_name = name
        if ext == ".webp":
            api_name = os.path.splitext(name)[0] + ".png"
            print("注意：webp 接口可能不认，已把名称改成", api_name, "若仍失败请先转成 jpg/png")

        rng = "%s!%s%d:%s%d" % (sid, col_letter(col), row, col_letter(col), row)
        print("上传 %s  %.2fMB  → %s" % (name, size_mb, rng))
        if size_mb > 8:
            print("  提示：体积较大，上传会较慢，失败可转高质量 JPG 后再试")

        body = {
            "range": rng,
            "image": list(raw),
            "name": api_name,
        }
        try:
            r = requests.post(
                BASE + "/sheets/v2/spreadsheets/%s/values_image" % ss_token,
                headers=headers,
                data=json.dumps(body),
                timeout=180,
            )
            try:
                data = r.json()
            except Exception:
                data = {"raw": r.text[:800]}
            if r.status_code == 200 and isinstance(data, dict) and data.get("code") == 0:
                print("  已写回", col_letter(col) + str(row))
                ok += 1
            else:
                print("  失败", r.status_code, data)
                fail += 1
            time.sleep(0.4)
        except Exception as e:
            print("  异常", e)
            fail += 1

    print()
    print("完成：成功 %d，失败/跳过 %d" % (ok, fail))
    pause_exit()


def main():
    banner()
    print("  [1] 导出  —— 从飞书表格下载图片到本地")
    print("  [2] 回传  —— 把本地图片写回飞书表格")
    print("  [0] 退出")
    print()
    print("使用前请先：pip install requests")
    hr()
    choice = input("请选择（1/2/0，回车=1）: ").strip() or "1"
    if choice in ("2", "回传", "上传", "导入"):
        run_upload()
    elif choice in ("0", "q", "Q", "退出"):
        print("已退出")
    else:
        run_export()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n已中断")
    except SystemExit:
        raise
    except Exception as e:
        print("出错：", e)
        pause_exit()