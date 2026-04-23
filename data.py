# -*- coding: utf-8 -*-
"""从 Excel 读取队伍汇总（按「队伍名称」分组，不暴露个人体重）。"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

XLSX_FILENAME = "团队健康大比拼0423.xlsx"


def resolve_xlsx_path(root: str | Path) -> Path:
    """优先使用最新周报表，回退到旧文件名。"""
    root_path = Path(root)
    latest = root_path / XLSX_FILENAME
    if latest.is_file():
        return latest
    fallback = root_path / "团队健康大比拼0408.xlsx"
    return fallback


def _team_label(value) -> str:
    if pd.isna(value):
        return "（未命名）"
    try:
        f = float(value)
        if f == int(f):
            return str(int(f))
    except (TypeError, ValueError):
        pass
    return str(value).strip() or "（未命名）"


def _weight_column(df: pd.DataFrame) -> str:
    if "初始体重KG（4/1日)" in df.columns:
        return "初始体重KG（4/1日)"
    for c in df.columns:
        s = str(c)
        if "体重" in s and "初始" in s:
            return c
    for c in df.columns:
        s = str(c)
        if "体重" in s and "第一次" not in s and not _is_weekly_weigh_column(s):
            return c
    for c in df.columns:
        if "体重" in str(c):
            return c
    raise ValueError("未找到体重列")


def _is_weekly_weigh_column(name: str) -> bool:
    """排除各周称重列，避免误当作基线。"""
    s = str(name)
    for token in ("4/8", "4/15", "4/22", "4/28", "4／8", "4／15"):
        if token in s and "体重" in s:
            return True
    return False


def _week_weight_column(df: pd.DataFrame, day_token: str) -> str | None:
    token_full = f"4/{day_token}"
    token_full_cn = f"4／{day_token}"
    for c in df.columns:
        s = str(c).strip()
        if "体重" in s and (token_full in s or token_full_cn in s):
            return c
    return None


def _week_definitions(df: pd.DataFrame) -> list[dict]:
    return [
        {"week_no": 1, "label": "第一周", "col": _week_weight_column(df, "8")},
        {"week_no": 2, "label": "第二周", "col": _week_weight_column(df, "15")},
        {"week_no": 3, "label": "第三周", "col": _week_weight_column(df, "22")},
    ]


def _pct_drop(initial: float, first: float) -> float | None:
    if initial is None or first is None:
        return None
    try:
        if pd.isna(initial) or pd.isna(first):
            return None
        init = float(initial)
        fst = float(first)
    except (TypeError, ValueError):
        return None
    if init <= 0:
        return None
    return (init - fst) / init * 100.0


def _fmt_pct(value: float | None) -> float | None:
    return round(value, 2) if value is not None else None


def load_teams(xlsx_path: str | Path) -> list[dict]:
    path = Path(xlsx_path)
    if not path.is_file():
        raise FileNotFoundError(
            f"未找到数据表：{path.name}，请将文件放在目录：{path.parent}"
        )
    df = pd.read_excel(path)
    if "队伍名称" not in df.columns:
        raise ValueError("表中缺少「队伍名称」列")
    if "姓名" not in df.columns:
        raise ValueError("表中缺少「姓名」列")

    weight_col = _weight_column(df)
    week_defs = _week_definitions(df)
    captain_col = "队长or队员" if "队长or队员" in df.columns else None

    teams: list[dict] = []
    for raw_key, g in df.groupby("队伍名称", sort=False):
        g = g.copy()
        g["_init"] = pd.to_numeric(g[weight_col], errors="coerce")
        week_cols: dict[int, str] = {}
        for wk in week_defs:
            col = wk["col"]
            if col:
                cast_col = f"_w{wk['week_no']}"
                g[cast_col] = pd.to_numeric(g[col], errors="coerce")
                week_cols[wk["week_no"]] = cast_col

        names = []
        members_by_week: dict[int, list[dict]] = {1: [], 2: [], 3: []}
        for _, row in g.iterrows():
            name = str(row["姓名"]).strip()
            cap = False
            if captain_col:
                v = row.get(captain_col)
                cap = (not pd.isna(v)) and str(v).strip() == "队长"
            names.append({"name": name, "captain": cap})
            prev_pct: float | None = None
            for wk in week_defs:
                week_no = wk["week_no"]
                cast_col = week_cols.get(week_no)
                current = row[cast_col] if cast_col else None
                pct = _pct_drop(row["_init"], current) if cast_col else None
                delta = None if prev_pct is None or pct is None else pct - prev_pct
                members_by_week[week_no].append(
                    {
                        "name": name,
                        "captain": cap,
                        "drop_pct": _fmt_pct(pct),
                        "delta_from_prev": _fmt_pct(delta),
                    }
                )
                prev_pct = pct

        captains = [m for m in names if m["captain"]]
        others = [m for m in names if not m["captain"]]
        names = captains + others

        cap_order = {m["name"]: (0 if m["captain"] else 1) for m in names}
        for w in (1, 2, 3):
            members_by_week[w].sort(key=lambda x: (cap_order.get(x["name"], 1), x["name"]))

        total_w = float(g["_init"].sum(skipna=True))
        week_reports: list[dict] = []
        has_any_week = False
        prev_team_pct: float | None = None
        for wk in week_defs:
            week_no = wk["week_no"]
            cast_col = week_cols.get(week_no)
            has_week = cast_col is not None
            both = g["_init"].notna() & g[cast_col].notna() if has_week else pd.Series([False] * len(g), index=g.index)
            init_sum_both = float(g.loc[both, "_init"].sum()) if has_week else 0.0
            curr_sum_both = float(g.loc[both, cast_col].sum()) if has_week else 0.0
            weighed = int(both.sum()) if has_week else 0
            team_drop_pct = ((init_sum_both - curr_sum_both) / init_sum_both * 100.0) if has_week and init_sum_both > 0 else None
            team_delta = None if prev_team_pct is None or team_drop_pct is None else team_drop_pct - prev_team_pct
            week_reports.append(
                {
                    "week_no": week_no,
                    "label": wk["label"],
                    "has_weigh": has_week,
                    "weighed_count": weighed,
                    "week_total_kg": round(curr_sum_both, 2) if has_week and weighed else None,
                    "team_drop_pct": _fmt_pct(team_drop_pct),
                    "team_delta_from_prev": _fmt_pct(team_delta),
                    "members": members_by_week[week_no],
                }
            )
            if has_week and team_drop_pct is not None:
                has_any_week = True
                prev_team_pct = team_drop_pct

        teams.append(
            {
                "name": _team_label(raw_key),
                "sort_key": raw_key,
                "count": len(g),
                "total_kg": round(total_w, 2),
                "members": names,
                "week_reports": week_reports,
                "has_any_week": has_any_week,
            }
        )

    def sort_key(t: dict):
        sk = t["sort_key"]
        try:
            return (0, float(sk))
        except (TypeError, ValueError):
            return (1, str(sk))

    teams.sort(key=sort_key)
    for t in teams:
        t.pop("sort_key", None)

    def _week_rate(team: dict, week_no: int) -> float | None:
        for w in team.get("week_reports", []):
            if w.get("week_no") == week_no:
                return w.get("team_drop_pct")
        return None

    rates = [_week_rate(t, 3) for t in teams]
    rates = [x for x in rates if x is not None]
    best = max(rates) if rates else None
    for t in teams:
        p = _week_rate(t, 3)
        t["is_first_place"] = bool(best is not None and p is not None and p == best)

    return teams


def build_week_rows(teams: list[dict]) -> list[dict]:
    """按周生成卡片行，并按当周下降率从高到低排名。"""
    week_rows: list[dict] = []
    week_defs = [
        (1, "第一周"),
        (2, "第二周"),
        (3, "第三周"),
    ]
    for week_no, label in week_defs:
        cards: list[dict] = []
        for t in teams:
            wr = next((w for w in t.get("week_reports", []) if w.get("week_no") == week_no), None)
            if not wr or not wr.get("has_weigh"):
                continue
            cards.append(
                {
                    "team_name": t["name"],
                    "count": t["count"],
                    "total_kg": t["total_kg"],
                    "members": t["members"],
                    "week_report": wr,
                }
            )

        def _sort_key(card: dict):
            pct = card["week_report"].get("team_drop_pct")
            return (pct is None, -(pct or -10**9), card["team_name"])

        cards.sort(key=_sort_key)
        for idx, c in enumerate(cards, start=1):
            c["rank"] = idx
            c["is_week_first"] = idx == 1

        week_rows.append(
            {
                "week_no": week_no,
                "label": label,
                "cards": cards,
            }
        )
    return week_rows


def build_trend_series(teams: list[dict]) -> list[dict]:
    """构建第一至第三周各队下降率折线图数据。"""
    series: list[dict] = []
    for t in teams:
        week1 = next((w for w in t.get("week_reports", []) if w.get("week_no") == 1), None)
        week2 = next((w for w in t.get("week_reports", []) if w.get("week_no") == 2), None)
        week3 = next((w for w in t.get("week_reports", []) if w.get("week_no") == 3), None)
        y1 = week1.get("team_drop_pct") if week1 and week1.get("has_weigh") else None
        y2 = week2.get("team_drop_pct") if week2 and week2.get("has_weigh") else None
        y3 = week3.get("team_drop_pct") if week3 and week3.get("has_weigh") else None
        if y1 is None and y2 is None and y3 is None:
            continue
        series.append(
            {
                "name": t["name"],
                "points": [y1, y2, y3],
            }
        )
    return series
