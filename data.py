# -*- coding: utf-8 -*-
"""从 Excel 读取队伍汇总（按「队伍名称」分组，不暴露个人体重）。"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

XLSX_FILENAME = "团队健康大比拼0408.xlsx"


def resolve_xlsx_path(root: str | Path) -> Path:
    """统一使用「团队健康大比拼0408.xlsx」。"""
    return Path(root) / XLSX_FILENAME


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


def _first_weight_column(df: pd.DataFrame) -> str | None:
    """第一次称体重列；无则返回 None。"""
    preferred = [
        "第一次称体重KG",
        "第一次称体重kg",
        "第一次称体重",
        "第一次体重KG（4/8日）",
        "第一次体重KG（4/8日)",
        "第一次体重",
        "体重KG（4/8日)",
        "体重KG（4/8日）",
    ]
    for name in preferred:
        if name in df.columns:
            return name
    for c in df.columns:
        s = str(c).strip()
        if ("4/8" in s or "4／8" in s) and "体重" in s:
            return c
    for c in df.columns:
        s = str(c).strip()
        if "第一次" in s and ("体重" in s or "KG" in s.upper() or "kg" in s.lower()):
            return c
    for c in df.columns:
        if "第一次" in str(c):
            return c
    return None


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
    first_col = _first_weight_column(df)
    captain_col = "队长or队员" if "队长or队员" in df.columns else None

    teams: list[dict] = []
    for raw_key, g in df.groupby("队伍名称", sort=False):
        g = g.copy()
        g["_init"] = pd.to_numeric(g[weight_col], errors="coerce")
        if first_col:
            g["_first"] = pd.to_numeric(g[first_col], errors="coerce")
        else:
            g["_first"] = pd.Series([pd.NA] * len(g), index=g.index, dtype="float64")

        names = []
        member_week1: list[dict] = []
        for _, row in g.iterrows():
            name = str(row["姓名"]).strip()
            cap = False
            if captain_col:
                v = row.get(captain_col)
                cap = (not pd.isna(v)) and str(v).strip() == "队长"
            names.append({"name": name, "captain": cap})
            pct = _pct_drop(row["_init"], row["_first"]) if first_col else None
            member_week1.append(
                {
                    "name": name,
                    "captain": cap,
                    "drop_pct": round(pct, 2) if pct is not None else None,
                }
            )

        captains = [m for m in names if m["captain"]]
        others = [m for m in names if not m["captain"]]
        names = captains + others

        cap_order = {m["name"]: (0 if m["captain"] else 1) for m in member_week1}
        member_week1.sort(key=lambda x: (cap_order.get(x["name"], 1), x["name"]))

        total_w = float(g["_init"].sum(skipna=True))
        both = g["_init"].notna() & g["_first"].notna()
        init_sum_both = float(g.loc[both, "_init"].sum())
        first_sum_both = float(g.loc[both, "_first"].sum())
        weighed = int(both.sum())
        team_drop_pct: float | None = None
        if first_col and init_sum_both > 0:
            team_drop_pct = (init_sum_both - first_sum_both) / init_sum_both * 100.0

        teams.append(
            {
                "name": _team_label(raw_key),
                "sort_key": raw_key,
                "count": len(g),
                "total_kg": round(total_w, 2),
                "first_total_kg": round(first_sum_both, 2) if first_col and weighed else None,
                "baseline_for_rate_kg": round(init_sum_both, 2) if first_col and weighed else None,
                "weighed_count": weighed if first_col else 0,
                "team_drop_pct": round(team_drop_pct, 2) if team_drop_pct is not None else None,
                "members": names,
                "member_week1": member_week1,
                "has_first_weigh": bool(first_col),
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

    rates = [t["team_drop_pct"] for t in teams if t.get("team_drop_pct") is not None]
    best = max(rates) if rates else None
    for t in teams:
        p = t.get("team_drop_pct")
        t["is_first_place"] = bool(best is not None and p is not None and p == best)

    return teams
