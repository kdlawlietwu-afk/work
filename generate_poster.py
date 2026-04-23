# -*- coding: utf-8 -*-
"""用与 Flask 相同的模板生成静态 健康大比拼海报.html（可选）。"""
from pathlib import Path

from flask import Flask, render_template

from data import build_trend_series, build_week_rows, load_teams, resolve_xlsx_path

ROOT = Path(__file__).resolve().parent
XLSX = resolve_xlsx_path(ROOT)
OUT = ROOT / "健康大比拼海报.html"

app = Flask(__name__, template_folder=str(ROOT / "templates"))


def main():
    teams = load_teams(XLSX)
    week_rows = build_week_rows(teams)
    trend_series = build_trend_series(teams)
    with app.app_context():
        html = render_template("index.html", teams=teams, week_rows=week_rows, trend_series=trend_series)
    OUT.write_text(html, encoding="utf-8")
    print(f"已写入: {OUT}")


if __name__ == "__main__":
    main()
