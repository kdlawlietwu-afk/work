# -*- coding: utf-8 -*-
"""团队健康大比拼 — Flask 展示（按队伍名称汇总）。"""
from pathlib import Path

from flask import Flask, render_template

from data import build_trend_series, build_week_rows, load_teams, resolve_xlsx_path

ROOT = Path(__file__).resolve().parent
XLSX = resolve_xlsx_path(ROOT)

app = Flask(__name__)


@app.route("/")
def index():
    teams = load_teams(XLSX)
    week_rows = build_week_rows(teams)
    trend_series = build_trend_series(teams)
    return render_template("index.html", teams=teams, week_rows=week_rows, trend_series=trend_series)


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
