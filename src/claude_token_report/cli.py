#!/usr/bin/env python3
"""Claude Code のログ (~/.claude/projects/**/*.jsonl) から日別トークン消費量と料金を集計する。"""
import argparse
import csv
import io
import json
import sys
from collections import defaultdict
from datetime import datetime, date
from pathlib import Path

from . import pricing
from . import xlsx_writer as xw

FIELDS = ["input", "output", "cache_read", "cache_write", "total"]


def parse_ts(s):
    return datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone()  # ローカルTZ


def iter_records(root):
    """重複排除済みの (date, model, tokens...) を返す。
    ストリーミングで同一 message.id が複数行出るため、最後の行を採用する。"""
    latest = {}
    for path in Path(root).expanduser().rglob("*.jsonl"):
        try:
            fh = path.open(encoding="utf-8", errors="replace")
        except OSError:
            continue
        with fh:
            for n, line in enumerate(fh):
                if '"usage"' not in line:
                    continue
                try:
                    d = json.loads(line)
                    if d.get("type") != "assistant":
                        continue
                    msg = d["message"]
                    u = msg["usage"]
                    model = msg.get("model") or ""
                    if not model or model.startswith("<"):
                        continue
                    day = parse_ts(d["timestamp"]).date()
                    cc = u.get("cache_creation") or {}
                    w5 = cc.get("ephemeral_5m_input_tokens")
                    w1 = cc.get("ephemeral_1h_input_tokens")
                    wtot = u.get("cache_creation_input_tokens") or 0
                    if w5 is None and w1 is None:
                        w5, w1 = wtot, 0  # 内訳なしは5分扱い
                    rec = (day, model, u.get("input_tokens") or 0, u.get("output_tokens") or 0,
                           u.get("cache_read_input_tokens") or 0, w5 or 0, w1 or 0)
                except (ValueError, KeyError, TypeError, AttributeError):
                    continue
                mid, rid = msg.get("id"), d.get("requestId")
                key = (mid, rid) if mid or rid else (str(path), n)
                latest[key] = rec
    return latest.values()


def aggregate(records, since=None, until=None, by_model=False):
    rows = defaultdict(lambda: dict.fromkeys(FIELDS, 0) | {"usd": 0.0, "unpriced": 0})
    unknown = set()
    for day, model, i, o, r, w5, w1 in records:
        if since and day < since or until and day > until:
            continue
        key = (day, model) if by_model else (day, None)
        row = rows[key]
        row["input"] += i
        row["output"] += o
        row["cache_read"] += r
        row["cache_write"] += w5 + w1
        row["total"] += i + o + r + w5 + w1
        c = pricing.cost_usd(model, i, o, r, w5, w1)
        if c is None:
            unknown.add(model)
            row["unpriced"] += i + o + r + w5 + w1
        else:
            row["usd"] += c
    return dict(sorted(rows.items(), key=lambda kv: (kv[0][0], kv[0][1] or ""))), unknown


def _scaled(n, base, suffix):
    if n < base:
        return f"{n:,}"
    v = round(n / base, 1)
    return f"{v:,.1f}".rstrip("0").rstrip(".") + suffix


def aggregate_detail(records, since=None, until=None):
    """Excel用。(日付, 単価ラベル) ごとに Input/Output/CacheRead/Write5m/Write1h を合算。
    単価未対応モデルは unknown に集め、明細には含めない。"""
    rows = defaultdict(lambda: [0, 0, 0, 0, 0])
    unknown = set()
    for day, model, i, o, r, w5, w1 in records:
        if since and day < since or until and day > until:
            continue
        label, _ = pricing.resolve(model, i + r + w5 + w1)
        if label is None:
            unknown.add(model)
            continue
        acc = rows[(day, label)]
        for k, v in enumerate((i, o, r, w5, w1)):
            acc[k] += v
    return dict(sorted(rows.items())), unknown


def write_excel(path, detail, rate):
    """シート: 日別集計 / 明細 (モデル別) / 単価表。料金は 単価 x トークン数 / 100万 の数式。"""
    prices = dict(pricing.price_table())
    n = len(detail)

    # --- 単価表 (A:F) ---
    price_rows = [[xw.Cell("モデル(単価ラベル)", "bold")] + [xw.Cell(h, "bold") for h in (
        "Input", "Cache Write 5m", "Cache Write 1h", "Cache Read", "Output")],
        [xw.Cell("単位: USD / MTok。ここを書き換えると料金が再計算されます。")]]
    for label, p in pricing.price_table():
        price_rows.append([label] + [xw.Cell(round(v, 6), "price") for v in p])
    price_rows = [price_rows[0]] + price_rows[2:] + [[], price_rows[1]]
    last_price = 1 + len(pricing.price_table())

    # --- 明細 ---
    head = ["日付", "モデル", "Input", "Output", "Cache Read", "Cache Write 5m", "Cache Write 1h",
            "単価 Input", "単価 CW 5m", "単価 CW 1h", "単価 Cache Read", "単価 Output", "料金(USD)"]
    det_rows = [[xw.Cell(h, "bold") for h in head]]
    cost_by_day = defaultdict(float)
    for r, ((day, label), (i, o, rd, w5, w1)) in enumerate(detail.items(), 2):
        p = prices[label]
        cost = (i * p[0] + w5 * p[1] + w1 * p[2] + rd * p[3] + o * p[4]) / 1_000_000
        cost_by_day[day] += cost
        look = lambda col: xw.Cell(round(p[col - 2], 6), "price",
                                   f"VLOOKUP($B{r},単価表!$A$2:$F${last_price},{col},FALSE)")
        det_rows.append([
            xw.Cell(day, "date"), label,
            xw.Cell(i, "int"), xw.Cell(o, "int"), xw.Cell(rd, "int"), xw.Cell(w5, "int"), xw.Cell(w1, "int"),
            look(2), look(3), look(4), look(5), look(6),
            xw.Cell(cost, "usd", f"(C{r}*H{r}+F{r}*I{r}+G{r}*J{r}+E{r}*K{r}+D{r}*L{r})/1000000"),
        ])
    last = n + 1

    # --- 日別集計 ---
    days = sorted({d for d, _ in detail})
    day_head = ["日付", "Input", "Output", "Cache Read", "Cache Write", "総処理量", "料金(USD)", "料金(JPY)"]
    day_rows = [[xw.Cell("円レート (JPY/USD)", "bold"), xw.Cell(rate, "rate")], [],
                [xw.Cell(h, "bold") for h in day_head]]
    sums = [0] * 5
    for k, day in enumerate(days):
        r = 4 + k
        agg = [0] * 5
        for (d, _), v in detail.items():
            if d == day:
                agg = [a + b for a, b in zip(agg, v)]
        i, o, rd, w5, w1 = agg
        for q in range(5):
            sums[q] += agg[q]

        def sumifs(col):
            return f"SUMIFS(明細!${col}$2:${col}${last},明細!$A$2:$A${last},$A{r})"
        usd = cost_by_day[day]
        day_rows.append([
            xw.Cell(day, "date"),
            xw.Cell(i, "int", sumifs("C")), xw.Cell(o, "int", sumifs("D")), xw.Cell(rd, "int", sumifs("E")),
            xw.Cell(w5 + w1, "int", f"{sumifs('F')}+{sumifs('G')}"),
            xw.Cell(i + o + rd + w5 + w1, "int", f"SUM(B{r}:E{r})"),
            xw.Cell(usd, "usd", sumifs("M")),
            xw.Cell(usd * rate, "jpy", f"G{r}*$B$1"),
        ])
    f, l, t = 4, 3 + len(days), 4 + len(days)
    total_usd = sum(cost_by_day.values())
    day_rows.append([
        xw.Cell("合計", "bold"),
        xw.Cell(sums[0], "boldint", f"SUM(B{f}:B{l})"), xw.Cell(sums[1], "boldint", f"SUM(C{f}:C{l})"),
        xw.Cell(sums[2], "boldint", f"SUM(D{f}:D{l})"), xw.Cell(sums[3] + sums[4], "boldint", f"SUM(E{f}:E{l})"),
        xw.Cell(sum(sums), "boldint", f"SUM(F{f}:F{l})"),
        xw.Cell(total_usd, "boldusd", f"SUM(G{f}:G{l})"),
        xw.Cell(total_usd * rate, "boldjpy", f"SUM(H{f}:H{l})"),
    ])
    day_rows += [[], [xw.Cell("※ API従量課金換算の目安。Batch割引・Fast mode・データレジデンシー・サーバーツール課金は含みません。")]]

    xw.write_xlsx(path, [
        ("日別集計", day_rows, [18, 14, 14, 14, 14, 14, 14, 14]),
        ("明細", det_rows, [12, 28, 14, 14, 14, 16, 16, 12, 12, 12, 14, 12, 14]),
        ("単価表", price_rows, [28, 12, 16, 16, 12, 12]),
    ])


def fmt_man(n):
    return _scaled(n, 10_000, "万")


def fmt_sen(n):
    return _scaled(n, 1_000, "千")


UNITS = {"raw": lambda n: f"{n:,}", "man": fmt_man, "sen": fmt_sen}


def width(s):
    return sum(2 if ord(c) > 0x7F else 1 for c in s)


def pad(s, w, right=True):
    sp = " " * (w - width(s))
    return sp + s if right else s + sp


def fmt_day(d, multi_year):
    return d.strftime("%Y/%m/%d") if multi_year else d.strftime("%m/%d")


def render_table(rows, rate, by_model, unit="man"):
    multi_year = len({k[0].year for k in rows}) > 1
    head = ["日付"] + (["モデル"] if by_model else []) + \
        ["Input", "Output", "Cache Read", "Cache Write", "総処理量", "USD", "JPY"]
    body = []
    tot = dict.fromkeys(FIELDS, 0) | {"usd": 0.0}
    for (day, model), r in rows.items():
        for f in tot:
            tot[f] += r[f]
        body.append(line(fmt_day(day, multi_year), model, r, rate, by_model, unit))
    body.append(line("合計", "" if by_model else None, tot, rate, by_model, unit))
    table = [head] + body
    ws = [max(width(r[c]) for r in table) for c in range(len(head))]
    lines = ["  ".join(pad(c, ws[i], right=i > (1 if by_model else 0) or (i == 1 and not by_model))
                        for i, c in enumerate(r)) for r in table]
    return "\n".join(lines)


def line(label, model, r, rate, by_model, unit="man"):
    cells = [label] + ([model] if by_model else []) + [UNITS[unit](r[f]) for f in FIELDS]
    cells += [f"${r['usd']:,.2f}", f"¥{r['usd'] * rate:,.0f}"]
    return cells


def render_raw(rows, rate, fmt, by_model):
    out = []
    for (day, model), r in rows.items():
        o = {"date": day.isoformat()}
        if by_model:
            o["model"] = model
        o.update({f: r[f] for f in FIELDS})
        o["usd"] = round(r["usd"], 6)
        o["jpy"] = round(r["usd"] * rate, 2)
        out.append(o)
    if fmt == "json":
        return json.dumps(out, ensure_ascii=False, indent=2)
    buf = io.StringIO()
    if out:
        w = csv.DictWriter(buf, fieldnames=list(out[0]))
        w.writeheader()
        w.writerows(out)
    return buf.getvalue().rstrip("\n")


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--project", help="特定プロジェクトのログディレクトリ (~/.claude/projects 配下の名前またはパス)")
    ap.add_argument("--since", type=date.fromisoformat, metavar="YYYY-MM-DD")
    ap.add_argument("--until", type=date.fromisoformat, metavar="YYYY-MM-DD")
    ap.add_argument("--format", choices=["table", "csv", "json", "excel"], default="table")
    ap.add_argument("-o", "--output", default="token_usage.xlsx",
                    help="--format excel の出力先 (既定 %(default)s)")
    ap.add_argument("--unit", choices=list(UNITS), default="man",
                    help="表の数値単位: raw=そのまま, man=万, sen=千 (既定 man。csv/json/excel は常に生の数値)")
    ap.add_argument("--rate", type=float, default=pricing.DEFAULT_JPY_PER_USD, help="JPY/USD (既定 %(default)s)")
    ap.add_argument("--by-model", action="store_true", help="モデル別内訳")
    ap.add_argument("--root", default="~/.claude/projects", help=argparse.SUPPRESS)
    a = ap.parse_args(argv)

    root = Path(a.root).expanduser()
    if a.project:
        p = Path(a.project).expanduser()
        root = p if p.is_dir() else root / a.project
    if not root.is_dir():
        sys.exit(f"ログディレクトリが見つかりません: {root}")

    if a.format == "excel":
        detail, unknown = aggregate_detail(iter_records(root), a.since, a.until)
        if not detail:
            sys.exit("該当するデータがありません")
        write_excel(a.output, detail, a.rate)
        print(f"出力しました: {a.output}")
        if unknown:
            print(f"警告: 単価未対応のモデルは明細に含めていません: {', '.join(sorted(unknown))}", file=sys.stderr)
        return

    rows, unknown = aggregate(iter_records(root), a.since, a.until, a.by_model)
    if not rows:
        sys.exit("該当するデータがありません")
    if a.format == "table":
        print(render_table(rows, a.rate, a.by_model, a.unit))
        print(f"\n※ API従量課金換算の目安 (レート ¥{a.rate:g}/USD)。Batch割引・Fast mode・データレジデンシー・サーバーツール課金は含みません。")
    else:
        print(render_raw(rows, a.rate, a.format, a.by_model))
    if unknown:
        print(f"警告: 単価未対応のモデルは料金に含めていません: {', '.join(sorted(unknown))}", file=sys.stderr)


if __name__ == "__main__":
    main()
