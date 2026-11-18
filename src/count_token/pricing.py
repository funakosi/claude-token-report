"""単価表 (USD / MTok)。出典: https://platform.claude.com/docs/en/about-claude/pricing
2026-11-18 時点。価格改定時はこのファイルだけ直す。"""

DEFAULT_JPY_PER_USD = 150.0  # 目安。--rate で上書き

# (input, cache_write_5m, cache_write_1h, cache_read, output)
def _p(inp, out, read_mult=0.1):
    return (inp, inp * 1.25, inp * 2, inp * read_mult, out)

# 前方一致で引く。長いキーを先に評価する。
PRICES = {
    "claude-fable-5-2": _p(8, 40, 0.025),
    "claude-mythos-5-2": _p(8, 40, 0.025),
    "claude-fable-5-1": _p(10, 50, 0.025),
    "claude-mythos-5-1": _p(10, 50, 0.025),
    "claude-fable-5": _p(10, 50),
    "claude-mythos-5": _p(10, 50),
    "claude-opus-5-5": _p(4, 20, 0.05),
    "claude-opus-5": _p(5, 25),
    "claude-opus-4-8": _p(5, 25),
    "claude-opus-4-7": _p(5, 25),
    "claude-opus-4-6": _p(5, 25),
    "claude-opus-4-5": _p(5, 25),
    "claude-opus-4-1": _p(15, 75),
    "claude-opus-4": _p(15, 75),
    "claude-sonnet-5-5": _p(2, 10, 0.05),
    "claude-sonnet-5": _p(2, 10),
    "claude-sonnet-4-6": _p(3, 15),
    "claude-sonnet-4-5": _p(3, 15),
    "claude-sonnet-4": _p(3, 15),
    "claude-haiku-4-5": _p(1, 5),
}

# Haiku 5.5 はプロンプト長 (input+cache read+cache write) で単価が変わる
HAIKU_5_5_PREFIX = "claude-haiku-5-5"
HAIKU_5_5_THRESHOLD = 100_000
HAIKU_5_5_LOW = _p(0.10, 0.50)
HAIKU_5_5_HIGH = _p(0.50, 2.50)


HAIKU_5_5_LOW_LABEL = HAIKU_5_5_PREFIX + " (<=100k)"
HAIKU_5_5_HIGH_LABEL = HAIKU_5_5_PREFIX + " (>100k)"


def resolve(model, prompt_tokens=0):
    """(単価ラベル, 単価タプル) を返す。未対応モデルは (None, None)。
    ラベルは日付サフィックス違いを同一視した単価表の行名。"""
    if model.startswith(HAIKU_5_5_PREFIX):
        if prompt_tokens > HAIKU_5_5_THRESHOLD:
            return HAIKU_5_5_HIGH_LABEL, HAIKU_5_5_HIGH
        return HAIKU_5_5_LOW_LABEL, HAIKU_5_5_LOW
    for key in sorted(PRICES, key=len, reverse=True):
        if model.startswith(key):
            return key, PRICES[key]
    return None, None


def lookup(model, prompt_tokens=0):
    """単価タプルを返す。未対応モデルは None。"""
    return resolve(model, prompt_tokens)[1]


def price_table():
    """単価表の全行 [(ラベル, 単価タプル)]。"""
    rows = list(PRICES.items())
    rows += [(HAIKU_5_5_LOW_LABEL, HAIKU_5_5_LOW), (HAIKU_5_5_HIGH_LABEL, HAIKU_5_5_HIGH)]
    return rows


def cost_usd(model, inp, out, read, w5m, w1h):
    """料金(USD)。未対応モデルは None。"""
    p = lookup(model, inp + read + w5m + w1h)
    if p is None:
        return None
    return (inp * p[0] + w5m * p[1] + w1h * p[2] + read * p[3] + out * p[4]) / 1_000_000
