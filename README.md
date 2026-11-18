# count-token

Claude Code のローカルログから日別トークン消費量と料金を集計する。仕様は `CLAUDE.md` を参照。

```
uv run count-token                      # 表 (万単位)
uv run count-token --format excel -o report.xlsx
uv run python -m unittest discover tests
```
