# claude-token-report

Claude Code のローカルログから日別トークン消費量と料金を集計する。仕様は `CLAUDE.md` を参照。

```
uv run claude-token-report                      # 表 (万単位)
uv run claude-token-report --format excel -o report.xlsx
uv run python -m unittest discover tests
```

## 前提・注意

- 料金は API 従量課金で換算した **目安** です。Pro/Max などサブスクリプションの実請求額とは異なります。
- 単価は <https://platform.claude.com/docs/en/about-claude/pricing> を参照し、2026-11-18 時点の値を `src/claude_token_report/pricing.py` に記載しています。価格改定時はこのファイルだけ修正してください。
- 円換算レートは固定値 150 JPY/USD（目安。為替サイト等からの取得ではなく、自動取得もしません）。`--rate` で上書きできます。
- Batch 割引（50%）、Fast mode、データレジデンシー（1.1x）、Web 検索などサーバーツールの追加課金は、ログから判別できないため考慮していません。
- 単価表に無いモデルはトークン数のみ集計し、合計料金には含めません。
- `~/.claude/projects/**/*.jsonl` を読み取りますが、トークン数・モデル名・タイムスタンプ以外は保持・出力しません。

## ライセンス

MIT（`LICENSE` を参照）
