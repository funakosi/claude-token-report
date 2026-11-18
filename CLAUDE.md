# count_token — Claude Code トークン消費量・料金集計スクリプト

Claude Code のローカルログから日別のトークン消費量を集計し、料金（USD / 円換算）を算出する Python スクリプトを作る。

## 技術方針

- 言語: Python 3.11+（標準ライブラリのみ。外部依存を増やさない）
- パッケージ管理: uv（src レイアウト、ビルドバックエンドは hatchling）
- 構成
  - `src/count_token/cli.py`: 集計・出力・CLI（`main`）
  - `src/count_token/pricing.py`: 単価表。価格改定時はここだけ直す
  - `src/count_token/xlsx_writer.py`: .xlsx 書き出し
  - `tests/`: unittest
- 実行: `uv run count-token [オプション]`（または `uv run python -m count_token`）
- テスト: `uv run python -m unittest discover tests`

## データソース

- `~/.claude/projects/**/*.jsonl` を再帰的に全て読む（全プロジェクト横断）
- `--project <path>` で特定プロジェクトのみに絞れるようにする
- 各行は JSON。`type == "assistant"` の行の `message.usage` を集計対象にする
  - `input_tokens` → Input
  - `output_tokens` → Output
  - `cache_read_input_tokens` → Cache Read
  - `cache_creation_input_tokens` → Cache Write（`cache_creation.ephemeral_5m_input_tokens` / `ephemeral_1h_input_tokens` があれば 5分/1時間を区別）
  - `message.model` → 料金計算に使うモデル名
  - `timestamp` → 日付判定（ISO8601 UTC。**ローカルタイムゾーン（JST）に変換して**日付集計）
- 重複排除: 同一レスポンスが複数行に書かれる場合があるため、`message.id` + `requestId` の組で重複を除く
- 壊れた行・`usage` のない行は黙ってスキップする（落とさない）
- `model` が `<synthetic>` 等の実課金のないものは除外する

## 集計と出力

- 日付単位で集計。列: `日付 / Input / Output / Cache Read / Cache Write / 総処理量 / 料金(USD) / 料金(円)`
- 総処理量 = Input + Output + Cache Read + Cache Write
- 数値の単位は `--unit raw|man|sen` で切り替える（表のみ。csv/json は常に生の数値）
  - `man`（既定）: 万単位（例: `12万`, `1.8万`, `102.8万`）。1万未満は整数表記
  - `sen`: 千単位（例: `120千`, `1.8千`）。1000未満は整数表記
  - `raw`: 桁区切りの数字そのまま（例: `1,028,000`）
  - 小数は最大1桁
- 日付表記は `MM/DD`（年をまたぐ場合は `YYYY/MM/DD`）
- 最終行に合計行を出す
- 出力例:

```
日付    Input    Output    Cache Read    Cache Write    総処理量    USD      JPY
11/16   12万     1.8万     85万          4万            102.8万     $12.34   ¥1,850
11/17   9.5万    2.5万     72万          3.5万          87.5万      $9.80    ¥1,470
合計    ...
```

- オプション
  - `--since YYYY-MM-DD` / `--until YYYY-MM-DD`: 期間指定
  - `--format table|csv|json|excel`: 出力形式（既定は table）。csv/json/excel では万単位表記にせず生の数値を出す
  - `-o, --output <path>`: excel の出力先（既定 `token_usage.xlsx`）
  - `--rate <JPY/USD>`: 円換算レート（既定は設定値。取得元・日付は README か設定に明記し、自動取得はしない）
  - `--by-model`: モデル別内訳も表示

## Excel 出力（`--format excel`）

- 外部ライブラリ不要。`xlsx_writer.py` が標準ライブラリ（zipfile）で .xlsx を直接書く。数式と計算済みキャッシュ値の両方を書き、開くと再計算される
- シート構成
  - `日別集計`: 日付 / Input / Output / Cache Read / Cache Write / 総処理量 / USD / JPY。`明細` を SUMIFS で集計する数式。B1 に円レート
  - `明細`: 日付 × 単価ラベルごとの行。トークン数と、`単価表` から VLOOKUP した単価（Input / CW 5m / CW 1h / Cache Read / Output）を並べ、料金 = `(Input×単価 + CW5m×単価 + CW1h×単価 + CacheRead×単価 + Output×単価) / 1,000,000` の数式
  - `単価表`: モデル別単価（USD/MTok）。ここを書き換えると全体が再計算される
- Haiku 5.5 はリクエスト単位の閾値があるため、単価ラベルを `claude-haiku-5-5 (<=100k)` / `(>100k)` に分けて集計する
- 単価未対応モデルは明細に含めず、stderr に警告を出す
- `--by-model` は不要（明細が常にモデル別）

## 料金計算

参照元: https://platform.claude.com/docs/en/about-claude/pricing

- ログの `model` ごとに単価を引いて行（リクエスト）単位で計算し、日別に合算する
- 料金 = トークン数 / 1,000,000 × 単価(USD/MTok)
- モデル名はバージョン日付サフィックス付き（例: `claude-sonnet-4-5-20250929`）でも前方一致で引けるようにする
- 単価表に無いモデルは料金を `未対応` として警告を出し、トークン数だけ集計する（合計料金には含めず、その旨を明記）
- Cache Write は 5分書き込みと 1時間書き込みで単価が異なる。内訳が取れない場合は 5分として扱う
- Cache Read（cache hits）はモデルにより倍率が異なる（下表の通り）

### 単価表（USD / MTok、2026-11-18 時点で取得。実装時に必ず公式ページで再確認する）

| モデル | Input | 5m Cache Write | 1h Cache Write | Cache Read | Output |
| :-- | --: | --: | --: | --: | --: |
| Fable 5.2 / Mythos 5.2 | 8 | 10 | 16 | 0.20 | 40 |
| Fable 5.1 / Mythos 5.1 | 10 | 12.50 | 20 | 0.25 | 50 |
| Fable 5 / Mythos 5 | 10 | 12.50 | 20 | 1 | 50 |
| Opus 5.5 | 4 | 5 | 8 | 0.20 | 20 |
| Opus 5 / 4.8 / 4.7 / 4.6 / 4.5 | 5 | 6.25 | 10 | 0.50 | 25 |
| Opus 4.1 / Opus 4 | 15 | 18.75 | 30 | 1.50 | 75 |
| Sonnet 5.5 | 2 | 2.50 | 4 | 0.10 | 10 |
| Sonnet 5 | 2 | 2.50 | 4 | 0.20 | 10 |
| Sonnet 4.6 / 4.5 / 4 | 3 | 3.75 | 6 | 0.30 | 15 |
| Haiku 4.5 | 1 | 1.25 | 2 | 0.10 | 5 |
| Haiku 5.5（プロンプト ≤100k トークン） | 0.10 | 0.125 | 0.20 | 0.01 | 0.50 |
| Haiku 5.5（プロンプト >100k トークン） | 0.50 | 0.625 | 1 | 0.05 | 2.50 |

規則:
- Cache Write 5m = Input × 1.25、1h = Input × 2、Cache Read は通常 Input × 0.1（Fable/Mythos 5.1・5.2 は 0.025、Opus 5.5 / Sonnet 5.5 は 0.05）
- Haiku 5.5 のみリクエスト単位でプロンプト長（input + cache read + cache write の合計）が 100,000 を超えると高い単価を適用
- それ以外の 4.6 以降のモデルは 1M コンテキストまで同一単価（長文割増なし）
- 本スクリプトでは Batch 割引（50%）、Fast mode、データレジデンシー（1.1x）は考慮しない。ログから判別できないため。前提として出力のフッターか README に明記する
- Web 検索などサーバーツールの追加課金は対象外

## 注意

- 出力される料金は API 従量課金で換算した「目安」。Pro/Max などサブスクリプション利用者の実請求額とは異なる
- 個人のログ（会話内容）は読み取るが、トークン数・モデル名・タイムスタンプ以外は保持・出力しない

## テスト

- `tests/` にサンプル JSONL（匿名化したダミー）を置き、集計・重複排除・万単位フォーマット・料金計算を検証する
- 実行: `uv run python -m unittest discover tests`
