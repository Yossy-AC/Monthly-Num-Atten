# Course-Stats-Analyzer

FastAPI から CLI/Skill へ移行した月次受講人数集計ツール。

lists フォルダ内の複数 Excel ファイル（生徒受講情報）を自動処理し、月別の受講人数を集計。Pivot 形式 Excel で出力します。

## 機能

- **自動集計**: lists/ 内のすべての Excel ファイル（`*_YYMM.xlsx`）を処理
- **Pivot形式**: 固定列（学年/教室/講座名/M/C/担当）× 月列（4月～3月）
- **CSV保管**: 月別結果を自動保存（`outputs/results/{YYYY-MM}.csv`）
- **Excel出力**: 全月データを統合した Pivot テーブル Excel（`monthly_stats.xlsx`）を出力

## クイックスタート

### 環境準備
```bash
uv sync
```

### 実行方法

#### 方法1: CLI スクリプト
```bash
python scripts/aggregate.py
```

#### 方法2: Claude Code Skill
Claude Code で以下を入力：
```
/aggregate-enrollment
```

## 入力ファイル仕様

### ファイル名形式
```
*_YYMM.xlsx
```

例：`〔定例報告〕2025AC受講者ﾘｽﾄ_2504.xlsx`

- **YY**: 年度下2桁（25 = 2025年度、26 = 2026年度など）
- **MM**: 月番号（01～12）

### 年号変換ルール

任意の年度に対応：
- month ≥ 4 → fiscal_year の暦年
- month < 4 → fiscal_year + 1 の暦年

例：
- _2504 → 2025-04（4月、2025年暦年）
- _2501 → 2026-01（1月、2026年暦年）
- _2604 → 2026-04（4月、2026年暦年）
- _2601 → 2027-01（1月、2027年暦年）

### Excel列構造

Row 4（0-indexed）がヘッダー。以下の列を使用：

| 列 | Index | 用途 |
|---|---|---|
| C | 2 | 受講追加日付 |
| G | 6 | 受講取消日付 |
| J | 9 | 講座名 |
| K | 10 | クラス区分（【マスター】【コア】） |
| L | 11 | 受講教室 |
| P | 15 | 学年コード（31=高1, 32=高2, 33=高3） |
| AA | 26 | 担当 |

## 集計仕様

### 対象者フィルタ
- **学年**: 高1/高2/高3 のみ（学年コード 31/32/33）
- **担当**: 「0」「-」「」（空）を除外

### 月別判定ルール
- **基準日**: 対象月の**前月末日**
  - 例：2025-05 月の集計 → 2025-04-30 時点
  - 判定式：`add_date <= cutoff AND (cancel_date IS NULL OR cancel_date > cutoff)`

### グループ化軸（5次元）
1. 学年
2. 教室
3. 講座名
4. M/C（マスター/コア）
5. 担当

### 出力形式
- **固定列**: 上記 5次元
- **月列**: 4月 → 3月（会計年度順）
- **セル値**: 受講人数（該当データがない月は出力されない）

## ファイル構成

```
Course-Stats-Analyzer/
├── scripts/
│   └── aggregate.py             # CLI メイン実行スクリプト
├── services/
│   └── aggregator.py            # 集計コアロジック
├── .claude/skills/
│   └── aggregate-enrollment/
│       └── SKILL.md             # Claude Code Skill定義
├── .claude/
│   └── settings.json            # ローカル設定
├── outputs/
│   ├── results/                 # 月別 CSV（{YYYY-MM}.csv）
│   └── monthly_stats.xlsx       # 最終出力 Excel（Pivot形式）
├── lists/                       # 入力 Excel ファイル（*_YYMM.xlsx）
├── uploads/                     # 一時保存（未使用）
├── pyproject.toml               # 依存関係定義
├── uv.lock                      # ロックファイル
└── CLAUDE.md                    # 開発者向け詳細仕様
```

## パフォーマンス

- 11ファイル（2025-04～2026-02）の集計・出力：約 5秒
- ベクトル化処理により行単位の繰り返し計算を廃止
- 出力ファイルサイズ: 9.2 KB

## 技術詳細

### コア関数（services/aggregator.py）

#### `parse_target_month(filename: str) -> pd.Period | None`
ファイル名から対象月を抽出。任意年度に対応。
```python
# 例
parse_target_month("file_2504.xlsx")  # → 2025-04
parse_target_month("file_2601.xlsx")  # → 2027-01
```

#### `aggregate(df, target_month) -> pd.DataFrame`
対象月の受講人数を集計。Pivot 準備形式で返す。
- グループ化軸：学年, 教室, 講座名, M/C, 担当
- ベクトル化処理で高速化

#### `build_pivot(results_dir) -> pd.DataFrame`
全月 CSV をマージして Pivot テーブル生成。
- `pd.concat()` + `groupby().sum()` で効率化
- MONTH_ORDER に従って月を整列

#### `to_excel_bytes(result) -> bytes`
DataFrame を Excel bytes に変換。openpyxl 使用。

### 会計年度ロジック

```python
fiscal_year = 2000 + year_suffix
year = fiscal_year if month >= 4 else fiscal_year + 1
```

この汎用ロジックにより、2025年度以降の任意年度に対応可能。

## 開発者向け情報

詳細は [CLAUDE.md](CLAUDE.md) を参照（実装仕様、関数説明、制約事項など）

---

**Last Updated**: 2026-02-26 (v4: CLI/Skill化、WebUI削除)
