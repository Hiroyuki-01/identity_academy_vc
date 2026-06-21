# ファンド提案チーム — 組織設計（agent team org）

> **目的**：実ファンド構築ではなく、**第3回課題（Solafuneへの投資を提案するVCピッチ）を勝たせる**ためのエージェントチーム。
> **設計憲法**：講義（`slides/IA_12期_セッションDay1.pdf` VCのメカニズム / `Day2.pdf` VCの評価点）＝「優秀な学生のスタートライン」に準拠。
> **最終意思決定**：すべて **ユーザー（このアカウント＝LP兼 最終決裁者）に集約**。どのエージェントも決裁しない。

## 階層図

```
                        YOU（LP兼 最終決裁者）★ go/no-go はすべてここ
                              ▲ IC提案メモ＋推奨を上申
   意思決定  managing-partner（IC議長）  ── 両班を統合し推奨を作る・決めない
                              ▲ diligence済みの版のみ採用
   確認      diligence-lead  ＋  vc-expert-judge  ＋  student-judge
              （整合・openpyxl検算・反トップダウン・ゲート ＋ 第三者採点）
              ▲                                   ▲
   ① ストーリー班（投資の妥当性を物語に）        ② ファンド中身班（4S設計＋財務）
      story-researcher（リサーチ）               researcher（市場/comps リサーチ）
      story-architect（制作）                    fund-architect（4S設計・制作）
                                               financial-modeler（xlsx・制作）
                          ↓ 物語＋確定数字
              slide-writer / scriptwriter-qa（下流：スライド・台本）
```

## ロスター（4機能 × 2班 ＋ 上位共通層）

| 階層 | エージェント | 機能 | 状態 |
|---|---|---|---|
| 意思決定 | `managing-partner` | IC提案メモ作成・ユーザーへ上申（決裁しない） | 新規 |
| 確認 | `diligence-lead` | 4S整合・モデル検算・反トップダウン・ゲート | 新規 |
| 確認 | `vc-expert-judge` / `student-judge` | 第三者採点（ロジック／印象） | 既存・統合 |
| ①ストーリー | `story-researcher` | 投資説得の素材（なぜ今/なぜ我々/競争/1号2号） | 新規 |
| ①ストーリー | `story-architect` | 投資ストーリー＋QA防御の制作 | 新規 |
| ②中身 | `researcher` | ボトムアップ市場規模・comps・必要売上 | 既存・統合 |
| ②中身 | `fund-architect` | 4S（Strategy/Economics/Portfolio/System）設計 | 新規 |
| ②中身 | `financial-modeler` | `02_財務モデル.xlsx`（数式連動・fund-returner判定） | 既存・統合 |
| 下流 | `slide-writer` / `scriptwriter-qa` | スライド骨子・台本・Q&A | 既存・温存 |

## 進め方（ボトムアップ → 上申）
1. **リサーチ**：`researcher`（定量）＋ `story-researcher`（定性）が素材を作る。
2. **制作**：`fund-architect`（4S）→ `financial-modeler`（数値）／ `story-architect`（物語）。
3. **確認**：`diligence-lead` が整合・検算でゲート → 通れば `vc-expert-judge`/`student-judge` が採点（`evals/採点ログ.md` に追記）。
4. **意思決定**：`managing-partner` が統合し **IC提案メモ＋推奨** を作成。
5. **決裁**：**ユーザー** が承認/修正。フィードバックは上から下へ差し戻し。
6. 下流の `slide-writer`/`scriptwriter-qa` で資料・台本に落とす。

> 実務上の dispatcher（各エージェントを順に起動し、最後の決裁をユーザーに渡す役）は、ユーザーの指示を受けたメインセッションが担う。サブエージェントは原則サブを起動しないため、階層は「論理的な権限・依存関係」を表す。

## 設計憲法（講義由来・全エージェント共通の前提）
- **4S**：Strategy / Economics / Portfolio / System（18項目）。
- **他人資本の4制約**：リターン(LP~20%＝10年6X)／時間軸(10年5+5)／金額(1社≤総額~10-13%)／機会(起業家>投資家)。
- **べき乗則**：ネット3X＝少数アウトライヤー(7-8X+)の積み上げ。fund-returner（1社150億超回収）が核。
- **市場はボトムアップ厳守**（トップダウンのシェア仮定は不可）。
- **財務**：詳細モデリング不要・ただし明確な前提＋積み上げ（5年売上/GM/Post-Val/Exitパス＆金額/セカンダリ割引）。
- **接続**：集中度上限↔LP構成、件数↔チーム天井。
- **評価基準①〜⑤**（`evals/eval.md`、45点）が最終の物差し。

## 既存6体の扱い（決定：統合・格上げ）
- `researcher`＝②中身班の定量リサーチに格上げ。`financial-modeler`＝②中身班の制作。`vc-expert-judge`/`student-judge`＝確認層の第三者採点。`slide-writer`/`scriptwriter-qa`＝下流の資料制作として温存（今回の2班スコープ外だが残す）。役割重複なし。
