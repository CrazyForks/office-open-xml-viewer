# 改善計画 実行チェックリスト（2026-07）

`docs/improvement-plan-2026-07.md` の実行トラッカー。項目 ID（A1, B2, C1, D5, E2 …）は計画本文の表を参照。
ローカルセッションで進める際は、完了した項目にチェックを入れてこのファイルを更新していく。

## 前提・引き継ぎ事項

- レビューはリモートセッションで実施。**主要指摘（god file 行数、`rawData.slice(0)`、pptx/xlsx `destroy()` の挙動、`parse_docx` のエラー整形、exports の条件順、CI typecheck ジョブの rust-cache 欠如）はコード上で裏取り済み。** それ以外の細部（行番号・定数値）は着手時に現物を再確認すること。
- VRT はローカル専用（`private/sample-*` が必要）。**レンダリングに触る修正は必ずローカルで `pnpm build:wasm && pnpm vrt` を回してから push。** リモートでは実行不可だった点に注意。
- 参照画像の更新（`UPDATE_REFS=1`）はユーザー明示指示のみ。
- 横断修正（Phase 3）は CLAUDE.md の規定どおり core / ooxml-common を含む 1 本の協調 PR にまとめてよい。commit は 1 関心ずつ。
- Phase 2 の各項目は **PR にベンチ前後比較を記載**（大型 docx / 200 スライド級 pptx / 多シート xlsx で計測）。
- **レビュー実施後に `DocxScrollViewer` / `PptxScrollViewer`（#650 系〜#656）が main に追加された。** 本計画の viewer 系指摘（C5/C6 の destroy ライフサイクル、C4 のスクロール描画、C7 の共有化）は着手時に ScrollViewer 系クラスも同じ観点で点検し、同種の穴があれば同一 PR で修正する。

## 検証コマンド early reference

```bash
pnpm build:wasm                          # parser を触ったら必須
pnpm test                                # vitest（unit）
pnpm typecheck
pnpm vrt                                 # ローカルのみ・全パッケージ
pnpm --filter @silurus/ooxml-docx vrt    # 単一パッケージ
cargo fmt --all --check && cargo clippy --all-targets -- -D warnings
cargo test
```

---

## Phase 1 — 安全網と即効修正 ✅ 完了（2026-07-02、PR #658–#664）

全 22 項目完了。各項目の注記は実装時の再定義・追加発見を含む（計画の指摘が現物と異なった項目: A6 / D7 / A8 / A9 / E5 barrel）。CI 運用知見: smoke ジョブの `playwright install chrome --with-deps` が apt/CDN 起因で 30 分超ハングする事象が 2 回連続発生 → プリインストール Chrome があれば skip + 全ジョブ timeout-minutes で恒久対策済み（#662 内）。

### CI / パッケージング

- [x] E1 (#658): smoke suite（`tests/smoke/layouts.spec.ts`）を CI に接続（WASM ビルド → Storybook 起動 → Playwright chromium）
- [x] E3 (#658): typecheck ジョブに `Swatinem/rust-cache@v2`（3 parser workspace 指定。rust ジョブ / publish.yml の記述を流用）
- [x] E8 (#658): devDep 化 + `OOXML_REQUIRE_SKIA=1` で silent skip を hard fail 化。pnpm 10 の build-script 承認（`onlyBuiltDependencies: [skia-canvas]`）が必須と判明し追加。初実行でテスト自体の実バグ 2 件（skia@2 API 不一致・到達不能アサーション）と pptx unhandled-rejection リークを検出・修正
- [x] E6 (#664): publish.yml に test + publint + attw `--profile esm-only`（fail-closed）+ tarball smoke（temp dir に npm i して 3 サブパス import）を追加
- [x] E2 (#664): types 先頭化 + default 追加。実測: tsc 4 象限は before も pass だが、attw が TS bug #50762 の「fallback condition 誤用」🐛 を bundler/node16-ESM 全 5 エントリで検出 → 修正後 🟢、publint 5 errors → clean
- [x] E5 (#664): sideEffects: false を root+4 パッケージに追加（grep 監査 + 独立監査 + DOM なし Node import 成功の三重確認）。「docx flat re-export 非対称」は計画の誤読と判明（root は対称、flat barrel は各パッケージ自身のもの）— API 変更なし

### 正しさ / 堅牢性（ユーザー可視）

- [x] C6 (#659): pptx `viewer.ts` `destroy()` — `wrapper.remove()` 前に caller の canvas を `insertBefore` で返還。検証: destroy → 同一 canvas で再生成
- [x] C5 (#659): xlsx `viewer.ts` `destroy()` — wrapper subtree と注入 `<style>` を除去（style は module-wide 1 回注入への変更でも可）。検証: Storybook で mount/unmount 繰り返し
- [x] C5/C6 追補 (#659): ScrollViewer ×2 は点検の結果、完全実装で修正不要。代わりに **DocxViewer に pptx と同一の canvas 喪失バグを発見し同 PR で修正**。レビューで stale-nextSibling の DOM 仕様違反（NotFoundError）も検出しガード追加、テスト DOM の二重所属欠陥も修正
- [x] A6 (#660) **再定義**: 「lgDash→solid」は誤診（dash.ts のテーブルは実は §20.1.10.82 ST_TextUnderlineType の下線用で、prstDash 値は流れない）。真バグ = paint.ts の `sysDashDotDot` 欠落（solid に化ける）を修正。preset テーブルを dash.ts へ `pptxPresetDashArray` として移設（byte-equivalence テスト付き）、下線側を `pptxUnderlineDashArray` に改名し誤 spec 引用を修正
- [x] D11 (#661): Result 化 + TS 受け口 2 箇所（worker.ts / render-worker.ts）の error フィールド probe 削除。壊れた入力で `docx-parser error: ...` throw を end-to-end 確認
- [x] D9(部分) (#661): map_err 化（panic による WASM インスタンス死を排除）

### 即効 perf

- [x] D1 (#662): 計画の 3 箇所に加え grep で 10 箇所追加発見、**計 13 箇所**を `index_for_name` に置換（bytes を実際に使う 3 箇所は正しく除外）。共有ヘルパーは不要と判断（archive ハンドル既保持のため inline が自然）
- [x] C1 (#662): `data` フィールドをプロトコルから**削除**（optional で残すより契約が明確 — parse 前の parseSheet は明示エラー）。worker-vs-main VRT 0.000% diff で挙動同一を証明
- [x] C3 (#663): `scrollableIndexAt` 新設で O(log n) 化。ベンチ: 50 万行 ×1000 ヒットで 9296ms → 0.378ms（~24,600×）、旧実装 verbatim コピーを oracle にした全数パリティ 0 差
- [x] B6 (#663): `WeakMap<fontFamilyClasses, Map<family, css>>`（identity キー、呼び出し面変更ゼロ）。font 再代入スキップは測定ループの 2 writer（measureText/strAdvance）を単一トラッカーで統一（draw パス 12 箇所は据え置き）
- [x] A8(前半) (#663): slice は **44 バイト**が正（EMF 判定に offset 40-43 が必要。計画の 8 バイトでは不足）。raster は Blob 直渡しで全体コピー消滅
- [x] D10(前半) (#662): 3 バイナリ計 1,967,032 → 1,903,734 bytes（**−3.22%**）。ビルド 23.8s → 33.4s。panic=abort でも console_error_panic_hook のスタック出力は保持（hook は abort 前に走る）

### 小粒衛生

- [x] D7 (#661): 実像は「docx にも私的ヘルパーあり」で**計 5 実装・~45 呼び出し面**を generic `read_zip_bytes/read_zip_string<R: Read+Seek>` に集約。cap 超過は Err 返却とし、画像パスの従来挙動（スキップ）は呼び出し側の明示的 `.ok()` で維持
- [x] A11 (#663): 4 export（edt1d/shadePixel/shadeParamsFor/fillDirFromKey）削除。materialClass/lightDirFromRig は pptx が実使用のため保持。テストは既に deep import で変更不要。dead export はゼロ（全数調査）
- [x] A9 (#663): 無条件直呼びは wmf.ts の 1 箇所のみだった（emf/dib はチェック済み）。createAuxCanvas を `core/canvas/aux-canvas.ts` へ移設し 3 ファイル統一 — EMF/DIB は main-thread fallback を獲得（strict superset）。worker 専用の OffscreenCanvas+transferToImageBitmap は正当につき対象外。pattern-bitmaps の重複 factory も統合
- [x] C11 (#663): pptx に加え **docx にも同型 monkey-patch を発見**（横断原則）— 両方 WeakMap 化。xlsx は同期描画で該当なし

## Phase 2 — WASM 境界とキャッシュ再設計 ✅ 完了（2026-07-03、PR #666–#672）

体感成果: 15MB pptx パース 30→11ms（#666+#669 複合）、xlsx シート切替 1.61×+転送ゼロ、docx ページ再訪 ~9×、xlsx スクロール 100 イベント→1 render、wedged worker の永久ハング解消。全 PR にベンチ/等価性証明（sha256 golden / oracle / byte-identity）添付。

- [x] ベンチ基盤 (#666/#671): `packages/node/src/bench-parse.mjs`（境界込み parse 時間、string/bytes 自動判別）+ `bench-handle.mjs`（parse 後の反復 work: 全シート切替/全画像抽出、--wasm-dir で before 計測）
- [x] C2+B7+D9 (#666): 4 parse 関数を Vec<u8>(JSON bytes) 化、worker は transferable 素通し、main で 1 回 decode+parse（render-worker は in-worker 消費なので現行維持）。15MB pptx median 30→15ms(~2×)。レビューの「WASM memory 全体転送」指摘は生成コードの .slice() 確認で却下
- [x] D2 (#671): stateful `#[wasm_bindgen]` アーカイブハンドル（`DocxArchive`/`PptxArchive`/`XlsxArchive` = 所有 `ZipArchive<Cursor<Vec<u8>>>` + `max` 保持）。bytes を WASM へ 1 回コピー・central directory を 1 回スキャンし `parse()`/`extract_image(path)`/`parse_sheet(i,name)` を retained archive 上で提供。フリー関数（`parse_*`/`extract_*`）は所有 archive を張る thin wrapper 化で温存（node/markdown/stories/MCP 無変更）。各 worker は `currentBuffer: Uint8Array` を `archive: *Archive|null` に置換し、構築後は JS 側 bytes を保持しない（メモリ二重化解消）。再 parse 時 `disposeArchive()` で明示 free（二重 free/UAF ガード）。ベンチ: pptx sample-4 (15MB, parse+5 media) **1.49×**、docx sample-9 (13 images) 1.10×
- [x] D3 (#671): xlsx `WorkbookShared`（workbook.xml/rels source + sheet list + theme palette + sharedStrings）をハンドル内で 1 回パースし全シート切替で再利用。`parse_sheet` 毎の sharedStrings/theme 全再パースを解消（`parse_sheet_with`/`parse_xlsx_inner_with` に分離、フリー関数は毎回 fresh build で従来コスト維持）。ベンチ: sample-12 (8 sheets) **1.61×**、sample-9 (3 sheets) 1.30×。roxmltree `Document` は借用のためキャッシュ不可 → cached string から都度再パース（inflate なしで安価）。drawing XML の parse-once 化は sheet 単位再パース解消を主目的とし本コミット範囲では見送り（別項）
- [x] D4 (#669): master 11 extractors+bg が単一 Document 共有、ParsedLayout+layout_cache で 4×S→1/distinct+1/slide。「slide XML 2 回パース」は誤認（1 回）。layout decorative は slide 固有 smartart 依存で意図的非キャッシュ（unsound 回避）。出力 sha256 byte-identical、-20%/-26% parse 回数、cfg(test) カウンタで 2+2N を regression 固定
- [x] B3 (#668): base bitmap は core キャッシュ、a:clrChange recolor は第 2 層（per-fetchImage WeakMap）でメモ化。ページ再訪の画像コスト ~9×（0.27→0.03ms）。destroy() で 3 キャッシュ全 drop
- [x] A8(後半) (#668): pptx 実装を core/image/bitmap-image-by-path.ts へ verbatim lift（LRU 256、同期 peek、eviction close、rejection-handler 規律をヘッダに明文化）。pptx は thin re-export。**xlsx は意図的非移行**（raster+SVG 同居の workbook 所有 Map で構造が異なる — false-abstraction 回避、二重 LRU と close 所有権競合を防ぐ）
- [x] C4 (#667): scheduleRender で scroll/resize/drag を rAF coalesce（100 イベント→1 render）、明示 API は同期維持。size guard は orchestrator 側にも（+setTransform 冪等化）。_renderSeq 世代で stale bitmap close。worker-vs-main VRT 0.000%
- [x] A4 (#672): innerShadow/softEdge(×3)/reflection の aux を bbox⊕margin に縮小（blur は 3σ=3·blur — pixel テストが margin バグを捕捉）。ピクセル −~99%（79–87×）。**プールは不採用**（bbox 化後はサイズ不一致でヒットせず）。**reflection は full-canvas に revert** — mirror blit のリサンプリングが crop 端で skia プラットフォーム依存（Linux CI で δ≤7 検出）。tripwire テストで再 crop を防止
- [x] A3 (#672) **縮小案採用**: region 限定（getImageData/loop/putImageData）のみ、GPU 合成は VRT 割れリスクで見送り。byte-identical 証明済み。正直な注記: pptx は shape 専用オフスクリーンのため現行パスの実利は pad 分のみ — seam は大 canvas 呼び出し向け
- [x] A5 (#672): {op, argTokens} へ lazy プリコンパイル（WeakMap、evaluatorForDef 単一チョークポイント）。186 プリセット全数 oracle 一致。split() 呼び出し −99.6%（12.9×）
- [x] A10 (#670): per-request timeout + AbortSignal + worker error/messageerror の pending 一括 reject（常時有効）。timeout は opt-in（LoadOptions.workerTimeoutMs、既定無制限 — 巨大ファイルの正当な長時間パースを壊さない）。viewer 5 箇所の明示 forwarding
- [~] B4: **Phase 4-1 に統合**。B3 で再訪コストの主犯（画像再デコード）は解消済み。LayoutLine[] キャッシュは B2 統一（compute-once 単一成果物）の副産物として実装するのが正道 — 中間キャッシュを別実装すると Phase 4-1 で捨てることになるため

## Phase 3 — 共有層への集約 ✅ 完了（2026-07-03、PR #677–#680, #682 + issues #674–#676）

- [x] D5 (#677): ThemeColorScheme(12 slot, srgb/sys/prst)+ThemeFonts+parse_ln_style_widths+preset_color を common 化。各 parser は thin adapter で出力 byte-stable。pptx lnStyleLst/objectDefaults は契約差で意図的 local 維持。xlsx/docx が prstClr を獲得（コーパス該当なしの latent 改善）
- [x] D6 (#677) **再定義**: leading-slash は実は 3 parser とも処理済み（計画が古い）。真の穴 = docx load_media_map の `../` 未正規化 → resolve_target 単一実装（leading-slash + ../ 正規化）で修正。BTreeMap 決定論化（serialize される HashMap 3 フィールド）も同 PR
- [x] D8 (#680) **再定義込み**: txBody は実質 2 重（docx は wps:txbx で対象外）。color-node=ColorSource+ThemeResolver trait（sysClr は lastClr.or(val) に統一）、fill=pptx の gradFill/pattFill を verbatim 移設（xlsx/docx への新機能展開は消費側が無くスコープ外）、BodyPr=spec デフォルト込み common 化
- [x] C7 (#678): decodeDataUrl 私的コピー削除 / google-fonts 単一レジストリ（oracle で byte-preserve、additive 差分のみ）/ preferVectorBlip type guard（4 call site）/ bidi は**共有可能な機械部分のみ**（RTL_GATE byte-identity 検証・buildVisualOrder）— docx classOverride / xlsx readingOrder 等の ECMA 根拠ある format 差は format 分岐関数を作らず各パッケージ維持
- [x] C10 (#682): ooxml_common::chart に core TS ChartModel と 1:1 の struct 群。pptx は el.chart 直渡し（60 フィールドコピー削除）、xlsx は From<ChartData>（adapter ロジック Rust 移設・対応表は commit に）。削除した旧実装をテスト内に凍結した oracle deep-equal で等価性証明。VRT chart サンプル不変
- [x] C8 (#680): xlsx parser が inset 4 属性をパース（spec デフォルト 91440/45720 EMU 常時 emit）、renderer は per-side inset（9.6/4.8px、asymmetric 対応）+ ascent は actualBoundingBoxAscent 実測（0.85 は fallback、pptx の inline 累積とは非同一につき core lift は原則どおり見送り）。×1.2 行送りは正確なので不変。VRT: committed 参照 byte-identical・閾値超過ゼロ（private 3 シートに spec 方向 ~2.6px の改善シフト — 参照再生成は任意の housekeeping）
- [x] A7 (#679): 10 サイトを幅ベース elide（'…' 付き二分探索、サロゲートペア安全化込み）に。凡例予算は equal-share 案を fidelity 理由で全幅に修正した結果、**現コーパス VRT 100% identical**（真にあふれる場合のみ省略）。CJK の文字数切断問題を解消
- [x] B12: issues #674（floating-table page-fit §17.4.57）/ #675（frame keep-with-anchor §17.3.1.11）/ #676（empty-mark 1em 閾値）を撤去条件つきで起票。「inside/outside margin 近似」は現物に存在せず（#676 に注記）。**ユーザー承認済み（2026-07-03）: Phase 4 完了後に #675→#674→#676 を自律実装（Word 逆解析承認込み）**

## Phase 4 — 大型構造リファクタ ✅ 一括バッチ完了（2026-07-04、PR #681, #684–#691。B2 後続のみ継続）

- [x] B1: **廃止（設計判断、2026-07-02 ユーザー裁定）** — 行数起点の機械的分割は独立作業として行わない。モジュール境界はドメイン構造が要求するときだけ引く。B2 の統一が生んだ相境界に沿った物理分割（verbatim + characterization、#635 手法）として B2 の最終段に吸収
- [x] B2 (#684, #689): measure/paint 統一 — **paragraph 完了**。Stage 1 (#684): paginator が layoutLines を stamp、renderParagraph が入力一致ガード（kinsoku 値等価・NUMPAGES/noteRef 除外）で再利用。Stage 2 (#689): **zoom 不変行分割** — scale-1 の行分割 partition のみ再利用し行ジオメトリは paint scale で再測定（rescaleLayoutLines。×scale はヒンティング非線形で drift するため不採用）、非 float fallback も scale-1 化、VRT 19/19 byte-identical。**継続: table → textbox（B5 吸収 = kinsoku/bidi/justify が textbox で効く挙動変更、reference 承認が絡む）→ 相境界での物理分割（B9/B10 のテスト再配置はここで）**
- [ ] B5/B8/B9/B10: B2 後続段に吸収（textbox 統合 / stamp コントラクト最終形 / `__test_*` 廃止・テスト再配置は物理分割時）— 上記継続項目
- [x] D12 (#686): pptx parser lib.rs 14,107 → 5,816 行 + 8 モジュール（types/markdown/chart/theme/fill/text/shape/master、xlsx 範型）。verbatim 9 commit、golden sha256 4 ファイル byte-identical、MasterBundle→ParsedMaster rename。テスト 109 は cross-module 統合のため lib.rs 残置（将来分割候補: 低レベル単体のみ）
- [x] A1 (#685): `computeChartFrame` + ファミリー毎 verbatim 定数の FrameParams（ドリフト表）、凡例実測化、axis/gridline painter 共有
- [x] A2 (#687): **数値照合ハーネス preset-parity.test.ts**（WHATWG canvas 意味論・≤0.35px サンプル・fill 走査線・157 live labels×3box×2adj）で証明できた 33 preset + 3 alias のみ spec エンジンへ委譲、VRT 152 枚バイト一致。残 120 labels は幾何近似 68（spec 側が正: legacy flowchartProcess の誤角丸等）/ 構造差 51 / arc 意図的、として switch 内に分類コメント。**将来バッチ（fill-equivalent 6 種 → 微小幾何差 → 可視差）は見た目が変わるため VRT reference 更新承認とセット**
- [x] C9: **見送り（実測判断）** — Phase 2 最適化後の実測で 209MB pptx の全スライドパース 19.4ms / 14MB 13.2ms（median）。lazy 化が節約できる時間に対し、meta 集計 4 消費者（render-worker meta / media-mime / google-fonts / MCP server）の再設計コストが見合わない。bitmap LRU は導入済み（core/image/bitmap-image-by-path.ts — 計画の誤診）
- [x] E4 (#681): `.wasm` 実アセット化（wasmAssetUrl プラグイン **apply:'build' 必須**（dev に emitFile 不在 — CI smoke が捕捉）、`wasmUrl`/`workerTimeoutMs` オプション、data-URL fallback 維持）。チャンク -85〜89%。worker チャンク 3 重複は誤認と実証
- [x] E9: **調査完了・ユーザー判断待ち** — フォント同梱でも macOS CoreText と Linux FreeType の rasterizer 差で pixel-perfect 可搬化は原理的に不可。案 A: 現状維持（コスト 0、private 100% カバー維持、推奨）/ 案 B: Noto 同梱 + demo のみ CI-VRT（2-3h + repo 6-8MB、検出力 demo 7-29%、reference 全置換承認必要）/ 案 C: 折衷（CI に甘い smoke VRT）
- [x] E10 (#688): site 0.43.0 / mcp-server crate 0.1.0 → 0.69.0 系列へ（rmcp は CARGO_PKG_VERSION を initialize handshake で MCP クライアントに報告 = ユーザー可視ドリフトだった）、CLAUDE.md 手順 5 拡張。mathjax 3MB bundle を gitignore 化し core `prepare` で `pnpm install` 時自動生成（全 CI ジョブ検証済み。配布は従来どおり math.mjs に inline — 将来候補: E4 同様の実アセット化）
- [x] B12 帰結（#690, #691, #676 コメント）: **#675 ✅ #690** frame keep-with-anchor（§17.3.1.11 は寸法のみ規定 — 既存 anchored-image float ガード §20.4.3.5 と同一意味論で明示化。vAnchor='text' のみ送り、parser 既定 'page' は据え置き）。**#674 ✅ #691** floating-table page-fit（frame と対称実装、compute-once 維持）。**#676 = 文書化済み 1-em 維持が正解と判断** — 閾値は non-empty 行と対称で一貫、Word 実規則（固定 em / 行高比率 / 字形幅 / metrics）の判別には 3 フォントサイズ×3 行間の検証 docx + Word PDF が必要（issue コメントに要件記録、blocked on validation）。#675/#674 のページ境界実サンプル検証も fixture 待ち（合成 vitest + red-check では検証済み）

## 推奨 PR 分割（Phase 1 の目安）

1. `ci/rendering-smoke` — E1 + E3 + E8（安全網を最優先で 1 本）
2. `fix/viewer-destroy` — C5 + C6（ユーザー可視のライフサイクルバグ）
3. `fix/core-dash-table` — A6（描画差の出る実バグ、VRT 必須）
4. `fix/parser-error-api` — D11 + D9(unwrap) + D7（Rust エラー/zip 衛生をまとめて）
5. `perf/zip-and-sheet-transfer` — D1 + C1 + D10(フラグ)（ベンチ添付）
6. `perf/hot-path-small` — C3 + B6 + A8(前半) + C11
7. `chore/packaging` — E2 + E5 + E6（npm 消費側の検証を伴うため独立）
