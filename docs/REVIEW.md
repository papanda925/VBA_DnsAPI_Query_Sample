# レビュー記録と観点

## レビューの数え方

下記10種類の利用者視点それぞれに、次の10テーマを適用し、`10ペルソナ × 10テーマ = 100チェックポイント`として静的レビューしました。

1. 初回導入と実行入口
2. DNS用語と教材としての説明
3. VBA7／Office 32bitのAPI型
4. VBA7／Office 64bitのポインターと配置
5. A／AAAA／CNAME／PTR／MX／TXTの解析
6. 入力、件数、長さ、ポインター境界
7. 正常・DNSエラー・VBA例外でのメモリ解放
8. TraceIdと結果表の追跡性
9. 数式注入、外部問い合わせ、ログ情報の安全性
10. 日本語、文字コード、試験手順、未確認事項

これは実在する100人のレビューやWindows Excelで100回実行した結果ではありません。確認可能な静的事項と、実機でのみ確認できる事項を分けるためのレビュー表です。

## 現時点の結論

ソースはMicrosoft Learnに公開された`DnsQuery_W`、`DNS_RECORDW`、各レコードデータ構造、`DnsRecordListFree`、`InetNtopW`の仕様を基準に静的確認しています。

この作成環境にはWindows版Excelがないため、VBEでのコンパイルとDNS APIの実行は未確認です。したがって、現段階を「実機試験済み」とは扱いません。公開前に[`TESTING.md`](TESTING.md)の32bit・64bit確認を行う必要があります。

## 10の利用者視点

| 視点 | 主な確認 | 反映した内容 |
|---|---|---|
| VBAを初めて触る学生 | 最初に何を実行するか分かるか | セットアップとローカルテストを分離 |
| 情報系の学生 | DNSの要求・応答を追えるか | TraceId、Side、Step、Directionを記録 |
| 学校の教員 | 授業で説明可能か | 用語、構造、公式資料へのリンクを追加 |
| 社内SE初心者 | 外部コマンド依存がないか | DnsQuery_Wを直接利用 |
| 32bit Office利用者 | LongLong依存がないか | LongPtrと条件付きオフセットを採用 |
| 64bit Office利用者 | ポインターが切り捨てられないか | API引数・戻り値・読み取りをLongPtr化 |
| セキュリティ担当 | DNS値が実行されないか | セルを文字列形式にし、自動実行なし |
| 障害解析担当 | 失敗箇所が分かるか | DNS_STATUS、エラー、メモリ解放を記録 |
| 保守担当 | API仕様の根拠が追えるか | READMEにMicrosoft Learnを列挙 |
| 上級VBA開発者 | 境界・解放・共用体が妥当か | 長さ検証、件数上限、finally相当の解放 |

## ポインターと構造体の静的確認

### DNS_RECORDW

| フィールド | 32bit offset | 64bit offset | 読み取り |
|---|---:|---:|---|
| pNext | 0 | 0 | LongPtr |
| pName | 4 | 8 | LongPtr |
| wType | 8 | 16 | unsigned 16bit相当 |
| wDataLength | 10 | 18 | unsigned 16bit相当 |
| dwTtl | 16 | 24 | unsigned 32bit相当をDoubleへ |
| Data | 24 | 32 | 種類別に解析 |

### 種類別データ

- A: Data先頭4バイトをネットワーク表現の4オクテットとして表示。
- AAAA: Data先頭16バイトを`InetNtopW(AF_INET6)`へ渡す。
- CNAME/PTR: Data先頭の`PWSTR`をUnicode文字列として読む。
- MX: `PWSTR`の直後にある`WORD wPreference`を読む。末尾のPadは使用しない。
- TXT: 先頭DWORDを文字列数とし、ポインター配列を32bitではoffset 4、64bitではC ABIのパディング後となるoffset 8から読む。

TXTの`wDataLength`検証は、Microsoft公式の`DNS_TEXT_RECORD_LENGTH`マクロ（`sizeof(DWORD) + StringCount * sizeof(PCHAR)`）に合わせ、`4 + ポインター数×幅`で計算します。これは物理配置のoffset 8とは別の「論理長」です。

## メモリ解放

- `DnsQuery_W`成功後の先頭ポインターを`resultList`に保持。
- 解析中は`currentRecord`だけを進め、解放用の先頭を失わない。
- 32bitの高位アドレスで符号付きLongの加算がオーバーフローしないよう、ポインター加算を専用関数へ集約。
- 正常終了、DNSエラー、VBA実行時エラーのすべてを`CleanUp`へ集約。
- `resultList <> 0`のときだけ`DnsFree(..., DnsFreeRecordList)`を呼ぶ。
- 解放後に`resultList = 0`として二重解放を避ける。

## 静的レビューのチェックリスト

- [x] `Option Explicit`が全モジュールにある。
- [x] `Shell`、PowerShell、`nslookup.exe`を使用していない。
- [x] APIポインターを`Long`へ格納していない。
- [x] 32bit/64bitでDNS_RECORDWのオフセットを分岐している。
- [x] 32bitの高位アドレスを考慮したポインター加算を使用している。
- [x] A/AAAA/CNAME/PTR/MX/TXTを実装している。
- [x] 実際に返った`wType`を見てDataを解析している。
- [x] `wDataLength`の最小長を検証している。
- [x] レコード数、TXT数、Unicode文字数に上限がある。
- [x] `pNext`自己参照と総レコード上限がある。
- [x] 先頭ポインターを保持し、すべての終了経路で解放する。
- [x] TraceIdで問い合わせ全体を結び付けている。
- [x] Traceと結果の文字列を数式として評価させない。
- [x] ローカルで始められるテストと環境差の説明がある。
- [x] 実機未確認であることを明記している。

## クロスレビューで修正した事項

- 32bit／64bit Office保守者の観点からTXT配置をWindows SDK定義と再照合。
- 64bit版の`DNS_TXT_DATAW.pStringArray`はC ABIのパディング後となるoffset 8、32bit版はoffset 4として分岐。
- `wDataLength`は物理offsetではなく、Microsoft Learnの`DNS_TEXT_RECORD_LENGTH`に従う論理長（`4 + count×pointer-size`）で検証。
- Trace／DnsResultsシートへ書けない場合も、DNS問い合わせとCollection構築を継続するよう分離。
- Traceはシート操作より先にイミディエイトウィンドウへ出し、シート保護時の診断消失を防止。

## 公開前に残る必須確認

1. 32bit版Excel VBA 7でのコンパイルと全種類テスト。
2. 64bit版Excel VBA 7でのコンパイルと全種類テスト。
3. 複数文字列を含むTXTレコードの確認。
4. CNAMEを含む応答で、返却リストの複数種類を正しく処理できること。
5. DNSエラー、タイムアウト、レコードなしでExcelが安定すること。
6. 連続問い合わせ後にメモリが解放されること。

レビュー回数を増やすことは、上記の実機証拠の代わりにはなりません。各指摘を記録し、再現可能なテスト結果に結び付けることを優先します。
