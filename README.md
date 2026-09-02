# VBA DnsAPI Query Sample

Excel VBAからWindows標準のDNS APIを直接呼び出し、名前解決の流れを観察する日本語教材です。

`nslookup.exe`、PowerShell、`WScript.Shell`は使いません。VBAから`DnsQuery_W`を呼び出し、A、AAAA、CNAME、PTR、MX、TXTレコードを読み取ります。処理結果だけでなく、問い合わせ開始、Windowsリゾルバーの応答、レコード解析、メモリ解放までを`Trace`シートへ記録します。

> [!IMPORTANT]
> このリポジトリはWindows版Excel・VBA 7向けです。Mac版ExcelではWindows DNS APIを呼び出せません。

## この教材で分かること

- DNSが「名前からIPアドレスを探す」だけではないこと
- AとAAAA、CNAME、PTR、MX、TXTの役割
- VBAからUnicode対応のWindows APIを呼ぶ方法
- 32bit版Officeと64bit版Officeでポインター幅が異なること
- OSが返す連結リストと共用体を、安全側に倒して読む方法
- OSが確保したメモリを、エラー時も含めて必ず解放する理由
- TraceIdで1回の問い合わせを追跡する方法

## クライアントと相手側

このサンプルはDNSサーバーを自作しません。VBAがクライアント、Windows標準のDNSリゾルバーが相手側です。

```text
Excel VBA
  │ DnsQuery_W("localhost", A)
  ▼
Windows DNSリゾルバー
  ├─ hosts
  ├─ DNSキャッシュ
  └─ PCに設定されたDNSサーバー（必要な場合）
  │ DNS_RECORDWの連結リスト
  ▼
Excel VBA
  ├─ レコード解析
  ├─ Trace / DnsResultsシートへ表示
  └─ DnsRecordListFree相当でメモリ解放
```

`localhost`から始められるので、最初のテストでは外部のDNSサーバーやAPIキーを用意する必要がありません。ただし、Windowsの設定・hosts・組織のネットワークポリシーによって結果は変わります。

## 対応するレコード

| 種類 | 主な意味 | 表示例 |
|---|---|---|
| A | 名前に対応するIPv4アドレス | `127.0.0.1` |
| AAAA | 名前に対応するIPv6アドレス | `::1` |
| CNAME | 別名から正式名への対応 | `www.example.net` |
| PTR | IPアドレス側から名前を調べる逆引き | `localhost` |
| MX | メールを受け取るサーバーと優先度 | `preference=10, exchange=...` |
| TXT | ドメインに関連付けられた文字列 | SPFなどの文字列 |

## 必要な環境

- Windows 10またはWindows 11
- Windows版Excel
- VBA 7（32bit版Officeまたは64bit版Office）
- マクロを実行できる、信頼済みの学習用ブック

外部レコードを調べる任意テストには、ネットワーク接続と組織の許可が必要です。

## 始め方

1. 空のExcelブックを作成します。
2. `.xlsx`ではなく「Excelマクロ有効ブック（`.xlsm`）」として保存します。
3. `Alt + F11`でVisual Basic Editorを開きます。
4. `ファイル`→`ファイルのインポート`で、`src`内の3つの`.bas`ファイルを読み込みます。
5. `デバッグ`→`VBAProjectのコンパイル`を実行します。
6. Excelへ戻り、`Alt + F8`から`DNS_初回セットアップ`を実行します。
7. 続けて`DNS_ローカルテスト`を実行します。
8. `Trace`シートと`DnsResults`シートを確認します。

### 日本語コメントが文字化けした場合

VBEの`.bas`インポートは、OfficeやWindowsの言語設定によりUTF-8の日本語を正しく読めない場合があります。その場合は、UTF-8対応エディターでファイルを開き、モジュールの内容をVBEへ貼り付けてください。コード自体より先に、日本語コメントが読めることを確認してください。

## 用意した実行マクロ

| マクロ | 用途 |
|---|---|
| `DNS_初回セットアップ` | Traceと結果シートを準備し、過去の行を消す |
| `DNS_ローカルテスト` | `localhost`のA/AAAAとPC名のAを調べる |
| `DNS_名前と種類を指定して検索` | 名前とレコード種類を入力して調べる |
| `DNS_IPv4逆引きテスト` | `127.0.0.1`をPTR用の名前へ変換して問い合わせる |
| `DNS_外部レコード例_任意実行` | `example.com`の各レコードを任意で試す |

## トレースの読み方

`Trace`シートでは、1回の問い合わせに同じ`TraceId`が付きます。

| Time | TraceId | Side | Step | Direction | Detail | Result | ElapsedMs |
|---|---|---|---|---|---|---|---:|
| 21:10:01.120 | DNS-... | CLIENT | VALIDATE | LOCAL | `name=localhost, type=A` | START | 0 |
| 21:10:01.122 | DNS-... | CLIENT | QUERY | OUT | `DnsQuery_W ...` | CALL | 2 |
| 21:10:01.130 | DNS-... | RESOLVER | RESOLVE | LOCAL | Windows resolver/cache/hosts | DNS_STATUS=0 | 10 |
| 21:10:01.132 | DNS-... | CLIENT | PARSE_RECORD | IN | returnedType=A | 127.0.0.1 | 12 |
| 21:10:01.133 | DNS-... | CLIENT | FREE | LOCAL | record list | OK | 13 |

同じ行はVBEのイミディエイトウィンドウにも出ます。VBEで`Ctrl + G`を押すと表示できます。

## 重要な実装ポイント

### なぜ`DnsQuery_W`なのか

末尾の`W`はUnicode版を意味します。VBAの文字列を`StrPtr`で渡し、日本語を含むWindowsの文字列形式と合わせています。

### なぜ構造体をそのままVBAで宣言しないのか

`DNS_RECORDW`の末尾は、レコード種類によって中身が変わるC言語の`union`です。VBAには同じ共用体機能がありません。本教材では、Microsoftの公開構造に従い、ポインター幅別のオフセットから必要なフィールドだけを読みます。

### 32bitと64bitで何が変わるのか

ポインターは32bit版Officeで4バイト、64bit版Officeで8バイトです。`LongPtr`と`#If Win64 Then`を使い、`pNext`、`pName`、レコードデータの開始位置を切り替えています。

TXTでは、メモリ上の物理配置と`wDataLength`の論理長を分けます。文字列ポインター配列は32bitではoffset 4、64bit C ABIではパディング後のoffset 8から読みます。一方、長さ検証はMicrosoft公式の`DNS_TEXT_RECORD_LENGTH`どおり、`4 + 文字列数 × ポインター幅`で計算します。

### なぜ最後にメモリを解放するのか

`DnsQuery_W`の成功時に返るレコードリストはWindows側で確保されます。使い終わったら`DnsRecordListFree`相当の処理が必要です。本教材ではWindows SDKのマクロの実体である`DnsFree`を、目的が分かる宣言名で呼び、`DnsFreeRecordList`を指定します。

## ファイル構成

```text
VBA_DnsAPI_Query_Sample/
├─ README.md
├─ SECURITY.md
├─ src/
│  ├─ modDnsApi.bas       Windows API宣言、レコード解析、メモリ解放
│  ├─ modDnsTrace.bas     Trace/DnsResultsシート、Debug.Print
│  └─ modDnsSamples.bas   初心者向けの実行入口
└─ docs/
   ├─ TESTING.md          実機での確認手順と期待結果
   └─ REVIEW.md           静的レビューの観点と既知の限界
```

## 制限事項

- 同期版の`DnsQuery_W`を使用します。非同期DNS APIの教材ではありません。
- DNSサーバーを新規実装するサンプルではありません。
- DNSSECの検証結果を表示するサンプルではありません。
- IDNの日本語ドメインを入力する場合のPunycode変換は実装していません。
- PTRレコード、PC名、CNAME、MX、TXTの有無は環境と対象ドメインに依存します。
- このリポジトリだけでは、32bit/64bit両方のExcel実機試験を代替できません。

実機試験の詳細は[`docs/TESTING.md`](docs/TESTING.md)を参照してください。

## 公式資料

- [DnsQuery_W function - Microsoft Learn](https://learn.microsoft.com/windows/win32/api/windns/nf-windns-dnsquery_w)
- [DNS_RECORDW structure - Microsoft Learn](https://learn.microsoft.com/windows/win32/api/windnsdef/ns-windnsdef-dns_recordw)
- [DnsRecordListFree macro - Microsoft Learn](https://learn.microsoft.com/windows/win32/api/windns/nf-windns-dnsrecordlistfree)
- [DNS_FREE_TYPE enumeration - Microsoft Learn](https://learn.microsoft.com/windows/win32/api/windns/ne-windns-dns_free_type)
- [InetNtopW function - Microsoft Learn](https://learn.microsoft.com/windows/win32/api/ws2tcpip/nf-ws2tcpip-inetntopw)
