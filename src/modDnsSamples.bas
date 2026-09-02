Attribute VB_Name = "modDnsSamples"
Option Explicit

' =============================================================================
' modDnsSamples.bas
' -----------------------------------------------------------------------------
' 初心者がF5キーですぐ試せる入口をまとめたモジュールです。
' まず「DNS_初回セットアップ」、次に「DNS_ローカルテスト」を実行してください。
' =============================================================================

' TraceシートとDnsResultsシートを作成し、古い結果を消します。
Public Sub DNS_初回セットアップ()
    DnsClearTrace
    DnsClearResults

    MsgBox "準備が完了しました。" & vbCrLf & _
           "次に「DNS_ローカルテスト」を実行してください。", _
           vbInformation, "DNS API Sample"
End Sub

' 外部インターネットへ依存しにくい、最初の動作確認です。
' 相手側は別途作るDNSサーバーではなく、Windows標準のDNSリゾルバーです。
' Windowsはhosts、キャッシュ、設定済みDNSサーバーなどを順に利用します。
Public Sub DNS_ローカルテスト()
    Dim computerName As String

    DnsClearTrace
    DnsClearResults

    RunAndPrint "localhost", "A"
    RunAndPrint "localhost", "AAAA"

    ' COMPUTERNAMEは通常、Windowsに設定されたPC名です。
    ' 組織のDNS登録状況によってはAレコードが見つからないこともあります。
    computerName = Trim$(Environ$("COMPUTERNAME"))
    If Len(computerName) > 0 Then RunAndPrint computerName, "A"

    MsgBox "ローカルテストが終了しました。" & vbCrLf & _
           "Traceシートで処理順を、DnsResultsシートで回答を確認してください。" & _
           vbCrLf & vbCrLf & _
           "注意: PC名の問い合わせ結果はネットワーク環境により異なります。", _
           vbInformation, "DNS API Sample"
End Sub

' 問い合わせ名と種類を入力して自由に試すサンプルです。
Public Sub DNS_名前と種類を指定して検索()
    Dim queryName As String
    Dim typeName As String

    queryName = InputBox( _
        "問い合わせる名前を入力してください。" & vbCrLf & _
        "例: localhost / example.com", _
        "DNS問い合わせ名", "localhost")
    If Len(Trim$(queryName)) = 0 Then Exit Sub

    typeName = InputBox( _
        "種類を入力してください。" & vbCrLf & _
        "A / AAAA / CNAME / PTR / MX / TXT", _
        "DNSレコード種類", "A")
    If Len(Trim$(typeName)) = 0 Then Exit Sub

    RunAndPrint queryName, typeName
End Sub

' 127.0.0.1のPTR（逆引き）を試します。
' PTR登録は環境依存なので、レコードなしでもコードの異常とは限りません。
Public Sub DNS_IPv4逆引きテスト()
    Dim answers As Collection
    Dim item As Variant

    Set answers = DnsQueryPtrFromIPv4("127.0.0.1")
    For Each item In answers
        Debug.Print CStr(item)
    Next item
End Sub

' 外部DNSを利用できる環境で、各種類の入口をまとめて試す任意テストです。
' 学校・会社のネットワークポリシーに従い、許可された環境だけで実行してください。
Public Sub DNS_外部レコード例_任意実行()
    RunAndPrint "example.com", "A"
    RunAndPrint "example.com", "AAAA"
    RunAndPrint "example.com", "MX"
    RunAndPrint "example.com", "TXT"
End Sub

Private Sub RunAndPrint(ByVal queryName As String, ByVal typeName As String)
    Dim answers As Collection
    Dim item As Variant

    Set answers = DnsQueryRecords(queryName, typeName)

    If answers.Count = 0 Then
        Debug.Print "回答なし: " & queryName & " / " & typeName & _
                    "（詳しくはTraceシートを確認）"
        Exit Sub
    End If

    For Each item In answers
        Debug.Print CStr(item)
    Next item
End Sub
