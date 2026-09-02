Attribute VB_Name = "modDnsTrace"
Option Explicit

' =============================================================================
' modDnsTrace.bas
' -----------------------------------------------------------------------------
' DNS問い合わせの「いつ・どちら側で・何が起きたか」を記録するモジュールです。
'
' Traceシートの列:
'   Time / TraceId / Side / Step / Direction / Detail / Result / ElapsedMs
'
' Debug.Printにも同じ内容を出すため、VBEのイミディエイトウィンドウ
' （Ctrl + G）でも処理の流れを追えます。
' =============================================================================

Private Const TRACE_SHEET_NAME As String = "Trace"
Private Const RESULT_SHEET_NAME As String = "DnsResults"
Private mTraceSequence As Long

' 1回の問い合わせをクライアント側とリゾルバー側で結び付けるIDを作ります。
' 同じ名前を連続して問い合わせても、TraceIdが異なるので混同しません。
Public Function DnsNewTraceId() As String
    mTraceSequence = mTraceSequence + 1

    DnsNewTraceId = "DNS-" & _
                    Format$(Now, "yyyymmdd-hhnnss") & "-" & _
                    Format$(CLng(Fix((Timer - Int(Timer)) * 1000#)), "000") & "-" & _
                    Format$(mTraceSequence Mod 10000, "0000")
End Function

' 経過時間の基準値を返します。Timerは、その日の午前0時からの秒数です。
Public Function DnsTimerStart() As Double
    DnsTimerStart = Timer
End Function

' Timerは午前0時に0へ戻るため、日付をまたいだ場合だけ1日分を補正します。
Public Function DnsElapsedMilliseconds(ByVal startedAt As Double) As Long
    Dim seconds As Double

    seconds = Timer - startedAt
    If seconds < 0# Then seconds = seconds + 86400#

    If seconds > 2147483# Then
        DnsElapsedMilliseconds = 2147483647
    Else
        DnsElapsedMilliseconds = CLng(seconds * 1000#)
    End If
End Function

' Traceシートとイミディエイトウィンドウへ、同じ1行を記録します。
' startedAtには問い合わせ開始時のDnsTimerStartの戻り値を渡します。
Public Sub DnsTraceLog( _
    ByVal traceId As String, _
    ByVal side As String, _
    ByVal stepName As String, _
    ByVal direction As String, _
    ByVal detail As String, _
    ByVal resultText As String, _
    ByVal startedAt As Double)

    Dim ws As Worksheet
    Dim nextRow As Long
    Dim timeText As String
    Dim elapsedMs As Long
    Dim traceErrorNumber As Long
    Dim traceErrorDescription As String

    elapsedMs = DnsElapsedMilliseconds(startedAt)
    timeText = DnsTraceTimeText()

    ' シートが保護されている場合も診断情報を失わないよう、Debug.Printを先に出す。
    Debug.Print timeText & vbTab & _
                traceId & vbTab & _
                side & vbTab & _
                stepName & vbTab & _
                direction & vbTab & _
                DnsTraceSafeText(detail) & vbTab & _
                DnsTraceSafeText(resultText) & vbTab & _
                CStr(elapsedMs) & " ms"

    ' Traceシートへ書けないことをDNS API自体の失敗として扱わない。
    On Error GoTo SheetUnavailable
    Set ws = EnsureTraceSheet()
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' 文字列として書き込みます。問い合わせ名やTXT値が「=」で始まっても、
    ' Excelの数式として評価されないよう、対象列を文字列形式にしています。
    ws.Range(ws.Cells(nextRow, 1), ws.Cells(nextRow, 7)).NumberFormat = "@"
    ws.Cells(nextRow, 1).Value2 = timeText
    ws.Cells(nextRow, 2).Value2 = DnsTraceSafeText(traceId)
    ws.Cells(nextRow, 3).Value2 = DnsTraceSafeText(side)
    ws.Cells(nextRow, 4).Value2 = DnsTraceSafeText(stepName)
    ws.Cells(nextRow, 5).Value2 = DnsTraceSafeText(direction)
    ws.Cells(nextRow, 6).Value2 = DnsTraceSafeText(detail)
    ws.Cells(nextRow, 7).Value2 = DnsTraceSafeText(resultText)
    ws.Cells(nextRow, 8).Value2 = elapsedMs

    Exit Sub

SheetUnavailable:
    traceErrorNumber = Err.Number
    traceErrorDescription = Err.Description
    Debug.Print DnsTraceTimeText() & vbTab & _
                "TRACE-WARNING" & vbTab & _
                "Traceシートへ記録できません: " & CStr(traceErrorNumber) & _
                " / " & traceErrorDescription
End Sub

' 問い合わせ結果を、人が読みやすい表としてDnsResultsシートへ追加します。
Public Sub DnsAppendResult( _
    ByVal traceId As String, _
    ByVal queryName As String, _
    ByVal requestedType As String, _
    ByVal returnedType As String, _
    ByVal ownerName As String, _
    ByVal ttlSeconds As Double, _
    ByVal recordValue As String)

    Dim ws As Worksheet
    Dim nextRow As Long
    Dim resultErrorNumber As Long
    Dim resultErrorDescription As String

    ' 結果シートの問題で、取得済みCollectionや残りレコードの解析を止めない。
    On Error GoTo SheetUnavailable
    Set ws = EnsureResultSheet()
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Range(ws.Cells(nextRow, 1), ws.Cells(nextRow, 5)).NumberFormat = "@"
    ws.Cells(nextRow, 7).NumberFormat = "@"
    ws.Cells(nextRow, 1).Value2 = DnsTraceSafeText(traceId)
    ws.Cells(nextRow, 2).Value2 = DnsTraceSafeText(queryName)
    ws.Cells(nextRow, 3).Value2 = DnsTraceSafeText(requestedType)
    ws.Cells(nextRow, 4).Value2 = DnsTraceSafeText(returnedType)
    ws.Cells(nextRow, 5).Value2 = DnsTraceSafeText(ownerName)
    ws.Cells(nextRow, 6).Value2 = ttlSeconds
    ws.Cells(nextRow, 7).Value2 = DnsTraceSafeText(recordValue)

    Exit Sub

SheetUnavailable:
    resultErrorNumber = Err.Number
    resultErrorDescription = Err.Description
    Debug.Print DnsTraceTimeText() & vbTab & _
                "RESULT-WARNING" & vbTab & _
                "DnsResultsシートへ記録できません: " & CStr(resultErrorNumber) & _
                " / " & resultErrorDescription
End Sub

' 明示的に実行したときだけ過去のトレースを消します。
' 問い合わせのたびに自動消去しないのは、複数回の結果を比較できるようにするためです。
Public Sub DnsClearTrace()
    Dim ws As Worksheet

    Set ws = EnsureTraceSheet()
    ws.Rows("2:" & CStr(ws.Rows.Count)).ClearContents
End Sub

Public Sub DnsClearResults()
    Dim ws As Worksheet

    Set ws = EnsureResultSheet()
    ws.Rows("2:" & CStr(ws.Rows.Count)).ClearContents
End Sub

Private Function EnsureTraceSheet() As Worksheet
    Dim ws As Worksheet

    Set ws = FindOrCreateSheet(TRACE_SHEET_NAME)

    If Len(CStr(ws.Cells(1, 1).Value2)) = 0 Then
        ws.Range("A1:H1").Value = Array( _
            "Time", "TraceId", "Side", "Step", _
            "Direction", "Detail", "Result", "ElapsedMs")
        ws.Rows(1).Font.Bold = True
        ws.Columns("A:G").NumberFormat = "@"
        ws.Columns("A:H").AutoFit
        ws.Columns("F:G").ColumnWidth = 48
        ws.Range("A1:H1").AutoFilter
    End If

    Set EnsureTraceSheet = ws
End Function

Private Function EnsureResultSheet() As Worksheet
    Dim ws As Worksheet

    Set ws = FindOrCreateSheet(RESULT_SHEET_NAME)

    If Len(CStr(ws.Cells(1, 1).Value2)) = 0 Then
        ws.Range("A1:G1").Value = Array( _
            "TraceId", "QueryName", "RequestedType", "ReturnedType", _
            "OwnerName", "TTLSeconds", "Value")
        ws.Rows(1).Font.Bold = True
        ws.Columns("A:E").NumberFormat = "@"
        ws.Columns("G:G").NumberFormat = "@"
        ws.Columns("A:G").AutoFit
        ws.Columns("E:E").ColumnWidth = 32
        ws.Columns("G:G").ColumnWidth = 60
        ws.Range("A1:G1").AutoFilter
    End If

    Set EnsureResultSheet = ws
End Function

Private Function FindOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set FindOrCreateSheet = ws
End Function

Private Function DnsTraceTimeText() As String
    Dim milliseconds As Long

    ' CLngは四捨五入するため、0.9996秒が1000になる可能性があります。
    ' Fixで切り捨て、常に000～999の3桁へ収めます。
    milliseconds = CLng(Fix((Timer - Int(Timer)) * 1000#))
    DnsTraceTimeText = Format$(Now, "yyyy-mm-dd hh:nn:ss") & "." & _
                       Format$(milliseconds, "000")
End Function

' 改行を1行表示用の文字へ変換し、Excelセルの最大長より少し短く切ります。
' 生の制御文字や巨大なTXTレコードでTraceシートを壊さないための処理です。
Private Function DnsTraceSafeText(ByVal value As String) As String
    value = Replace(value, vbCr, "\r")
    value = Replace(value, vbLf, "\n")
    value = Replace(value, vbTab, "\t")

    If Len(value) > 30000 Then
        value = Left$(value, 30000) & "...(省略)"
    End If

    DnsTraceSafeText = value
End Function
