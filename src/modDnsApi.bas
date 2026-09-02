Attribute VB_Name = "modDnsApi"
Option Explicit

' =============================================================================
' modDnsApi.bas
' -----------------------------------------------------------------------------
' Windows標準のDNS APIをExcel VBAから直接呼び出す学習用モジュールです。
'
' 重要:
'   * nslookup.exeやPowerShell、WScript.Shellは使用しません。
'   * DnsQuery_Wを使うため、問い合わせ名はUnicodeで渡します。
'   * 32bit版Officeと64bit版Officeの両方で使えるようLongPtrを使用します。
'   * DnsQuery_Wが確保したメモリは、成功後に必ずDnsRecordListFree相当の
'     DnsFree(..., DnsFreeRecordList)で解放します。
'   * DNS_RECORDWは共用体を含み、VBAのユーザー定義型で完全再現しにくいため、
'     公式構造体のオフセットに従い、必要なフィールドだけ慎重に読み取ります。
' =============================================================================

#If VBA7 Then
    Private Declare PtrSafe Function DnsQuery_W Lib "dnsapi.dll" ( _
        ByVal pszName As LongPtr, _
        ByVal wType As Integer, _
        ByVal Options As Long, _
        ByVal pExtra As LongPtr, _
        ByRef ppQueryResults As LongPtr, _
        ByVal pReserved As LongPtr) As Long

    ' DnsRecordListFreeはWindows SDK上ではマクロです。
    ' 実体のDnsFreeを、教材内では目的が分かる名前で宣言しています。
    Private Declare PtrSafe Sub DnsRecordListFree Lib "dnsapi.dll" Alias "DnsFree" ( _
        ByVal pRecordList As LongPtr, _
        ByVal freeType As Long)

    Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
        ByVal destination As LongPtr, _
        ByVal source As LongPtr, _
        ByVal byteCount As LongPtr)

    Private Declare PtrSafe Function lstrlenW Lib "kernel32.dll" ( _
        ByVal textPointer As LongPtr) As Long

    Private Declare PtrSafe Function InetNtopW Lib "ws2_32.dll" ( _
        ByVal addressFamily As Long, _
        ByVal addressPointer As LongPtr, _
        ByVal stringBuffer As LongPtr, _
        ByVal stringBufferCharacters As LongPtr) As LongPtr
#End If

' DNS record type。値はwindns.hで定義されている「ホストバイト順」の値です。
Private Const DNS_TYPE_A As Long = 1
Private Const DNS_TYPE_CNAME As Long = 5
Private Const DNS_TYPE_PTR As Long = 12
Private Const DNS_TYPE_MX As Long = 15
Private Const DNS_TYPE_TEXT As Long = 16
Private Const DNS_TYPE_AAAA As Long = 28

Private Const DNS_QUERY_STANDARD As Long = 0
Private Const DNS_FREE_RECORD_LIST As Long = 1
Private Const AF_INET6 As Long = 23

' 壊れたポインターリストで無限ループしないための教育用上限です。
' 通常のDNS応答がこの件数へ達することはほとんどありません。
Private Const MAX_RECORDS_PER_QUERY As Long = 512
Private Const MAX_TXT_STRINGS As Long = 256
Private Const MAX_UNICODE_CHARS As Long = 32767

' DNS_RECORDWの配置はポインター幅で変わります。
' 64bit: pNext(8), pName(8), WORD(2), WORD(2), DWORD×3 → Dataは32バイト目
' 32bit: pNext(4), pName(4), WORD(2), WORD(2), DWORD×3 → Dataは24バイト目
#If Win64 Then
    Private Const POINTER_SIZE As Long = 8
    Private Const OFFSET_RECORD_NAME As Long = 8
    Private Const OFFSET_RECORD_TYPE As Long = 16
    Private Const OFFSET_DATA_LENGTH As Long = 18
    Private Const OFFSET_TTL As Long = 24
    Private Const OFFSET_RECORD_DATA As Long = 32
#Else
    Private Const POINTER_SIZE As Long = 4
    Private Const OFFSET_RECORD_NAME As Long = 4
    Private Const OFFSET_RECORD_TYPE As Long = 8
    Private Const OFFSET_DATA_LENGTH As Long = 10
    Private Const OFFSET_TTL As Long = 16
    Private Const OFFSET_RECORD_DATA As Long = 24
#End If

' DNS_TXT_DATAWの物理配置とwDataLengthの論理長は分けて考えます。
' 64bit C ABIでは、DWORDの後に4バイトのパディングが入り、PWSTR配列は
' offset 8から始まります。32bitではoffset 4です。
' 一方、公式DNS_TEXT_RECORD_LENGTHはパディングを数えず
'   sizeof(DWORD) + StringCount * sizeof(PCHAR)
' と計算します。後段の長さ検証では、この論理長を別に計算します。
#If Win64 Then
    Private Const OFFSET_TXT_POINTER_ARRAY As Long = 8
#Else
    Private Const OFFSET_TXT_POINTER_ARRAY As Long = 4
#End If
Private Const TXT_LOGICAL_HEADER_BYTES As Long = 4

' 指定した名前と種類をDNSへ問い合わせます。
' 戻り値は「種類 | owner=... | ttl=... | value=...」形式のCollectionです。
' 同じ内容をDnsResultsシートへ列分けして書き込みます。
Public Function DnsQueryRecords( _
    ByVal queryName As String, _
    ByVal requestedTypeName As String) As Collection

    Dim answers As Collection
    Dim traceId As String
    Dim startedAt As Double
    Dim normalizedName As String
    Dim requestedType As Long
    Dim status As Long
    Dim resultList As LongPtr
    Dim currentRecord As LongPtr
    Dim nextRecord As LongPtr
    Dim recordCount As Long
    Dim actualType As Long
    Dim dataLength As Long
    Dim ownerPointer As LongPtr
    Dim ownerName As String
    Dim ttlSeconds As Double
    Dim recordValue As String
    Dim answerLine As String
    Dim errorNumber As Long
    Dim errorDescription As String

    Set answers = New Collection
    Set DnsQueryRecords = answers

    traceId = DnsNewTraceId()
    startedAt = DnsTimerStart()

    On Error GoTo QueryFailed

    normalizedName = Trim$(queryName)
    requestedTypeName = UCase$(Trim$(requestedTypeName))

    DnsTraceLog traceId, "CLIENT", "VALIDATE", "LOCAL", _
                "name=" & normalizedName & ", type=" & requestedTypeName, _
                "START", startedAt

    If Len(normalizedName) = 0 Then
        Err.Raise vbObjectError + 2101, "DnsQueryRecords", _
                  "問い合わせ名が空です。"
    End If

    requestedType = DnsTypeFromName(requestedTypeName)
    If requestedType = 0 Then
        Err.Raise vbObjectError + 2102, "DnsQueryRecords", _
                  "未対応の種類です。A / AAAA / CNAME / PTR / MX / TXT を指定してください。"
    End If

    DnsTraceLog traceId, "CLIENT", "QUERY", "OUT", _
                "DnsQuery_W name=" & normalizedName & _
                ", type=" & requestedTypeName, _
                "CALL", startedAt

    ' StrPtrはVBA文字列のUnicodeバッファを指します。
    ' DnsQuery_Wの末尾「_W」は、引数がUnicodeであることを表します。
    status = DnsQuery_W( _
        StrPtr(normalizedName), _
        CInt(requestedType), _
        DNS_QUERY_STANDARD, _
        0, _
        resultList, _
        0)

    DnsTraceLog traceId, "RESOLVER", "RESOLVE", "LOCAL", _
                "Windows DNS resolver/cache/hosts", _
                "DNS_STATUS=" & CStr(status) & " (" & DnsStatusText(status) & ")", _
                startedAt

    If status <> 0 Then
        DnsTraceLog traceId, "CLIENT", "COMPLETE", "IN", _
                    "name=" & normalizedName & ", type=" & requestedTypeName, _
                    "ERROR: " & DnsStatusText(status), startedAt
        GoTo CleanUp
    End If

    currentRecord = resultList

    Do While currentRecord <> 0
        recordCount = recordCount + 1
        If recordCount > MAX_RECORDS_PER_QUERY Then
            Err.Raise vbObjectError + 2103, "DnsQueryRecords", _
                      "DNSレコード数が安全上限を超えました。"
        End If

        actualType = ReadUInt16(currentRecord, OFFSET_RECORD_TYPE)
        dataLength = ReadUInt16(currentRecord, OFFSET_DATA_LENGTH)
        ownerPointer = ReadPointer(currentRecord, OFFSET_RECORD_NAME)
        ownerName = PointerToUnicodeString(ownerPointer)
        ttlSeconds = ReadUInt32AsDouble(currentRecord, OFFSET_TTL)

        recordValue = ParseRecordValue( _
            currentRecord, actualType, dataLength)

        answerLine = DnsTypeName(actualType) & _
                     " | owner=" & ownerName & _
                     " | ttl=" & Format$(ttlSeconds, "0") & _
                     " | value=" & recordValue
        answers.Add answerLine

        DnsAppendResult traceId, normalizedName, requestedTypeName, _
                        DnsTypeName(actualType), ownerName, ttlSeconds, recordValue

        DnsTraceLog traceId, "CLIENT", "PARSE_RECORD", "IN", _
                    "record=" & CStr(recordCount) & _
                    ", returnedType=" & DnsTypeName(actualType) & _
                    ", dataLength=" & CStr(dataLength), _
                    recordValue, startedAt

        nextRecord = ReadPointer(currentRecord, 0)
        If nextRecord = currentRecord Then
            Err.Raise vbObjectError + 2104, "DnsQueryRecords", _
                      "DNSレコードのpNextが自分自身を指しています。"
        End If
        currentRecord = nextRecord
    Loop

    DnsTraceLog traceId, "CLIENT", "COMPLETE", "IN", _
                "name=" & normalizedName & ", type=" & requestedTypeName, _
                CStr(recordCount) & " record(s)", startedAt

CleanUp:
    FreeDnsResultList resultList, traceId, startedAt

    Set DnsQueryRecords = answers
    Exit Function

QueryFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    DnsTraceLog traceId, "CLIENT", "ERROR", "LOCAL", _
                "Err.Number=" & CStr(errorNumber), _
                errorDescription, startedAt
    ' エラー処理中にさらにエラーが起きても、OSから受け取ったメモリの解放を
    ' 試みられるよう、この経路ではOn Error Resume Nextのまま後始末します。
    FreeDnsResultList resultList, traceId, startedAt
    Set DnsQueryRecords = answers
End Function

' DnsQuery_Wが成功して返したリストは、途中で解析エラーが起きても解放します。
' この処理を忘れると、問い合わせのたびにExcelプロセス内でメモリリークします。
Private Sub FreeDnsResultList( _
    ByRef resultList As LongPtr, _
    ByVal traceId As String, _
    ByVal startedAt As Double)

    If resultList = 0 Then Exit Sub

    DnsRecordListFree resultList, DNS_FREE_RECORD_LIST
    resultList = 0

    DnsTraceLog traceId, "CLIENT", "FREE", "LOCAL", _
                "DnsRecordListFree (DnsFreeRecordList)", _
                "OK", startedAt
End Sub

' IPv4アドレスをPTR問い合わせ用の逆引き名へ変換して検索します。
' 例: 127.0.0.1 → 1.0.0.127.in-addr.arpa
Public Function DnsQueryPtrFromIPv4(ByVal ipv4Address As String) As Collection
    Set DnsQueryPtrFromIPv4 = DnsQueryRecords( _
        IPv4ToReverseLookupName(ipv4Address), "PTR")
End Function

Public Function IPv4ToReverseLookupName(ByVal ipv4Address As String) As String
    Dim parts() As String
    Dim index As Long

    ipv4Address = Trim$(ipv4Address)
    parts = Split(ipv4Address, ".")

    If UBound(parts) - LBound(parts) + 1 <> 4 Then
        Err.Raise vbObjectError + 2110, "IPv4ToReverseLookupName", _
                  "IPv4アドレスは、192.0.2.1のように4つの10進数で指定してください。"
    End If

    For index = LBound(parts) To UBound(parts)
        If Not IsValidIPv4Octet(parts(index)) Then
            Err.Raise vbObjectError + 2111, "IPv4ToReverseLookupName", _
                      "IPv4アドレスの各数値は0～255で指定してください。"
        End If
    Next index

    IPv4ToReverseLookupName = _
        CStr(CLng(parts(3))) & "." & _
        CStr(CLng(parts(2))) & "." & _
        CStr(CLng(parts(1))) & "." & _
        CStr(CLng(parts(0))) & ".in-addr.arpa"
End Function

Private Function ParseRecordValue( _
    ByVal recordPointer As LongPtr, _
    ByVal recordType As Long, _
    ByVal dataLength As Long) As String

    Dim dataPointer As LongPtr
    Dim namePointer As LongPtr
    Dim preference As Long

    dataPointer = PointerAdd(recordPointer, OFFSET_RECORD_DATA)

    Select Case recordType
        Case DNS_TYPE_A
            RequireDataLength dataLength, 4, "A"
            ParseRecordValue = IPv4BytesToString(dataPointer)

        Case DNS_TYPE_AAAA
            RequireDataLength dataLength, 16, "AAAA"
            ParseRecordValue = IPv6BytesToString(dataPointer)

        Case DNS_TYPE_CNAME, DNS_TYPE_PTR
            RequireDataLength dataLength, POINTER_SIZE, DnsTypeName(recordType)
            namePointer = ReadPointer(dataPointer, 0)
            ParseRecordValue = PointerToUnicodeString(namePointer)

        Case DNS_TYPE_MX
            ' DNS_MX_DATAW = PWSTR pNameExchange + WORD wPreference + WORD Pad
            RequireDataLength dataLength, POINTER_SIZE + 4, "MX"
            namePointer = ReadPointer(dataPointer, 0)
            preference = ReadUInt16(dataPointer, POINTER_SIZE)
            ParseRecordValue = "preference=" & CStr(preference) & _
                               ", exchange=" & PointerToUnicodeString(namePointer)

        Case DNS_TYPE_TEXT
            ParseRecordValue = ParseTxtData(dataPointer, dataLength)

        Case Else
            ' DnsQuery_Wは、問い合わせた種類以外の関連レコードを返す場合があります。
            ' 未対応型の共用体を誤って読むより、種類と長さだけ安全に表示します。
            ParseRecordValue = "(未解析の種類: type=" & CStr(recordType) & _
                               ", dataLength=" & CStr(dataLength) & ")"
    End Select
End Function

Private Function ParseTxtData( _
    ByVal dataPointer As LongPtr, _
    ByVal dataLength As Long) As String

    Dim stringCount As Long
    Dim logicalRequiredBytes As Double
    Dim index As Long
    Dim itemPointer As LongPtr
    Dim itemText As String
    Dim combined As String

    RequireDataLength dataLength, 4, "TXT"
    stringCount = ReadLong(dataPointer, 0)

    If stringCount < 0 Or stringCount > MAX_TXT_STRINGS Then
        Err.Raise vbObjectError + 2120, "ParseTxtData", _
                  "TXT文字列数が安全範囲外です: " & CStr(stringCount)
    End If

    ' wDataLengthは物理パディングを含まない公式マクロの論理長です。
    ' したがって「4 + 文字列数×ポインター幅」で検証し、ポインターを読む
    ' 物理位置だけはOFFSET_TXT_POINTER_ARRAY（64bit=8、32bit=4）を使います。
    logicalRequiredBytes = CDbl(TXT_LOGICAL_HEADER_BYTES) + _
                           (CDbl(stringCount) * CDbl(POINTER_SIZE))
    If logicalRequiredBytes > CDbl(dataLength) Then
        Err.Raise vbObjectError + 2121, "ParseTxtData", _
                  "TXTデータ長が文字列数と一致しません。"
    End If

    For index = 0 To stringCount - 1
        itemPointer = ReadPointer( _
            dataPointer, OFFSET_TXT_POINTER_ARRAY + (index * POINTER_SIZE))
        itemText = PointerToUnicodeString(itemPointer)

        If index > 0 Then combined = combined & " | "
        combined = combined & itemText
    Next index

    ParseTxtData = combined
End Function

Private Function IPv4BytesToString(ByVal addressPointer As LongPtr) As String
    Dim octet0 As Byte
    Dim octet1 As Byte
    Dim octet2 As Byte
    Dim octet3 As Byte

    octet0 = ReadByte(addressPointer, 0)
    octet1 = ReadByte(addressPointer, 1)
    octet2 = ReadByte(addressPointer, 2)
    octet3 = ReadByte(addressPointer, 3)

    IPv4BytesToString = CStr(octet0) & "." & CStr(octet1) & "." & _
                        CStr(octet2) & "." & CStr(octet3)
End Function

Private Function IPv6BytesToString(ByVal addressPointer As LongPtr) As String
    Dim buffer As String
    Dim returnedPointer As LongPtr
    Dim nullPosition As Long

    ' Microsoftの定数に合わせて65文字分を確保します。
    buffer = String$(65, vbNullChar)
    returnedPointer = InetNtopW(AF_INET6, addressPointer, StrPtr(buffer), Len(buffer))

    If returnedPointer = 0 Then
        Err.Raise vbObjectError + 2130, "IPv6BytesToString", _
                  "InetNtopWでIPv6アドレスを文字列化できませんでした。"
    End If

    nullPosition = InStr(1, buffer, vbNullChar, vbBinaryCompare)
    If nullPosition > 0 Then buffer = Left$(buffer, nullPosition - 1)
    IPv6BytesToString = buffer
End Function

Private Function PointerToUnicodeString(ByVal textPointer As LongPtr) As String
    Dim characterCount As Long
    Dim value As String

    If textPointer = 0 Then
        PointerToUnicodeString = ""
        Exit Function
    End If

    ' ポインターはWindows DNS APIが返したレコードリスト内のものだけを読みます。
    ' 外部から任意の数値を渡して使用してはいけません。
    characterCount = lstrlenW(textPointer)
    If characterCount < 0 Or characterCount > MAX_UNICODE_CHARS Then
        Err.Raise vbObjectError + 2140, "PointerToUnicodeString", _
                  "Unicode文字列長が安全範囲外です。"
    End If

    If characterCount = 0 Then
        PointerToUnicodeString = ""
        Exit Function
    End If

    value = String$(characterCount, vbNullChar)
    CopyMemory StrPtr(value), textPointer, characterCount * 2
    PointerToUnicodeString = value
End Function

Private Function ReadPointer( _
    ByVal basePointer As LongPtr, _
    ByVal byteOffset As Long) As LongPtr

    Dim value As LongPtr
    CopyMemory VarPtr(value), PointerAdd(basePointer, byteOffset), POINTER_SIZE
    ReadPointer = value
End Function

Private Function ReadLong( _
    ByVal basePointer As LongPtr, _
    ByVal byteOffset As Long) As Long

    Dim value As Long
    CopyMemory VarPtr(value), PointerAdd(basePointer, byteOffset), 4
    ReadLong = value
End Function

Private Function ReadUInt16( _
    ByVal basePointer As LongPtr, _
    ByVal byteOffset As Long) As Long

    Dim value As Integer
    CopyMemory VarPtr(value), PointerAdd(basePointer, byteOffset), 2
    ReadUInt16 = CLng(value) And &HFFFF&
End Function

Private Function ReadByte( _
    ByVal basePointer As LongPtr, _
    ByVal byteOffset As Long) As Byte

    Dim value As Byte
    CopyMemory VarPtr(value), PointerAdd(basePointer, byteOffset), 1
    ReadByte = value
End Function

' ポインターへ小さなバイトオフセットを加えます。
'
' 32bit版VBAのLongPtrは符号付きLongです。アドレスが&H80000000以上の場合、
' 単純な「basePointer + byteOffset」はオーバーフローする可能性があります。
' 符号ビットを一時的に反転して加算することで、32bitのビット列を保ったまま
' アドレス計算します。64bit版では通常のLongPtr加算で十分です。
Private Function PointerAdd( _
    ByVal basePointer As LongPtr, _
    ByVal byteOffset As Long) As LongPtr

#If Win64 Then
    PointerAdd = basePointer + byteOffset
#Else
    PointerAdd = ((basePointer Xor &H80000000) + byteOffset) Xor &H80000000
#End If
End Function

Private Function ReadUInt32AsDouble( _
    ByVal basePointer As LongPtr, _
    ByVal byteOffset As Long) As Double

    Dim signedValue As Long
    signedValue = ReadLong(basePointer, byteOffset)

    If signedValue < 0 Then
        ReadUInt32AsDouble = CDbl(signedValue) + 4294967296#
    Else
        ReadUInt32AsDouble = CDbl(signedValue)
    End If
End Function

Private Sub RequireDataLength( _
    ByVal actualLength As Long, _
    ByVal minimumLength As Long, _
    ByVal typeName As String)

    If actualLength < minimumLength Then
        Err.Raise vbObjectError + 2150, "RequireDataLength", _
                  typeName & "レコードのデータ長が短すぎます。actual=" & _
                  CStr(actualLength) & ", minimum=" & CStr(minimumLength)
    End If
End Sub

Private Function DnsTypeFromName(ByVal typeName As String) As Long
    Select Case UCase$(Trim$(typeName))
        Case "A": DnsTypeFromName = DNS_TYPE_A
        Case "AAAA": DnsTypeFromName = DNS_TYPE_AAAA
        Case "CNAME": DnsTypeFromName = DNS_TYPE_CNAME
        Case "PTR": DnsTypeFromName = DNS_TYPE_PTR
        Case "MX": DnsTypeFromName = DNS_TYPE_MX
        Case "TXT": DnsTypeFromName = DNS_TYPE_TEXT
        Case Else: DnsTypeFromName = 0
    End Select
End Function

Private Function DnsTypeName(ByVal recordType As Long) As String
    Select Case recordType
        Case DNS_TYPE_A: DnsTypeName = "A"
        Case DNS_TYPE_AAAA: DnsTypeName = "AAAA"
        Case DNS_TYPE_CNAME: DnsTypeName = "CNAME"
        Case DNS_TYPE_PTR: DnsTypeName = "PTR"
        Case DNS_TYPE_MX: DnsTypeName = "MX"
        Case DNS_TYPE_TEXT: DnsTypeName = "TXT"
        Case Else: DnsTypeName = "TYPE" & CStr(recordType)
    End Select
End Function

Private Function DnsStatusText(ByVal status As Long) As String
    Select Case status
        Case 0: DnsStatusText = "ERROR_SUCCESS"
        Case 5: DnsStatusText = "ERROR_ACCESS_DENIED"
        Case 87: DnsStatusText = "ERROR_INVALID_PARAMETER"
        Case 123: DnsStatusText = "ERROR_INVALID_NAME"
        Case 1460: DnsStatusText = "ERROR_TIMEOUT"
        Case 9002: DnsStatusText = "DNS_ERROR_RCODE_SERVER_FAILURE"
        Case 9003: DnsStatusText = "DNS_ERROR_RCODE_NAME_ERROR (名前が存在しません)"
        Case 9501: DnsStatusText = "DNS_INFO_NO_RECORDS (該当種類のレコードなし)"
        Case Else: DnsStatusText = "DNS_STATUS " & CStr(status)
    End Select
End Function

Private Function IsValidIPv4Octet(ByVal value As String) As Boolean
    Dim index As Long
    Dim character As String
    Dim numericValue As Long

    If Len(value) = 0 Or Len(value) > 3 Then Exit Function

    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        If character < "0" Or character > "9" Then Exit Function
    Next index

    numericValue = CLng(value)
    IsValidIPv4Octet = (numericValue >= 0 And numericValue <= 255)
End Function
