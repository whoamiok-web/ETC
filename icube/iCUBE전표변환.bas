Attribute VB_Name = "iCUBE전표변환"
'==============================================================================
' 더존 iCUBE 전표 업로드용 변환 매크로
'  - [원본] 시트의 통장/카드 거래내역을 읽어
'  - [계정매핑] 키워드로 계정과목을 자동 지정하고
'  - [업로드] 시트에 차변/대변 2줄짜리 대체전표 형태로 만들어 줍니다.
'  - 출력 열 순서/제목은 [설정] 시트의 "출력 열" 표에서 바꿀 수 있습니다.
'==============================================================================
Option Explicit

Private Const SH_SET As String = "설정"
Private Const SH_MAP As String = "계정매핑"
Private Const SH_SRC As String = "원본"
Private Const SH_OUT As String = "업로드"

' 출력 필드 키
Private Const F_DATE As String = "일자"
Private Const F_NO As String = "번호"
Private Const F_GUBUN As String = "구분"
Private Const F_ACCT As String = "계정코드"
Private Const F_ACCTNM As String = "계정명"
Private Const F_CUST As String = "거래처코드"
Private Const F_CUSTNM As String = "거래처명"
Private Const F_MEMO As String = "적요"
Private Const F_DR As String = "차변"
Private Const F_CR As String = "대변"

Private gSet As Collection  ' 설정값 (키=항목)
Private gKeys() As String    ' 출력 열 필드키
Private gTitles() As String  ' 출력 열 제목
Private gNCols As Long

'------------------------------------------------------------------------------
' 메인: 변환 실행
'------------------------------------------------------------------------------
Public Sub 전표변환()
    Dim wsSrc As Worksheet, wsMap As Worksheet, wsOut As Worksheet
    Dim srcType As String, startRow As Long, lastRow As Long, r As Long
    Dim outRow As Long, nTrans As Long, nUnmapped As Long, nSkip As Long
    Dim sumDr As Double, sumCr As Double
    Dim prevDate As String, seq As Long

    On Error GoTo EH
    Set wsSrc = ThisWorkbook.Worksheets(SH_SRC)
    Set wsMap = ThisWorkbook.Worksheets(SH_MAP)
    Set wsOut = ThisWorkbook.Worksheets(SH_OUT)

    LoadSettings
    LoadOutputColumns

    srcType = Trim$(CStr(Setting("원본유형")))          ' 통장 / 카드
    If srcType <> "통장" And srcType <> "카드" Then
        MsgBox "[설정] 원본유형은 '통장' 또는 '카드'로 입력하세요.", vbExclamation
        Exit Sub
    End If
    startRow = CLng(Val(Setting("원본시작행")))
    If startRow < 1 Then startRow = 2

    Application.ScreenUpdating = False

    ' 업로드 시트 초기화 + 제목행
    wsOut.Cells.Clear
    Dim c As Long
    For c = 1 To gNCols
        wsOut.Cells(1, c).Value = gTitles(c)
        wsOut.Cells(1, c).Font.Bold = True
        wsOut.Cells(1, c).Interior.Color = RGB(221, 235, 247)
    Next c
    outRow = 2

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, ColNum(Setting("열_일자"))).End(xlUp).Row

    For r = startRow To lastRow
        Dim dt As Variant, memo As String, cust As String
        Dim amtIn As Double, amtOut As Double, amt As Double

        dt = ParseDate(wsSrc.Cells(r, ColNum(Setting("열_일자"))).Value)
        If IsEmpty(dt) Then GoTo NextRow

        memo = Trim$(CStr(CellAt(wsSrc, r, Setting("열_적요"))))
        cust = Trim$(CStr(CellAt(wsSrc, r, Setting("열_거래처"))))
        If cust = "" Then cust = memo

        If srcType = "통장" Then
            amtIn = ParseAmount(CellAt(wsSrc, r, Setting("열_입금")))
            amtOut = ParseAmount(CellAt(wsSrc, r, Setting("열_출금")))
        Else
            amtOut = ParseAmount(CellAt(wsSrc, r, Setting("열_금액")))
            amtIn = 0
            If amtOut < 0 Then amtIn = -amtOut: amtOut = 0   ' 카드 취소(음수)
        End If
        If amtIn = 0 And amtOut = 0 Then nSkip = nSkip + 1: GoTo NextRow

        ' 계정 매핑
        Dim m As Variant, mapped As Boolean
        ' 카드 취소는 원래 비용계정을 되돌리므로 '출금' 매핑을 사용
        m = FindMapping(wsMap, memo & " " & cust, (srcType = "통장" And amtIn > 0), mapped)
        If Not mapped Then nUnmapped = nUnmapped + 1

        Dim dKey As String: dKey = Format$(dt, "yyyymmdd")
        If dKey <> prevDate Then seq = 0: prevDate = dKey
        seq = seq + 1

        Dim memoOut As String
        memoOut = CStr(m(3)): If memoOut = "" Then memoOut = memo
        memoOut = Left$(memoOut, 80)

        ' 상대계정(통장=보통예금, 카드=미지급금)
        Dim baseAcct As String, baseNm As String, baseCust As String, baseCustNm As String
        If srcType = "통장" Then
            baseAcct = Setting("통장_계정코드"): baseNm = Setting("통장_계정명")
            baseCust = Setting("통장_거래처코드"): baseCustNm = Setting("통장_거래처명")
        Else
            baseAcct = Setting("카드_계정코드"): baseNm = Setting("카드_계정명")
            baseCust = Setting("카드_거래처코드"): baseCustNm = Setting("카드_거래처명")
        End If

        amt = amtOut + amtIn
        Dim custCode As String, custNm As String
        custCode = CStr(m(2)): custNm = cust

        If amtOut > 0 Then
            ' 지출: (차) 비용계정  / (대) 보통예금·미지급금
            If UCase$(CStr(m(4))) = "Y" Then
                Dim vat As Double: vat = Round(amt / 11, 0)
                WriteLine wsOut, outRow, dt, seq, "3", m(0), m(1), custCode, custNm, memoOut, amt - vat, 0, Not mapped
                WriteLine wsOut, outRow, dt, seq, "3", Setting("부가세_계정코드"), Setting("부가세_계정명"), custCode, custNm, memoOut, vat, 0, Not mapped
            Else
                WriteLine wsOut, outRow, dt, seq, "3", m(0), m(1), custCode, custNm, memoOut, amt, 0, Not mapped
            End If
            WriteLine wsOut, outRow, dt, seq, "4", baseAcct, baseNm, baseCust, baseCustNm, memoOut, 0, amt, False
        Else
            ' 입금/카드취소: (차) 보통예금·미지급금 / (대) 수익·상대계정
            WriteLine wsOut, outRow, dt, seq, "3", baseAcct, baseNm, baseCust, baseCustNm, memoOut, amt, 0, False
            If UCase$(CStr(m(4))) = "Y" Then
                Dim vat2 As Double: vat2 = Round(amt / 11, 0)
                WriteLine wsOut, outRow, dt, seq, "4", m(0), m(1), custCode, custNm, memoOut, 0, amt - vat2, Not mapped
                WriteLine wsOut, outRow, dt, seq, "4", Setting("부가세_계정코드"), Setting("부가세_계정명"), custCode, custNm, memoOut, 0, vat2, Not mapped
            Else
                WriteLine wsOut, outRow, dt, seq, "4", m(0), m(1), custCode, custNm, memoOut, 0, amt, Not mapped
            End If
        End If
        sumDr = sumDr + amt: sumCr = sumCr + amt
        nTrans = nTrans + 1
NextRow:
    Next r

    ' 서식
    Dim cDr As Long, cCr As Long, cDt As Long
    cDr = FieldCol(F_DR): cCr = FieldCol(F_CR): cDt = FieldCol(F_DATE)
    If outRow > 2 Then
        If cDr > 0 Then wsOut.Range(wsOut.Cells(2, cDr), wsOut.Cells(outRow - 1, cDr)).NumberFormat = "#,##0"
        If cCr > 0 Then wsOut.Range(wsOut.Cells(2, cCr), wsOut.Cells(outRow - 1, cCr)).NumberFormat = "#,##0"
    End If
    wsOut.Columns.AutoFit

    Application.ScreenUpdating = True
    wsOut.Activate

    MsgBox "변환 완료" & vbCrLf & vbCrLf & _
           "거래 " & nTrans & "건 → 전표 " & (outRow - 2) & "줄" & vbCrLf & _
           "차변 합계 " & Format$(sumDr, "#,##0") & " / 대변 합계 " & Format$(sumCr, "#,##0") & vbCrLf & _
           "계정 미지정(노란색) " & nUnmapped & "건, 금액 0 건너뜀 " & nSkip & "건" & vbCrLf & vbCrLf & _
           "노란색 줄의 계정을 확인한 뒤 [업로드 파일 저장]을 실행하세요.", vbInformation
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "오류: " & Err.Description & " (원본 " & r & "행 근처)", vbCritical
End Sub

'------------------------------------------------------------------------------
' 업로드 시트만 별도 .xlsx 로 저장 (iCUBE에서 불러올 파일)
'------------------------------------------------------------------------------
Public Sub 업로드파일저장()
    Dim wsOut As Worksheet, wb As Workbook, path As Variant
    Set wsOut = ThisWorkbook.Worksheets(SH_OUT)
    If Application.WorksheetFunction.CountA(wsOut.Rows(2)) = 0 Then
        MsgBox "먼저 [전표 변환]을 실행하세요.", vbExclamation: Exit Sub
    End If
    path = Application.GetSaveAsFilename( _
        InitialFileName:="iCUBE_전표업로드_" & Format$(Date, "yyyymmdd") & ".xls", _
        FileFilter:="Excel 97-2003 (*.xls),*.xls,Excel 통합문서 (*.xlsx),*.xlsx")
    If VarType(path) = vbBoolean Then Exit Sub

    wsOut.Copy
    Set wb = ActiveWorkbook
    wb.Worksheets(1).Cells.Interior.Pattern = xlNone   ' 표시용 색 제거
    Application.DisplayAlerts = False
    If LCase$(Right$(path, 4)) = ".xls" Then
        wb.SaveAs Filename:=path, FileFormat:=56       ' xlExcel8
    Else
        wb.SaveAs Filename:=path, FileFormat:=51       ' xlOpenXMLWorkbook
    End If
    Application.DisplayAlerts = True
    wb.Close False
    MsgBox "저장했습니다: " & vbCrLf & path, vbInformation
End Sub

'------------------------------------------------------------------------------
' 원본/업로드 비우기
'------------------------------------------------------------------------------
Public Sub 원본비우기()
    Dim ws As Worksheet, sr As Long
    If MsgBox("[원본] 시트의 데이터(시작행 이후)를 지울까요?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    LoadSettings
    sr = CLng(Val(Setting("원본시작행"))): If sr < 1 Then sr = 2
    Set ws = ThisWorkbook.Worksheets(SH_SRC)
    ws.Range(ws.Rows(sr), ws.Rows(ws.Rows.Count)).ClearContents
    ThisWorkbook.Worksheets(SH_OUT).Cells.Clear
End Sub

'==============================================================================
' 내부 함수
'==============================================================================
Private Sub WriteLine(ByVal ws As Worksheet, ByRef outRow As Long, ByVal dt As Variant, ByVal seq As Long, _
                      ByVal gubun As String, ByVal acct As Variant, ByVal acctNm As Variant, ByVal custCd As Variant, _
                      ByVal custNm As Variant, ByVal memo As String, ByVal dr As Double, ByVal cr As Double, ByVal warn As Boolean)
    Dim c As Long, fkey As String, v As Variant, fmtDate As String
    fmtDate = CStr(Setting("일자형식")): If fmtDate = "" Then fmtDate = "yyyymmdd"
    For c = 1 To gNCols
        fkey = gKeys(c)
        Select Case fkey
            Case F_DATE:   v = Format$(dt, fmtDate)
            Case F_NO:     v = seq
            Case F_GUBUN:  v = gubun
            Case F_ACCT:   v = acct
            Case F_ACCTNM: v = acctNm
            Case F_CUST:   v = custCd
            Case F_CUSTNM: v = custNm
            Case F_MEMO:   v = memo
            Case F_DR:     v = dr
            Case F_CR:     v = cr
            Case Else:     v = ""
        End Select
        If fkey = F_DATE Or fkey = F_ACCT Or fkey = F_CUST Then
            ws.Cells(outRow, c).NumberFormat = "@"     ' 앞자리 0 보존
        End If
        ws.Cells(outRow, c).Value = v
    Next c
    If warn Then ws.Range(ws.Cells(outRow, 1), ws.Cells(outRow, gNCols)).Interior.Color = RGB(255, 242, 153)
    outRow = outRow + 1
End Sub

' 반환: Array(계정코드, 계정명, 거래처코드, 적요, 부가세공제Y/N)
Private Function FindMapping(wsMap As Worksheet, text As String, isIncome As Boolean, ByRef mapped As Boolean) As Variant
    Dim r As Long, lastRow As Long, kws As Variant, k As Long, kw As String, direction As String
    lastRow = wsMap.Cells(wsMap.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        direction = Trim$(CStr(wsMap.Cells(r, 2).Value))  ' 출금/입금/공통
        If direction = "" Then direction = "공통"
        If direction = "공통" Or (direction = "입금" And isIncome) Or (direction = "출금" And Not isIncome) Then
            kws = Split(CStr(wsMap.Cells(r, 1).Value), "|")
            For k = LBound(kws) To UBound(kws)
                kw = Trim$(kws(k))
                If kw <> "" Then
                    If InStr(1, text, kw, vbTextCompare) > 0 Then
                        FindMapping = Array(CStr(wsMap.Cells(r, 3).Value), CStr(wsMap.Cells(r, 4).Value), _
                                            CStr(wsMap.Cells(r, 5).Value), CStr(wsMap.Cells(r, 6).Value), _
                                            CStr(wsMap.Cells(r, 7).Value))
                        mapped = True
                        Exit Function
                    End If
                End If
            Next k
        End If
    Next r
    mapped = False
    If isIncome Then
        FindMapping = Array(Setting("미지정_입금_계정코드"), Setting("미지정_입금_계정명"), "", "", "N")
    Else
        FindMapping = Array(Setting("미지정_출금_계정코드"), Setting("미지정_출금_계정명"), "", "", "N")
    End If
End Function

Private Sub LoadSettings()
    Dim ws As Worksheet, r As Long, lastRow As Long, k As String
    Set gSet = New Collection
    Set ws = ThisWorkbook.Worksheets(SH_SET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 1 To lastRow
        k = Trim$(CStr(ws.Cells(r, 1).Value))
        If k <> "" Then
            On Error Resume Next            ' 같은 항목이 두 번 있으면 처음 값 사용
            gSet.Add ws.Cells(r, 2).Value, k
            On Error GoTo 0
        End If
    Next r
End Sub

Private Function Setting(key As String) As Variant
    Dim v As Variant
    v = ""
    On Error Resume Next
    v = gSet.Item(key)
    On Error GoTo 0
    If IsEmpty(v) Then v = ""
    Setting = v
End Function

' [설정] 시트 D:E 열 = 필드키 / 출력제목 (D1 머리글, 2행부터)
Private Sub LoadOutputColumns()
    Dim ws As Worksheet, r As Long, lastRow As Long, k As String, t As String
    Set ws = ThisWorkbook.Worksheets(SH_SET)
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    gNCols = 0
    If lastRow < 2 Then lastRow = 2
    ReDim gKeys(1 To lastRow): ReDim gTitles(1 To lastRow)
    For r = 2 To lastRow
        k = Trim$(CStr(ws.Cells(r, 4).Value))
        t = Trim$(CStr(ws.Cells(r, 5).Value))
        If k <> "" Then
            If t = "" Then t = k
            gNCols = gNCols + 1
            gKeys(gNCols) = k: gTitles(gNCols) = t
        End If
    Next r
    If gNCols = 0 Then Err.Raise vbObjectError + 1, , "[설정] 출력 열(D:E)이 비어 있습니다."
End Sub

Private Function FieldCol(key As String) As Long
    Dim c As Long
    For c = 1 To gNCols
        If gKeys(c) = key Then FieldCol = c: Exit Function
    Next c
End Function

' "A", "C", 3 등 → 열 번호
Private Function ColNum(v As Variant) As Long
    Dim s As String, i As Long, n As Long
    s = UCase$(Trim$(CStr(v)))
    If s = "" Then ColNum = 0: Exit Function
    If IsNumeric(s) Then ColNum = CLng(s): Exit Function
    For i = 1 To Len(s)
        n = n * 26 + (Asc(Mid$(s, i, 1)) - 64)
    Next i
    ColNum = n
End Function

Private Function CellAt(ws As Worksheet, r As Long, colSpec As Variant) As Variant
    Dim c As Long: c = ColNum(colSpec)
    If c = 0 Then CellAt = "" Else CellAt = ws.Cells(r, c).Value
End Function

Private Function ParseAmount(v As Variant) As Double
    Dim s As String
    If IsNumeric(v) And Not IsEmpty(v) Then
        If VarType(v) <> vbString Then ParseAmount = CDbl(v): Exit Function
    End If
    s = Trim$(CStr(v))
    s = Replace(s, ",", ""): s = Replace(s, "원", ""): s = Replace(s, " ", "")
    If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then s = "-" & Mid$(s, 2, Len(s) - 2)
    If s = "" Or s = "-" Then ParseAmount = 0: Exit Function
    If IsNumeric(s) Then ParseAmount = CDbl(s) Else ParseAmount = 0
End Function

' 2026-11-17 / 2026.11.17 / 2026/11/17 13:22 / 20261117 / 날짜값 → Date (실패 시 Empty)
Private Function ParseDate(v As Variant) As Variant
    Dim s As String, y As Long, mo As Long, d As Long, parts As Variant
    ParseDate = Empty
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then ParseDate = DateValue(v): Exit Function
    If IsNumeric(v) And VarType(v) <> vbString Then
        If v > 19000101 And v < 21001231 Then
            s = CStr(CLng(v))
        ElseIf v > 20000 And v < 80000 Then
            ParseDate = DateValue(CDate(v)): Exit Function   ' 엑셀 일련번호
        Else
            Exit Function
        End If
    Else
        s = Trim$(CStr(v))
    End If
    If InStr(s, " ") > 0 Then s = Left$(s, InStr(s, " ") - 1)
    s = Replace(Replace(Replace(s, ".", "-"), "/", "-"), "년", "-")
    s = Replace(Replace(s, "월", "-"), "일", "")
    If Right$(s, 1) = "-" Then s = Left$(s, Len(s) - 1)
    If Len(s) = 8 And IsNumeric(s) Then
        y = CLng(Left$(s, 4)): mo = CLng(Mid$(s, 5, 2)): d = CLng(Right$(s, 2))
    Else
        parts = Split(s, "-")
        If UBound(parts) <> 2 Then Exit Function
        If Not (IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2))) Then Exit Function
        y = CLng(parts(0)): mo = CLng(parts(1)): d = CLng(parts(2))
        If y < 100 Then y = y + 2000
    End If
    If y < 1900 Or mo < 1 Or mo > 12 Or d < 1 Or d > 31 Then Exit Function
    ParseDate = DateSerial(y, mo, d)
End Function
