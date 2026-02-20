Attribute VB_Name = "Module1"
Option Explicit

' =========================
' 현재(원본) 파일 시트명
' =========================
Private Const SHEET_RAW As String = "Rawdata"
Private Const SHEET_EQINFO As String = "eqInfo"
Private Const SHEET_CODEINFO As String = "CodeInfo"

' =========================
' 외부(새 분석표) 파일명 (같은 폴더)
' =========================
Private Const NEW_REPORT_FILE As String = "치수변화율-원단시험분석표2_v1.1_20260128.xlsx"  '필요시 수정

' =========================
' Rawdata 컬럼(중요)
' =========================
Private Const COL_AKEY As Long = 1      'A열: @접수@시료,방법
Private Const COL_BEFORE As Long = 2    'B열: 기준값(전/Before/SPEC)
Private Const COL_DATE As Long = 10     'J열: 날짜(대표행)

' 측정값(시험후 After)
Private Const COL_LEN_K As Long = 11    'K
Private Const COL_LEN_L As Long = 12    'L
Private Const COL_LEN_M As Long = 13    'M
Private Const COL_WID_N As Long = 14    'N
Private Const COL_WID_O As Long = 15    'O
Private Const COL_WID_P As Long = 16    'P
Private Const SHEET_ROUNDING As String = "roundingInfo"


' =========================
' 메인
' =========================
Public Sub Print_NewReport_ToExternalTemplate()

    Dim wsR As Worksheet, wsEq As Worksheet, wsCode As Worksheet
    Set wsR = ThisWorkbook.Worksheets(SHEET_RAW)
    Set wsEq = ThisWorkbook.Worksheets(SHEET_EQINFO)
    Set wsCode = ThisWorkbook.Worksheets(SHEET_CODEINFO)
    Dim wsRound As Worksheet
    Set wsRound = ThisWorkbook.Worksheets(SHEET_ROUNDING)

    If ActiveSheet.Name <> SHEET_RAW Then
        MsgBox "Rawdata 시트에서 출력할 행을 선택한 뒤 실행하세요.", vbExclamation
        Exit Sub
    End If
    If TypeName(Selection) <> "Range" Then
        MsgBox "출력할 행(범위)을 먼저 선택하세요.", vbExclamation
        Exit Sub
    End If
    If Selection.Rows.Count = 0 Then
        MsgBox "선택된 범위가 없습니다.", vbExclamation
        Exit Sub
    End If

    ' groups(key) = Dictionary(sampleNo(Long) -> Collection(rowNums))
    Dim groups As Object: Set groups = CreateObject("Scripting.Dictionary")
    Dim coreByKey As Object: Set coreByKey = CreateObject("Scripting.Dictionary")

    Dim r As Range, rowNum As Long
    Dim receipt As String, core As String, methodNo As String
    Dim sampleNo As Long, key As String

    ' 1) 선택 영역 파싱
    For Each r In Selection.Rows
        rowNum = r.Row

        If ParseAKey(wsR.Cells(rowNum, COL_AKEY).Value, receipt, core, sampleNo, methodNo) Then
            If sampleNo >= 1 Then
                key = receipt & "|" & methodNo

                If Not groups.Exists(key) Then
                    groups.Add key, CreateObject("Scripting.Dictionary")
                    coreByKey.Add key, core
                End If

                Dim d As Object: Set d = groups(key)
                If Not d.Exists(sampleNo) Then
                    Dim colRows As Collection: Set colRows = New Collection
                    d.Add sampleNo, colRows
                End If
                d(sampleNo).Add rowNum
            End If
        End If
    Next r

    If groups.Count = 0 Then
        MsgBox "선택 범위에서 유효한 A열 키(@접수@시료,방법)를 찾지 못했습니다.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim wbNew As Workbook
    Dim newPath As String
    newPath = ThisWorkbook.Path & "\" & NEW_REPORT_FILE

    On Error GoTo CleanFail

    If Dir$(newPath) = "" Then
        MsgBox "새 분석표 파일을 찾지 못했습니다:" & vbCrLf & newPath, vbExclamation
        GoTo CleanExit
    End If

    Set wbNew = Workbooks.Open(newPath, ReadOnly:=False)

    ' 2) 키들을 '접수번호 순'으로 정렬하여 출력
    Dim keysArr() As String
    keysArr = DictKeysToArray(groups)
    QuickSortKeys keysArr, LBound(keysArr), UBound(keysArr)

    Dim i As Long
    Dim grp As Object
    Dim firstRow As Long
    Dim methodName As String
    Dim langSheetName As String
    Dim wsOut As Worksheet

    Dim measureMap As Variant
    Dim resultMap As Variant
    measureMap = SamplePosMap_UpTo6_Measures()
    resultMap = SamplePosMap_UpTo6_Results()

    For i = LBound(keysArr) To UBound(keysArr)

        key = keysArr(i)
        Set grp = groups(key)

        receipt = Split(key, "|")(0)
        methodNo = Split(key, "|")(1)
        core = coreByKey(key)

        firstRow = GetFirstRowFromGroup(grp)

        ' 시험방법명(txt) + 끝 ":" 제거
        methodName = GetMethodNameFromTxt(core, methodNo)
        methodName = RemoveTrailingColon(methodName)
        Dim roundStep As Double
        roundStep = GetRoundingStep(methodName, wsRound, 0.1) '기본 0.1 (원하면 기본값 변경)


        ' 출력 시트 결정(국문/영문/일문)
        langSheetName = ResolveLangSheet(wsCode, receipt, core)
        Set wsOut = wbNew.Worksheets(langSheetName)

        ' 시트 초기화
        ClearNewReport wsOut

        ' ========= 공통값 입력 =========
        ' 실무자
        SetValueSafe wsOut, "K4", wsR.Range("I2").Value

        ' 온습도(값/범위)
        'SetValueSafe wsOut, "C38", wsR.Range("B2").Value
        'SetValueSafe wsOut, "E38", wsR.Range("C2").Value
        'SetValueSafe wsOut, "C39", wsR.Range("B3").Value
        'SetValueSafe wsOut, "E39", wsR.Range("C3").Value

        ' 접수번호/시험방법
        SetValueSafe wsOut, "B5", receipt
        SetValueSafe wsOut, "B6", methodName

        ' 날짜(J열 기반)
        'SetValueSafe wsOut, "J38", FormatDate_yyyy_mm_dd(wsR.Cells(firstRow, COL_DATE).Value)

        ' eqInfo: B27 / D27
        Dim eqB As String, eqC As String
        GetEqInfoForMethod methodName, wsEq, eqB, eqC

        SetValueSafe wsOut, "B27", eqB
        SetValueSafe wsOut, "D27", eqC

        ' 샘플 오름차순 최대 6개
        Dim sampleList() As Long
        sampleList = GetSortedSampleNos(grp)

        Dim sIdx As Long, sNo As Long
        Dim rr0 As Long
        Dim beforeV As Double
        Dim specVal As Variant

        Dim L1 As Double, L2 As Double, L3 As Double
        Dim W1 As Double, W2 As Double, W3 As Double
        Dim Lx As Double, Ly As Double, Lz As Double, Lm As Double
        Dim Wx As Double, Wy As Double, Wz As Double, Wm As Double

        Dim totalSamples As Long
        totalSamples = UBound(sampleList) - LBound(sampleList) + 1
        
        Dim pageIdx As Long
        Dim startIdx As Long, endIdx As Long
        
        Dim sampleNoMap As Variant
        sampleNoMap = SampleNoPosMap_UpTo6()
        
        For pageIdx = 0 To (totalSamples - 1) \ 6
        
            ' 페이지마다 양식 초기화
            ClearNewReport wsOut
        
            ' ===== 공통값 다시 입력 =====
            SetValueSafe wsOut, "K4", wsR.Range("I2").Value
            'SetValueSafe wsOut, "C38", wsR.Range("B2").Value
            'SetValueSafe wsOut, "E38", wsR.Range("C2").Value
            'SetValueSafe wsOut, "C39", wsR.Range("B3").Value
            'SetValueSafe wsOut, "E39", wsR.Range("C3").Value
            SetValueSafe wsOut, "B5", receipt
            SetValueSafe wsOut, "B6", methodName
            'SetValueSafe wsOut, "J38", FormatDate_yyyy_mm_dd(wsR.Cells(firstRow, COL_DATE).Value)
            SetValueSafe wsOut, "B27", eqB
            SetValueSafe wsOut, "D27", eqC
        
            startIdx = pageIdx * 6
            endIdx = Application.Min(startIdx + 5, totalSamples - 1)
        
            For sIdx = startIdx To endIdx
        
                Dim localIdx As Long
                localIdx = sIdx - startIdx   ' 0~5
        
                sNo = sampleList(LBound(sampleList) + sIdx)
                rr0 = CLng(grp(sNo)(1))
        
                ' ===== 시료번호 표시 =====
                SetValueSafe wsOut, sampleNoMap(localIdx), sNo
        
                ' ===== 기준값 =====
                specVal = wsR.Cells(rr0, COL_BEFORE).Value
                Dim washCnt As Variant
                washCnt = wsR.Cells(rr0, 8).Value   ' H열 = 8
                
                SetSpecTextCells wsOut, specVal, washCnt, eqB

                If IsNumeric(specVal) And specVal <> 0 Then
                    beforeV = CDbl(specVal)
                Else
                    beforeV = 0#
                End If
        
                ' ===== 측정값 읽기 =====
                L1 = NzD(wsR.Cells(rr0, COL_LEN_K).Value)
                L2 = NzD(wsR.Cells(rr0, COL_LEN_L).Value)
                L3 = NzD(wsR.Cells(rr0, COL_LEN_M).Value)
                W1 = NzD(wsR.Cells(rr0, COL_WID_N).Value)
                W2 = NzD(wsR.Cells(rr0, COL_WID_O).Value)
                W3 = NzD(wsR.Cells(rr0, COL_WID_P).Value)
        
                ' ===== 측정값 입력 =====
                SetValueSafe wsOut, measureMap(localIdx)(0), Fmt1(L1)
                SetValueSafe wsOut, measureMap(localIdx)(1), Fmt1(L2)
                SetValueSafe wsOut, measureMap(localIdx)(2), Fmt1(L3)
                SetValueSafe wsOut, measureMap(localIdx)(3), Fmt1(W1)
                SetValueSafe wsOut, measureMap(localIdx)(4), Fmt1(W2)
                SetValueSafe wsOut, measureMap(localIdx)(5), Fmt1(W3)
        
                ' ===== 치수변화율 계산 =====
                CalcShrinkXYZ beforeV, L1, L2, L3, Lx, Ly, Lz, Lm
                CalcShrinkXYZ beforeV, W1, W2, W3, Wx, Wy, Wz, Wm
        
                ' ===== 결과 입력 =====
                SetPctSafe wsOut, resultMap(localIdx)(0), FmtPct1(Lx)
                SetPctSafe wsOut, resultMap(localIdx)(1), FmtPct1(Ly)
                SetPctSafe wsOut, resultMap(localIdx)(2), FmtPct1(Lz)
                SetPctSafe wsOut, resultMap(localIdx)(3), FmtPct1(Lm)
        
                SetPctSafe wsOut, resultMap(localIdx)(4), FmtPct1(Wx)
                SetPctSafe wsOut, resultMap(localIdx)(5), FmtPct1(Wy)
                SetPctSafe wsOut, resultMap(localIdx)(6), FmtPct1(Wz)
                SetPctSafe wsOut, resultMap(localIdx)(7), FmtPct1(Wm)
        
                ' ===== 수치맺음 AVG =====
                SetPctSafe wsOut, OffsetCell1Down(resultMap(localIdx)(3)), _
                             FmtPct1(RoundToStep(Lm, roundStep))
                SetPctSafe wsOut, OffsetCell1Down(resultMap(localIdx)(7)), _
                             FmtPct1(RoundToStep(Wm, roundStep))
        
            Next sIdx
        
            ' ===== 페이지별 출력 =====
            wsOut.PrintOut
        
        Next pageIdx


    Next i

CleanExit:
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "완료: 새 분석표 파일(국문/영문/일문)에 채운 후 인쇄했습니다.", vbInformation
    Exit Sub

CleanFail:
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "오류: " & Err.Description, vbExclamation
End Sub

'=========================================================
' [계산] before + 3개의 after -> x,y,z,avg
'=========================================================
Private Sub CalcShrinkXYZ(ByVal beforeV As Double, ByVal a1 As Double, ByVal a2 As Double, ByVal a3 As Double, _
                          ByRef x As Double, ByRef y As Double, ByRef z As Double, ByRef meanV As Double)
    If beforeV = 0# Then
        x = 0#: y = 0#: z = 0#: meanV = 0#
        Exit Sub
    End If
    x = (a1 - beforeV) / beforeV * 100#
    y = (a2 - beforeV) / beforeV * 100#
    z = (a3 - beforeV) / beforeV * 100#
    meanV = (x + y + z) / 3#
End Sub

Private Function NzD(ByVal v As Variant) As Double
    If IsNumeric(v) Then NzD = CDbl(v) Else NzD = 0#
End Function

Private Function FmtPct1(ByVal v As Double) As String
    ' 양수면 + 붙임, 0은 0.0, 음수는 - 유지
    If v > 0 Then
        FmtPct1 = "+" & Format$(v, "0.0")
    Else
        FmtPct1 = Format$(v, "0.0")
    End If
End Function




Private Function Fmt1(ByVal v As Variant) As String
    If IsNumeric(v) Then
        Fmt1 = Format$(CDbl(v), "0.0")
    Else
        Fmt1 = CStr(v)
    End If
End Function

'=========================================================
' SPEC 문구 치환해서 셀에 입력 + G34/I34는 숫자 그대로
' - specValue: Rawdata B열
' - washCount: Rawdata H열 (N1)
' - eqHeader : eqInfo에서 선택된 헤더(=B27에 들어가는 eqB)
'=========================================================
Private Sub SetSpecTextCells(ByVal ws As Worksheet, ByVal specValue As Variant, ByVal washCount As Variant, ByVal eqHeader As String)

    Dim specStr As String
    specStr = CStr(specValue)

    Dim n1 As Long
    If IsNumeric(washCount) Then
        n1 = CLng(washCount)
    Else
        n1 = 0
    End If

    Dim n2 As String
    n2 = OrdinalEn(n1)   ' 1 -> 1st, 2 -> 2nd ...

    Dim baseTxt As String

    ' --- 조건 템플릿 선택(요청 반영) ---
    If InStr(1, eqHeader, "스팀프레스", vbTextCompare) > 0 Then
        baseTxt = "N1 회 다리미질 후 (  SPEC mm )" & vbCrLf & "After N2 Steam press"

    ElseIf (InStr(1, eqHeader, "퍼클로로에틸렌", vbTextCompare) > 0) Or _
           (InStr(1, eqHeader, "석유계", vbTextCompare) > 0) Then
        baseTxt = "N1 회 드라이클리닝 후 (  SPEC mm )" & vbCrLf & "After N2 Drycleaned"

    ElseIf InStr(1, eqHeader, "세탁", vbTextCompare) > 0 Then
        baseTxt = "N1 회 세탁 후 (  SPEC mm )" & vbCrLf & "After N2 Washing"

    Else
        ' 기본값(원래 문구)
        baseTxt = "N1 회 세탁/드라이클리닝 후 (  SPEC mm )" & vbCrLf & "After N2 Washing/Drycleaned"
    End If

    Dim txt As String
    txt = Replace(Replace(Replace(baseTxt, "N1", CStr(n1)), "SPEC", specStr), "N2", n2)

    ' 문장 들어갈 셀
    Dim addrsText As Variant, i As Long
    addrsText = Array("B10", "F10", "J10", "B19", "F19", "J19")

    For i = LBound(addrsText) To UBound(addrsText)
        SetValueSafe ws, CStr(addrsText(i)), txt
    Next i

    ' G34 / I34는 숫자 그대로
    SetValueSafe ws, "G34", specValue
    SetValueSafe ws, "I34", specValue

End Sub

'=========================================================
' 영어 서수 변환: 1->1st, 2->2nd, 3->3rd, 4->4th ...
'=========================================================
Private Function OrdinalEn(ByVal n As Long) As String

    If n <= 0 Then
        OrdinalEn = ""
        Exit Function
    End If

    Dim mod100 As Long, mod10 As Long, suf As String
    mod100 = n Mod 100
    mod10 = n Mod 10

    If mod100 >= 11 And mod100 <= 13 Then
        suf = "th"
    Else
        Select Case mod10
            Case 1: suf = "st"
            Case 2: suf = "nd"
            Case 3: suf = "rd"
            Case Else: suf = "th"
        End Select
    End If

    OrdinalEn = CStr(n) & suf

End Function


'=========================================================
' 새 분석표 초기화 범위 (요청 반영)
'=========================================================
Private Sub ClearNewReport(ByVal ws As Worksheet)

    ' 접수번호/시험방법
    ws.Range("B5").ClearContents
    ws.Range("B6").ClearContents

    ' 실무자/온습도/날짜
    ws.Range("K4").ClearContents
    'ws.Range("C38,E38,C39,E39").ClearContents
    'ws.Range("J38").ClearContents

    ' eqInfo
    ws.Range("B27,D27").ClearContents

    ' SPEC(지정 8셀)
    ws.Range("B10,F10,J10,B19,F19,J19,G34,I34").ClearContents

    ' 측정값(길이/폭) 6개 블록
    ws.Range("B13:B15,D13:D15,F13:F15,H13:H15,J13:J15,L13:L15").ClearContents
    ws.Range("B22:B24,D22:D24,F22:F24,H22:H24,J22:J24,L22:L24").ClearContents

    ' 결과(치수변화율) 6개 블록
    ws.Range("C13:C16,E13:E16,G13:G16,I13:I16,K13:K16,M13:M16").ClearContents
    ws.Range("C22:C25,E22:E25,G22:G25,I22:I25,K22:K25,M22:M25").ClearContents
    
    ' 수치맺음 AVG가 들어가는 행(한 줄 아래) 추가 초기화
    ws.Range("C17,E17,G17,I17,K17,M17").ClearContents
    ws.Range("C26,E26,G26,I26,K26,M26").ClearContents
    
    ' ? [추가] 시료번호 표시 셀 초기화 (페이지 넘어갈 때 잔상 제거)
    ws.Range("C9,G9,K9,C18,G18,K18").ClearContents
    

End Sub

'=========================================================
' 샘플 1~6 [측정값] 위치 매핑
'  (0)(1)(2)=길이 K,L,M / (3)(4)(5)=폭 N,O,P
'=========================================================
Private Function SamplePosMap_UpTo6_Measures() As Variant
    Dim m(0 To 5) As Variant

    m(0) = Array("B13", "B14", "B15", "D13", "D14", "D15")
    m(1) = Array("F13", "F14", "F15", "H13", "H14", "H15")
    m(2) = Array("J13", "J14", "J15", "L13", "L14", "L15")
    m(3) = Array("B22", "B23", "B24", "D22", "D23", "D24")
    m(4) = Array("F22", "F23", "F24", "H22", "H23", "H24")
    m(5) = Array("J22", "J23", "J24", "L22", "L23", "L24")

    SamplePosMap_UpTo6_Measures = m
End Function

'=========================================================
' 샘플 1~6 [결과값(치수변화율)] 위치 매핑
'  (0)(1)(2)(3)=길이 x,y,z,avg / (4)(5)(6)(7)=폭 x,y,z,avg
'=========================================================
Private Function SamplePosMap_UpTo6_Results() As Variant
    Dim m(0 To 5) As Variant

    m(0) = Array("C13", "C14", "C15", "C16", "E13", "E14", "E15", "E16")
    m(1) = Array("G13", "G14", "G15", "G16", "I13", "I14", "I15", "I16")
    m(2) = Array("K13", "K14", "K15", "K16", "M13", "M14", "M15", "M16")
    m(3) = Array("C22", "C23", "C24", "C25", "E22", "E23", "E24", "E25")
    m(4) = Array("G22", "G23", "G24", "G25", "I22", "I23", "I24", "I25")
    m(5) = Array("K22", "K23", "K24", "K25", "M22", "M23", "M24", "M25")

    SamplePosMap_UpTo6_Results = m
End Function

'=========================================================
' 병합셀 안전 쓰기
'=========================================================
Private Sub SetValueSafe(ByVal ws As Worksheet, ByVal addr As String, ByVal v As Variant)
    With ws.Range(addr)
        If .MergeCells Then
            .MergeArea.Cells(1, 1).Value = v
        Else
            .Value = v
        End If
    End With
End Sub

'=========================================================
' [언어 시트 결정]
' 기본=국문
' CodeInfo!B열 키워드 포함 -> 영문
' CodeInfo!C열 키워드 포함 -> 일문
'=========================================================
Private Function ResolveLangSheet(ByVal wsCode As Worksheet, ByVal receipt As String, ByVal core As String) As String

    ResolveLangSheet = "국문"

    Dim lastRow As Long, r As Long
    lastRow = wsCode.Cells(wsCode.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then Exit Function

    For r = 1 To lastRow
        Dim kwB As String, kwC As String
        kwB = Trim$(CStr(wsCode.Cells(r, 2).Value)) 'B
        kwC = Trim$(CStr(wsCode.Cells(r, 3).Value)) 'C

        If Len(kwB) > 0 Then
            If InStr(1, receipt, kwB, vbTextCompare) > 0 Or InStr(1, core, kwB, vbTextCompare) > 0 Then
                ResolveLangSheet = "영문"
                Exit Function
            End If
        End If

        If Len(kwC) > 0 Then
            If InStr(1, receipt, kwC, vbTextCompare) > 0 Or InStr(1, core, kwC, vbTextCompare) > 0 Then
                ResolveLangSheet = "일문"
                Exit Function
            End If
        End If
    Next r
End Function

'=========================================================
' eqInfo 반환(기존 규칙 유지)
'=========================================================
Private Sub GetEqInfoForMethod(ByVal methodName As String, ByVal wsEq As Worksheet, ByRef outB As String, ByRef outC As String)

    outB = ""
    outC = ""

    Dim headerKey As String
    headerKey = ""

    If (InStr(1, methodName, "세탁", vbTextCompare) > 0) And (InStr(1, methodName, "손세탁", vbTextCompare) = 0) Then
        If ContainsAnyToken(methodName, Array("4B", "5B", "6B", "7B", "8B", "9B", "11B")) Then
            headerKey = "B형세탁기"
        Else
            headerKey = "A형세탁기"
        End If
        PickEqInfoByHeader wsEq, headerKey, outB, outC
        Exit Sub
    End If

    Dim keys As Variant: keys = Array("석유계", "퍼클로로에틸렌", "손세탁", "다리미질", "KS K 0642")
    Dim k As Variant
    For Each k In keys
        If InStr(1, methodName, CStr(k), vbTextCompare) > 0 Then
            PickEqInfoByHeader wsEq, CStr(k), outB, outC
            Exit Sub
        End If
    Next k
End Sub

Private Sub PickEqInfoByHeader(ByVal wsEq As Worksheet, ByVal headerKey As String, ByRef outB As String, ByRef outC As String)

    Dim lastCol As Long, c As Long
    lastCol = wsEq.Cells(1, wsEq.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If InStr(1, CStr(wsEq.Cells(1, c).Value), headerKey, vbTextCompare) > 0 Then
            outB = CStr(wsEq.Cells(1, c).Value)
            outC = PickRandomNonEmptyFromColumn(wsEq, c, 2)
            Exit Sub
        End If
    Next c
End Sub

Private Function PickRandomNonEmptyFromColumn(ByVal ws As Worksheet, ByVal col As Long, ByVal startRow As Long) As String
    Dim lastRow As Long, r As Long, v As String
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    Dim arr() As String, cnt As Long
    cnt = 0

    If lastRow < startRow Then
        PickRandomNonEmptyFromColumn = ""
        Exit Function
    End If

    For r = startRow To lastRow
        v = Trim$(CStr(ws.Cells(r, col).Value))
        If Len(v) > 0 Then
            cnt = cnt + 1
            ReDim Preserve arr(1 To cnt)
            arr(cnt) = v
        End If
    Next r

    If cnt = 0 Then
        PickRandomNonEmptyFromColumn = ""
        Exit Function
    End If

    Randomize
    PickRandomNonEmptyFromColumn = arr(WorksheetFunction.RandBetween(1, cnt))
End Function

Private Function ContainsAnyToken(ByVal s As String, ByVal tokens As Variant) As Boolean
    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
        If InStr(1, s, CStr(tokens(i)), vbTextCompare) > 0 Then
            ContainsAnyToken = True
            Exit Function
        End If
    Next i
    ContainsAnyToken = False
End Function

'=========================================================
' 날짜 포맷
'=========================================================
Private Function FormatDate_yyyy_mm_dd(ByVal v As Variant) As String
    If IsDate(v) Then
        FormatDate_yyyy_mm_dd = Format$(CDate(v), "yyyy-mm-dd")
    Else
        FormatDate_yyyy_mm_dd = ""
    End If
End Function

'=========================================================
' A열 키 파싱: @C2312600337@1,23  -> receipt/core/sample/method
' receipt: C231-26-00337
' core:    C2312600337
'=========================================================
Public Function ParseAKey(ByVal s As String, ByRef receipt As String, ByRef core As String, ByRef sampleNo As Long, ByRef methodNo As String) As Boolean
    On Error GoTo Fail

    Dim p1 As Long, p2 As Long, arr() As String

    p1 = InStr(1, s, "@", vbTextCompare)
    p2 = InStr(p1 + 1, s, "@", vbTextCompare)
    If p1 = 0 Or p2 = 0 Then GoTo Fail

    core = Mid$(s, p1 + 1, p2 - p1 - 1)
    arr = Split(Replace(Mid$(s, p2 + 1), " ", ""), ",")

    sampleNo = CLng(Val(arr(0)))
    methodNo = Trim$(CStr(arr(1)))

    receipt = Left$(core, 4) & "-" & Mid$(core, 5, 2) & "-" & Right$(core, 5)

    ParseAKey = True
    Exit Function

Fail:
    ParseAKey = False
End Function

'=========================================================
' txt에서 시험방법명 가져오기
'=========================================================
Private Function GetMethodNameFromTxt(ByVal core As String, ByVal methodNo As String) As String
    On Error GoTo Fail

    Dim f As String
    f = ThisWorkbook.Path & "\testNumber\2025\" & core & ".txt"
    If Dir$(f) = "" Then
        GetMethodNameFromTxt = methodNo
        Exit Function
    End If

    Dim ff As Integer: ff = FreeFile
    Open f For Input As #ff

    Dim line As String, a() As String
    Do While Not EOF(ff)
        Line Input #ff, line
        line = Trim$(line)
        If Len(line) = 0 Then GoTo ContinueLine

        If InStr(line, vbTab) > 0 Then
            a = Split(line, vbTab)
        ElseIf InStr(line, ",") > 0 Then
            a = Split(line, ",")
        Else
            GoTo ContinueLine
        End If

        If UBound(a) >= 3 Then
            If Trim$(CStr(a(1))) = Trim$(methodNo) Then
                GetMethodNameFromTxt = InsertBeforeParen(CStr(a(2)), CStr(a(3)))
                Close #ff
                Exit Function
            End If
        End If
ContinueLine:
    Loop

    Close #ff
    GetMethodNameFromTxt = methodNo
    Exit Function

Fail:
    On Error Resume Next
    Close #ff
    GetMethodNameFromTxt = methodNo
End Function

Private Function InsertBeforeParen(ByVal text As String, ByVal ins As String) As String
    Dim p As Long: p = InStr(1, text, "(", vbTextCompare)
    If p > 0 And Len(Trim$(ins)) > 0 Then
        InsertBeforeParen = Trim$(Left$(text, p - 1)) & " " & Trim$(ins) & " " & Mid$(text, p)
    Else
        InsertBeforeParen = text
    End If
End Function

Private Function RemoveTrailingColon(ByVal s As String) As String
    s = Trim$(s)
    If Right$(s, 1) = ":" Then
        RemoveTrailingColon = Trim$(Left$(s, Len(s) - 1))
    Else
        RemoveTrailingColon = s
    End If
End Function

'=========================================================
' 정렬/그룹 유틸
'=========================================================
Public Function DictKeysToArray(ByVal dict As Object) As String()
    Dim arr() As String
    Dim i As Long: i = 0
    Dim k As Variant
    ReDim arr(0 To dict.Count - 1)
    For Each k In dict.keys
        arr(i) = CStr(k)
        i = i + 1
    Next k
    DictKeysToArray = arr
End Function

Private Sub QuickSortKeys(ByRef a() As String, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim pivot As String, tmp As String

    i = lo: j = hi
    pivot = a((lo + hi) \ 2)

    Do While i <= j
        Do While CompareKey(a(i), pivot) < 0
            i = i + 1
        Loop
        Do While CompareKey(a(j), pivot) > 0
            j = j - 1
        Loop

        If i <= j Then
            tmp = a(i): a(i) = a(j): a(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop

    If lo < j Then QuickSortKeys a, lo, j
    If i < hi Then QuickSortKeys a, i, hi
End Sub

Private Function CompareKey(ByVal k1 As String, ByVal k2 As String) As Long
    Dim p1() As String, p2() As String
    Dim r1 As String, r2 As String
    Dim m1 As Long, m2 As Long

    p1 = Split(k1, "|")
    p2 = Split(k2, "|")

    r1 = Replace(CStr(p1(0)), "-", "")
    r2 = Replace(CStr(p2(0)), "-", "")

    If r1 < r2 Then
        CompareKey = -1: Exit Function
    ElseIf r1 > r2 Then
        CompareKey = 1: Exit Function
    End If

    m1 = CLng(Val(CStr(p1(1))))
    m2 = CLng(Val(CStr(p2(1))))

    If m1 < m2 Then
        CompareKey = -1
    ElseIf m1 > m2 Then
        CompareKey = 1
    Else
        CompareKey = 0
    End If
End Function

Private Function GetFirstRowFromGroup(ByVal grp As Object) As Long
    Dim s As Long, minS As Long
    minS = 999999
    Dim k As Variant
    For Each k In grp.keys
        s = CLng(k)
        If s < minS Then minS = s
    Next k
    GetFirstRowFromGroup = CLng(grp(minS)(1))
End Function

Private Function GetSortedSampleNos(ByVal grp As Object) As Long()
    Dim k As Variant, i As Long
    Dim arr() As Long

    ReDim arr(0 To grp.Count - 1)
    i = 0
    For Each k In grp.keys
        arr(i) = CLng(k)
        i = i + 1
    Next k

    QuickSortLong arr, LBound(arr), UBound(arr)
    GetSortedSampleNos = arr
End Function

Private Sub QuickSortLong(ByRef a() As Long, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim pivot As Long, tmp As Long

    i = lo: j = hi
    pivot = a((lo + hi) \ 2)

    Do While i <= j
        Do While a(i) < pivot: i = i + 1: Loop
        Do While a(j) > pivot: j = j - 1: Loop
        If i <= j Then
            tmp = a(i): a(i) = a(j): a(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop

    If lo < j Then QuickSortLong a, lo, j
    If i < hi Then QuickSortLong a, i, hi
End Sub


'=========================================================
' roundingInfo에서 methodName에 맞는 수치맺음(step) 찾기
'  - 1행: 0.5, 0.1, 0.2 ... (숫자)
'  - 2행~: 해당 수치맺음에 해당하는 "시험방법 키워드" 목록
'  - methodName에 키워드가 포함되면 그 열의 step 선택
'=========================================================
Private Function GetRoundingStep(ByVal methodName As String, ByVal wsRound As Worksheet, ByVal defaultStep As Double) As Double

    GetRoundingStep = defaultStep

    Dim lastCol As Long, c As Long
    lastCol = wsRound.Cells(1, wsRound.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Function

    Dim lastRow As Long
    lastRow = wsRound.Cells(wsRound.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    For c = 1 To lastCol

        Dim stepV As Variant
        stepV = wsRound.Cells(1, c).Value
        If Not IsNumeric(stepV) Then GoTo ContinueCol
        If CDbl(stepV) <= 0 Then GoTo ContinueCol

        Dim r As Long
        For r = 2 To lastRow
            Dim kw As String
            kw = Trim$(CStr(wsRound.Cells(r, c).Value))
            If Len(kw) > 0 Then
                If InStr(1, methodName, kw, vbTextCompare) > 0 Then
                    GetRoundingStep = CDbl(stepV)
                    Exit Function
                End If
            End If
        Next r

ContinueCol:
    Next c
End Function

Private Function RoundToStep(ByVal v As Double, ByVal stepV As Double) As Double

    If stepV <= 0 Then
        RoundToStep = v
        Exit Function
    End If

    ' 0은 그대로
    If v = 0 Then
        RoundToStep = 0
        Exit Function
    End If

    ' ★ 핵심 규칙 ★
    ' step=0.5이고 |v| < step 인 경우 → 무조건 ±step
    If stepV = 0.5 Then
        If Abs(v) < stepV Then
            If v > 0 Then
                RoundToStep = stepV
            Else
                RoundToStep = -stepV
            End If
            Exit Function
        End If
    End If

    ' 그 외는 기존 반올림
    Dim q As Double
    q = v / stepV

    If q >= 0 Then
        RoundToStep = Int(q + 0.5) * stepV
    Else
        RoundToStep = -Int(Abs(q) + 0.5) * stepV
    End If

End Function


'=========================================================
' 주소 문자열(예: "C16")을 받아 한 줄 아래("C17") 주소 반환
'=========================================================
Private Function OffsetCell1Down(ByVal addr As String) As String
    Dim rg As Range
    Set rg = Range(addr)
    OffsetCell1Down = rg.Offset(1, 0).Address(False, False)
End Function

'=========================================================
' 시료번호 표시 위치 (페이지 내 6개)
'=========================================================
Private Function SampleNoPosMap_UpTo6() As Variant
    Dim m(0 To 5) As String
    m(0) = "C9"
    m(1) = "G9"
    m(2) = "K9"
    m(3) = "C18"
    m(4) = "G18"
    m(5) = "K18"
    SampleNoPosMap_UpTo6 = m
End Function

Private Sub SetPctSafe(ByVal ws As Worksheet, ByVal addr As String, ByVal v As Double)
    With ws.Range(addr)
        If .MergeCells Then
            With .MergeArea.Cells(1, 1)
                .NumberFormat = "+0.0;-0.0;0.0"
                .Value = v
            End With
        Else
            .NumberFormat = "+0.0;-0.0;0.0"
            .Value = v
        End If
    End With
End Sub

