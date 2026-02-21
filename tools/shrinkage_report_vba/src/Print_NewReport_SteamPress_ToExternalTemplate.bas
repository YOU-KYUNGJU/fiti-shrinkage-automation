Attribute VB_Name = "Module3"
'===============================
' Module: modShrink_SteamPress
' 목적:
' - A열 키가 "@core@sample,method" 또는 "@core@sample,method,#0/#1" 형태일 때
' - "#0/#1"가 붙은 경우를 "스팀프레스(1시료=3회 측정)"로 간주하여
'   같은 시료번호의 3회 측정값을 기존 시험 1~3(또는 4~6) 슬롯에 채우고,
'   시료번호는 G9(1번째 시료), G18(2번째 시료)에 표시
' - 스팀프레스 시료가 3~4개면 2개씩 페이지를 나눠 출력
'
' 주의:
' - 기존 모듈과 중복 선언(함수명/전역변수) 피하려고 이 모듈은 Steam_ 접두어 사용
' - 기존 템플릿 좌표/계산 규칙은 질문에 첨부한 코드와 동일하게 사용
'===============================
Option Explicit

' =========================
' 현재(원본) 파일 시트명
' =========================
Private Const SHEET_RAW As String = "Rawdata"
Private Const SHEET_EQINFO As String = "eqInfo"
Private Const SHEET_CODEINFO As String = "CodeInfo"
Private Const SHEET_ROUNDING As String = "roundingInfo"

' =========================
' 외부(새 분석표) 파일명 (같은 폴더)
' =========================
Private Const NEW_REPORT_FILE As String = "치수변화율-원단시험분석표2_v1.1_20260128.xlsx" '필요시 수정

' =========================
' Rawdata 컬럼(중요)
' =========================
Private Const COL_AKEY As Long = 1      'A열: @접수@시료,방법(,#0/#1 가능)
Private Const COL_BEFORE As Long = 2    'B열: 기준값(전/Before/SPEC)
Private Const COL_WASHCNT As Long = 8   'H열: 세탁횟수
Private Const COL_DATE As Long = 10     'J열: 날짜(대표행)

' 측정값(시험후 After)
Private Const COL_LEN_K As Long = 11    'K
Private Const COL_LEN_L As Long = 12    'L
Private Const COL_LEN_M As Long = 13    'M
Private Const COL_WID_N As Long = 14    'N
Private Const COL_WID_O As Long = 15    'O
Private Const COL_WID_P As Long = 16    'P

'=========================
' 메인(스팀프레스 대응)
'=========================
Public Sub Print_NewReport_SteamPress_ToExternalTemplate()

    Dim wsR As Worksheet, wsEq As Worksheet, wsCode As Worksheet, wsRound As Worksheet
    Set wsR = ThisWorkbook.Worksheets(SHEET_RAW)
    Set wsEq = ThisWorkbook.Worksheets(SHEET_EQINFO)
    Set wsCode = ThisWorkbook.Worksheets(SHEET_CODEINFO)
    Set wsRound = ThisWorkbook.Worksheets(SHEET_ROUNDING)

    If ActiveSheet.Name <> SHEET_RAW Then
        MsgBox "Rawdata 시트에서 출력할 행을 선택한 뒤 실행하세요.", vbExclamation
        Exit Sub
    End If
    If TypeName(Selection) <> "Range" Or Selection.Rows.Count = 0 Then
        MsgBox "출력할 행(범위)을 먼저 선택하세요.", vbExclamation
        Exit Sub
    End If

    '---------------------------------------
    ' 1) 선택 영역 파싱
    '  - normal: key=receipt|method  -> sampleNo -> rows(Collection)
    '  - steam : key=receipt|method  -> sampleNo -> trials(Dictionary(1..3 -> rowNum))
    '---------------------------------------
    Dim normalGroups As Object: Set normalGroups = CreateObject("Scripting.Dictionary")
    Dim steamGroups As Object:  Set steamGroups = CreateObject("Scripting.Dictionary")
    Dim coreByKey As Object:    Set coreByKey = CreateObject("Scripting.Dictionary")

    Dim r As Range, rowNum As Long
    Dim receipt As String, core As String, methodNo As String
    Dim sampleNo As Long, key As String
    Dim isSteam As Boolean, trialNo As Long

    For Each r In Selection.Rows
        rowNum = r.Row

        If Steam_ParseAKeyEx(wsR.Cells(rowNum, COL_AKEY).Value, receipt, core, sampleNo, methodNo, isSteam, trialNo) Then
            If sampleNo >= 1 Then
                key = receipt & "|" & methodNo

                If Not coreByKey.Exists(key) Then coreByKey.Add key, core

                ' 무# 포함 3회 측정 세트이므로 전부 steamGroups로 적재
                If Not steamGroups.Exists(key) Then steamGroups.Add key, CreateObject("Scripting.Dictionary")
                Dim dS As Object: Set dS = steamGroups(key) ' sampleNo -> trials(dict)
                
                If Not dS.Exists(sampleNo) Then dS.Add sampleNo, CreateObject("Scripting.Dictionary")
                Dim trials As Object: Set trials = dS(sampleNo) ' 1..3 -> rowNum
                
                If trialNo < 1 Or trialNo > 3 Then trialNo = 1
                trials(trialNo) = rowNum

            End If
        End If
    Next r

    If normalGroups.Count = 0 And steamGroups.Count = 0 Then
        MsgBox "선택 범위에서 유효한 A열 키(@접수@시료,방법)를 찾지 못했습니다.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim wbNew As Workbook, newPath As String
    newPath = ThisWorkbook.Path & "\" & NEW_REPORT_FILE

    On Error GoTo CleanFail

    If Dir$(newPath) = "" Then
        MsgBox "새 분석표 파일을 찾지 못했습니다:" & vbCrLf & newPath, vbExclamation
        GoTo CleanExit
    End If

    Set wbNew = Workbooks.Open(newPath, ReadOnly:=False)

    Dim measureMap As Variant, resultMap As Variant
    measureMap = Steam_SamplePosMap_UpTo6_Measures()
    resultMap = Steam_SamplePosMap_UpTo6_Results()

    '---------------------------------------
    ' 2) 출력 (normal 먼저, steam 다음) - 필요하면 순서 바꿔도 됨
    '---------------------------------------
    If normalGroups.Count > 0 Then
        Steam_PrintNormalGroups normalGroups, coreByKey, wsR, wsEq, wsCode, wsRound, wbNew, measureMap, resultMap
    End If

    If steamGroups.Count > 0 Then
        Steam_PrintSteamGroups steamGroups, coreByKey, wsR, wsEq, wsCode, wsRound, wbNew, measureMap, resultMap
    End If

CleanExit:
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "완료: 새 분석표(국문/영문/일문)에 채운 후 인쇄했습니다.", vbInformation
    Exit Sub

CleanFail:
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "오류: " & Err.Description, vbExclamation
End Sub

'=========================================================
' [Normal] 기존 로직(샘플당 1행) - 원본과 동일하게 6개/페이지
'=========================================================
Private Sub Steam_PrintNormalGroups(ByVal groups As Object, ByVal coreByKey As Object, _
                                   ByVal wsR As Worksheet, ByVal wsEq As Worksheet, ByVal wsCode As Worksheet, ByVal wsRound As Worksheet, _
                                   ByVal wbNew As Workbook, ByVal measureMap As Variant, ByVal resultMap As Variant)

    Dim keysArr() As String
    keysArr = Steam_DictKeysToArray(groups)
    Steam_QuickSortKeys keysArr, LBound(keysArr), UBound(keysArr)

    Dim i As Long, key As String
    Dim grp As Object
    Dim receipt As String, methodNo As String, core As String
    Dim firstRow As Long, methodName As String
    Dim langSheetName As String, wsOut As Worksheet
    Dim eqB As String, eqC As String
    Dim roundStep As Double

    For i = LBound(keysArr) To UBound(keysArr)

        key = keysArr(i)
        Set grp = groups(key)

        receipt = Split(key, "|")(0)
        methodNo = Split(key, "|")(1)
        core = coreByKey(key)

        firstRow = Steam_GetFirstRowFromGroup_Normal(grp)

        methodName = Steam_GetMethodNameFromTxt(core, methodNo)
        methodName = Steam_RemoveTrailingColon(methodName)

        roundStep = Steam_GetRoundingStep(methodName, wsRound, 0.1)

        langSheetName = Steam_ResolveLangSheet(wsCode, receipt, core)
        Set wsOut = wbNew.Worksheets(langSheetName)

        Steam_GetEqInfoForMethod methodName, wsEq, eqB, eqC

        Dim sampleList() As Long
        sampleList = Steam_GetSortedSampleNos(grp)

        Dim totalSamples As Long
        totalSamples = UBound(sampleList) - LBound(sampleList) + 1

        Dim pageIdx As Long, startIdx As Long, endIdx As Long
        Dim sampleNoMap As Variant
        sampleNoMap = Steam_SampleNoPosMap_UpTo6()

        For pageIdx = 0 To (totalSamples - 1) \ 6

            Steam_ClearNewReport wsOut
            ' 공통값
            Steam_SetValueSafe wsOut, "K4", wsR.Range("I2").Value
            Steam_SetValueSafe wsOut, "B5", receipt
            Steam_SetValueSafe wsOut, "B6", methodName
            Steam_SetValueSafe wsOut, "B27", eqB
            Steam_SetValueSafe wsOut, "D27", eqC

            startIdx = pageIdx * 6
            endIdx = WorksheetFunction.Min(startIdx + 5, totalSamples - 1)

            Dim sIdx As Long, localIdx As Long, sNo As Long, rr0 As Long
            Dim specVal As Variant, washCnt As Variant
            Dim beforeV As Double

            For sIdx = startIdx To endIdx

                localIdx = sIdx - startIdx '0~5
                sNo = sampleList(LBound(sampleList) + sIdx)
                rr0 = CLng(grp(sNo)(1)) '대표행

                Steam_SetValueSafe wsOut, sampleNoMap(localIdx), sNo

                specVal = wsR.Cells(rr0, COL_BEFORE).Value
                washCnt = wsR.Cells(rr0, COL_WASHCNT).Value
                Steam_SetSpecTextCells wsOut, specVal, washCnt, eqB

                If IsNumeric(specVal) And specVal <> 0 Then
                    beforeV = CDbl(specVal)
                Else
                    beforeV = 0#
                End If

                Dim L1 As Double, L2 As Double, L3 As Double
                Dim W1 As Double, W2 As Double, W3 As Double
                L1 = Steam_NzD(wsR.Cells(rr0, COL_LEN_K).Value)
                L2 = Steam_NzD(wsR.Cells(rr0, COL_LEN_L).Value)
                L3 = Steam_NzD(wsR.Cells(rr0, COL_LEN_M).Value)
                W1 = Steam_NzD(wsR.Cells(rr0, COL_WID_N).Value)
                W2 = Steam_NzD(wsR.Cells(rr0, COL_WID_O).Value)
                W3 = Steam_NzD(wsR.Cells(rr0, COL_WID_P).Value)

                Steam_SetValueSafe wsOut, measureMap(localIdx)(0), Steam_Fmt1(L1)
                Steam_SetValueSafe wsOut, measureMap(localIdx)(1), Steam_Fmt1(L2)
                Steam_SetValueSafe wsOut, measureMap(localIdx)(2), Steam_Fmt1(L3)
                Steam_SetValueSafe wsOut, measureMap(localIdx)(3), Steam_Fmt1(W1)
                Steam_SetValueSafe wsOut, measureMap(localIdx)(4), Steam_Fmt1(W2)
                Steam_SetValueSafe wsOut, measureMap(localIdx)(5), Steam_Fmt1(W3)

                Dim Lx As Double, Ly As Double, Lz As Double, Lm As Double
                Dim Wx As Double, Wy As Double, Wz As Double, Wm As Double
                Steam_CalcShrinkXYZ beforeV, L1, L2, L3, Lx, Ly, Lz, Lm
                Steam_CalcShrinkXYZ beforeV, W1, W2, W3, Wx, Wy, Wz, Wm

                Steam_SetPctSafe wsOut, resultMap(localIdx)(0), Steam_FmtPct1(Lx)
                Steam_SetPctSafe wsOut, resultMap(localIdx)(1), Steam_FmtPct1(Ly)
                Steam_SetPctSafe wsOut, resultMap(localIdx)(2), Steam_FmtPct1(Lz)
                Steam_SetPctSafe wsOut, resultMap(localIdx)(3), Steam_FmtPct1(Lm)

                Steam_SetPctSafe wsOut, resultMap(localIdx)(4), Steam_FmtPct1(Wx)
                Steam_SetPctSafe wsOut, resultMap(localIdx)(5), Steam_FmtPct1(Wy)
                Steam_SetPctSafe wsOut, resultMap(localIdx)(6), Steam_FmtPct1(Wz)
                Steam_SetPctSafe wsOut, resultMap(localIdx)(7), Steam_FmtPct1(Wm)

                Steam_SetPctSafe wsOut, Steam_OffsetCell1Down(resultMap(localIdx)(3)), Steam_FmtPct1(Steam_RoundToStep(Lm, roundStep))
                Steam_SetPctSafe wsOut, Steam_OffsetCell1Down(resultMap(localIdx)(7)), Steam_FmtPct1(Steam_RoundToStep(Wm, roundStep))

            Next sIdx

            wsOut.PrintOut
        Next pageIdx
    Next i

End Sub
'=========================================================
' [SteamPress 완성본]
'  - 1시료 = 3회(Trial 1~3) 결과를 "한 시험편 블록"에 다음처럼 채움
'    (상단 시험편)
'      1회차: C13~C16(경사), E13~E16(위사)
'      2회차: G13~G16(경사), I13~I16(위사)
'      3회차: K13~K16(경사), M13~M16(위사)
'      평균의 평균: G17(경사), I17(위사) = (각 회차 평균값들의 평균) 후 수치맺음 적용
'
'    (하단 시험편)
'      1회차: C22~C25(경사), E22~E25(위사)
'      2회차: G22~G25(경사), I22~I25(위사)
'      3회차: K22~K25(경사), M22~M25(위사)
'      평균의 평균: G26(경사), I26(위사)
'
'  - 페이지당 시험편(시료) 2개 출력
'  - 시료번호 표기(요청): 첫번째 시료 -> G9, 두번째 시료 -> G18
'
' 전제:
'  - groups(key) 구조: key=receipt|method  -> sampleNo -> trials(Dictionary(1..3 -> rowNum))
'  - 핵심 헬퍼는 모듈에 이미 존재:
'    Steam_GetMethodNameFromTxt, Steam_RemoveTrailingColon, Steam_GetEqInfoForMethod,
'    Steam_GetRoundingStep, Steam_ResolveLangSheet, Steam_ClearNewReport,
'    Steam_SetValueSafe, Steam_SetPctSafe, Steam_SetSpecTextCells,
'    Steam_CalcShrinkXYZ, Steam_RoundToStep, Steam_FmtPct1, Steam_Fmt1, Steam_NzD,
'    Steam_DictKeysToArray, Steam_QuickSortKeys, Steam_GetSortedSampleNos_FromDict,
'    Steam_GetAnyTrialRow, Steam_GetTrialRow
'=========================================================
Private Sub Steam_PrintSteamGroups(ByVal groups As Object, ByVal coreByKey As Object, _
                                  ByVal wsR As Worksheet, ByVal wsEq As Worksheet, ByVal wsCode As Worksheet, ByVal wsRound As Worksheet, _
                                  ByVal wbNew As Workbook, ByVal measureMap As Variant, ByVal resultMap As Variant)

    Dim keysArr() As String
    keysArr = Steam_DictKeysToArray(groups)
    Steam_QuickSortKeys keysArr, LBound(keysArr), UBound(keysArr)

    Dim i As Long, key As String
    Dim receipt As String, methodNo As String, core As String
    Dim methodName As String, langSheetName As String
    Dim wsOut As Worksheet
    Dim eqB As String, eqC As String
    Dim roundStep As Double

    For i = LBound(keysArr) To UBound(keysArr)

        key = keysArr(i)
        receipt = Split(key, "|")(0)
        methodNo = Split(key, "|")(1)
        core = coreByKey(key)

        methodName = Steam_GetMethodNameFromTxt(core, methodNo)
        methodName = Steam_RemoveTrailingColon(methodName)

        roundStep = Steam_GetRoundingStep(methodName, wsRound, 0.1)

        langSheetName = Steam_ResolveLangSheet(wsCode, receipt, core)
        Set wsOut = wbNew.Worksheets(langSheetName)

        Steam_GetEqInfoForMethod methodName, wsEq, eqB, eqC

        Dim dS As Object
        Set dS = groups(key) ' sampleNo -> trials(dict)

        Dim sampleList() As Long
        sampleList = Steam_GetSortedSampleNos_FromDict(dS)

        Dim totalSamples As Long
        totalSamples = UBound(sampleList) - LBound(sampleList) + 1

        Dim pageIdx As Long, startIdx As Long, endIdx As Long

        ' 페이지당 2시료(=시험편 2개)
        For pageIdx = 0 To (totalSamples - 1) \ 2

            Steam_ClearNewReport wsOut

            ' ----- 공통값 -----
            Steam_SetValueSafe wsOut, "K4", wsR.Range("I2").Value
            Steam_SetValueSafe wsOut, "B5", receipt
            Steam_SetValueSafe wsOut, "B6", methodName
            Steam_SetValueSafe wsOut, "B27", eqB
            Steam_SetValueSafe wsOut, "D27", eqC

            startIdx = pageIdx * 2
            endIdx = WorksheetFunction.Min(startIdx + 1, totalSamples - 1)

            Dim sIdx As Long
            For sIdx = startIdx To endIdx

                Dim sampleNo As Long
                sampleNo = sampleList(LBound(sampleList) + sIdx)

                Dim sampleOnPage As Long
                sampleOnPage = sIdx - startIdx  '0(상단), 1(하단)

                ' ----- 시료번호 표기(요청) -----
                If sampleOnPage = 0 Then
                    Steam_SetValueSafe wsOut, "G9", sampleNo
                Else
                    Steam_SetValueSafe wsOut, "G18", sampleNo
                End If

                Dim trials As Object
                Set trials = dS(sampleNo) ' 1..3 -> rowNum

                ' ----- SPEC/세탁횟수는 대표행에서 1회만 -----
                Dim rrSpec As Long
                rrSpec = Steam_GetAnyTrialRow(trials)
                If rrSpec = 0 Then GoTo ContinueSample

                Dim specVal As Variant, washCnt As Variant
                specVal = wsR.Cells(rrSpec, COL_BEFORE).Value
                washCnt = wsR.Cells(rrSpec, COL_WASHCNT).Value
                Steam_SetSpecTextCells wsOut, specVal, washCnt, eqB

                Dim beforeV As Double
                If IsNumeric(specVal) And specVal <> 0 Then
                    beforeV = CDbl(specVal)
                Else
                    beforeV = 0#
                End If

                ' ----- 블록 기준(상단/하단) -----
                Dim baseRow As Long, avgRow As Long
                If sampleOnPage = 0 Then
                    baseRow = 13   '상단 시작
                    avgRow = 17    '상단 평균의 평균
                Else
                    baseRow = 22   '하단 시작
                    avgRow = 26    '하단 평균의 평균
                End If

                ' 3회차 평균을 모아서 평균의 평균 계산
                Dim LmArr(1 To 3) As Double, WmArr(1 To 3) As Double
                Dim hasTrial(1 To 3) As Boolean

                Dim t As Long
                For t = 1 To 3

                    Dim rr0 As Long
                    rr0 = Steam_GetTrialRow(trials, t)
                    If rr0 = 0 Then GoTo ContinueTrial

                    hasTrial(t) = True

                    ' ---- 측정값 읽기 ----
                    Dim L1 As Double, L2 As Double, L3 As Double
                    Dim W1 As Double, W2 As Double, W3 As Double
                    L1 = Steam_NzD(wsR.Cells(rr0, COL_LEN_K).Value)
                    L2 = Steam_NzD(wsR.Cells(rr0, COL_LEN_L).Value)
                    L3 = Steam_NzD(wsR.Cells(rr0, COL_LEN_M).Value)
                    W1 = Steam_NzD(wsR.Cells(rr0, COL_WID_N).Value)
                    W2 = Steam_NzD(wsR.Cells(rr0, COL_WID_O).Value)
                    W3 = Steam_NzD(wsR.Cells(rr0, COL_WID_P).Value)

                    ' ---- 해당 trial의 출력 열 결정 ----
                    Dim colWarp As String, colWeft As String
                    Select Case t
                        Case 1
                            colWarp = "C": colWeft = "E"
                        Case 2
                            colWarp = "G": colWeft = "I"
                        Case 3
                            colWarp = "K": colWeft = "M"
                    End Select

                    ' ---- (선택) 측정값도 템플릿에 기입해야 하면 아래 매핑이 필요 ----
                    ' 질문에서 요구는 "%결과 셀" 위주였지만,
                    ' 기존과 동일하게 "측정값(길이/폭)"도 넣고 싶으면
                    ' t=1..3을 slotIdx=0..2(상단) 또는 3..5(하단)로 매핑하여 measureMap에 넣는다.
                    Dim slotIdx As Long
                    slotIdx = sampleOnPage * 3 + (t - 1) '0..2 또는 3..5

                    Steam_SetValueSafe wsOut, measureMap(slotIdx)(0), Steam_Fmt1(L1)
                    Steam_SetValueSafe wsOut, measureMap(slotIdx)(1), Steam_Fmt1(L2)
                    Steam_SetValueSafe wsOut, measureMap(slotIdx)(2), Steam_Fmt1(L3)
                    Steam_SetValueSafe wsOut, measureMap(slotIdx)(3), Steam_Fmt1(W1)
                    Steam_SetValueSafe wsOut, measureMap(slotIdx)(4), Steam_Fmt1(W2)
                    Steam_SetValueSafe wsOut, measureMap(slotIdx)(5), Steam_Fmt1(W3)

                    ' ---- 치수변화율 계산 ----
                    Dim Lx As Double, Ly As Double, Lz As Double, Lm As Double
                    Dim Wx As Double, Wy As Double, Wz As Double, Wm As Double
                    Steam_CalcShrinkXYZ beforeV, L1, L2, L3, Lx, Ly, Lz, Lm
                    Steam_CalcShrinkXYZ beforeV, W1, W2, W3, Wx, Wy, Wz, Wm

                    LmArr(t) = Lm
                    WmArr(t) = Wm

                    ' ---- % 결과(요구 셀에 고정 입력) ----
                    ' 경사 x,y,z
                    Steam_SetPctSafe wsOut, colWarp & baseRow, Steam_FmtPct1(Lx)
                    Steam_SetPctSafe wsOut, colWarp & (baseRow + 1), Steam_FmtPct1(Ly)
                    Steam_SetPctSafe wsOut, colWarp & (baseRow + 2), Steam_FmtPct1(Lz)
                    ' 경사 평균
                    Steam_SetPctSafe wsOut, colWarp & (baseRow + 3), Steam_FmtPct1(Lm)

                    ' 위사 x,y,z
                    Steam_SetPctSafe wsOut, colWeft & baseRow, Steam_FmtPct1(Wx)
                    Steam_SetPctSafe wsOut, colWeft & (baseRow + 1), Steam_FmtPct1(Wy)
                    Steam_SetPctSafe wsOut, colWeft & (baseRow + 2), Steam_FmtPct1(Wz)
                    ' 위사 평균
                    Steam_SetPctSafe wsOut, colWeft & (baseRow + 3), Steam_FmtPct1(Wm)

ContinueTrial:
                Next t

                ' ----- 평균의 평균(G17/I17 or G26/I26) -----
                Dim cnt As Long: cnt = 0
                Dim sumLm As Double: sumLm = 0#
                Dim sumWm As Double: sumWm = 0#

                For t = 1 To 3
                    If hasTrial(t) Then
                        cnt = cnt + 1
                        sumLm = sumLm + LmArr(t)
                        sumWm = sumWm + WmArr(t)
                    End If
                Next t

                If cnt > 0 Then
                    Dim grandLm As Double, grandWm As Double
                    grandLm = sumLm / cnt
                    grandWm = sumWm / cnt

                    ' 수치맺음 적용 후 넣기(요청: 평균값에 대한 평균값)
                    Steam_SetPctSafe wsOut, "G" & avgRow, Steam_FmtPct1(Steam_RoundToStep(grandLm, roundStep))
                    Steam_SetPctSafe wsOut, "I" & avgRow, Steam_FmtPct1(Steam_RoundToStep(grandWm, roundStep))
                End If

ContinueSample:
            Next sIdx

            wsOut.PrintOut
        Next pageIdx

    Next i

End Sub

' resultMap(slotIdx) = Array(경사 x,y,z,avg,  위사 x,y,z,avg)  총 8개 셀주소
Private Sub Steam_ClearResultBlock(ByVal ws As Worksheet, ByVal oneSlotResultMap As Variant)
    Dim i As Long
    For i = LBound(oneSlotResultMap) To UBound(oneSlotResultMap)
        ws.Range(CStr(oneSlotResultMap(i))).ClearContents
    Next i
    ' 수치맺음(AVG 한 줄 아래)도 같이 지우기(있으면)
    ws.Range(Steam_OffsetCell1Down(CStr(oneSlotResultMap(3)))).ClearContents
    ws.Range(Steam_OffsetCell1Down(CStr(oneSlotResultMap(7)))).ClearContents
End Sub

'=========================================================
' A열 키 파싱(신규):
'  @core@sample,method          -> trial 1
'  @core@sample,method#0        -> trial 2
'  @core@sample,method#1        -> trial 3
'
' (과거형도 호환)
'  @core@sample,method,#0 / #1  -> trial 2/3
'=========================================================
Private Function Steam_ParseAKeyEx(ByVal s As String, _
                                  ByRef receipt As String, ByRef core As String, _
                                  ByRef sampleNo As Long, ByRef methodNo As String, _
                                  ByRef isSteam As Boolean, ByRef trialNo As Long) As Boolean
    On Error GoTo Fail

    Dim p1 As Long, p2 As Long
    p1 = InStr(1, s, "@", vbTextCompare)
    p2 = InStr(p1 + 1, s, "@", vbTextCompare)
    If p1 = 0 Or p2 = 0 Then GoTo Fail

    core = Mid$(s, p1 + 1, p2 - p1 - 1)
    receipt = Left$(core, 4) & "-" & Mid$(core, 5, 2) & "-" & Right$(core, 5)

    Dim tail As String
    tail = Replace(Mid$(s, p2 + 1), " ", "")   ' 예: "1,8#0" / "1,8" / "1,8,#0"

    ' sampleNo는 첫 번째 콤마 전까지
    Dim commaPos As Long
    commaPos = InStr(1, tail, ",")
    If commaPos = 0 Then GoTo Fail

    sampleNo = CLng(Val(Left$(tail, commaPos - 1)))

    ' 나머지(방법+옵션)
    Dim rest As String
    rest = Mid$(tail, commaPos + 1)           ' 예: "8#0" / "8,#0" / "8"

    ' 혹시 "8,#0"처럼 콤마가 한 번 더 있으면 제거해서 합침
    rest = Replace(rest, ",", "")             ' "8#0" 형태로 정리

    Dim sharpPos As Long
    sharpPos = InStr(1, rest, "#")

    methodNo = rest
    isSteam = True
    trialNo = 1

    If sharpPos > 0 Then
        methodNo = Left$(rest, sharpPos - 1)  ' "8"
        Dim idx As Long
        idx = CLng(Val(Mid$(rest, sharpPos + 1))) ' 0/1
        trialNo = idx + 2                         ' #0->2, #1->3
    End If

    methodNo = Trim$(methodNo)
    If Len(methodNo) = 0 Then GoTo Fail

    Steam_ParseAKeyEx = True
    Exit Function

Fail:
    Steam_ParseAKeyEx = False
End Function



'=========================================================
' txt에서 시험방법명 가져오기(원본과 동일)
'=========================================================
Private Function Steam_GetMethodNameFromTxt(ByVal core As String, ByVal methodNo As String) As String
    On Error GoTo Fail

    Dim f As String
    f = ThisWorkbook.Path & "\testNumber\2025\" & core & ".txt"
    If Dir$(f) = "" Then
        Steam_GetMethodNameFromTxt = methodNo
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
                Steam_GetMethodNameFromTxt = Steam_InsertBeforeParen(CStr(a(2)), CStr(a(3)))
                Close #ff
                Exit Function
            End If
        End If
ContinueLine:
    Loop

    Close #ff
    Steam_GetMethodNameFromTxt = methodNo
    Exit Function

Fail:
    On Error Resume Next
    Close #ff
    Steam_GetMethodNameFromTxt = methodNo
End Function

Private Function Steam_InsertBeforeParen(ByVal text As String, ByVal ins As String) As String
    Dim p As Long: p = InStr(1, text, "(", vbTextCompare)
    If p > 0 And Len(Trim$(ins)) > 0 Then
        Steam_InsertBeforeParen = Trim$(Left$(text, p - 1)) & " " & Trim$(ins) & " " & Mid$(text, p)
    Else
        Steam_InsertBeforeParen = text
    End If
End Function

Private Function Steam_RemoveTrailingColon(ByVal s As String) As String
    s = Trim$(s)
    If Right$(s, 1) = ":" Then
        Steam_RemoveTrailingColon = Trim$(Left$(s, Len(s) - 1))
    Else
        Steam_RemoveTrailingColon = s
    End If
End Function

'=========================================================
' eqInfo 반환(원본 규칙 유지)
'=========================================================
Private Sub Steam_GetEqInfoForMethod(ByVal methodName As String, ByVal wsEq As Worksheet, ByRef outB As String, ByRef outC As String)

    outB = ""
    outC = ""

    Dim headerKey As String
    headerKey = ""

    If (InStr(1, methodName, "세탁", vbTextCompare) > 0) And (InStr(1, methodName, "손세탁", vbTextCompare) = 0) Then
        If Steam_ContainsAnyToken(methodName, Array("4B", "5B", "6B", "7B", "8B", "9B", "11B")) Then
            headerKey = "B형세탁기"
        Else
            headerKey = "A형세탁기"
        End If
        Steam_PickEqInfoByHeader wsEq, headerKey, outB, outC
        Exit Sub
    End If

    Dim keys As Variant: keys = Array("석유계", "퍼클로로에틸렌", "손세탁", "다리미질", "KS K 0642")
    Dim k As Variant
    For Each k In keys
        If InStr(1, methodName, CStr(k), vbTextCompare) > 0 Then
            Steam_PickEqInfoByHeader wsEq, CStr(k), outB, outC
            Exit Sub
        End If
    Next k
End Sub

Private Sub Steam_PickEqInfoByHeader(ByVal wsEq As Worksheet, ByVal headerKey As String, ByRef outB As String, ByRef outC As String)

    Dim lastCol As Long, c As Long
    lastCol = wsEq.Cells(1, wsEq.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If InStr(1, CStr(wsEq.Cells(1, c).Value), headerKey, vbTextCompare) > 0 Then
            outB = CStr(wsEq.Cells(1, c).Value)
            outC = Steam_PickRandomNonEmptyFromColumn(wsEq, c, 2)
            Exit Sub
        End If
    Next c
End Sub

Private Function Steam_PickRandomNonEmptyFromColumn(ByVal ws As Worksheet, ByVal col As Long, ByVal startRow As Long) As String
    Dim lastRow As Long, r As Long, v As String
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    Dim arr() As String, cnt As Long
    cnt = 0

    If lastRow < startRow Then
        Steam_PickRandomNonEmptyFromColumn = ""
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
        Steam_PickRandomNonEmptyFromColumn = ""
        Exit Function
    End If

    Randomize
    Steam_PickRandomNonEmptyFromColumn = arr(WorksheetFunction.RandBetween(1, cnt))
End Function

Private Function Steam_ContainsAnyToken(ByVal s As String, ByVal tokens As Variant) As Boolean
    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
        If InStr(1, s, CStr(tokens(i)), vbTextCompare) > 0 Then
            Steam_ContainsAnyToken = True
            Exit Function
        End If
    Next i
    Steam_ContainsAnyToken = False
End Function

'=========================================================
' 언어 시트 결정(원본 규칙 유지)
'=========================================================
Private Function Steam_ResolveLangSheet(ByVal wsCode As Worksheet, ByVal receipt As String, ByVal core As String) As String

    Steam_ResolveLangSheet = "국문"

    Dim lastRow As Long, r As Long
    lastRow = wsCode.Cells(wsCode.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then Exit Function

    For r = 1 To lastRow
        Dim kwB As String, kwC As String
        kwB = Trim$(CStr(wsCode.Cells(r, 2).Value)) 'B
        kwC = Trim$(CStr(wsCode.Cells(r, 3).Value)) 'C

        If Len(kwB) > 0 Then
            If InStr(1, receipt, kwB, vbTextCompare) > 0 Or InStr(1, core, kwB, vbTextCompare) > 0 Then
                Steam_ResolveLangSheet = "영문"
                Exit Function
            End If
        End If

        If Len(kwC) > 0 Then
            If InStr(1, receipt, kwC, vbTextCompare) > 0 Or InStr(1, core, kwC, vbTextCompare) > 0 Then
                Steam_ResolveLangSheet = "일문"
                Exit Function
            End If
        End If
    Next r
End Function

'=========================================================
' ClearNewReport (원본 + 시료번호표시 6셀 초기화 포함)
'=========================================================
Private Sub Steam_ClearNewReport(ByVal ws As Worksheet)

    ws.Range("B5").ClearContents
    ws.Range("B6").ClearContents

    ws.Range("K4").ClearContents
    ws.Range("B27,D27").ClearContents

    ws.Range("B10,F10,J10,B19,F19,J19,G34,I34").ClearContents

    ws.Range("B13:B15,D13:D15,F13:F15,H13:H15,J13:J15,L13:L15").ClearContents
    ws.Range("B22:B24,D22:D24,F22:F24,H22:H24,J22:J24,L22:L24").ClearContents

    ws.Range("C13:C16,E13:E16,G13:G16,I13:I16,K13:K16,M13:M16").ClearContents
    ws.Range("C22:C25,E22:E25,G22:G25,I22:I25,K22:K25,M22:M25").ClearContents

    ws.Range("C17,E17,G17,I17,K17,M17").ClearContents
    ws.Range("C26,E26,G26,I26,K26,M26").ClearContents

    ws.Range("C9,G9,K9,C18,G18,K18").ClearContents
End Sub

'=========================================================
' SPEC 문구 치환 + G34/I34 숫자 (원본 유지)
'=========================================================
Private Sub Steam_SetSpecTextCells(ByVal ws As Worksheet, ByVal specValue As Variant, ByVal washCount As Variant, ByVal eqHeader As String)

    Dim specStr As String
    specStr = CStr(specValue)

    Dim n1 As Long
    If IsNumeric(washCount) Then
        n1 = CLng(washCount)
    Else
        n1 = 0
    End If

    Dim n2 As String
    n2 = Steam_OrdinalEn(n1)

    Dim baseTxt As String
    If InStr(1, eqHeader, "다리미질", vbTextCompare) > 0 Then
        baseTxt = "다리미질 후 (  SPEC mm )" & vbCrLf & "After Ironing"
    ElseIf (InStr(1, eqHeader, "퍼클로로에틸렌", vbTextCompare) > 0) Or _
           (InStr(1, eqHeader, "석유계", vbTextCompare) > 0) Then
        baseTxt = "N1 회 드라이클리닝 후 (  SPEC mm )" & vbCrLf & "After N2 Drycleaned"
    ElseIf InStr(1, eqHeader, "세탁", vbTextCompare) > 0 Then
        baseTxt = "N1 회 세탁 후 (  SPEC mm )" & vbCrLf & "After N2 Washing"
    Else
        baseTxt = "N1 회 세탁/드라이클리닝 후 (  SPEC mm )" & vbCrLf & "After N2 Washing/Drycleaned"
    End If

    Dim txt As String
    txt = Replace(Replace(Replace(baseTxt, "N1", CStr(n1)), "SPEC", specStr), "N2", n2)

    Dim addrsText As Variant, i As Long
    addrsText = Array("B10", "F10", "J10", "B19", "F19", "J19")

    For i = LBound(addrsText) To UBound(addrsText)
        Steam_SetValueSafe ws, CStr(addrsText(i)), txt
    Next i

    Steam_SetValueSafe ws, "G34", specValue
    Steam_SetValueSafe ws, "I34", specValue
End Sub

Private Function Steam_OrdinalEn(ByVal n As Long) As String
    If n <= 0 Then
        Steam_OrdinalEn = ""
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

    Steam_OrdinalEn = CStr(n) & suf
End Function

'=========================================================
' [계산] before + 3개의 after -> x,y,z,avg
'=========================================================
Private Sub Steam_CalcShrinkXYZ(ByVal beforeV As Double, ByVal a1 As Double, ByVal a2 As Double, ByVal a3 As Double, _
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

Private Function Steam_NzD(ByVal v As Variant) As Double
    If IsNumeric(v) Then Steam_NzD = CDbl(v) Else Steam_NzD = 0#
End Function

Private Function Steam_Fmt1(ByVal v As Variant) As String
    If IsNumeric(v) Then
        Steam_Fmt1 = Format$(CDbl(v), "0.0")
    Else
        Steam_Fmt1 = CStr(v)
    End If
End Function

Private Function Steam_FmtPct1(ByVal v As Double) As Double
    ' SetPctSafe에서 NumberFormat을 주니까 값은 Double로 넣는 편이 안전
    Steam_FmtPct1 = CDbl(Format$(v, "0.0"))
End Function

'=========================================================
' 병합셀 안전 쓰기
'=========================================================
Private Sub Steam_SetValueSafe(ByVal ws As Worksheet, ByVal addr As String, ByVal v As Variant)
    With ws.Range(addr)
        If .MergeCells Then
            .MergeArea.Cells(1, 1).Value = v
        Else
            .Value = v
        End If
    End With
End Sub

Private Sub Steam_SetPctSafe(ByVal ws As Worksheet, ByVal addr As String, ByVal v As Double)
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

'=========================================================
' 주소 문자열(예: "C16")을 받아 한 줄 아래("C17") 주소 반환
'=========================================================
Private Function Steam_OffsetCell1Down(ByVal addr As String) As String
    Dim rg As Range
    Set rg = Range(addr)
    Steam_OffsetCell1Down = rg.Offset(1, 0).Address(False, False)
End Function

'=========================================================
' 샘플 1~6 [측정값] 위치 매핑
'=========================================================
Private Function Steam_SamplePosMap_UpTo6_Measures() As Variant
    Dim m(0 To 5) As Variant
    m(0) = Array("B13", "B14", "B15", "D13", "D14", "D15")
    m(1) = Array("F13", "F14", "F15", "H13", "H14", "H15")
    m(2) = Array("J13", "J14", "J15", "L13", "L14", "L15")
    m(3) = Array("B22", "B23", "B24", "D22", "D23", "D24")
    m(4) = Array("F22", "F23", "F24", "H22", "H23", "H24")
    m(5) = Array("J22", "J23", "J24", "L22", "L23", "L24")
    Steam_SamplePosMap_UpTo6_Measures = m
End Function

'=========================================================
' 샘플 1~6 [결과값] 위치 매핑
'=========================================================
Private Function Steam_SamplePosMap_UpTo6_Results() As Variant
    Dim m(0 To 5) As Variant
    m(0) = Array("C13", "C14", "C15", "C16", "E13", "E14", "E15", "E16")
    m(1) = Array("G13", "G14", "G15", "G16", "I13", "I14", "I15", "I16")
    m(2) = Array("K13", "K14", "K15", "K16", "M13", "M14", "M15", "M16")
    m(3) = Array("C22", "C23", "C24", "C25", "E22", "E23", "E24", "E25")
    m(4) = Array("G22", "G23", "G24", "G25", "I22", "I23", "I24", "I25")
    m(5) = Array("K22", "K23", "K24", "K25", "M22", "M23", "M24", "M25")
    Steam_SamplePosMap_UpTo6_Results = m
End Function

'=========================================================
' 시료번호 표시 위치(기존 6개)
'=========================================================
Private Function Steam_SampleNoPosMap_UpTo6() As Variant
    Dim m(0 To 5) As String
    m(0) = "C9"
    m(1) = "G9"
    m(2) = "K9"
    m(3) = "C18"
    m(4) = "G18"
    m(5) = "K18"
    Steam_SampleNoPosMap_UpTo6 = m
End Function

'=========================================================
' 정렬/그룹 유틸
'=========================================================
Private Function Steam_DictKeysToArray(ByVal dict As Object) As String()
    Dim arr() As String, i As Long, k As Variant
    ReDim arr(0 To dict.Count - 1)
    i = 0
    For Each k In dict.keys
        arr(i) = CStr(k)
        i = i + 1
    Next k
    Steam_DictKeysToArray = arr
End Function

Private Sub Steam_QuickSortKeys(ByRef a() As String, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim pivot As String, tmp As String

    i = lo: j = hi
    pivot = a((lo + hi) \ 2)

    Do While i <= j
        Do While Steam_CompareKey(a(i), pivot) < 0
            i = i + 1
        Loop
        Do While Steam_CompareKey(a(j), pivot) > 0
            j = j - 1
        Loop

        If i <= j Then
            tmp = a(i): a(i) = a(j): a(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop

    If lo < j Then Steam_QuickSortKeys a, lo, j
    If i < hi Then Steam_QuickSortKeys a, i, hi
End Sub

Private Function Steam_CompareKey(ByVal k1 As String, ByVal k2 As String) As Long
    Dim p1() As String, p2() As String
    Dim r1 As String, r2 As String
    Dim m1 As Long, m2 As Long

    p1 = Split(k1, "|")
    p2 = Split(k2, "|")

    r1 = Replace(CStr(p1(0)), "-", "")
    r2 = Replace(CStr(p2(0)), "-", "")

    If r1 < r2 Then
        Steam_CompareKey = -1: Exit Function
    ElseIf r1 > r2 Then
        Steam_CompareKey = 1: Exit Function
    End If

    m1 = CLng(Val(CStr(p1(1))))
    m2 = CLng(Val(CStr(p2(1))))

    If m1 < m2 Then
        Steam_CompareKey = -1
    ElseIf m1 > m2 Then
        Steam_CompareKey = 1
    Else
        Steam_CompareKey = 0
    End If
End Function

Private Function Steam_GetFirstRowFromGroup_Normal(ByVal grp As Object) As Long
    ' grp: sampleNo -> Collection(rowNums)
    Dim s As Long, minS As Long, k As Variant
    minS = 999999
    For Each k In grp.keys
        s = CLng(k)
        If s < minS Then minS = s
    Next k
    Steam_GetFirstRowFromGroup_Normal = CLng(grp(minS)(1))
End Function

Private Function Steam_GetSortedSampleNos(ByVal grp As Object) As Long()
    ' grp: sampleNo -> Collection
    Dim k As Variant, i As Long
    Dim arr() As Long
    ReDim arr(0 To grp.Count - 1)
    i = 0
    For Each k In grp.keys
        arr(i) = CLng(k)
        i = i + 1
    Next k
    Steam_QuickSortLong arr, LBound(arr), UBound(arr)
    Steam_GetSortedSampleNos = arr
End Function

Private Function Steam_GetSortedSampleNos_FromDict(ByVal grp As Object) As Long()
    ' grp: sampleNo -> trials(dict)
    Dim k As Variant, i As Long
    Dim arr() As Long
    ReDim arr(0 To grp.Count - 1)
    i = 0
    For Each k In grp.keys
        arr(i) = CLng(k)
        i = i + 1
    Next k
    Steam_QuickSortLong arr, LBound(arr), UBound(arr)
    Steam_GetSortedSampleNos_FromDict = arr
End Function

Private Sub Steam_QuickSortLong(ByRef a() As Long, ByVal lo As Long, ByVal hi As Long)
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

    If lo < j Then Steam_QuickSortLong a, lo, j
    If i < hi Then Steam_QuickSortLong a, i, hi
End Sub

'=========================================================
' roundingInfo(step) 찾기 + 반올림 규칙(원본 유지)
'=========================================================
Private Function Steam_GetRoundingStep(ByVal methodName As String, ByVal wsRound As Worksheet, ByVal defaultStep As Double) As Double

    Steam_GetRoundingStep = defaultStep

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
                    Steam_GetRoundingStep = CDbl(stepV)
                    Exit Function
                End If
            End If
        Next r

ContinueCol:
    Next c
End Function

Private Function Steam_RoundToStep(ByVal v As Double, ByVal stepV As Double) As Double

    If stepV <= 0 Then
        Steam_RoundToStep = v
        Exit Function
    End If

    If v = 0 Then
        Steam_RoundToStep = 0
        Exit Function
    End If

    If stepV = 0.5 Then
        If Abs(v) < stepV Then
            If v > 0 Then
                Steam_RoundToStep = stepV
            Else
                Steam_RoundToStep = -stepV
            End If
            Exit Function
        End If
    End If

    Dim q As Double
    q = v / stepV

    If q >= 0 Then
        Steam_RoundToStep = Int(q + 0.5) * stepV
    Else
        Steam_RoundToStep = -Int(Abs(q) + 0.5) * stepV
    End If
End Function

'=========================================================
' Steam trials 유틸
'=========================================================
Private Function Steam_GetTrialRow(ByVal trials As Object, ByVal trialNo As Long) As Long
    ' trials: 1..3 -> rowNum
    If trials.Exists(trialNo) Then
        Steam_GetTrialRow = CLng(trials(trialNo))
    Else
        ' 1회차가 없고 2/3만 있는 데이터도 있을 수 있으니,
        ' 우선 아무거나라도 있으면 채우고 싶다면 아래 로직을 바꾸면 됨.
        Steam_GetTrialRow = 0
    End If
End Function

Private Function Steam_GetAnyTrialRow(ByVal trials As Object) As Long
    ' 대표행(스펙/세탁횟수 용): 1이 있으면 1, 없으면 2, 없으면 3
    If trials.Exists(1) Then
        Steam_GetAnyTrialRow = CLng(trials(1))
    ElseIf trials.Exists(2) Then
        Steam_GetAnyTrialRow = CLng(trials(2))
    ElseIf trials.Exists(3) Then
        Steam_GetAnyTrialRow = CLng(trials(3))
    Else
        Steam_GetAnyTrialRow = 0
    End If
End Function


