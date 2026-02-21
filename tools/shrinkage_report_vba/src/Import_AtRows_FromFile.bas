Attribute VB_Name = "Module2"
Option Explicit

'외부 파일에서 A6부터 "@" 형식이 있는 마지막 행까지(해당 행의 모든 열) 현재 시트로 가져오기
Public Sub Import_AtRows_FromFile()

    Dim filePath As Variant
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet

    Dim firstRow As Long: firstRow = 6
    Dim lastRow As Long, lastCol As Long
    Dim rngCopy As Range

    '1) 가져올 파일 선택
    filePath = Application.GetOpenFilename("Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", , "가져올 파일 선택")
    If filePath = False Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo CleanFail

    Set wsDst = ActiveSheet

    '? (추가) 가져오기 전에 현재 시트 기존 데이터 삭제(값+서식 포함)
    ClearDstFromA6 wsDst

    '2) 소스 파일 열기(읽기 전용)
    Set wbSrc = Workbooks.Open(CStr(filePath), ReadOnly:=True)

    '3) 소스 시트 선택: Rawdata 있으면 그 시트, 없으면 첫 시트
    On Error Resume Next
    Set wsSrc = wbSrc.Worksheets("Rawdata")
    On Error GoTo 0
    If wsSrc Is Nothing Then Set wsSrc = wbSrc.Worksheets(1)

    '4) A열에서 "@" 형식이 있는 마지막 행 찾기 (A6부터 아래로)
    lastRow = GetLastAtRow(wsSrc, firstRow, 1) '1 = A열
    If lastRow < firstRow Then
        MsgBox "A6부터 아래에 '@' 형식 데이터가 없습니다.", vbExclamation
        GoTo CleanExit
    End If

    '5) 소스 시트에서 마지막 사용 열 찾기(행 전체 복사 범위용)
    lastCol = GetLastUsedCol(wsSrc)
    If lastCol < 1 Then lastCol = 1

    '6) 복사 범위 설정
    Set rngCopy = wsSrc.Range(wsSrc.Cells(firstRow, 1), wsSrc.Cells(lastRow, lastCol))

    '7) 현재 시트의 A6부터 붙여넣기 (값+서식 함께)
    rngCopy.Copy Destination:=wsDst.Range("A6")

CleanExit:
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

CleanFail:
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "오류: " & Err.Description, vbExclamation
End Sub

'? (추가) 현재 시트의 A6부터 아래 전체(사용영역) 삭제(값+서식)
Private Sub ClearDstFromA6(ByVal ws As Worksheet)
    Dim lastRow As Long, lastCol As Long

    ' 사용된 마지막 행/열 기준으로 A6부터 아래 범위 계산
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 6 Then lastRow = 6

    lastCol = GetLastUsedCol(ws)
    If lastCol < 1 Then lastCol = 1

    ' A6~(lastRow,lastCol) 범위 "완전 삭제"
    ws.Range(ws.Cells(6, 1), ws.Cells(lastRow, lastCol)).Clear
End Sub

'--- A열에서 "@형식"이 있는 마지막 행을 찾는다 (startRow~마지막행 중)
Private Function GetLastAtRow(ByVal ws As Worksheet, ByVal startRow As Long, ByVal col As Long) As Long
    Dim r As Long
    Dim lastAny As Long
    Dim v As String

    lastAny = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If lastAny < startRow Then
        GetLastAtRow = 0
        Exit Function
    End If

    '아래에서 위로 올라오며 "@형식" 찾기
    For r = lastAny To startRow Step -1
        v = CStr(ws.Cells(r, col).Value)
        If IsAtKey(v) Then
            GetLastAtRow = r
            Exit Function
        End If
    Next r

    GetLastAtRow = 0
End Function

'--- "@...@" 형태(최소 2개의 @)인지 판단
Private Function IsAtKey(ByVal s As String) As Boolean
    s = Trim$(s)
    If Len(s) < 3 Then Exit Function
    If Left$(s, 1) <> "@" Then Exit Function
    If InStr(2, s, "@") = 0 Then Exit Function
    IsAtKey = True
End Function

'--- 시트에서 마지막 사용 열(Used Range 기준) 찾기
Private Function GetLastUsedCol(ByVal ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                          LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If f Is Nothing Then
        GetLastUsedCol = 0
    Else
        GetLastUsedCol = f.Column
    End If
End Function


