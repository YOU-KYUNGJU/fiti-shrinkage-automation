#NoEnv
#SingleInstance, Force

SetTitleMatchMode, 2
CoordMode, Mouse, Window

; FITI 창이 없으면 실행
IfWinNotExist, FITI
{
    Run, %ComSpec% /c C:\fiti\fiti.exe, hide
    Sleep, 5000  ; 실행 후 대기


    ; "시험관리시스템 로그인" 창이 나타날 때까지 대기
    WinWait, 시험관리시스템 로그인, , 10
    if ErrorLevel  ; 10초 동안 창이 안 나타나면 종료
    {
        MsgBox, 16, 오류, "시험관리시스템 로그인" 창을 찾을 수 없습니다.
        ExitApp
    }

    WinActivate  ; 로그인 창 활성화
    ControlSetText, ThunderRT6TextBox1, FB08030433
    ControlSetText, ThunderRT6TextBox2, FB08030433!!
    Send, {Enter}
    Send, {space}
    Send, {space}
    Sleep, 5000  ; 로그인 후 대기

    ; FITI 창이 나타날 때까지 대기
    Loop
    {
        IfWinExist, FITI
            break
        MsgBox, 0, 자동 로그인, FITI시험관리시스템 실행 중, 1
    }

    ; FITI 창 정보 저장
    WinGet, fitiHwnd, ID, FITI

    ; FITI 창 위치 및 크기 조정
    WinActivate, FITI
    WinMove, FITI, , 200, 0, 1300, 1000
}
WinActivate, FITI
; FITI 창 정보 저장
WinGet, fitiHwnd, ID, FITI
MouseClick, left, 280, 40, 1, 0
send,{Up}{Up}{Up}{Up}{Up}
send,{Right}{Up}{Up}
send,{Enter}
Sleep,500
WinMove, FITI, , 200, 0 , 1300, 1000 ; 창 크기 맞추기
; 엑셀 강제 종료
Process, Close, EXCEL.EXE

; --- 1. ThunderRT6MDIForm 창 활성화 ---
WinActivate, ahk_class ThunderRT6MDIForm
WinWaitActive, ahk_class ThunderRT6MDIForm

; --- 2. ComboBox 값 선택 ---
;ThunderRT6ComboBox1 가공성능평가팀
Control, Choose, 3, ThunderRT6ComboBox1, ahk_class ThunderRT6MDIForm
Control, ChooseString, 접수일, ThunderRT6ComboBox5, ahk_class ThunderRT6MDIForm
Control, ChooseString, 수축율, ThunderRT6ComboBox2, ahk_class ThunderRT6MDIForm

; --- 3. 날짜 계산 ---
today := A_Now

; 오늘 날짜 -12일 (tempDate)
tempDate := today
EnvAdd, tempDate, -12, days

; 오늘 날짜 +1일 (tempDate2)
tempDate2 := today
EnvAdd, tempDate2, 1, days

; --- 4. DTPicker 컨트롤에 날짜 설정 (PostMessage 사용: 필요시 사용) ---
SetDatePicker("DTPicker20WndClass1", "ahk_class ThunderRT6MDIForm", tempDate)
SetDatePicker("DTPicker20WndClass2", "ahk_class ThunderRT6MDIForm", tempDate2)

; --- 4.1 DTPicker20WndClass1 직접 입력 방식 조정 (예상: 오늘 - 5일) ---
FormatTime, expectedDate1, %tempDate%, yyyy-MM-dd
; 예상값에서 연, 월, 일을 분리 (문자열 그대로 사용)
expectedYear1 := SubStr(expectedDate1, 1, 4)
expectedMonth1 := SubStr(expectedDate1, 6, 2)
expectedDay1 := SubStr(expectedDate1, 9, 2)

ControlFocus, DTPicker20WndClass1, ahk_class ThunderRT6MDIForm
ControlClick, DTPicker20WndClass1, ahk_class ThunderRT6MDIForm, , Left, 2
Sleep, 500
; 월 입력
ControlSend, DTPicker20WndClass1, %expectedMonth1%, ahk_class ThunderRT6MDIForm
Sleep, 500
ControlSend, DTPicker20WndClass1, {Right}, ahk_class ThunderRT6MDIForm
Sleep, 500
; 일 입력
ControlSend, DTPicker20WndClass1, %expectedDay1%, ahk_class ThunderRT6MDIForm
Sleep, 500
ControlSend, DTPicker20WndClass1, {Right}, ahk_class ThunderRT6MDIForm
Sleep, 500
; 년 입력
ControlSend, DTPicker20WndClass1, %expectedYear1%, ahk_class ThunderRT6MDIForm
Sleep, 500
; 최종 텍스트 확인
ControlGetText, currentText1, DTPicker20WndClass1, ahk_class ThunderRT6MDIForm
if (currentText1 = expectedDate1)
    ToolTip, DTPicker20WndClass1 날짜 일치: %currentText1%
else
    ToolTip, DTPicker20WndClass1 날짜 불일치:n예상: %expectedDate1%n현재: %currentText1%
Sleep, 2000
ToolTip

; --- 4.2 DTPicker20WndClass2 직접 입력 방식 조정 (예상: 오늘 + 1일) ---
FormatTime, expectedDate2, %tempDate2%, yyyy-MM-dd
expectedYear2 := SubStr(expectedDate2, 1, 4)
expectedMonth2 := SubStr(expectedDate2, 6, 2)
expectedDay2 := SubStr(expectedDate2, 9, 2)

ControlFocus, DTPicker20WndClass2, ahk_class ThunderRT6MDIForm
ControlClick, DTPicker20WndClass2, ahk_class ThunderRT6MDIForm, , Left, 2
Sleep, 500
; 월 입력
ControlSend, DTPicker20WndClass2, %expectedMonth2%, ahk_class ThunderRT6MDIForm
Sleep, 500
ControlSend, DTPicker20WndClass2, {Right}, ahk_class ThunderRT6MDIForm
Sleep, 500
; 일 입력
ControlSend, DTPicker20WndClass2, %expectedDay2%, ahk_class ThunderRT6MDIForm
Sleep, 500
ControlSend, DTPicker20WndClass2, {Right}, ahk_class ThunderRT6MDIForm
Sleep, 500
; 년 입력
ControlSend, DTPicker20WndClass2, %expectedYear2%, ahk_class ThunderRT6MDIForm
Sleep, 500
; 최종 텍스트 확인
ControlGetText, currentText2, DTPicker20WndClass2, ahk_class ThunderRT6MDIForm
if (currentText2 = expectedDate2)
    ToolTip, DTPicker20WndClass2 날짜 일치: %currentText2%
else
    ToolTip, DTPicker20WndClass2 날짜 불일치:n예상: %expectedDate2%n현재: %currentText2%
Sleep, 2000
ToolTip

; --- 5. 마우스 클릭 (창 기준 좌표) ---
MouseClick, left, 220, 90, 1, 0
Sleep, 1000
MouseClick, left, 300, 90, 1, 0
Sleep, 30000  ; 30초 대기

; --- 6. "엑셀"이 포함된 창 활성화 (5분 제한) ---
excelActivated := false
elapsedTime := 0

Loop, 15 {  ; 5분 / 20초 = 15번 반복
    WinActivate, ahk_exe EXCEL.EXE
    if WinExist("ahk_exe EXCEL.EXE") {
        WinActivate, ahk_exe EXCEL.EXE
        excelActivated := true
        Break
    }
    Sleep, 20000
    elapsedTime += 20
}

if (!excelActivated) {
    MsgBox, 16, 에러, Excel 창 활성화에 실패했습니다. 프로그램을 종료합니다.
    ExitApp
}

WinActivate, ahk_exe EXCEL.EXE
Sleep, 5000  ; Excel 창 활성화 후 5초 대기

try {
    xl := ComObjActive("Excel.Application")
    wb := xl.ActiveWorkbook
    ws := wb.Sheets(1)  ; 첫 번째 시트 사용
    ; H열 필터 적용 ("pH" 포함된 값만 표시)
    ;ws.Range("H3").AutoFilter(8, "*pH*", 2, true)  ;8번째 열(H)에서 "pH" 필터링
    ; 첫 번째 파일 저장 (pH 포함된 값만 남김)
    basePath := A_ScriptDir
    FilePath1 := "\\192.168.1.7\유해물질시험팀\3. 폼알데히드,pH파트\자동화프로그램 개발\2025\가공성능평가팀_수축율\start.xlsx"
    wb.SaveAs(FilePath1)
    ; 필터 해제 후 다시 적용 (pH 포함된 값 제외)
    if ws.AutoFilterMode  ; 기존 필터가 있으면 제거
        ws.AutoFilterMode := False
    Sleep, 1000  ; 필터 제거 후 잠시 대기
    ;ws.Range("H3").AutoFilter
    ;ws.Range("H3").AutoFilter(8, "<>*pH*", 2, true)  ; "pH"가 포함되지 않은 값만 필터링
    ; 두 번째 파일 저장 (pH 제외한 값만 남김)
    ;FilePath2 := "\\192.168.1.7\유해물질시험팀\3. 폼알데히드,pH파트\양식\2025\통계_시험실별시험항목접수현황_HCHO.xlsx"
    ;wb.SaveAs(FilePath2)
    ; 필터 해제
    ;ws.AutoFilterMode := False
} catch e {
    MsgBox, Excel 파일 저장 중 오류가 발생했습니다.`n%e%
}

; --- 7. 모든 작업 완료 후 FITI 창 종료 ---
if fitiHwnd
{
    PostMessage, 0x10, 0, 0, , ahk_id %fitiHwnd%
    Sleep, 2000
}


; --- 8. 엑셀 창 종료 ---
if (xl) {
    xl.Quit()
    xl := ""
}

; --- 모든 작업 완료 메시지 ---
MsgBox,,, 모든 작업이 성공적으로 완료되었습니다.,1
ExitApp
Return

; --- 함수: DTPicker 컨트롤에 날짜를 설정 (SYSTEMTIME 구조체 사용) ---
SetDatePicker(ctrl, winTitle, dt) {
    FormatTime, year, %dt%, yyyy
    FormatTime, month, %dt%, MM
    FormatTime, day, %dt%, dd
    VarSetCapacity(st, 16, 0)
    NumPut(year, st, 0, "UShort")
    NumPut(month, st, 2, "UShort")
    NumPut(0, st, 4, "UShort")  ; 요일 (0)
    NumPut(day, st, 6, "UShort")
    NumPut(0, st, 8, "UShort")  ; 시간
    NumPut(0, st, 10, "UShort") ; 분
    NumPut(0, st, 12, "UShort") ; 초
    NumPut(0, st, 14, "UShort") ; 밀리초
    DTM_SETSYSTEMTIME := 0x1002
    PostMessage, %DTM_SETSYSTEMTIME%, 1, &st,, %ctrl%, %winTitle%
}
return



Esc::
    if (IsObject(Xl))  ; Xl이 존재하면
    {
        try {
            Xl.Visible := true
            Xl.DisplayAlerts := true
            Xl.ActiveWorkbook.Close(0)
            Xl.Quit
        }
        catch e {
            MsgBox, 16, 오류, Xl을 종료하는 중 오류 발생: %e%
        }
        Xl := ""  ; 객체 변수 해제
    }
    ExitApp
return

