#NoEnv
#Warn
SendMode Input
SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2
global xl := 0  ; 미리 선언
global wb := ""  ; start.xlsx workbook 객체
global startXlsxTemp := ""  ; 임시 복사본 경로
global yymmToday, log_dir_error
yymmToday := SubStr(A_Now, 1, 8)  ; initialize
log_dir_error := ""
; === 공휴일 목록 로드 (CSV 파일: holiday_2025.csv) ===
global gHolidays := []  ; 공휴일 배열
sampleArray := []    ; ← make sure this runs before you ever reference sampleArray
global pyoFlag := false
; CSV 파일 형식: 각 줄에 "yyyyMMdd" 혹은 "yyyy-MM-dd" 가 있다고 가정
holidayFile := A_ScriptDir "\holiday_2025.csv"
if FileExist(holidayFile)
{
    FileRead, holidayContent, %holidayFile%
    Loop, Parse, holidayContent, `n, `r
    {
        line := trim(A_LoopField)
        line := StrSplit(line, ",")[1]
        if (line != "")
        {
            ; 공백제거 후 배열에 추가 (만약 "-"가 포함되어 있다면 제거)
            ;line := StrReplace(line, "-", "")
            gHolidays.Push(line)
            /* 확인용
            for i, v in gHolidays
            {
                MsgBox, % "공휴일[" i "]: " v
            }
            */
        }
    }
}
else
{
    MsgBox, 16, 오류, 공휴일 파일을 찾을 수 없습니다: %holidayFile%
    ExitApp
}

; === 함수: 공휴일 여부 확인 ===

isHoliday(yyyyMMdd)
{
    ;MsgBox, %yyyyMMdd%,3
    global gHolidays
    for i, holiday in gHolidays
    {
        if (holiday = yyyyMMdd)
            return true
    }
    return false
}

; === 함수: 기준일로부터 offset번째 업무일(평일 & 공휴일 제외) 구하기 ===
getWorkingDay(offset := 0, refDate := "") {
    ; 기준일이 제공되지 않으면 현재 시간(A_Now)을 사용
    if (refDate = "")
        refDate := A_Now
    Loop {
        FormatTime, ymd, %refDate%, yyyy-MM-dd  ; 내부 비교용 YYYYMMDD
        FormatTime, dow, %refDate%, WDay     ; 일요일=1, 토요일=7
        ; 평일(월~금)이고 공휴일이 아니면 업무일로 인정
        if (dow > 1 && dow < 7 && !isHoliday(ymd)) {
            if (offset = 0)
                return ymd
            offset--
        }
        ; 기준일 하루 전으로 이동
        refDate += -1, days
    }
    ;MsgBox, refDate
}

; ------------------ FITI 실행 및 로그인 ------------------
IfWinNotExist, FITI
{
    Run, %ComSpec% /c C:\fiti\fiti.exe, hide
    Sleep, 5000

    WinWait, 시험관리시스템 로그인, , 15
    if ErrorLevel
    {
        yymmToday := SubStr(A_Now, 1, 8)
        log_dir_error = qrLog_input\%yymmToday%\%yymmToday%_qrLog.txt
        FileCreateDir, qrLog_input\%yymmToday%\
        timeRecord_error("시험관리시스템 로그인 창을 찾을 수 없습니다.")
        MsgBox, , 오류, "시험관리시스템 로그인" 창을 찾을 수 없습니다., 1
        ExitApp
    }
    WinActivate
    ControlSetText, ThunderRT6TextBox1, FB08030433
    ControlSetText, ThunderRT6TextBox2, FB08030433!!
    Send, {Enter}
    Send, {space}
    Send, {space}
    Sleep, 5000

    Loop
    {
        IfWinExist, FITI
            break
        MsgBox, 0, 자동 로그인, FITI시험관리시스템 실행 중, 1
    }
    WinGet, fitiHwnd, ID, FITI
    Sleep, 500
    MouseClick, left, 220, 40, 1, 0
    send,{down}{down}
    send,{enter}
    Sleep, 500
    MouseClick, left, 220, 40, 1, 0
    send,{down}{down}
    send,{enter}
    WinMove, FITI, , 200, 0, 1300, 1000
    ControlMove, ThunderRT6FormDC1, 10,120,1300,800,FITI
}
WinActivate, FITI
WinGet, fitiHwnd, ID, FITI
Sleep, 500
MouseClick, left, 220, 40, 1, 0
send,{down}{down}
send,{enter}
WinMove, FITI, , 200, 0, 1300, 1000
ControlMove, ThunderRT6FormDC1, 10,120,1300,800,FITI
Sleep, 500

Control, Choose, 4, ThunderRT6ComboBox2, FITI
Control, Choose, 1, ThunderRT6ComboBox1, FITI
Sleep, 1000
ControlSetText, ThunderRT6TextBox7, N, FITI
ControlSetText, ThunderRT6TextBox8, 231, FITI
ControlSetText, ThunderRT6TextBox9, 24, FITI
ControlSetText, ThunderRT6TextBox10, 1, FITI
Sleep, 500
Click 220, 80, 1
Sleep, 500
send, {Space}
send, {Space}
Click 220, 80, 1
Sleep, 500
send, {Space}
send, {Space}
Sleep, 500

Sleep, 500
MouseMove, 1000, 1000

basePath := A_ScriptDir
FilePath1 := "\\192.168.1.7\유해물질시험팀\3. 폼알데히드,pH파트\자동화프로그램 개발\2025\가공성능평가팀_수축율\start.xlsx"
startXlsxTemp := A_Temp "\start_" A_Now "_" A_TickCount ".xlsx"
FileCopy, %FilePath1%, %startXlsxTemp%, 1
xl := ComObjCreate("Excel.Application")
wb := xl.Workbooks.Open(startXlsxTemp, 0, true)  ; ReadOnly := true
xl.Visible := false
xl.DisplayAlerts := false

; === 업무일 기준 날짜 계산 ===
; 기본적으로 오늘, 어제, 2일전(업무일 기준)을 구합니다.
; 단, 만약 오늘이 월요일(A_WDay=2)라면 어제(실제 금요일, 혹은 금요일이 공휴일이면 그 이전 업무일)를 사용합니다.
;A_WDay1 := A_WDay  ; A_WDay 내장변수 (일=1, 월=2, …, 토=7)
;if (A_WDay1 = 2) {
;    acceptableDate := getWorkingDay(1)  ; 월요일이면 어제 업무일 (보통 금요일)
;} else {
    acceptableDate := getWorkingDay(0) . "," . getWorkingDay(1) . "," . getWorkingDay(2) . "," . getWorkingDay(3) . "," . getWorkingDay(4) . "," . getWorkingDay(5)
;}
;MsgBox, %acceptableDate%
FormatTime, todayFormatted, , yyyy-MM-dd


; ------------------ Excel 행 반복 ------------------
prevReceiptNumber := ""
row := 4
loop
{

    receiptNumber := wb.Worksheets(1).Range("A" . row).Value
    if (receiptNumber = "")
        break
    ;MsgBox, %receiptNumber%
    if (receiptNumber = prevReceiptNumber) {
        ;MsgBox, %receiptNumber%
        row++
        continue
    }
    prevReceiptNumber := receiptNumber
    saveNumber := StrReplace(receiptNumber, "-", "")  ; 하이픈 제거
    filePath := "\\192.168.1.7\유해물질시험팀\3. 폼알데히드,pH파트\자동화프로그램 개발\2025\가공성능평가팀_수축율\testNumber\2025\" saveNumber ".txt"

    ; 파일이 이미 존재하면 다음 반복으로
    if FileExist(filePath) {
        row++
        continue
    }
    ; --- B열 날짜 조건 ---
    ;MsgBox, %acceptableDate%
    cellB := wb.Worksheets(1).Range("B" . row).Text
    ;MsgBox, %cellB%
    ; acceptableDate 예: "20250421,20250420,20250419"
    found := false
    Loop, Parse, acceptableDate, `,
    {
        if (A_LoopField = cellB) {
            found := true
            break
        }
    }

    if (!found) {
        row++
        continue
    }



    ;-------

    ; 1) 접수번호 입력 로직
    ControlSetText, ThunderRT6TextBox7, % SubStr(receiptNumber, 1, 1), FITI
    ControlSetText, ThunderRT6TextBox8, % SubStr(receiptNumber, 2, 3), FITI
    ControlSetText, ThunderRT6TextBox9, % SubStr(receiptNumber, 6, 2), FITI
    ControlSetText, ThunderRT6TextBox10, % SubStr(receiptNumber, 9, 5), FITI

    ; 접수번호 검증 (재입력)
    ControlGetText, check2A, ThunderRT6TextBox7, FITI
    ControlGetText, check2B, ThunderRT6TextBox8, FITI
    ControlGetText, check2C, ThunderRT6TextBox10, FITI
    check1A := SubStr(receiptNumber, 1, 1)
    check1B := SubStr(receiptNumber, 2, 3)
    check1C := SubStr(receiptNumber, 9, 5)
    if (check1A != check2A or check1B != check2B or check1C != check2C)
    {
        MsgBox,,, 접수 번호 보정중.., 2
        ControlSetText, ThunderRT6TextBox7, % check1A, FITI
        ControlSetText, ThunderRT6TextBox8, % check1B, FITI
        ControlSetText, ThunderRT6TextBox9, % SubStr(receiptNumber, 6, 2), FITI
        ControlSetText, ThunderRT6TextBox10, % check1C, FITI
        ControlGetText, check2A, ThunderRT6TextBox7, FITI
        ControlGetText, check2B, ThunderRT6TextBox8, FITI
        ControlGetText, check2C, ThunderRT6TextBox10, FITI
        if (check1A != check2A or check1B != check2B or check1C != check2C)
        {
            MsgBox,,, 접수번호 보정 실패, 1
            row++
            continue
        }
    }

    ; 접수번호조회 창 닫기
    WinGetTitle, a, FITI
    WinActivate, %a%
    MouseClick, left, 600, 300
    WinGetTitle, a, A
    Sleep, 500
    If (a = "접수번호조회")
    {
        ControlClick, Button3, ahk_class ThunderRT6FormDC
    }

    ; 조회 버튼 클릭
    Sleep, 1500
    ImageSearch, vx, vy, 0, 0, A_ScreenWidth, A_ScreenHeight, *70 %A_ScriptDir%/image/lookup.png
    if (ErrorLevel = 0)
    {
        MouseClick, Left, %vx%, %vy%
        Sleep, 1000
    }
    else
    {
        MsgBox,,, 조회 버튼 찾기 실패, 1
        row++
        continue
    }
    ImageSearch, okx, oky, 0, 0, 250, 130, *100 %A_ScriptDir%/image/ok.png
    if (ErrorLevel = 0)
    {
        Send, {Space}{Space}
        ; ★ OK 누른 뒤 그리드 로딩/안정화 대기
        WaitGridAfterOK(fitiHwnd, 9000)
    }
    else
    {
        ; OK가 없어도, 느린 조회 대비로 한번 대기(짧게)
        WaitGridAfterOK(fitiHwnd, 3000)
    }

    ; --- 복사 영역 클릭 및 선택 ---
    ; FITI 창 기준 좌표 (60, 280) 클릭
    WinActivate, ahk_id %fitiHwnd%

    ;WinGetPos, fitiX, fitiY, , , ahk_id %fitiHwnd%
    ;clickX := fitiX + 60
    ;clickY := fitiY + 280
    ;MouseClick, left, %clickX%, %clickY%
    Sleep, 300

    ; Shift + PgDn으로 복사 영역 선택
    Send, +{PgDn}
    Sleep, 300

    ; Ctrl + C로 복사
    Send, ^c
    Sleep, 500

    ; 클립보드 내용 가져오기
    ClipWait, 2
    if (!ErrorLevel)
    {
        clipboardData := Clipboard

        ; ? 추가: "빨랫줄건조" 다음 줄바꿈 + ", 스파크세제"를 한 줄로 붙이기
        clipboardData := FixDetergentLineBreak(clipboardData)

        ; 파일명: saveNumber.txt
            ; → ThunderRT6TextBox5에서 텍스트 가져오기
        ControlGetText, b1Value, ThunderRT6TextBox5, ahk_id %fitiHwnd%
        filePath := "\\192.168.1.7\유해물질시험팀\3. 폼알데히드,pH파트\자동화프로그램 개발\2025\가공성능평가팀_수축율\testNumber\2025\" saveNumber ".txt"
        FileCreateDir, "\\192.168.1.7\유해물질시험팀\3. 폼알데히드,pH파트\자동화프로그램 개발\2025\가공성능평가팀_수축율\testNumber\2025"
        FileDelete, %filePath%
        FileAppend, %b1Value%, %filePath%
        FileAppend, %clipboardData%, %filePath%
    }
    else
    {
        MsgBox,,, 클립보드 복사 실패!, 1
    }
    row++
    Sleep, 100

}
if fitiHwnd
{
    PostMessage, 0x10, 0, 0, , ahk_id %fitiHwnd%
    Sleep, 2000
}

if (IsObject(wb))
{
    wb.Close(0)
    wb := ""
}
if (IsObject(xl))
{
    xl.Quit
    xl := ""
}
if (startXlsxTemp != "" && FileExist(startXlsxTemp))
    FileDelete, %startXlsxTemp%

MsgBox,,, 작업이 완료되었습니다., 1
ExitApp
return

; ====== 함수 정의: 배열에 값이 있는지 확인 ======
ArrayContains(arr, item) {
    Local i ;
    for i, val in arr {
        if (val = item)
            return true
    }
    return false
}

timeRecord_error(sentence_error){
    global log_dir_error
    FileAppend, [%A_Year% %A_Mon%/%A_Mday% %A_Hour%:%A_Min% %A_sec%][%sentence_error%]`n, %log_dir_error%
}

FixDetergentLineBreak(text)
{
    ; 케이스: "빨랫줄건조`n, 스파크세제" 또는 "빨랫줄건조`r`n, 스파크세제"
    ; 공백/탭이 섞여도 "빨랫줄건조, 스파크세제"로 정규화  ; v1에서는 다중라인 함수호출이 깨질 수 있어 한 줄로 작성
    return RegExReplace(text, "빨랫줄건조\s*`r?`n\s*,\s*스파크세제", "빨랫줄건조, 스파크세제")
}

WaitGridAfterOK(fitiHwnd, maxWaitMs := 8000)
{
    ; OK 눌렀다고 가정하고, 그리드가 갱신/안정화될 때까지 기다림
    ; "그리드 복사 결과가 유의미해질 때"를 신호로 사용

    WinActivate, ahk_id %fitiHwnd%
    Sleep, 120

    WinGetPos, fX, fY, , , ahk_id %fitiHwnd%
    gx := fX + 60
    gy := fY + 280

    start := A_TickCount
    last := ""
    stableCount := 0

    Loop
    {
        txt := PeekGridText(gx, gy)

        ; 너무 짧으면(로딩중/빈그리드) 계속 대기
        if (StrLen(txt) < 40) {
            stableCount := 0
        } else {
            ; 내용이 2번 연속 동일하면 "안정화"로 판단
            if (txt = last)
                stableCount++
            else
                stableCount := 0

            last := txt
            if (stableCount >= 2)
                return true
        }

        if (A_TickCount - start > maxWaitMs)
            return false

        Sleep, 200
    }
}


PeekGridText(gx, gy)
{
    Clipboard :=
    ; 혹시 다른 컨트롤에 포커스가 가있을 수 있으니 Tab/Shift+Tab로 그리드로 복귀하는 방법도 가능
    ;Send, {Tab}  ; 필요하면 한 번만
    ; 좌클릭 대신 우클릭 (대부분 팝업 안 뜸)
    Click, %gx%, %gy%, left
    Send, +{PgDn}
    Sleep, 300
    Sleep, 80
    Send, ^c
    ClipWait, 0.5
    if (ErrorLevel)
        return ""

    t := Clipboard
    ; 공백/개행 약간 정리
    t := RegExReplace(t, "\s+", " ")
    return t
}



F3::
Pause
return

Esc::
    if (IsObject(wb))
    {
        try {
            wb.Close(0)
        }
        catch {
        }
        wb := ""
    }
    if (IsObject(xl))
    {
        try {
            xl.Visible := true
            xl.DisplayAlerts := true
            xl.Quit
        }
        catch {
            log_dir_error = qrLog_input\%yymmToday%\%yymmToday%_qrLog.txt
            FileCreateDir, qrLog_input\%yymmToday%\
            timeRecord_error("오류 발생!")
            MsgBox,,, 오류 발생!, 1
        }
        xl := ""
    }
    if (startXlsxTemp != "" && FileExist(startXlsxTemp))
        FileDelete, %startXlsxTemp%
    ExitApp
return



