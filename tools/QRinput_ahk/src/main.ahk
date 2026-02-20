#NoEnv
#Persistent
#SingleInstance, Force
#Include Gdip_All.ahk
SetBatchLines, -1

SetTitleMatchMode, 2
;CoordMode, Mouse, Window
global xl := ""  ; 미리 선언
;────────────────────────────────────────────
; [공통] 에러 시 프로그램 종료 대신 ESC 키가 눌릴 때까지 대기하도록 변경
HandleError(msg) {
    MsgBox,, 오류, %msg%`n(프로그램 종료를 원하시면 ESC 키를 눌러주세요.), 1
    KeyWait, Esc  ; ESC가 눌릴 때까지 대기
    return
}

;────────────────────────────────────────────
; 전역 변수 초기화
global csvFileList := []  ; CSV파일 목록
global TaskCount := 0, TaskTotal := 0, FileCount1 := 0
global currentDate := 0
;────────────────────────────────────────────
; [시간대별 폴더 경로 구성]
; 현재 시간을 기준으로 날짜 결정
FormatTime, currentHour, A_Now, HH
FormatTime, currentDate, A_Now, yyyyMMdd

subtractMonth := 1

if (currentHour < 10)   ; 오전 10시 이전 -> baseMonthPath 내에서 현재 날짜 바로 이전 날짜 폴더 찾기
{
    ; currentDate는 "yyyyMMdd" 형식이므로, day를 추출합니다.
    day := SubStr(currentDate, 7, 2)

    ; 만약 day가 "01"이면, 전월 폴더에서 찾도록 처리
    if (day = "01") {
        ; currentDate에 "000000" (시간)을 덧붙여 전체 날짜/시간 문자열로 만들어줍니다.
        tempDate := currentDate . "000000"
        ; 한 달(subtract 1 month)을 빼서 전월 날짜 계산 (EnvSub는 full datetime 형식을 처리할 수 있습니다)
        EnvSub, tempDate, %tempDate%, %subtractMonth%, Months
        FormatTime, prevYear, %tempDate%, yyyy
        FormatTime, prevMonth, %tempDate%, MM
        baseMonthPath := "\\192.168.1.7\유해물질시험팀\3. 폼알데히드,pH파트\자동화프로그램 개발\2025\가공성능평가팀_수축율\" . prevYear . "\" . prevMonth . "\"
    }
    else {
        year := SubStr(currentDate, 1, 4)
        month := SubStr(currentDate, 5, 2)
        baseMonthPath := "\\192.168.1.7\유해물질시험팀\3. 폼알데히드,pH파트\자동화프로그램 개발\2025\가공성능평가팀_수축율\" . year . "\" . month . "\"
    }

    prevDate := ""
    ; baseMonthPath 아래의 모든 폴더를 순회
    Loop, Files, % baseMonthPath "\*", D
    {
        folderName := A_LoopFileName
        ; 폴더명이 8자리 숫자인지 확인 (예: 20250430)
        if (RegExMatch(folderName, "^\d{8}$"))
        {
            ; 현재 날짜(currentDate)보다 작은 폴더들 중, 가장 큰 값을 prevDate에 저장
            if (folderName < currentDate)
            {
                if (prevDate = "" or folderName > prevDate)
                    prevDate := folderName
            }
        }
    }
    if (prevDate != "")
        currentDate := prevDate  ; 찾은 이전 날짜 폴더명을 currentDate로 대체
}

/*
else if (currentHour >= 17)  ; 오후 5시 이후 -> 내일 대신 폴더 내에서 currentDate 보다 큰 폴더명 찾기
{
    year := SubStr(currentDate, 1, 4)
    month := SubStr(currentDate, 5, 2)
    baseMonthPath := "\\192.168.1.7\유해물질시험팀\3. 폼알데히드,pH파트\자동화프로그램 개발\2025\가공성능평가팀_수축율\" . year . "\" . month . "\"
    newDate := ""
    ; 현재 월(baseMonthPath) 아래의 모든 폴더를 순회
    Loop, Files, % baseMonthPath "\*", D
    {
        folderName := A_LoopFileName
        ; 폴더명이 8자리 숫자인지 확인 (예: 20250411)
        if (RegExMatch(folderName, "^\d{8}$"))
        {
            ; 현재 날짜(currentDate)보다 큰 폴더명인지 검사
            if (folderName > currentDate)
            {
                ; 아직 newDate가 비어 있거나 folderName이 현재까지의 newDate보다 작으면 갱신 (즉, 가장 가까운 미래 날짜)
                if (newDate = "" or folderName < newDate)
                    newDate := folderName
            }
        }
    }
    if (newDate != "")
        currentDate := newDate  ; 현재 월에서 찾은 폴더명으로 currentDate 변경
    else {
        ; 현재 월에서 해당 폴더가 없으면, 다음 월 폴더로 검색
        tempDate := currentDate . "000000"  ; "yyyyMMdd000000" 형태로 만듦
        EnvAdd, tempDate, 1, Months      ; 1개월 추가하여 다음 월 날짜 계산
        FormatTime, nextYear, %tempDate%, yyyy
        FormatTime, nextMonth, %tempDate%, MM
        nextBaseMonthPath := "\\192.168.1.7\유해물질시험팀\3. 폼알데히드,pH파트\자동화프로그램 개발\2025\가공성능평가팀_수축율\" . nextYear . "\" . nextMonth . "\"
        newDate := ""
        Loop, Files, % nextBaseMonthPath "\*", D
        {
            folderName := A_LoopFileName
            if (RegExMatch(folderName, "^\d{8}$"))
            {
                ; 다음 월의 폴더들 중 가장 작은 값을 찾음 (이미 다음 월의 폴더이므로 비교 없이 가장 가까운 폴더)
                if (newDate = "" or folderName < newDate)
                    newDate := folderName
            }
        }
        if (newDate != "")
            currentDate := newDate  ; 다음 월에서 찾은 폴더명으로 currentDate 변경
        ; 만약 다음 월에서도 찾지 못하면 currentDate는 그대로 유지
    }
}
*/
currentDate := SubStr(currentDate, 1,8)
;MsgBox, %currentDate%
; 11시 ~ 17시는 당일 사용
;MsgBox, %currentDate%
yymmToday := SubStr(A_Now, 1, 8)
year := SubStr(currentDate, 1, 4)
month := SubStr(currentDate, 5, 2)
day := SubStr(currentDate, 7, 2)
folderBase := "\\192.168.1.7\유해물질시험팀\3. 폼알데히드,pH파트\자동화프로그램 개발\2025\가공성능평가팀_수축율\" . year . "\" . month . "\" . currentDate

if not FileExist(folderBase)
{
    FileCreateDir, qrLog_input\%yymmToday%\
    log_dir_error = qrLog_input\%yymmToday%\%yymmToday%_qrLog.txt
    FileAppend, 지정한 폴더가 존재하지 않습니다. [" folderBase "], %log_dir_error%
    HandleError("지정한 폴더가 존재하지 않습니다. [" folderBase "]")
    return
}

; 폴더 내의 모든 CSV 파일을 목록에 추가 (수정시간 비교 없이)
Loop, Files, % folderBase "\*.csv", F
{
    FileGetTime, createdTime, %A_LoopFileFullPath%, C  ; C = Created (생성 시간)
    FormatTime, fileHour, %createdTime%, HH

    ; 오전 10시 이전이라면 생성 시간이 17시 이후인 파일만 처리
    if (currentHour < 10) {
        if (fileHour >= 17) {
            csvFileList.Push(A_LoopFileFullPath)
        }
    } else {
        ; 그 외 시간대는 모두 포함
        csvFileList.Push(A_LoopFileFullPath)
    }
}

if (csvFileList.MaxIndex() = 0)
{
    FileCreateDir, qrLog_input\%yymmToday%\
    log_dir_error = qrLog_input\%yymmToday%\%yymmToday%_qrLog.txt
    FileAppend, 지정한 폴더 내에 CSV 파일이 존재하지 않습니다., %log_dir_error%
    HandleError("지정한 폴더 내에 CSV 파일이 존재하지 않습니다.")
    return
}

;────────────────────────────────────────────
; [FITI 실행 및 자동 QR 처리 시작]
; FITI 창이 없으면 실행 (QR스캔 블록)
IfWinNotExist, QR스캔
{
    Run, %ComSpec% /c C:\fiti\fiti.exe, hide
    Sleep, 5000  ; 실행 후 대기

    ; "시험관리시스템 로그인" 창이 나타날 때까지 대기 (최대 10초)
    WinWait, 시험관리시스템 로그인, , 15
    if ErrorLevel
    {
        yymmToday := SubStr(A_Now, 1, 8)
        FileCreateDir, qrLog_input\%yymmToday%\
        log_dir_error = qrLog_input\%yymmToday%\%yymmToday%_qrLog.txt
        FileAppend, 시험관리시스템 로그인 창을 찾을 수 없습니다., %log_dir_error%
        HandleError("시험관리시스템 로그인 창을 찾을 수 없습니다.")
        return
    }
    WinActivate
    ControlSetText, ThunderRT6TextBox1, FB15041858
    ControlSetText, ThunderRT6TextBox2, FB15041858!
    Send, {Enter}
    Send, {space}
    Send, {space}
    Sleep, 5000  ; 로그인 후 대기

    ; FITI 창 대기
    Loop
    {
        IfWinExist, FITI
            break
        MsgBox, 0, 자동 로그인, FITI시험관리시스템 실행 중, 1
    }

    WinActivate, FITI
    WinMove, FITI, , 200, 0, 1300, 1000
    Sleep, 500
    MouseClick, left, 730, 80, 1, 0
    Sleep, 1000
    MouseClick, left, 730, 80, 1, 0
    Sleep, 1000
}

;────────────────────────────────────────────
; [처리 GUI 생성]
;Gui, +DisableTheming
SysGet, MonitorHeight, 2
GuiHeight := 0
GuiY := MonitorHeight - Abs(GuiHeight) - 10

Gui, Processing:New, +AlwaysOnTop -SysMenu, 입고 자동 처리
Gui, Processing:Add, Text, x10 y10 w170 h20, 현재 작업 진행 상황:
Gui, Processing:Add, Progress, x10 y40 w220 h20 vProgress, 0
Gui, Processing:Add, Text, x10 y70 w170 h20 vCurrentTask, 현재 작업: 없음
Gui, Processing:Add, Edit, x10 y100 w220 h40 vLog ReadOnly, 로그 출력...

; ----- 기존 좌측 순번 ListBox, 우측 파일명 ListBox 대신 단일 ListView 사용 -----
; ListView 컨트롤은 열(column)을 여러 개 사용 가능하므로, 예제에서는 "순번"과 "파일명" 두 개의 열을 정의합니다.
Gui, Processing:Add, ListView, x10 y150 w300 h400 vFileListView ReadOnly, 순번|파일명|완료

Gui, Processing:Add, Text, x10 y630 w220 h20 vFileCount, 파일 처리 개수: 0
Gui, Processing:Add, Button, x10 y570 w100 h30 gStartProcess, 시작
Gui, Processing:Add, Button, x120 y570 w100 h30 gStopProcess, 중지

Gui, Processing:Show, x10 y%GuiY% w240 h680, 입고 자동 처리

; CSV파일 목록을 ListView에 추가 - 각 행에 순번과 파일명을 채워 넣기
; (csvFileList 는 이미 초기화된 전역 배열)
LV_Delete()  ; 기존 행 삭제 (혹은 초기 상태)
for index, filePath in csvFileList {
    SplitPath, filePath, fileName
    LV_Add("", index, fileName, "")  ; 첫 번째 열: index, 두 번째 열: fileName ; 세 번째 열 "완료"는 비워둠
}
LV_ModifyCol(1, 40)  ; 열 자동 크기 조절
LV_ModifyCol(2, 140)  ; 열 자동 크기 조절
LV_ModifyCol(3, 40)  ; 열 자동 크기 조절


; 각 행에 배경색을 설정 (예시: 짝수 행은 연한 분홍색, 홀수 행은 연한 녹색)
rowCount := LV_GetCount()
Loop, %rowCount%
{
    row := A_Index
    if (Mod(row, 2) = 0)
        LV_Modify(row, "ColBackColor1 0xFFCCCC")  ; 열1에 대해 적용 (색상값 0xFFCCCC)
    else
        LV_Modify(row, "ColBackColor1 0xCCFFCC")
}

; 자동 시작: 시작 버튼을 누를 필요 없이 바로 작업을 시작
StartProcess()

;────────────────────────────────────────────
; [프로세스 시작 함수]
StartProcess() {

    GuiControl, Processing:, Log, 작업을 시작합니다...
    global csvFileList, TaskTotal, TaskCount, FileCount1
    TaskTotal := csvFileList.MaxIndex()
    Loop, %TaskTotal%
    {
        TaskCount := A_Index
        ; 처리할 파일의 인덱스를 저장 (ListView 행 번호)
        fileRow := TaskCount

        progressVal := (A_Index / TaskTotal) * 100
        GuiControl, Processing:, Progress, %progressVal%
        GuiControl, Processing:, CurrentTask, 현재 작업: %A_Index% / %TaskTotal%
        GuiControl, Processing:, FileCount, 파일 처리 개수: %FileCount1%
        Sleep, 1000  ; QR 처리 작업 지연 (실제 코드로 교체)

        ; 현재 CSV 파일 처리 - 파일 경로 가져오기
        path1 := csvFileList[A_Index]
        GuiControl, Processing:, Log, % "작업 진행 중... " A_Index "/" TaskTotal "`n" path1

        ; --- 기존 QR 스캔 처리 (엑셀 열기, QR 전송 등) ---
        X1 := ComObjCreate("Excel.Application")
        wb := X1.Workbooks.Open(path1)
        ws := wb.Worksheets(1)
        TaskTotal := ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        ; path 문자열 가공 예제
        /*
        MsgBox, %path1%
        path2 := SubStr(currentDate, -15)
        path3 := SubStr(currentDate, -7)
        StringReplace, path1, path1, 즉시, S, All
        StringReplace, path1, path1, 에스윈, SWin, All
        StringReplace, path1, path1, 동일, Same, All
        StringReplace, path1, path1, 만, only, All
        StringReplace, path1, path1, 포름, HCHO, All
        path1 := RegExReplace(path1, "[^a-zA-Z0-9]")
        path2 := RegExReplace(path2, "[^a-zA-Z0-9]")
        StringTrimRight, path1, path1, 3
        StringReplace, path1, path1, 192168173,, All
        path1 := RegExReplace(path1, path2, "")
        path1 = %path3%_%path1%
        */
        ; CSV 파일 경로(path1)에서 파일명만 추출하여 변수 fileName에 저장
        SplitPath, path1, fileName

        ; currentDate와 fileName을 "_"로 연결하여 새로운 로그 문자열 생성
        logText := currentDate . "_" . fileName
        yymmToday := SubStr(A_Now, 1, 8)
        X1.Visible := false
        X1.DisplayAlerts := false

        ; QR 처리 루프 (간략화된 예제)
        Loop {
            BVarLine := "B" A_Index
            cB2 := ws.Range(BVarLine).Value
            if !cB2 {
                ; 시트 입고 처리 완료 시 해당 CSV파일의 ListView 행(순번)의 완료 열에 "V" 표시
                MsgBox,,, %path1% 시트 입고 처리 완료.,1
                TaskTotal := %TaskTotal% - %TaskCount%
                TaskCount := %TaskCount% - %TaskCount%
                FileCount1++
                GuiControl, Processing:, FileCount, 파일 처리 개수: %FileCount1%
                ; ListView에 완료 표시 ("V")
                LV_Modify(fileRow, "Col3", "V")  ; 완료 열에 V 기록
                ; 완료된 파일을 완료폴더로 이동
                doneFolder := folderBase "\완료된파일"
                IfNotExist, %doneFolder%
                    FileCreateDir, %doneFolder%

                SplitPath, path1, fileName
                newPath := doneFolder "\" fileName
                FileMove, %path1%, %newPath%, 1  ; 마지막 1은 덮어쓰기 허용
                ; 완료된 경우 해당 열의 배경색을 녹색으로 변경 (예, 0x90EE90: LightGreen)
                LV_Modify(fileRow, "LineBackgroundColor 0x90EE90")
                LV_Modify(fileRow, "Select Vis")  ; ?? 자동 스크롤: 해당 행으로 이동
                break
            }
            ; QR 코드 생성, 전송, 검증 등 (생략)
            WinActivate, FITI
            Sleep, 500

            setformat, float, 05
            cA2 := ws.Range("A" A_Index).Value
            BVarLine := "B" A_Index
            cB2 := ws.Range(BVarLine).Value
            setformat, float, 0
            CVarLine = C%A_index%
            cC2 := X1.Range(CVarLine).Value
            setformat, float, 0
            DVarLine = D%A_index%
            cD2 := X1.Range(DVarLine).Value
            if (StrLen(cD2) = 0)
            {
                cD2 := SubStr(currentDate, 3,2)
            }
            TN = @%cA2%%cD2%%cB2%@
            if (StrLen(cB2) = 13)
                TN = %cB2%
            if (TN = BeforeTN)
            {
                MsgBox,,,다른 접수 번호 찾는 중,1
                TaskCount++
                ; Log 컨트롤에 출력 (예: "QR 코드 처리 중: TaskCount/TaskTotal" 이후 새로운 로그 문자열 출력)
                GuiControl, Processing:, Log, % "QR 코드 처리 중: " A_Index "/" TaskTotal "`n" logText
                GuiControl, Processing:, Progress, % (A_Index / TaskTotal) * 100
                GuiControl, Processing:, CurrentTask, 현재 작업: %A_Index% / %TaskTotal%
                continue
            }
            BeforeTN = %TN%
            WinActivate, %a%
            TaskCount++
            ; Log 컨트롤에 출력 (예: "QR 코드 처리 중: TaskCount/TaskTotal" 이후 새로운 로그 문자열 출력)
            GuiControl, Processing:, Log, % "QR 코드 처리 중: " A_Index "/" TaskTotal "`n" logText
            GuiControl, Processing:, Progress, % (A_Index / TaskTotal) * 100
            GuiControl, Processing:, CurrentTask, 현재 작업: %A_Index% / %TaskTotal%
            send, %TN%
            ControlGetText, check1A, ThunderRT6TextBox2, QR스캔
            ControlGetText, check2A, ThunderRT6TextBox3, QR스캔
            ControlGetText, check3A, ThunderRT6TextBox4, QR스캔
            ControlGetText, check4A, ThunderRT6TextBox5, QR스캔
            checkTN1 = @%check1A%%check2A%%check3A%%check4A%@
            if (checkTN1 != TN)
            {
                Sleep, 1000
                send, %TN%
                MsgBox,,,전산 시스템이 느려서 접수 번호 보정중 입니다.`n잠시만 기다려 주세요. 2초_check1, 1
                MsgBox,,,전산 시스템이 느려서 접수 번호 보정중 입니다.`n잠시만 기다려 주세요. 1초_check1, 1
                ControlGetText, check1A, ThunderRT6TextBox2, QR스캔
                ControlGetText, check2A, ThunderRT6TextBox3, QR스캔
                ControlGetText, check3A, ThunderRT6TextBox4, QR스캔
                ControlGetText, check4A, ThunderRT6TextBox5, QR스캔
                checkTN1 = @%check1A%%check2A%%check3A%%check4A%@
                if (checkTN1 != TN)
                {
                    send, %TN%
                    MsgBox,,,전산 시스템이 느려서 접수 번호 보정중 입니다.`n잠시만 기다려 주세요. 5초_check1, 1
                    MsgBox,,,전산 시스템이 느려서 접수 번호 보정중 입니다.`n잠시만 기다려 주세요. 4초_check1, 1
                    MsgBox,,,전산 시스템이 느려서 접수 번호 보정중 입니다.`n잠시만 기다려 주세요. 3초_check1, 1
                    MsgBox,,,전산 시스템이 느려서 접수 번호 보정중 입니다.`n잠시만 기다려 주세요. 2초_check1, 1
                    MsgBox,,,전산 시스템이 느려서 접수 번호 보정중 입니다.`n잠시만 기다려 주세요. 1초_check1, 1
                    ControlGetText, check1A, ThunderRT6TextBox2, QR스캔
                    ControlGetText, check2A, ThunderRT6TextBox3, QR스캔
                    ControlGetText, check3A, ThunderRT6TextBox4, QR스캔
                    ControlGetText, check4A, ThunderRT6TextBox5, QR스캔
                    checkTN1 = @%check1A%%check2A%%check3A%%check4A%@
                    if (checkTN1 != TN)
                    {
                        log_dir_error = qrLog_input\%yymmToday%\%yymmToday%_qrLog.txt
                        FileCreateDir, qrLog_input\%yymmToday%\
                        timeRecord_error("시험관리시스템 작동이 느려서 자동QR스캔을 종료합니다._check1.")
                        MsgBox,,, 시험관리시스템 작동이 느려서 자동QR스캔을 종료하지 않고 진행합니다.._check1, 1
                        X1.Visible := true
                        X1.DisplayAlerts := true
                        X1.ActiveWorkBook.Close(0)
                        X1.Quit
                        continue
                    }
                }
            }
; --------------------------------
            Sleep,500
            PixelGetColor, set, 825, 171, RGB
            if set = 0xFF0000
            {
            MsgBox,,,%logText% %A_index% 번째 입고 처리 완료1,1
            send, %TN%
            #include Gdip_All.ahk
            pToken := Gdip_StartUp()
            pBitmap := Gdip_BitmapFromScreen("0|0|1400|800")
            IfExist, %A_ScriptDir%\qrLog_input\%yymmToday%\%logText%
            {
            }
            else
            {
                FileCreateDir,%A_ScriptDir%\qrLog_input\%yymmToday%\%logText%
            }
            name = %A_ScriptDir%/qrLog_input/%yymmToday%/%logText%/%A_index%_%TN%_QR.png
            Gdip_SaveBitmapToFile(pBitmap,  name)
            Gdip_DisposeImage(pBitmap)
            Gdip_Shutdown(pToken)
            msgbox, 0, 확인창, 다음 접수번호를 작업합니다., 1

            continue
            }
            if set = 0xC0C0C0
            {
            WinGetTitle, Title, A
            ControlClick, Button6, %Title%
            ControlClick, Button2, %Title%  ; Button2(저장) 으로 변경할 것
            Sleep, 500
            PixelGetColor, set, 825, 171, RGB
                if set = 0xFF0000
                {
                MsgBox,,,%logText% %A_index% 번째 입고 처리 완료.`n저장중 입니다. 2초,1
                MsgBox,,,%logText% %A_index% 번째 입고 처리 완료.`n저장중 입니다. 1초,1
                }
                else
                {
                FileCreateDir, qrLog_input\%yymmToday%\
                log_dir_error = qrLog_input\%yymmToday%\%TN%_qrLog.txt
                FileAppend, 접수번호 검색 안됨 / 접수QR 출고 또는 다른 부서인지 확인, %log_dir_error%
                }
            }
            ;comfirm2 = @H1112400000@
            ;send, %comfirm2%
            ControlClick, Button8, %Title%
            PixelGetColor, set, 725, 171, RGB
                if set = 0xC0C0C0
                {
                MsgBox,,,화면 저장중입니다.,1
                }
                if set = 0xFFFFFF
                {
                ;send, %comfirm2%
                ControlClick, Button8, %Title%
                Sleep, 2000
                PixelGetColor, set, 725, 171, RGB
                    if set = 0xC0C0C0
                    {
                    MsgBox,,,화면 저장중입니다.,1
                    }
                    if set = 0xFFFFFF
                    {
                    ControlClick, Button8, %Title%
                    MsgBox,,,화면 저장중입니다. 3초_check2,1
                    MsgBox,,,화면 저장중입니다. 2초_check2,1
                    MsgBox,,,화면 저장중입니다. 1초_check2,1
                    }
                    PixelGetColor, set, 725, 171, RGB
                    if set = 0xFFFFFF
                    {
                    FileCreateDir, qrLog_input\%yymmToday%\
                    log_dir_error = qrLog_input\%yymmToday%\%yymmToday%_qrLog.txt
                    FileAppend, 시험관리시스템 작동이 느려서 자동QR스캔을 종료합니다._check2., %log_dir_error%
                    MsgBox,,, 시험관리시스템 작동이 느려서 자동QR스캔을 종료합니다._check2,1
                    X1.Visible := true
                    X1.DisplayAlerts := true
                    X1.ActiveWorkBook.Close(0)
                    X1.Quit
                    ExitApp
                    return
                    }
                }
            send, %TN%
            ControlGetText, check1A, ThunderRT6TextBox2, QR스캔
            ControlGetText, check2A, ThunderRT6TextBox3, QR스캔
            ControlGetText, check3A, ThunderRT6TextBox4, QR스캔
            ControlGetText, check4A, ThunderRT6TextBox5, QR스캔
            checkTN1 = @%check1A%%check2A%%check3A%%check4A%@
            if (checkTN1 != TN)
            {
            Sleep, 1000
            send, %TN%
            MsgBox,,,전산 시스템이 느려서 접수 번호 보정중 입니다.`n잠시만 기다려 주세요. 2초_check3, 1
            MsgBox,,,전산 시스템이 느려서 접수 번호 보정중 입니다.`n잠시만 기다려 주세요. 1초_check3, 1
            ControlGetText, check1A, ThunderRT6TextBox2, QR스캔
            ControlGetText, check2A, ThunderRT6TextBox3, QR스캔
            ControlGetText, check3A, ThunderRT6TextBox4, QR스캔
            ControlGetText, check4A, ThunderRT6TextBox5, QR스캔
            checkTN1 = @%check1A%%check2A%%check3A%%check4A%@
                if (checkTN1 != TN)
                {
                send, %TN%
                MsgBox,,,전산 시스템이 느려서 접수 번호 보정중 입니다.`n잠시만 기다려 주세요. 5초_check3, 1
                MsgBox,,,전산 시스템이 느려서 접수 번호 보정중 입니다.`n잠시만 기다려 주세요. 4초_check3, 1
                MsgBox,,,전산 시스템이 느려서 접수 번호 보정중 입니다.`n잠시만 기다려 주세요. 3초_check3, 1
                MsgBox,,,전산 시스템이 느려서 접수 번호 보정중 입니다.`n잠시만 기다려 주세요. 2초_check3, 1
                MsgBox,,,전산 시스템이 느려서 접수 번호 보정중 입니다.`n잠시만 기다려 주세요. 1초_check3, 1
                ControlGetText, check1A, ThunderRT6TextBox2, QR스캔
                ControlGetText, check2A, ThunderRT6TextBox3, QR스캔
                ControlGetText, check3A, ThunderRT6TextBox4, QR스캔
                ControlGetText, check4A, ThunderRT6TextBox5, QR스캔
                checkTN1 = @%check1A%%check2A%%check3A%%check4A%@
                    if (checkTN1 != TN)
                    {
                    FileCreateDir, qrLog_input\%yymmToday%\
                    log_dir_error = qrLog_input\%yymmToday%\%yymmToday%_qrLog.txt
                    FileAppend, 시험관리시스템 작동이 느려서 자동QR스캔을 종료합니다._check3., %log_dir_error%
                    MsgBox,,, 시험관리시스템 작동이 느려서 자동QR스캔을 종료합니다._check3,1
                    X1.Visible := true
                    X1.DisplayAlerts := true
                    X1.ActiveWorkBook.Close(0)
                    X1.Quit
                    ExitApp
                    return
                    }
                }
            }
; ---------------------------------
            Sleep,500
            pToken := Gdip_StartUp()
            pBitmap := Gdip_BitmapFromScreen("0|0|1400|800")
            logFolder := A_ScriptDir "\qrLog_input\" SubStr(A_Now, 1, 8) "\" logText
            if not FileExist(logFolder)
                FileCreateDir, %logFolder%
            name := logFolder "\" A_Index "_" TN "_QR.png"
            Gdip_SaveBitmapToFile(pBitmap, name)
            Gdip_DisposeImage(pBitmap)
            Gdip_Shutdown(pToken)
            MsgBox, 0, 확인창, 다음 접수번호를 작업합니다., 1
        }
        X1.Visible := true
        X1.DisplayAlerts := true
        X1.ActiveWorkBook.Close(0)
        X1.Quit
    }
    GuiControl, Processing:, Log, 모든 작업이 완료되었습니다!
    ; QR 작업 완료 후, FITI 창 닫기
    IfWinExist, FITI
        WinClose, FITI
    ; 프로그램 종료
    ExitApp
}

timeRecord_error(sentence_error){
    global log_dir_error
    FileAppend, [%A_Year% %A_Mon%/%A_Mday% %A_Hour%:%A_Min% %A_sec%][%sentence_error%]`n, %log_dir_error%
}

;────────────────────────────────────────────
; [프로세스 중지 함수]
StopProcess() {
    GuiControl, Processing:, Log, 작업을 중지했습니다.
}

;────────────────────────────────────────────
; [GUI 닫기 처리 ? 창 닫더라도 스크립트는 ESC 키를 눌러야 종료]
GuiClose:
    MsgBox,,, 창을 닫더라도 스크립트는 ESC 키를 눌러야 종료됩니다., 1
    return

;────────────────────────────────────────────
; [ESC 키 핫키 ? ESC 누르면 종료]
Esc::
    if (IsObject(X1)) {
        try {
            X1.Visible := true
            X1.DisplayAlerts := true
            X1.ActiveWorkbook.Close(0)
            X1.Quit
        } catch e {
            MsgBox, 16, 오류, X1 종료 중 오류 발생: %e%
        }
        X1 := ""
    }
    ExitApp
return

;────────────────────────────────────────────
; [F3 키: 일시정지 토글]
F3::
    Pause
return
