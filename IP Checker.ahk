/*
    IP Checker
    Author: Daniel Thomas

    This script monitors sets of IP addresses and keep a record of "last seen"
        dates and times.

    Known Issues:
        ListView will not update a record if it's message is received while the
        user is interacting with the UI (Scrolling the LV).
*/
;<=====  System Settings  =====================================================>
#SingleInstance Force
#NoEnv
SetBatchLines, -1

;<=====  Startup  =============================================================>
OnMessage(0x4a, "Receive_WM_COPYDATA")
DllCall("AllocConsole")
WinHide % "ahk_id " DllCall("GetConsoleWindow", "ptr")
threadCount := 0
SetTimer, CheckTime, 60000

;<=====  Load XML  ===========================================================>
file := fileOpen(A_ScriptDir . "\IPChecker.xml", "r")
xml := file.read()
file.close()

doc := ComObjCreate("MSXML2.DOMdocument.6.0")
doc.async := false
if !doc.loadXML(xml)
{
    MsgBox, % "Failed to load XML!"
}

;<=====  Build Settings Object  ===============================================>
settings := Object()
for Node in doc.selectNodes("/IPChecker/settings/*")
{
	settings[Node.tagName] := Node.text
}

settings.scanTimes := Object()
for Node in doc.selectNodes("/IPChecker/settings/rowset[@id='scanTimes']/*")
{
	settings.scanTimes[Node.getAttribute("id")] := Node.getAttribute("minute")
}

settings.netmasks := Object()
for Node in doc.selectNodes("/IPChecker/settings/rowset[@id='netmasks']/*")
{
    settings.netmasks[Node.getAttribute("id")] := Node.getAttribute("netmask")
}

;<=====  Main  ================================================================>
buildGUI()
global imported := 0
global importFile := ""
Gui, Show,, IP Checker
Gui, +MinSize
return

;<=====  Subs  ================================================================>
AddRange:
    subnet := new SubNetwork(IPCtrlGetAddress(hNetmaskControl), IPCtrlGetAddress(hIPControl))
    For each, ip in subnet
    {
        LV_Add("", ip)
        ;IPs.Push
    }
    return

ClearList:
    LV_Delete()
    return

Exit:
GuiClose:
    if imported
    {
        MsgBox, 4,, % "File " . importFile . " was imported.`nSave before exit?"
        IfMsgBox, Yes
        {
            GoSub ExportList
        }
    }
    ExitApp

ExportList:
    strout := """IP"",""Ping"",""Host"",""Last Seen"",""Description""`n"
    Loop, % LV_GetCount()
    {
        row := A_Index
        LV_GetText(IP, row, 1)
        LV_GetText(Ping, row, 2)
        LV_GetText(Host, row, 3)
        LV_GetText(Seen, row, 4)
        LV_GetText(Desc, row, 5)
        strout .= """" . IP . """,""" . Ping . """,""" . Host . """,""" . Seen
            . """,""" . Desc . """`n"
    }
    FileSelectFile, saveFile, S16, % A_MyDocuments, Save IP scan results, CSV (*.csv)
    if instr(saveFile, ".csv")
    {
        saveFile := strReplace(saveFile, ".csv") 
    }
    file := fileOpen(saveFile . ".csv", "w")
    file.write(strout)
    file.close()
    MsgBox, % "Exported list to " . saveFile
    Return

GuiSize:
	AutoXYWH(lv1, "wh")
    AutoXYWH(btn1, "y")
    AutoXYWH(btn2, "y")
    AutoXYWH(btn3, "xy")
	return

NetmaskDDL:
    GuiControlGet, NetmaskDDL
    StringTrimLeft, NetmaskDDL, NetmaskDDL, 1
    IPCtrlSetAddress(hNetmaskControl, settings.netmasks[NetmaskDDL])
    return

;<=====  Functions  ===========================================================>
buildGUI(){
	Global

	Gui, -DPIScale
	Gui, +Resize
	Gui, Margin, 5, 5

    Gui, Add, Text, x5 y8 w60, IP Address:
    Gui, Add, Custom, x+0 yp-3 ClassSysIPAddress32 r1 w150 hwndhIPControl
    IPCtrlSetAddress(hIPControl, A_IPAddress1)
    Gui, Add, Text, x+5 yp+3 w50, Netmask:
    Gui, Add, DropDownList, x+0 yp-3 r10 w50 gNetmaskDDL vNetmaskDDL, /4|/5|/6|/7|/8|/9|/10|/11|/12|/13|/14|/15|/16|/17|/18|/19|/20|/21|/22|/23|/24||/25|/26|/27|/28|/29|/30
    Gui, Add, Custom, x+5 yp ClassSysIPAddress32 r1 w150 hwndhNetmaskControl +Disabled
    IPCtrlSetAddress(hNetmaskControl, settings.netmasks[24])
    Gui, Add, Button, x+5p yp-1 gAddRange vAddRange, Add Range
    Gui, Add, Button, x+5 yp gImportList vImportList, Import List
    Gui, Add, ListView, x5 y+5 w605 h500 HWNDlv1 gHostList vHostList, IP|Ping|Hostname|Last Seen|Comment
    Gui, Add, Button, x5 y+5 w100 HWNDbtn1 gCheckHosts vCheckHosts, Check Now
    Gui, Add, Button, x+5 yp w100 HWNDbtn2 gClearList vClearList, Clear List
    Gui, Add, Button, x+300 yp w100 HWNDbtn3 gExportList vExportList, Export

    Gui, Add, StatusBar

    Gui, 1:Default

    SB_SetParts(150,150)
    SB_SetText("Ready", 1)
    LV_ModifyCol(1, "100")
    LV_ModifyCol(2, "75")
    LV_ModifyCol(3, "150")
    LV_ModifyCol(4, "150")
}

CheckHosts(){
    Global

    ;Setup Timers
    SetTimer, CheckTime, Off
    SetTimer, StaleThreads, Off
    SetTimer, UpdateSB, 100

    ;Modify statusbar
    SB_SetText("Pinging", 1)

    ;Setup local variables
    scannedHosts := 0
    activeCount := 0
    totalHosts := LV_GetCount()
    threads := Object()
    hosts := Object()

    ;Disable GUI controls
    GuiControl, Disable, hIPControl
    GuiControl, Disable, NetmaskDDL
    GuiControl, Disable, hNetmaskControl
    GuiControl, Disable, AddRange
    GuiControl, Disable, ImportList
    GuiControl, Disable, CheckHosts
    GuiControl, Disable, ClearList
    GuiControl, Disable, ExportList

    ;Do the thing
    Loop, % LV_GetCount()
    {
        ;Wait for a free thread slot
        while (threads.MaxIndex() >= settings["maxThreads"])
        {
            sleep, 250
        }

        ;Start thread to check host
        row := A_Index
        LV_GetText(IP, row, 1)
        LV_Modify(row,,,"Pinging...")
        Run, %A_ScriptDir%\lib\PingMsg.ahk "%A_ScriptName%" %IP% %row%,,Hide, threadID

        ;Add thread ID to arrays for tracking
        hosts[A_Index, "threadID"] := threadID
        threads.Push(threadID)
    }

    ; Allow 1 minute for threads to return
    SetTimer, StaleThreads, 60000
    
    while (threads.MaxIndex() > 0)
    {
        sleep, 100
    }
    
    ;Enable GUI controls    
    GuiControl, Enable, hIPControl
    GuiControl, Enable, NetmaskDDL
    GuiControl, Enable, hNetmaskControl
    GuiControl, Enable, AddRange
    GuiControl, Enable, ImportList
    GuiControl, Enable, CheckHosts
    GuiControl, Enable, ClearList
    GuiControl, Enable, ExportList

    ;Modify LV columns to fit content
    Loop, % 5
    {
        LV_ModifyCol(A_Index, "AutoHDR")
    }

    ;Toggle timers
    SetTimer, UpdateSB, Off
    SetTimer, CheckTime, On

    ;Count active hosts and update SB
    Loop, % LV_GetCount()
    {
        LV_GetText(rowStatus, A_Index, 2)
        if ((rowStatus != "TIMEOUT")&&(rowStatus != "Pinging..."))
            activeCount++
    }
    SB_SetText("Ready", 1)
    SB_SetText(activeCount . "/" . LV_GetCount() . " hosts alive", 2)
    SB_SetText("", 3)
    return
}

CheckTime(){
    Global settings
    for each, scanTime in settings.scanTimes
    {
        if (scanTime == A_Min)
        {
            CheckHosts()
        }
    }
    return
}

HostList(CtrlHwnd, GuiEvent, EventInfo, ErrLevel := ""){
    if (GuiEvent == "DoubleClick")
    {
        row := LV_GetNext()
        LV_GetText(IP, row)
        InputBox, desc,, % "Enter description for " . IP
        if ErrorLevel
            return
        else if IP
        {
            LV_Modify(row,,,,,,desc)
        }
    }
    else if (GuiEvent == "RightClick")
    {
        ;show context menu here
        return
    }
    return
}

importList(){
    FileSelectFile, fileIn,,, Select CSV file to import, IP List (*.csv)
    file := fileOpen(fileIn, "r")
    ;iterate file adding fields to LV
    while !file.AtEOF
    {
        line := StrSplit(file.ReadLine(), ",", """")
        if(A_Index == 1)
        {
            Continue
        }
        Else
        {
            LV_Add("", line[1], line[2], line[3], line[4], StrReplace(StrReplace(line[5], """"), "`n"))
        }
    }
    file.close()

    Loop, 5
    {
        LV_ModifyCol(A_Index, "AutoHDR")
    }
    imported := 1
    return
}

IPCtrlSetAddress(hControl, ipaddress){
    static WM_USER := 0x400
    static IPM_SETADDRESS := WM_USER + 101

    ; Pack the IP address into a 32-bit word for use with SendMessage.
    ipaddrword := 0
    Loop, Parse, ipaddress, .
    {
        ipaddrword := (ipaddrword * 256) + A_LoopField
    }
    SendMessage IPM_SETADDRESS, 0, ipaddrword,, ahk_id %hControl%
}

IPCtrlGetAddress(hControl){
    static WM_USER := 0x400
    static IPM_GETADDRESS := WM_USER + 102

    VarSetCapacity(addrword, 4)
    SendMessage IPM_GETADDRESS, 0, &addrword,, ahk_id %hControl%
    return NumGet(addrword, 3, "UChar") "." NumGet(addrword, 2, "UChar") "." NumGet(addrword, 1, "UChar") "." NumGet(addrword, 0, "UChar")
}

Receive_WM_COPYDATA(wParam, lParam){
    Global
    Critical
    StringAddress := NumGet(lParam + 2*A_PtrSize)
    CopyOfData := StrGet(StringAddress)

    ;hostID|HostName|PingTime|IP
    reply := StrSplit(CopyOfData, "|")

    ;Clear thread from threads array
    Loop, % threads.MaxIndex()
    {
        if (threads[A_Index] == hosts[reply[1], "threadID"])
            threads.removeAt(A_Index)
    }

    ;Process reply
    if (reply[2] == reply[4])
    {
        reply[2] := "<No DNS Resolution>"
    }

    if (reply[3] != "TIMEOUT")
    {
        FormatTime, scanTime, A_Now, HH:mm MM/dd/yyyy
        LV_Modify(reply[1],,,reply[3] . "ms",reply[2], scanTime)
    } else {
        LV_Modify(reply[1],,,reply[3])
    }
    scannedHosts++
    return true
}

StaleThreads(){
    Global

    ;Loop hosts array
    for host in hosts
    {
        ;Store index in hostID for use in nested loop
        hostID := A_Index

        ;Get host thread ID
        hostThreadID := hosts[hostID, "threadID"]

        ;Check if hosts threadID is still in the threads array
        for each, tid in threads
        {
            ;If found, mark thread as stale and remove from array
            LV_Modify(reply[1],,,"TIMEOUT")
            threads.RemoveAt(A_Index)
        }
    }
}

UpdateSB(){
    Global
    SB_SetText(threads.MaxIndex() . "/" . settings["maxThreads"] . " threads", 2)
    SB_SetText(scannedHosts . "/" . totalHosts . " scanned", 3)
    return
}

;<=====  Includes  ============================================================>
#Include %A_ScriptDir%\lib
#Include AutoXYWH.ahk
#Include Common Functions.ahk
#Include SubNetwork.ahk
