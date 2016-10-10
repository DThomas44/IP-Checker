/*
    IP Checker
    Author: Daniel Thomas

    This script monitors sets of IP addresses and keep a record of "last seen"
        dates and times. IP addresses can be added individually, or by masks
        (i.e. 192.168.1.0/24)

    Known Issues:
        ListView will not update a record if it's message is received while the
        user is interacting with the UI (Scrolling the LV).

    To Do:
        Move to arrays of IPs.
            Update array entry onmsg.
            Update LV entries based on IP onmsg.
        Allow saving of scan information into XML to repopulate on next launch.
        Add right click menu for interacting with an IP (SSH, Telnet, tracert, etc)
*/
;<=====  System Settings  =====================================================>
#SingleInstance Force
#NoEnv
SetBatchLines, -1
Debug := True

;<=====  Startup  =============================================================>
OnMessage(0x4a, "Receive_WM_COPYDATA")
DllCall("AllocConsole")
WinHide % "ahk_id " DllCall("GetConsoleWindow", "ptr")
threadCount := 0
SetTimer, CheckIPs, 60000

;<=====  Load XML  ===========================================================>
file := fileOpen(A_ScriptDir . "\IPChecker.xml", "r")
xml := file.read()
file.close()

doc := ComObjCreate("MSXML2.DOMdocument.6.0")
doc.async := false
if !doc.loadXML(xml)
    MsgBox, % "Failed to load XML!"

;<=====  Build Object  =======================================================>
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
Gui, Show,, IP Checker
Gui, +MinSize
return

;<=====  Subs  ================================================================>
AddRange:
    subnet := new SubNetwork(IPCtrlGetAddress(hNetmaskControl), IPCtrlGetAddress(hIPControl))
    For each, ip in subnet
        LV_Add("", ip)
    return

CheckIPs:
    for each, scanTime in settings.scanTimes
    {
        if (scanTime == A_Min)
            GoSub, CheckNow
    }
    return

CheckNow:
    SetTimer, CheckIPs, Off
    SetTimer, UpdateSB, 100
    SB_SetText("Pinging", 1)
    scannedHosts := 0
    totalHosts := LV_GetCount()
    GuiControl, Disable, CheckNow
    GuiControl, Disable, ClearList
    ; do the thing
    Loop, % LV_GetCount()
    {
        ; wait for a free thread slot
        while (threadCount >= settings["maxThreads"])
            sleep, 250

        row := A_Index
        LV_GetText(IP, row, 1)
        LV_Modify(row,,,"Pinging...")
        Run, %A_ScriptDir%\lib\PingMsg.ahk "%A_ScriptName%" %IP% %row%,,Hide
        threadCount++
    }
    GuiControl, Enable, CheckNow
    GuiControl, Enable, ClearList
    Sleep, 5000
    if threadCount
    {
        threadCount := 0
    }
    SetTimer, UpdateSB, Off
    SetTimer, CheckIPs, On
    activeCount := 0
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

ClearList:
    LV_Delete()
    return

Exit:
GuiClose:
    ExitApp

GuiSize:
	AutoXYWH(lv1, "wh")
    AutoXYWH(btn1, "y")
    AutoXYWH(btn2, "y")
	return

NetmaskDDL:
    GuiControlGet, NetmaskDDL
    StringTrimLeft, NetmaskDDL, NetmaskDDL, 1
    IPCtrlSetAddress(hNetmaskControl, settings.netmasks[NetmaskDDL])
    return

UpdateSB:
    SB_SetText(threadCount . "/" . settings["maxThreads"] . " threads", 2)
    SB_SetText(scannedHosts . "/" . totalHosts . " scanned", 3)
    return

;<=====  Classes  =============================================================>
/*
    Awesome SubNetwork class by nnnik on #ahkscript - 1/26/2016
    Input subnet mask and base IP and returns an array of all IPs in range.
*/
 class SubNetwork {

	__New(SubNetMask,IP)
	{
		This.SubNetMask := This.IPStringToNr(SubNetMask)
		This.IPMask := (This.IPStringToNr(IP)&This.SubNetMask)
		This.IPs := 1
		Loop 32
			If !(This.SubNetMask&(1<<(A_Index-1)))
				This.IPs*=2
		This.IPs -= 2
	}

	__Get(ipNr)
	{
		if (ipNr>This.IPs)
			return
		ip := 0
		Loop, 32
		{
			if !(This.SubNetMask&(1<<(A_Index-1)))
			{
				ip := ip|((ipNr&1)<<(A_Index-1))
				ipNr := ipNr>>1
			}
		}
		return This.NrToIPString(ip|This.IPMask)
	}

	IPStringToNr(String)
	{
		ip := 0
		RegExMatch(String,"(\d+)\.(\d+)\.(\d+)\.(\d+)",IPNr)
		Loop 4
			ip := ip|(IPNr%A_Index%)<<(8*(4-A_Index))
		return ip
	}

	NrToIPString(Nr)
	{
		ip := ""
		Loop % 4
			ip .= ((Nr>>(8*(4-A_Index))) & 0xFF) . "."
		return substr(ip,1,-1)
	}

	_NewEnum()
	{
		_Enum := This.Clone()
		_Enum.Iteration := 1
		return _Enum
	}

	Next(byref Iteration,byref IP)
	{
		Iteration := This.Iteration
		IP := This[Iteration]
		This.Iteration++
		return IP
	}

}

;<=====  Functions  ===========================================================>
AutoXYWH(ctrl_list, Attributes, Redraw = False){
    static cInfo := {}, New := []
    Loop, Parse, ctrl_list, |
    {
        ctrl := A_LoopField
        if ( cInfo[ctrl]._x = "" )
        {
            GuiControlGet, i, Pos, %ctrl%
            _x := A_GuiWidth  - iX
            _y := A_GuiHeight - iY
            _w := A_GuiWidth  - iW
            _h := A_GuiHeight - iH
            _a := RegExReplace(Attributes, "i)[^xywh]")
            cInfo[ctrl] := { _x:(_x), _y:(_y), _w:(_w), _h:(_h), _a:StrSplit(_a) }
        }
        else
        {
            if ( cInfo[ctrl]._a.1 = "" )
                Return
            New.x := A_GuiWidth  - cInfo[ctrl]._x
            New.y := A_GuiHeight - cInfo[ctrl]._y
            New.w := A_GuiWidth  - cInfo[ctrl]._w
            New.h := A_GuiHeight - cInfo[ctrl]._h
            Loop, % cInfo[ctrl]._a.MaxIndex()
            {
                ThisA   := cInfo[ctrl]._a[A_Index]
                Options .= ThisA New[ThisA] A_Space
            }
            GuiControl, % Redraw ? "MoveDraw" : "Move", % ctrl, % Options
        }
    }
}

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
    Gui, Add, Button, x+5p yp-1 gAddRange, Add Range
    Gui, Add, ListView, x5 y+5 w605 h500 HWNDlv1 vHostList, IP|Ping|Hostname|Last Seen
    Gui, Add, Button, x5 y+5 w100 HWNDbtn1 gCheckNow vCheckNow, Check Now
    Gui, Add, Button, x+5 yp w100 HWNDbtn2 gClearList vClearList, Clear List

    Gui, Add, StatusBar

    Gui, 1:Default

    SB_SetParts(150,150)
    SB_SetText("Ready", 1)
    LV_ModifyCol(1, "100")
    LV_ModifyCol(2, "75")
    LV_ModifyCol(3, "150")
    LV_ModifyCol(4, "150")
}

enumObj(Obj, indent := 0){
	if !isObject(Obj)
		MsgBox, Not an object!`n-->%obj%<--
	srtOut := ""
	for key, val in Obj
	{
		if isObject(val)
		{
			loop, %indent%
				strOut .= A_Tab
			strOut .= key . ":`n" . enumObj(val, indent + 1)
		}
		else
		{
			loop, %indent%
				strOut .= A_Tab
			strOut .= key . ": " . val . "`n"
		}
	}
	return strOut
}

IPCtrlSetAddress(hControl, ipaddress)
{
    static WM_USER := 0x400
    static IPM_SETADDRESS := WM_USER + 101

    ; Pack the IP address into a 32-bit word for use with SendMessage.
    ipaddrword := 0
    Loop, Parse, ipaddress, .
        ipaddrword := (ipaddrword * 256) + A_LoopField
    SendMessage IPM_SETADDRESS, 0, ipaddrword,, ahk_id %hControl%
}

IPCtrlGetAddress(hControl)
{
    static WM_USER := 0x400
    static IPM_GETADDRESS := WM_USER + 102

    VarSetCapacity(addrword, 4)
    SendMessage IPM_GETADDRESS, 0, &addrword,, ahk_id %hControl%
    return NumGet(addrword, 3, "UChar") "." NumGet(addrword, 2, "UChar") "." NumGet(addrword, 1, "UChar") "." NumGet(addrword, 0, "UChar")
}

Receive_WM_COPYDATA(wParam, lParam){
    global
    StringAddress := NumGet(lParam + 2*A_PtrSize)
    CopyOfData := StrGet(StringAddress)
    reply := StrSplit(CopyOfData, "|")
    if (reply[3] != "TIMEOUT")
    {
        FormatTime, scanTime, A_Now, HH:mm MM/dd/yyyy
        LV_Modify(reply[1],,,reply[3],reply[2], scanTime)
    } else {
        LV_Modify(reply[1],,,reply[3])
    }
    threadCount--
    scannedHosts++
    return true
}
