/*
    Awesome SubNetwork class by nnnik on #ahkscript - 1/26/2016
    Input subnet mask and base IP and returns an array of all IPs in range.
*/
;<=====  Don't run me  ========================================================>
if (A_ScriptName == "SubNetwork.ahk"){
    MsgBox, % "This file is not a stand alone component.`nPlease run the main script."
    ExitApp
}

;<=====  SubNetwork  ==========================================================>
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