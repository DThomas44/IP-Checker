;<=====  Don't run me  ========================================================>
if (A_ScriptName == "AutoXYWH.ahk"){
    MsgBox, % "This file is not a stand alone component.`nPlease run the main script."
    ExitApp
}

;<=====  AutoXYWH  ============================================================>
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