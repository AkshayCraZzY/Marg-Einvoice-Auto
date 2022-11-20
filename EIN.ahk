#NoEnv    
#Persistent
#SingleInstance, Force
#Include <WaitPixelColor>
#include <Gdip_All>
#include <csv>
DetectHiddenWindows, On
SetWorkingDir %A_ScriptDir%
sGst=%A_WorkingDir%\files\temp_data\sGst.jpg
gst=%A_WorkingDir%\files\temp_data\gst.txt
debug_gst=%A_WorkingDir%\files\temp_data\debug_gst.txt
debug=0
;sGst=%A_WorkingDir%\files\temp_data\sGst.jpg
;MsgBox, 38 ,MargAuto E-Invoice,Is E-Invoice window open?
if not WinExist("ahk_exe margwin.exe")
{

}
Else
{
    WinActivate
}
sleep 300
Initial:
    If not (debug)
        Run, gst.py,,HIDE,pyid
    ;msgbox, % pyid
    PixelSearch, Px, Py, 259,143, 260,144, 0xFFE9D2, 2, Fast
    If ErrorLevel
    {
        InputBox, eindate, E-Invoice Date, Please enter date or E-Invoicing., , 210, 150,,,,,
        if ErrorLevel
        {
            MsgBox,Cancelled search.
            exitapp
        }
        ;eindate = 7
        if not WinExist("ahk_exe margwin.exe")
        {

        }
        Else
        {
            WinActivate
        }

        Send {Alt}
        sleep 300
        Send {G}
        sleep 300
        Send {G}
        sleep 300
        Send {Enter}
        sleep 300
        Send {Down}
        sleep 300
        Send {Enter}
        sleep 300
        SEND, % eindate
        SLEEP 300
        SEND {ENTER}
        SLEEP 300
        SEND, % eindate
        SLEEP 300
        SEND {ENTER}
        SLEEP 300
        SEND {ENTER}
        SLEEP 300
        MouseMove, 687,875
        varr := WaitPixelColor(0xC8C8C8,687,875,600)
        If ( varr = 0 )
        {
            SLEEP 300
            SEND {ESC}
        }
        Else If ( varr = 1 )
        {
            Msgbox Error!
        }

        SLEEP 300
        PixelSearch, Px, Py, 22,106, 24,108, 0x000000, 2, Fast
        if ErrorLevel
        {
            MouseClick,, 24, 107,,0
        }
        else
        {
            MouseClick,, 24, 107,,0
            SLEEP 500
            MouseClick,, 24, 107,,0
        }

        SLEEP 500
        MouseClick,,1816,90,,0
        sleep 300
        send {AltDown}
        sleep 300
        send Y
        sleep 300
        send {AltUp}
        SLEEP 600
    }

LControl & LShift::
    Gui, Destroy
    Send {up}
    sleep 200
    Gosub, Check
    return

Check:
{
    Loop
    {
        PixelGetColor, color, 1040,740
        IF(color=0x48000)
        {
            send {Down}
            SLEEP 150
        }
        else if(color=0x0000FF)
        {
            MouseClick,, 1040,740,,0
            ;MSGBOX, HERE
            WaitPixelColor(0xE8E8E8,900,155)
           ; sleep 400
            Ocr:    
                pToken:=Gdip_Startup()
                pBitmap:=Gdip_BitmapFromScreen()
                pGst:=Gdip_CloneBitmapArea(pBitmap, 904,90,183,25)
                Gdip_SaveBitmapToFile(pGst,sGst)
                clipboard := ""
        
                ;MouseGetPos,Tx,Ty
                Tooltip, Please Wait`n Fetching GST data..,920,540
                ReplaceSystemCursor("IDC_ARROW","IDC_WAIT")
                Runwait,ocr.py,,hide,pyid
                ;RunWait %comspec% /c  "%A_WorkingDir%\c2t\Capture2Text_CLI.exe  -i %sGst% > %gst%",,hide,Pid
                FileReadLine, gstno, %gst%, 1
                StringReplace gstno, gstno,.,, All
                StringReplace gstno, gstno,*,, All
                StringReplace gstno, gstno,#,, All
                ;StringReplace gstno, gstno,:,, All
                /*
                StringReplace gstno, gstno,\,, All
                StringReplace gstno, gstno,’,, All
                StringReplace gstno, gstno,|,, All
                StringReplace gstno, gstno,â,, All
                StringReplace gstno, gstno,€,, All
                StringReplace gstno, gstno,™,, All
                */

                file := FileOpen(gst, "w") 
                file.write(gstno)
                file.close()
            ;exitapp
            If(StrLen(gstno)!=15)
            {
                ReplaceSystemCursor()
                ToolTip,
                MsgBox, 54 ,MargAuto E-Invoice, Unable to get proper GST. `n[%gstno%]`nContinue to enter manually.
                IfMsgBox Cancel
                {
                    Gui, Destroy
                    break
                }
                IfMsgBox TryAgain
                {
                    ;Gui, Destroy
                    Gosub Ocr
                    return 
                }
                IfMsgBox Continue
                {
                    InputBox, gstno, GST Number, Please enter the GST number., , 210, 150,,,,,%gstno%
                    if ErrorLevel
                    {
                        MsgBox,Cancelled search.
                        Gui, Destroy
                        BREAK
                    }
                    else
                    {
                        file := FileOpen(gst, "w") 
                        file.write(gstno)
                        file.close()
                        ;sleep 200
                        Tooltip, Please Wait`nGetting details.,920,540
                        ReplaceSystemCursor("IDC_ARROW","IDC_WAIT")
                        ;Gui, Destroy
                    }
                        ;Gosub Ocr 
                }
            }
            ;Tooltip, Fetching GST data.,920,540
            sleep 200
            global miss_addr2=0
            global miss_addr1=0
            global miss_loc=0
            global miss_pin=0
            global miss=
            FileDelete, %A_WorkingDir%\files\temp_data\*.csv
           
           
           
            PixelSearch, Px, Py, 903, 281, 1190,295, 0xA0D1FF, 5, Fast
            if ErrorLevel
            {
                miss_addr1=1
                miss=%miss%Address1,
            }
            PixelSearch, Px, Py, 903,306, 1190,319, 0xA0D1FF, 5, Fast
            if ErrorLevel
            {
                miss_addr2=1
                miss=%miss%Address2,
            }
            PixelSearch, Px, Py, 903,330, 1190,345, 0xA0D1FF, 5, Fast
            if ErrorLevel
            {
                miss_loc=1
                miss=%miss%Location,
            }
            PixelSearch, Px, Py, 1135,355, 1192,367, 0xA0D1FF, 5, Fast
            if ErrorLevel
            {
                miss_pin=1
                miss=%miss%Pincode
            }
            If not (debug)
            {
                ToolTip,Please Wait. Fetching GST data.`nMissing Data:`n%miss%,920,540
                While ! FileExist(A_WorkingDir "\files\temp_data\*.csv")
                    Sleep 250
                Loop, %A_WorkingDir%\files\temp_data\*.csv
                    gstCSVPath=%A_Loopfilefullpath%
                disGST(gstCSVPath)
            }
            Else
            {
                FileAppend, %gstno%, %debug_gst%
                WinActivate
                sleep 200
                send, {Esc}
                sleep 200
                Send {Down}
                sleep 200
                Send {Down}
                SLEEP 100
                Gosub, Check
            }
            ReplaceSystemCursor()
            ToolTip,
            if FileExist(sGst)
                FileRecycle, %sGst%
            BREAK
        }
    }
    return
}
GuiEscape:
GuiClose:
        Gui, Destroy
        return

disGST(FilePath)
{
    global Gst
    global Pin
    global Loc
    global Stat
    global Addr

    Gui,+AlwaysOnTop +Caption +LastFound +ToolWindow
    Gui, Font,s11 Bold
    Gui Add, Text, x0 y0 w166 h24 +0x200 +Center +Border, PINCODE
    Gui Add, Text, x166 y0 w166 h24 +0x200 +Center +Border, LOCATION
    Gui Add, Text, x332 y0 w100 h24 +0x200 +Center +Border, STATUS
    Gui Add, Text, x432 y0 w550 h24 +0x200 +Center +Border, ADDRESS
    Gui, Font,s11 Normal
    Gui Add, Edit, x0 y24 w166 h21 +Readonly +Center +Readonly vPin
    Gui Add, Edit, x166 y24 w166 h21 +Readonly +Center +Readonly vLoc
    Gui Add, Edit, x332 y24 w100 h21 +wrap +Center +Readonly vStat
    Gui Add, Edit, x432 y24 w550 h21 -wrap +Readonly vAddr
    sheight := A_ScreenHeight-505
    CSV_Load(FilePath,"data")
    isvalid_csv:=CSV_ReadCell("data",1,1)
    whyinvalid_csv:=CSV_ReadCell("data",1,2)
    if(isvalid_csv="INVALID")
    {
        MsgBox, 48 ,MargAuto E-Invoice,Invalid GST`nUnable to get data`nReason: %whyinvalid_csv%.
        Gui, Destroy
        return
    }

    gst_csv:=CSV_ReadCell("data",2,1)
    pin_csv:=CSV_ReadCell("data",2,2)
    loc_csv:=CSV_ReadCell("data",2,3)
    stat_csv:=CSV_ReadCell("data",2,5)
    addr_csv:=CSV_ReadCell("data",2,4)



    GuiControl,, Pin, %pin_csv%
    GuiControl,, Loc, %loc_csv%
    if (stat_csv="Active")
    {
        Gui, Font, cGreen
        GuiControl, Font, Stat
        ;return
    }
    else
    {
        Gui, Font, cRed
        GuiControl, Font, Stat
        ;return
    }
    GuiControl,, Stat, %stat_csv%
    GuiControl,, Addr, %addr_csv%

    Gui Show, y%sheight% w982 h51, %gst_csv%
    ;clipboard := "" 
    ;clipboard := pin_csv
    if not WinExist("ahk_exe margwin.exe")
    {

    }
    Else
    {
        WinActivate
    }
    
    if(miss_addr1=1 or miss_addr2=1 or miss_loc=1 or miss_pin=1)
    {
        sleep 100
        send {Enter}
        sleep 100
        send {Tab}
        sleep 100
        send {Tab}
        sleep 100
        send {Tab}
        sleep 100
        send {Tab}
        sleep 100
        send {Tab}
        sleep 100
        if(miss_addr1)
            SendInput, %addr_csv%
        sleep 100
        send {Tab}
        sleep 100
        if(miss_addr2)
            SendInput, %loc_csv%
        sleep 100
        send {Tab}
        sleep 100
        if(miss_loc)
            SendInput, %loc_csv%
        sleep 100
        send {Tab}
        sleep 100
        send {Tab}
        sleep 100
        if(miss_pin)
            SendInput, %pin_csv%
        SLEEP 100
    }
    Else
    {
        Send {Tab}
        sleep 100
        Send {Enter}
        Sleep 100
    }
    Return
    
}
ReplaceSystemCursor(old := "", new := "")
{
   static IMAGE_CURSOR := 2, SPI_SETCURSORS := 0x57
        , exitFunc := Func("ReplaceSystemCursor").Bind("", "")
        , setOnExit := false
        , SysCursors := { IDC_APPSTARTING: 32650
                        , IDC_ARROW      : 32512
                        , IDC_CROSS      : 32515
                        , IDC_HAND       : 32649
                        , IDC_HELP       : 32651
                        , IDC_IBEAM      : 32513
                        , IDC_NO         : 32648
                        , IDC_SIZEALL    : 32646
                        , IDC_SIZENESW   : 32643
                        , IDC_SIZENWSE   : 32642
                        , IDC_SIZEWE     : 32644
                        , IDC_SIZENS     : 32645 
                        , IDC_UPARROW    : 32516
                        , IDC_WAIT       : 32514 }
   if !old {
      DllCall("SystemParametersInfo", "UInt", SPI_SETCURSORS, "UInt", 0, "UInt", 0, "UInt", 0)
      OnExit(exitFunc, 0), setOnExit := false
   }
   else  {
      hCursor := DllCall("LoadCursor", "Ptr", 0, "UInt", SysCursors[new], "Ptr")
      hCopy := DllCall("CopyImage", "Ptr", hCursor, "UInt", IMAGE_CURSOR, "Int", 0, "Int", 0, "UInt", 0, "Ptr")
      DllCall("SetSystemCursor", "Ptr", hCopy, "UInt", SysCursors[old])
      if !setOnExit
         OnExit(exitFunc), setOnExit := true
   }
}
^X::
    Process, Close , %pyid%
    WinClose, ahk_exe chromedriver.exe
    WinGet, active_id, ID, ahk_exe py.exe
    Process, Close , %active_id%
    WinClose, ahk_exe py.exe
    WinGet, chromeWindows, List, ahk_exe chrome.exe
    Loop, % chromeWindows
        WinClose, % "ahk_id " chromeWindows%A_Index%
    MSGBOX,STOPPED %pyid% %active_id%
    Gui, Destroy
    exitapp
^z::
    Process, Close , %pyid%
    WinClose, ahk_exe chromedriver.exe
    WinGet, active_id, ID, ahk_exe py.exe
    Process, Close , %active_id%
    WinClose, ahk_exe py.exe
    WinGet, chromeWindows, List, ahk_exe chrome.exe
    Loop, % chromeWindows
        WinClose, % "ahk_id " chromeWindows%A_Index%
    Reload


