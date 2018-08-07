; Authors: Manadar, GarryFrost
; Contributor: WideBoyDixon

#include <WinAPI.au3>
#include <WindowsConstants.au3>
#include <Array.au3>

#include-once

Local $hStub_KeyProc = DllCallbackRegister("_HotString_KeyProc", "long", "int;wparam;lparam")
Local $hmod = _WinAPI_GetModuleHandle(0)
Local $hHook = _WinAPI_SetWindowsHookEx($WH_KEYBOARD_LL, DllCallbackGetPtr($hStub_KeyProc), $hmod)
Local $buffer = ""
Local $hotstrings[1]
Local $hotfuncs[1]
Local $hotargs[1]
Local $hWnd = GUICreate("")
GUIRegisterMsg($WM_KEYDOWN, "_HotString_GUIKeyProc")
OnAutoItExitRegister("_HotString_OnAutoItExit")

Func _HotStringSet($hotstring, $func, $args)
    _ArrayAdd($hotstrings, $hotstring)
    _ArrayAdd($hotfuncs, $func)
    _ArrayAdd($hotargs,$args)
EndFunc

Func _HotString_EvaluateKey($keycode)
    If (($keycode > 64) And ($keycode < 91)) _ ; A - Z
            Or (($keycode > 47) And ($keycode < 58)) Then ; 0 - 9
        $buffer &= Chr($keycode)
        $a = _ArraySearch($hotstrings, $buffer)
		If ( $a >= 0 ) Then
			Call($hotfuncs[$a], $hotargs[$a])
		EndIf
    ElseIf ($keycode > 159) And ($keycode < 164) Then
        Return
    Else
        $buffer = ""
    EndIf
EndFunc   ;==>EvaluateKey

Func _HotString_GUIKeyProc($hWnd, $Msg, $wParam, $lParam)
	_HotString_EvaluateKey(Number($wParam))
EndFunc

Func _HotString_KeyProc($nCode, $wParam, $lParam)
    Local $tKEYHOOKS
    $tKEYHOOKS = DllStructCreate($tagKBDLLHOOKSTRUCT, $lParam)
    If $nCode < 0 Then
        Return _WinAPI_CallNextHookEx($hHook, $nCode, $wParam, $lParam)
    EndIf
    If $wParam = $WM_KEYDOWN Then
		$vkKey = DllStructGetData($tKEYHOOKS, "vkCode")
		_WinAPI_PostMessage($hWnd, $WM_KEYDOWN, $vkKey, 0)
    EndIf
    Return _WinAPI_CallNextHookEx($hHook, $nCode, $wParam, $lParam)
EndFunc   ;==>_KeyProc

Func _HotString_OnAutoItExit()
    _WinAPI_UnhookWindowsHookEx($hHook)
    DllCallbackFree($hStub_KeyProc)
EndFunc   ;==>OnAutoItExit

;; ==========================
;; This is my script.
;; ==========================

Dim $ExpandedText[100]

$ExpandedText[0] = "CallArgArray"
$ExpandedText[1] = "hi, Good Morning "
$ExpandedText[2] = "by the way, "
$ExpandedText[3] = "Let me know if you need further details."
$ExpandedText[4] = "Let me know if you need further information on this."
$ExpandedText[5] = "Regards,+{ENTER}Sathya Sivam"
$ExpandedText[6] = "Thanks,+{ENTER}Sathya Sivam"
$ExpandedText[7] = "Office 365 Excel "
$ExpandedText[8] = "Excel 2016 vba "
$ExpandedText[9] = "SharePoint "


_HotStringSet("hgm","ExpandText",1)
_HotStringSet("btw", "ExpandText",2)
_HotStringSet("lfdd", "ExpandText",3)
_HotStringSet("lfii", "ExpandText",4)
_HotStringSet("rgs", "ExpandText",5)
_HotStringSet("tks", "ExpandText",6)
_HotStringSet("xll", "ExpandText",7)
_HotStringSet("xlv", "ExpandText",8)
_HotStringSet("spp", "ExpandText",9)


While 1
    Sleep(10)
WEnd

Func ExpandText($ArgIndx)
	 ; MsgBox(0, "AutoIt Example", $hotstrings[$ArgIndx])
    Send("{BACKSPACE " & StringLen($hotstrings[$ArgIndx]) & "}")
	
	Switch $ArgIndx
		Case 16
			Send("1568925{TAB}")
		Case 23
			Send("1568925{TAB}:){ENTER}")
		Case Else
			Send($ExpandedText[$ArgIndx])
	EndSwitch
    ; if $ArgIndx = 16 then
		; Send("1568925{TAB}ppp{ENTER}")
	; else
		; Send($ExpandedText[$ArgIndx])
	; endif
	; MsgBox(0, "AutoIt Example", $ExpandedText[$ArgIndx])
EndFunc
