; date: 2017.5.10
; author: junericlin

#include <MsgBoxConstants.au3>
#include <Debug.au3>

Global $isDebugOut=false;
   Global $outlookPath="C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE"
   Global $outlookClass="[CLASS:rctrl_renwnd32]";


main();
exit;

;*******************Functioin definition********************
Func main()
init()
mainLoop();
EndFunc

Func init()
If $isDebugOut Then
_DebugSetup("Keep Outlook always on")
EndIf
EndFunc


Func trace($msg)
   Local $exeName="keepOutlookAlwaysOn.exe";
   Local $tag=$exeName&"::Debug> ";
   If $isDebugOut Then
	  _DebugOut($tag&$msg);
   EndIf
   ToolTip($tag&$msg)
EndFunc


Func clearTooTip()
   ToolTip("");
EndFunc

Func test()
   Local $hwnd=WinWait($outlookClass, "", 5); wait 5 seconds
   If $hwnd =0 Then
	  trace("Can't find outlook windows Class:"&$outlookClass);
	  break;
   EndIf
   trace("try to resize and move to pos outlook-window");
   WinMove($hwnd,"", 84, 60, 1432, 844);

   trace("try to minimize outlook-window");
   ;WinSetState($hwnd,"", @SW_MINIMIZE);
   Exit;
EndFunc

Func mainLoop()
   ;test();
   While 1
	  If ProcessExists("OUTLOOK.EXE") Then ; Check if the Notepad process is running.
		 ;trace("OUTLOOK.EXE is running")
	  Else
		 trace("OUTLOOK.EXE is not running, try to open it.");

		 Local $pid=Run($outlookPath);
		 If $pid = 0 Then ; error
			trace("Can't run program: "&$outlookPath);
			break;
		 EndIf
		 Sleep(1500);//wait for the window to appear.
		 Local $hwnd=WinWait($outlookClass, "", 5); wait 5 seconds
		 If $hwnd =0 Then
			trace("Can't find outlook windows Class:"&$outlookClass);
			break;
		 EndIf
		 trace("try to resize and move to pos outlook-window");
		 WinMove($hwnd,"", 84, 60, 1432, 844);
		 Sleep(500)
		 trace("try to minimize outlook-window");
		 WinSetState($hwnd,"", @SW_MINIMIZE);
		 ;WinSetState($hwnd,"", @SW_HIDE);
	  EndIf
	  Sleep(1000);
	  clearTooTip();
   WEnd
EndFunc


;Local $aPos = WinGetPos($hWnd);
;WinMove($hwnd,"", 0, 0);
;Local $dx=10, $dy=10;
;WinMove($hwnd,"", $aPos[0]+$dx, $aPos[1]+$dy);
