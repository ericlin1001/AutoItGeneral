#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here


#include <Debug.au3>

#include <Array.au3>




If 0 Then
Global $qq1="3503525770"
Global $pass1="donny123"
Else
Global $qq1="2042485997"
Global $pass1="1q2w{#}E$R"
EndIf

Global $qq2="3526489809"
Global $pass2="1q2w{#}E$R"

Global $path1="D:\Chess\trunk\chess_client_proj\trunk\unity\Platform\win\winPack"
Global $path2="D:\Chess\trunk\chess_client_proj\trunk\unity\Platform\win\winPack2"
Global $exe="chess.exe"
Global $h1=0
Global $h2=0;
Global $delay=100;
Global $runDelay=5300
Global $isShowLog=false

Global $map1[8][8][2];
Global $map2[8][8][2];


Global $TL[2];  54, 273
Global $BR[2]; 472, 673
$TL[0]=54
$TL[1]=273
$BR[0]=472
$BR[1]=673


For $i=0 to 7
   for $j=0 to 7
	  $map1[$i][$j][0]=$TL[0]+$j*($BR[0]-$TL[0])/7;
	  $map1[$i][$j][1]=$TL[1]+$i*($BR[1]-$TL[1])/7;

	  $map2[$i][$j][0]=528+$map1[$i][$j][0];
	  $map2[$i][$j][1]= $map1[$i][$j][1];
   Next
Next

Func tryClick()
mClickPieace(0, 0);
mClickPieace(7, 0);
EndFunc


Func mclick($x, $y)
   MouseClick("left", $x, $y);
   Sleep($delay);
EndFunc

Func mClickPieace($r, $c)
mclick($map1[$r][$c][0], $map1[$r][$c][1] );
EndFunc


Func init()
if $isShowLog Then
_DebugSetup("Check exe")
EndIf


HotKeySet("{F9}", "terminate");
EndFunc

Func terminate()
   Exit
EndFunc



Func run2()

Run($path1&"\"& $exe);
;Sleep(2000);

Run($path2&"\" &$exe);
WinWait("腾讯国际象棋");
Sleep($runDelay);
EndFunc


Func trace($mess)
   if $isShowLog Then
   _DebugOut("debug> "&$mess);
   EndIf
EndFunc

Func sendKey($key)
Send($key);
Sleep($delay);
EndFunc


;run2();

init();
ensue2Win()
WinMove($h1, "", 0, 0, 528, 946);
WinMove($h2, "", 528, 0, 528, 946);
WinActivate($h1);
WinActivate($h2);

; click play with QQ
mclick(266, 710);
mclick(783, 712);

; click QQ1
WinActivate($h1);
mclick(400, 465)
mclick(400, 465)
sendKey("{END}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}")
sendKey ($qq1);
mclick(400, 515);
mclick(400, 515);
sendKey("{END}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}")
sendKey ($pass1);
mclick(261, 582);

; click QQ2
WinActivate($h2);
mclick(930, 460);qq
mclick(930, 460);qq
sendKey("{END}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}")
sendKey($qq2);
mclick(930, 517);pass
mclick(930, 517);pass
sendKey("{END}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}{BS}")
sendKey($pass2);
mclick(785, 590);//enter
Sleep(500);
; after login

mclick(362, 468);h1
mclick(900, 490);/h2


mclick(324, 317);
mclick(811, 330)

Sleep(15000);
tryClick();
exit;





Func ensue2Win()
    ; Retrieve a list of window handles using a regular expression. The regular expression looks for titles that contain the word SciTE or Internet Explorer.
    Local $aWinList = WinList("腾讯国际象棋")

   if  not($aWinList[0][0] = 2) Then
	   trace("try to close all, $aWinList[0][0]="&$aWinList[0][0]);
		for $i=1 to $aWinList[0][0]
			trace("ok , winclose:"&WinClose($aWinList[$i][1]));
		 Next
		 run2();
	  Else
		  trace("there are eactaly two win, $aWinList[0][0]="&$aWinList[0][0]);
   EndIf
   $aWinList = WinList("腾讯国际象棋");
   if  not($aWinList[0][0] = 2) Then
	  trace("fails ....");
	  trace("$aWinList[0][0]="&$aWinList[0][0]);
   Else
	  trace("try to get h1, h2, $aWinList[0][0]="&$aWinList[0][0]);
	    trace("h1="&$h1);
	  trace("h2="&$h2);
	  $h2=$aWinList[1][1];
	  $h1=$aWinList[2][1];
   EndIf
EndFunc   ;==>Example


Func getState()
EndFunc



