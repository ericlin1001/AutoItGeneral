#cs
#
Author: junericlin
Data: 2017.5.23
# What does it do?
Every 1 hour and 45 mintues, this script, visit url:http://auth-proxy.oa.com/DevNetTempVisit.aspx, and click the apply button(申请访问)。　
简单的说，　这个脚本，每２小时左右，　自动申请外网访问
#ce
#include <Array.au3>
#include <MsgBoxConstants.au3>
#include <Debug.au3>
Global $isDebugOut=true;
Global $winText="IT服务-申请临时访问 - Windows Internet Explorer";
Global Const $MAX_NUM_EXCEPT_WEEK_DAYS=7
Global Const $MAX_NUM_EXCEPT_TIME=100;
Global $iNumExceptWeekDays=0;
Global $aExceptWeekDays[$MAX_NUM_EXCEPT_WEEK_DAYS];
Global $iNumExceptTimes=0;
Global $aExceptTimes[$MAX_NUM_EXCEPT_TIME][2][2];//for each [i]:time{[0]:startTiem{[0]:hour, [1]:minutue}, [1]:endTime{[0]:hour, [1]:minutue}}

Global $isShowTrace=False;

Global $checkEveryTime=(60+55)*60*1000 ;in milliseconds. 1h45min=(60+45)min=(60+45)*60sec=(60+45)*60*1000millisec
;Global $checkEveryTime=2*1000 ;in milliseconds. 1h45min=(60+45)min=(60+45)*60sec=(60+45)*60*1000millisec
Global $isRealDo=true;
init();

main();
clearTooTip()

Exit

Func parseTime($srcStr)
   trace("parseTime(srcStr="&$srcStr&")");
   $srcStr=StringStripWS($srcStr,$STR_STRIPLEADING +$STR_STRIPTRAILING )
   Local $aSubStr=StringSplit($srcStr,":");
   If @error =1 Then
	  SetError(1);
   ElseIf not $aSubStr[0]=2 Then ; must contain only a hour, and mintues.
	  SetError(1);
   Else
	  Local $hourTime=Number($aSubStr[1]);
	  Local $mintueTime= Number($aSubStr[2]);
	  Local $aTmp[2];
	  $aTmp[0]=Number($hourTime);
	  $aTmp[1]=Number($mintueTime);
	  trace("result:"&$hourTime&":"&$mintueTime);
	  return $aTmp;
   EndIf
EndFunc

Func parseStartEndTime($srcStr)
   trace("parseStartEndTime(srcStr="&$srcStr&")");
   Local $aSubStr=StringSplit($srcStr,"~");
   If @error =1 Then

	  SetError(1);
   ElseIf Not $aSubStr[0]=2 Then ; must contain only a start time,and end time.
	  SetError(1);
   Else
	  Local $startTime=parseTime($aSubStr[1]);
	  If @error Then
		 SetError(1);
	  EndIf

	  Local $endTime= parseTime($aSubStr[2]);
	  If @error Then
		 SetError(1);
	  EndIf
	  Local $aTmp[2][2];
	  $aTmp[0][0]=$startTime[0];
	  $aTmp[0][1]=$startTime[1];
	  $aTmp[1][0]=$endTime[0];
	  $aTmp[1][1]=$endTime[1];

   trace("$startTime:"&$startTime[0]&":"&$startTime[1]);
   trace("$endTime:"&$endTime[0]&":"&$endTime[1]);

	  Return $aTmp
   EndIf
EndFunc


Func readIni()
   Local $iniFile="tempVisitSetting.ini";
   Local $aGeneralSection=IniReadSection($iniFile,"General");

   If Not @error Then
	  For $i = 1 To $aGeneralSection[0][0]; for each key=value.
		 Local $key=$aGeneralSection[$i][0];
		 Local $value=$aGeneralSection[$i][1];
		 Local $aSubStr;

		 If $key= "exceptWeekDays" Then
			$aSubStr=StringSplit($value, ",");
			If @error=1 Then
			   trace("value:"&$value&" has wrong format!!");
			   ContinueLoop;
			Else
			   For $j=1 to $aSubStr[0]
				  ;push into $aExceptWeekDays
				  Local $token=StringStripWS($aSubStr[$j],$STR_STRIPLEADING +$STR_STRIPTRAILING );
				  If not $token="" Then
				  $aExceptWeekDays[$iNumExceptWeekDays]=Number($token)
				  $iNumExceptWeekDays=$iNumExceptWeekDays+1;
				  EndIf
				  Next
			EndIf
		 ElseIf $key="exceptTimes" Then
			trace("parsing exceptTimes");
			$aSubStr=StringSplit($value, ",");
			If @error=1 Then
			   trace("value:"&$value&" has wrong format!!");
			   ContinueLoop;
			Else
			   trace("$aSubStr[0]"&$aSubStr[0]);
			   For $j=1 to $aSubStr[0]
				  Local $token=StringStripWS($aSubStr[$j],$STR_STRIPLEADING +$STR_STRIPTRAILING );

				  If not $token="" Then
					 Local $tmp[2][2]=parseStartEndTime($token);
					 If @error=0 Then
						;push into $aExceptTimes

						For $a=0 to 1
						   for $b=0 to 1
							  ;$aExceptTimes[$iNumExceptTimes]=$tmp;
							  $aExceptTimes[$iNumExceptTimes][$a][$b]=$tmp[$a][$b];
						   Next
						Next
						$iNumExceptTimes=$iNumExceptTimes+1;
					 Else
						trace("can't parse value:"&$token);
					 EndIf
				  Else
					 trace("skip empty token:"&$token);
				  EndIf
			   Next
			EndIf
		 Else
			trace("Error: Can't parse this key:"&$key);
		 EndIf
		 trace("key:"&$key&" value:"&$value);
	  Next
	  #cs
	  trace("after reading ini file....");
	  trace("$iNumExceptWeekDays:"&$iNumExceptWeekDays);
	  trace("$iNumExceptTimes:"&$iNumExceptTimes);


	  _ArrayDisplay($aExceptWeekDays);
	  For $i =0 to $iNumExceptTimes-1
		 trace( "$iNumExceptTimes["&$i&"]:");
		 For $a=0 to 1
			   for $b=0 to 1
				  trace($aExceptTimes[$i][$a][$b]);
			   Next
			Next
			;Local $abc[2][2]=$aExceptTimes[$i];

	  Next
	  _ArrayDisplay($aExceptTimes, "$iNumExceptTimes["&$i&"]", Default, 1);
	  #ce
   Else
	  trace("can't read ini file:"&$iniFile);
   EndIf

EndFunc


Func checkIfNeedToDo()

   For $i=0 to $iNumExceptWeekDays-1
	  if ($aExceptWeekDays[$i]+1)=@WDAY Or ($aExceptWeekDays[$i]+1-7)=@WDAY Then
		 return False
	  EndIf

   Next
   For $i=0 to $iNumExceptTimes-1
	  Local $curH=@HOUR;
	  Local $curM=@MIN;
	  Local $curTime=$curH*60+$curM;
	  Local $sH=$aExceptTimes[$i][0][0];
	  Local $sM=$aExceptTimes[$i][0][1];
	  Local $sTime=$sH*60+$sM;
	  Local $eH=$aExceptTimes[$i][1][0];
	  Local $eM=$aExceptTimes[$i][1][1];
	  Local $eTime=$eH*60+$eM;
	  If $sTime<=$curTime And $curTime<=$eTime then
		 return False
	  EndIf

   Next
   Return True;
EndFunc




Func init()
If $isDebugOut Then
_DebugSetup("AutoApplyForTempVisit.exe")
readIni();
EndIf


EndFunc
Func clearTooTip()
   ToolTip("");
EndFunc


Func trace($msg)
   Local $exeName="keepOutlookAlwaysOn.exe";
   Local $tag=$exeName&"::Debug> ";
   if $isShowTrace Then
   If $isDebugOut Then
	  _DebugOut($tag&$msg);
   Else
	  ToolTip($tag&$msg)
   EndIf
   EndIf

EndFunc

Func mclick($x, $y)
    MouseClick($MOUSE_CLICK_LEFT, $x, $y);
	Sleep(20);
 EndFunc



Func main()

   Local $lastDo=TimerInit();
   Local $isFirstTime=true;
   While 1
	  Local $timeDiff=TimerDiff($lastDo);
	  If $timeDiff > $checkEveryTime Or $isFirstTime Then

		 If checkIfNeedToDo() Then

			   If $isRealDo Then
				  doApplyForTempVisit();

			   Else
				  trace("simulate(not real do) doApplyForTempVisit()");
			   EndIf
			   $lastDo=TimerInit()

		 Else
			trace("Do NOT need to doApplyForTempVisit(), because time is in exceptTime");
		 EndIf
	  Else
			trace("Dont' need to do, becase, time="&($timeDiff/1000/60)&"mintues < 2h");
	  EndIf


	  clearTooTip()
	  Sleep(1000*60);//check every mintues.
	  ;Sleep(1000*1);//check every mintues.
	  $isFirstTime=False;
   WEnd

EndFunc



Func doApplyForTempVisit()
   Local $hwnd=0;


   Local $tryCount=0;
   Local $maxTryCount=5;

   ;try to get winhandle
   $hwnd=WinGetHandle("IT服务-申请临时访问 - Windows Internet Explorer", $winText);
   While $hwnd=0
	  If $tryCount> $maxTryCount Then
		 Break
	  EndIf
	  trace("try to run ie to open ...");
	  Run("C:\Program Files (x86)\Internet Explorer\iexplore.exe http://auth-proxy.oa.com/DevNetTempVisit.aspx");
	  ;Sleep(1000);
	  WinWait("IT服务-申请临时访问 - Windows Internet Explorer", $winText, 2);   //wait 2 seconds.
	  $hwnd=WinGetHandle("IT服务-申请临时访问 - Windows Internet Explorer", $winText);
	  $tryCount=$tryCount+1;
   WEnd

   ;after getting handle, do the
   If $hwnd=0 Then
	  trace("Can't find the window...");
   Else
	   trace("hwnd:"&$hwnd);
	  ;WinActivate($hwnd, $winText); Don't need to activate the windows.
	  Local $ret=WinMove($hwnd, $winText, 0, 0, 844, 749);
	  trace("$ret:"&$ret);
	  Sleep(500);
	  ControlSend("[CLASS:IEFrame]","","[CLASS:Internet Explorer_Server; INSTANCE:1]","{F5}"); refresh the web page.
	  Sleep(500);
	  ControlClick("[CLASS:IEFrame]","","[CLASS:Internet Explorer_Server; INSTANCE:1]","left",1,408, 331); click the button, try 1
	  Sleep(200);
	  ControlClick("[CLASS:IEFrame]","","[CLASS:Internet Explorer_Server; INSTANCE:1]","left",1,408, 331); click the button, try 2
	  Sleep(300);
	  ControlClick("[CLASS:IEFrame]","","[CLASS:Internet Explorer_Server; INSTANCE:1]","left",1,408, 331); click the button, try3, to ensure it will be clicked.
   EndIf

   Sleep(2000);
   tryCloseAllWin();
   clearTooTip()
EndFunc
Func tryCloseAllWin()
   Local $maxTryCloseCount=10;
   While $maxTryCloseCount>0
	  $maxTryCloseCount=$maxTryCloseCount-1;
	  WinGetHandle($winText,$winText)
	  if @error Then
		 $maxTryCloseCount=0
	  EndIf
	  WinClose($winText, $winText);
   WEnd
EndFunc

