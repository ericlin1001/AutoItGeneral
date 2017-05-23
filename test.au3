;WinMove("腾讯国际象棋","", 100,0);
;MouseClick("left", 171, 772)

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
   Local $aMove=[1,0, 3, 0, 1,1, 3, 1, 1,2, 3, 2];
   For $i=0 to 11 step 4
   mClickPieace1($aMove[$i], $aMove[$i+1]);
   mClickPieace1($aMove[$i+2], $aMove[$i+3]);
   mClickPieace2($aMove[$i], $aMove[$i+1]);
   mClickPieace2($aMove[$i+2], $aMove[$i+3]);
   Next

EndFunc


Func mclick($x, $y)
   MouseClick("left", $x, $y);
;   Sleep($delay);
EndFunc

Func mClickPieace1($r, $c)
   $r=7-$r;
   Local $d=200;
mclick($map1[$r][$c][0], $map1[$r][$c][1] );
Sleep($d);
mclick($map1[$r][$c][0], $map1[$r][$c][1] );
Sleep($d);
EndFunc

Func mClickPieace2($r, $c)
   $r=7-$r;
   Local $d=300;
   mclick($map2[$r][$c][0], $map2[$r][$c][1] );
   Sleep($d);
   mclick($map2[$r][$c][0], $map2[$r][$c][1] );
   Sleep($d);
EndFunc

tryClick();