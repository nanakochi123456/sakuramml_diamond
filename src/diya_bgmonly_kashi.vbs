' そのまま実行できます。
'
main

Sub main()
	Dim lp
	Dim Bgm
	Dim title
	Dim vtitle
	Dim kasi

	title="「Let's Mining」歌詞ジェネレータ (c)2014 NEET株式会社"
	vtitle="ダイヤを掘ろう！「Let's Mining」VOCALOID用歌詞" & vbcrlf & "Copyright 2014 NEET株式会社" & vbcrlf

	Bgm=MsgBox("BGMモードですか？(はい) / 歌付ですか？(いいえ)",vbYesNo,title)
	If Bgm=vbYes Then
		lp=InputBox("ループ数を数字で入力して下さい",title)
	Else
		lp=2
	End If

	On Error Resume Next
	If lp=0 And Bgm=vbYes Then
		MsgBox "値が0なの", vbOKOnly + vbCritical,title
		Exit Sub
	End If

	If lp>99 Then
		MsgBox "値が" & lp & "と巨大すぎて、あたしには処理できないの。ぐすぐす。。。", vbOKOnly + vbCritical,title
		Exit Sub
	End If

	On Error Goto 0

	kasi=makekashi(Bgm,lp)

	Dim filename
	filename="diamond_vocaloid_kashi.txt"
	Set objFso = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFso.OpenTextFile(filename, 2, True)

	If Err.Number > 0 Then
		MsgBox filename & "に書き込みできなかったなの；；ぐすぐす；；", vbOKOnly + vbCritical ,title
	Else
		objFile.Write vtitle
		objFile.Write kasi
		objFile.Close
		Set objFile = Nothing
		Set objFso = Nothing
		MsgBox filename & "に書き込みできたなの", vbOKOnly + vbInformation,title
	End If
End Sub

Function skashi(str,count)
	Dim kashi
	For i=1 To count
		kashi=kashi+str
	Next
	skashi=kashi
End Function

Function kashititle(track)
	Dim str
	str="-----------------------------------------------" & vbcrlf
	str=str+"Track " & track & vbcrlf & vbcrlf
	kashititle=str
End Function

Function makekashi(bgm,lp)
	Dim kashi
	Dim mkashi


	If bgm=vbNo Then
		mkashi=mkashi+"たてわ　すこぷ よこわ"
		mkashi=mkashi+"つるはし ほりすすめ"
		mkashi=mkashi+"てらせた　いまつ　だい"
		mkashi=mkashi+"やのため ほりすすめ"
		mkashi=mkashi+"かさねて すこぷ"
		mkashi=mkashi+"ひろえた いま  つ"
		mkashi=mkashi+"ほりすすめ"
		mkashi=mkashi+"たてに　すこぷ よこに"
		mkashi=mkashi+"つるはし ほりすすめ"
		mkashi=mkashi+"ひろえた いまつ だい"
		mkashi=mkashi+"やのため ほりすすめ"
		mkashi=mkashi+"コンボわ きまる"
		mkashi=mkashi+"ロマンわ ひかる"
		mkashi=mkashi+"パズルみたく さいくつをしよお"
		mkashi=mkashi+"ジャクわ　どこでも"
		mkashi=mkashi+"クイわ　ななめ"
		mkashi=mkashi+"キングよ つらぬけ"
		mkashi=mkashi+"ジョカわ つよい"
		mkashi=mkashi+"すこおぷ　つかえ"
		mkashi=mkashi+"つるはし　つかえ"
		mkashi=mkashi+"がんばん つらぬき"
		mkashi=mkashi+"ダイヤをさがせ"
		mkashi=mkashi+"きいのすこぷ　てつの"
		mkashi=mkashi+"すこおぷ ほりすすめ"
		mkashi=mkashi+"きんのすこぷ　だいや"
		mkashi=mkashi+"のすこぷ　ほりすすめ"
		mkashi=mkashi+"かさねて つるはし"
		mkashi=mkashi+"もえよたいまつ"
		mkashi=mkashi+"ほりすすめ"
		mkashi=mkashi+"きいのつ るはし　てつの"
		mkashi=mkashi+"つるはし ほりすすめ"
		mkashi=mkashi+"きんのつ　るはし　だいや"
		mkashi=mkashi+"のつるはし ほりすすめ"
		mkashi=mkashi+"コンボを きめて"
		mkashi=mkashi+"ロマンを ひからせ"
		mkashi=mkashi+"パズルみたく さいくつをしよお"
		mkashi=mkashi+"エスで　じつわ"
		mkashi=mkashi+"キングを　ほれる"
		mkashi=mkashi+"コンボを きめて"
		mkashi=mkashi+"ぎゃくてんねらえ"
		mkashi=mkashi+"スコップ　つかえ"
		mkashi=mkashi+"つるはし　つかえ"
		mkashi=mkashi+"がんばん つらぬき"
		mkashi=mkashi+"ダイヤをさがせ"

		kashi=kashi+kashititle(1)
		kashi=kashi+"ほろお"
		kashi=kashi+mkashi
		kashi=kashi+vbcrlf

		kashi=kashi+kashititle(2)
		kashi=kashi+"だいや"
		kashi=kashi+mkashi
		kashi=kashi+vbcrlf

		kashi=kashi+kashititle("3/4")
		kashi=kashi+"だいや"
		kashi=kashi+vbcrlf

	End If
	'Track 7/8/9
	kashi=kashi+kashititle("5/6/7")

	Dim Hatsu

	If bgm=vbNo Then
		Hatsu="ちゃ"
	Else
		Hatsu="る"
	End If

	'Intro
	kashi=kashi+skashi("ちゃ",3*4)
	kashi=kashi+skashi("た",22)

	For i=1 To lp
		If i<>1 Then
			'sub
			kashi=kashi+skashi(Hatsu,6)
			kashi=kashi+skashi("た",10)
		End If

		'melody1
		kashi=kashi+skashi(Hatsu,7*4)
		kashi=kashi+skashi("た",12)
		kashi=kashi+skashi(Hatsu,2)
		kashi=kashi+skashi("た",4)

		'melody2
		kashi=kashi+skashi(Hatsu,4)
		kashi=kashi+skashi("た",8)
		kashi=kashi+skashi(Hatsu,8)
		kashi=kashi+skashi("た",4)
	Next

	'end
	kashi=kashi+skashi("ちゃ",14)
	kashi=kashi+skashi("ら",1)

	kashi=kashi+vbcrlf+vbcrlf

	'Track 8
	kashi=kashi+kashititle(8)
	kashi=kashi+skashi("ら",1)
	kashi=kashi+vbcrlf+vbcrlf

	'Track 9
	kashi=kashi+kashititle(9)

	'Intro
	kashi=kashi+skashi("とぅ",32)

	For i=1 To lp
		If i<>1 Then
			'sub
			kashi=kashi+skashi("とぅ",16)
		End If

		'melody1
		kashi=kashi+skashi("とぅ",15*4+12)

		'melody2
		kashi=kashi+skashi("とぅ",5+4+4+4+5+5+4+4)
	Next

	'end
	kashi=kashi+skashi("とぅ",6*4+2+8)
	kashi=kashi+skashi("ら",1)
	kashi=kashi+vbcrlf+vbcrlf

	makekashi=kashi
End Function


' たわごと・・・
' SAKURAMMLでも、この機能できそうなんですが、
' デバッグウィンドウではなく、ちゃんとテキストエディタで
' 開いた方が楽なので・・・
