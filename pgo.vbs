Option Explicit
On Error Resume Next
'#┌──────────────────────────────────────
'#│  自動サーチ v0.0. (2017/06/06)
'#│  pgo.vbs
'#└──────────────────────────────────────
'#
'# [ 更新履歴 ]
'# 2017/06/06 -> v0.0.9
'#  ピゴサのボタンのIDが変わったので変更
'# 
'# 2017/05/16 -> v0.0.8
'#  ピゴサのボタンのIDが変わったので変更
'# 
'# 2017/05/14 -> v0.0.7
'#  ピゴサのボタンのIDが変わったので変更
'# 
'# 2017/05/13 -> v0.0.6
'#  サーチ開始時刻よりサーチ終了時刻の方が早い場合の処理を追加
'# 
'# 2017/05/09 -> v0.0.5
'#  サーチ開始時刻、サーチ終了時刻、サーチ実行曜日を追加
'# 
'# 2017/04/18 -> v0.0.4
'#  csvファイルの３列目に名前を追加
'# 
'# 2017/03/31 -> v0.0.3
'#  IEがフリーズ時の処理追加
'# 
'# 2017/03/24 -> v0.0.2
'#  IEがフリーズ時の処理追加
'# 
'# 2017/03/14 -> v0.0.1
'#  書式の完成
'# 
'# [ 使い方 ]
'#  このファイルを同階層にlist.csvというテキストファイルを置いてください。
'#  list.csvにはサーチしたい場所をカンマ区切りで緯度,経度,名前を書きます。
'#  一行一箇所になりますので、二箇所目は二行目に書きます。
'#  list.csvができたら、このファイルをダブルクリックして実行してください。
'#  ＩＥは非表示状態で裏で動いていますので、別のウィンドウでピゴサの状況を確認してください。
'# 
'# [ サーチ制限 ]
'#  ピゴサで現在（2017/03/24）同じIPアドレスで一日でサーチできる回数は840回程度らしいです。
'#  回数制限に達するとサーチがしばらくできなくなります。
'#  840回を24時間で割ると2分毎にサーチする計算になります。
'#  10分毎に特定の座標をサーチしたい場合は5ヶ所をlist.csvに記載してください。
'#  30分毎に特定の座標をサーチしたい場合は15ヶ所をlist.csvに記載してください。
'#  なおピゴサに聖地・巣情報として登録されている場所ではふしぎなちからによって
'#  24時間自動サーチがされていますので、その場所をサーチする必要はありません。
'# 
'# [ IEの確認 ]
'#  普段IEを使っていない方は、一度IEを手動で起動して、初期メッセージなどが表示されないか確認をしてください。
'#  また一度、ピゴサにアクセスして、ポケモンやポケストップの表示を全部非表示にします。
'#  自動サーチではサーチボタンが押せればいいだけなので、極力処理を早くするためです。
'# 
'# [ 時間の指定 ]
'#  時間の指定がしたい場合は基本設定を変更してください。
'#
'# [ 複雑な指定 ]
'#  pgo.vbsをコピーしてリネームすれば、時間指定と組み合わせることで、複雑な指定が可能です。
'#  例えば、時間帯でサーチする場所を変更したり、休日のサーチする場所を変更したりできます。
'#  コピーしたpgo.vbsをすべて起動してください。指定時間ではないものは時間になるまで待機モードになります。
'#
'#==============================================================================
'# 設定値をここに記載する。
Dim Conf
Set Conf = CreateObject( "Scripting.Dictionary" )

'#-- [ 基本設定 ] --------------------------------------------------------------

Conf.Add "URL",  "https://pmap.kuku.lu/#" ' P-GO SEARCH URL
Conf.Add "BTN",  "area_buttonresearch" ' サーチボタンID
Conf.Add "WAIT", 120*1000 ' サーチ後待機秒
Conf.Add "READ", 5*1000 ' 読込待機秒
Conf.Add "IE",   True ' IEを表示するか、表示：True, 非表示：False
Conf.Add "LIST", "list.csv" ' サーチする座標が書かれたテキストファイル

Conf.Add "START", "0:00" ' サーチ開始時刻
Conf.Add "END",   "24:00" ' サーチ終了時刻

Conf.Add "WD1", True ' 日曜日に実行するか、実行：True, 非実行：False
Conf.Add "WD2", True ' 月曜日に実行するか、実行：True, 非実行：False
Conf.Add "WD3", True ' 火曜日に実行するか、実行：True, 非実行：False
Conf.Add "WD4", True ' 水曜日に実行するか、実行：True, 非実行：False
Conf.Add "WD5", True ' 木曜日に実行するか、実行：True, 非実行：False
Conf.Add "WD6", True ' 金曜日に実行するか、実行：True, 非実行：False
Conf.Add "WD7", True ' 土曜日に実行するか、実行：True, 非実行：False

'#==============================================================================
'# グローバル変数宣言, 設定値結合, ライブラリ読込

'#------------------------------------------------------------------------------
Main()

'#=============================================================================#
'#                               [ メイン関数 ]                                #
'#=============================================================================#
Function Main()
	Dim wsh, fso
	Set wsh = WScript.CreateObject("WScript.Shell")
	Set fso = CreateObject("Scripting.FileSystemObject")
	'---------------------------------------------------------------------------
	' CSCRIPT.EXE で無い場合に切り替える

	Dim strPath, strTarget, strMyPath, strParam, j
	strPath   = WScript.FullName ' フルパス
	strTarget = Right( strPath, 11 )
	strTarget = Ucase( strTarget )

	if strTarget <> "CSCRIPT.EXE" then

		' 自分自身ののフルパス
		strMyPath = WScript.ScriptFullName
		strParam  = " "

		' 引数を全て読みだして、スペースが含まれていた場合は再度引き渡す為に " で囲み直しています
		For j = 0 to Wscript.Arguments.Count - 1
			If instr(Wscript.Arguments(j), " ") < 1 Then
				strParam = strParam & Wscript.Arguments(j) & " "
			Else
				strParam = strParam & """" & Wscript.Arguments(j) & """ "
			End If
		Next

		' CSCRIPT.EXE に引き継いで終了
		Call wsh.Run( "cscript.exe //nologo """ & strMyPath & """" & strParam )
		WScript.Quit

	end if
	'---------------------------------------------------------------------------

	' CSV読み込み
	Dim	objStream
	Dim	strLine
	Dim	arrFields

	Dim latArray, lngArray, nameArray
	Set latArray = CreateObject("System.Collections.ArrayList")
	Set lngArray = CreateObject("System.Collections.ArrayList")
	Set nameArray = CreateObject("System.Collections.ArrayList")

	Set objStream = fso.OpenTextFile(Conf("LIST"), 1)
	Dim i
	i = 0
	Do Until objStream.AtEndOfStream
		strLine = objStream.ReadLine
		arrFields = Split(strLine,",")
		latArray.add arrFields(0)
		lngArray.add arrFields(1)
		nameArray.add arrFields(2)
		i = i + 1
	Loop
	Dim r: r = latArray.Count

	objStream.Close
	Set	objStream = Nothing
	Set	fso = Nothing

	'---------------------------------------------------------------------------
	WScript.Echo "サーチ開始。このウィンドウを閉じれば終了します。"
	Dim ie, elm, b, n
	i = 0
	n = 1

	Dim staSec, endSec, nowSec
	staSec = DaySecond(Conf("START"))
	endSec = DaySecond(Conf("END"))

	Dim offFlg: offFlg = False
	Dim re, m
	Set re = new regexp
	re.Pattern = Conf("BTN") & "_[a-zA-Z0-9]*"

	Do
		nowSec = DaySecond(CStr(Hour(Now))&":"&CStr(Minute(Now)))
		If (_
			(staSec < endSec And staSec =< nowSec And endSec > nowSec) OR _
			(staSec > endSec And (staSec =< nowSec Or endSec > nowSec))_
		) And Conf("WD"&CStr(Weekday(Now))) Then
			offFlg = False

			Set ie = CreateObject("InternetExplorer.Application")
			ie.Visible = Conf("IE")
			ie.Navigate Conf("URL") & latArray(i) & "," & lngArray(i)

			WScript.Sleep Conf("READ")

			Set m = re.Execute(ie.Document.Body.InnerHtml)
			Set elm = ie.document.getElementById(m(0).Value)
			elm.Click
			WScript.Echo "サーチ：" & CStr(n) & "回目(" & Time & ") " & nameArray(i) & "(" & CStr(i+1) & "行目) "
			WScript.Sleep Conf("WAIT") - Conf("READ")
			ie.Quit
			Set ie = Nothing

			If Err.Number <> 0 Then
				WScript.Echo "エラー番号 " & CStr(Err.Number) & " " & Err.Description
				Err.Clear
				WScript.Sleep 5000
			Else
				i = i + 1
				n = n + 1
				If i = r Then: i = 0: End If
			End If
		Else
			If Not offFlg Then
				WScript.Echo "サーチ実行時刻になるまで待機中です。"
				offFlg = True
			End If
		End If

	Loop

	Main = True
'#------------------------------------------------------------------------------
End Function

' 秒に変換
Function DaySecond(c)
	Dim	cFields, h, m, s
	cFields = Split(c, ":")
	h = CInt(cFields(0))
	m = CInt(cFields(1))
	s = 0

	Dim t:t = 0
	t = t + (h*3600)
	t = t + (m*60)
	t = t + s
	DaySecond = t
'#------------------------------------------------------------------------------
End Function