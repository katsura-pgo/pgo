Option Explicit
On Error Resume Next
'#┌──────────────────────────────────────
'#│  自動サーチ v0.0.2 (2017/03/24)
'#│  pgo.vbs 
'#└──────────────────────────────────────
'#
'# [ 更新履歴 ]
'# 2017/03/24 -> v0.0.2
'#  IEがフリーズ時の処理追加
'# 
'# 2017/03/14 -> v0.0.1
'#  書式の完成
'# 
'# [ 使い方 ]
'#  このファイルを同階層にlist.csvというテキストファイルを置いてください。
'#  list.csvにはサーチしたい場所をカンマ区切りで緯度,経度を書きます。
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
'#
'#==============================================================================
'# 設定値をここに記載する。
Dim CMN
Set CMN = CreateObject( "Scripting.Dictionary" )

'#-- [ 基本設定 ] --------------------------------------------------------------

Call CMN.Add("URL",  "https://pmap.kuku.lu/#") ' P-GO SEARCH URL
Call CMN.Add("BTN",  "area_buttonsearch") ' サーチボタンID
Call CMN.Add("WAIT",120*1000) ' サーチ後待機秒
Call CMN.Add("IE",   True) ' IEを表示するか、表示：True, 非表示：False
Call CMN.Add("LIST", "list.csv") ' サーチする座標が書かれたテキストファイル

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

	Dim latArray, lngArray
	Set latArray = CreateObject("System.Collections.ArrayList")
	Set lngArray = CreateObject("System.Collections.ArrayList")

	Set objStream = fso.OpenTextFile(CMN("LIST"), 1)
	Dim i
	i = 0
	Do Until objStream.AtEndOfStream
		strLine = objStream.ReadLine
		arrFields = Split(strLine,",")
		latArray.add arrFields(0)
		lngArray.add arrFields(1)
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

	Do
		Set ie = CreateObject("InternetExplorer.Application")
		ie.Visible = CMN("IE")
		ie.Navigate CMN("URL") & latArray(i) & "," & lngArray(i)

		b = 0
		Do While ie.Busy = True Or ie.ReadyState <> 4 Or b < 50
			WScript.Sleep 100
			b = b + 1
		Loop

		Set elm = ie.document.getElementById(CMN("BTN"))
		elm.Click
		WScript.Echo "サーチ：" & CStr(n) & "回目(" & Time & ") " & CStr(i+1) & "行目(" & latArray(i) & "," & lngArray(i) & ") "
		WScript.Sleep CMN("WAIT")
		ie.Quit

		If Err.Number <> 0 Then
			WScript.Echo "エラー番号 " & CStr(Err.Number) & " " & Err.Description
			Err.Clear
			WScript.Sleep 5000
		Else
			i = i + 1
			n = n + 1
			If i = r Then: i = 0: End If
		End If
	Loop

	Main = True
'#------------------------------------------------------------------------------
End Function

