Option Explicit
On Error Resume Next
'#������������������������������������������������������������������������������
'#��  �����T�[�` v0.0. (2017/06/06)
'#��  pgo.vbs
'#������������������������������������������������������������������������������
'#
'# [ �X�V���� ]
'# 2017/06/06 -> v0.0.9
'#  �s�S�T�̃{�^����ID���ς�����̂ŕύX
'# 
'# 2017/05/16 -> v0.0.8
'#  �s�S�T�̃{�^����ID���ς�����̂ŕύX
'# 
'# 2017/05/14 -> v0.0.7
'#  �s�S�T�̃{�^����ID���ς�����̂ŕύX
'# 
'# 2017/05/13 -> v0.0.6
'#  �T�[�`�J�n�������T�[�`�I�������̕��������ꍇ�̏�����ǉ�
'# 
'# 2017/05/09 -> v0.0.5
'#  �T�[�`�J�n�����A�T�[�`�I�������A�T�[�`���s�j����ǉ�
'# 
'# 2017/04/18 -> v0.0.4
'#  csv�t�@�C���̂R��ڂɖ��O��ǉ�
'# 
'# 2017/03/31 -> v0.0.3
'#  IE���t���[�Y���̏����ǉ�
'# 
'# 2017/03/24 -> v0.0.2
'#  IE���t���[�Y���̏����ǉ�
'# 
'# 2017/03/14 -> v0.0.1
'#  �����̊���
'# 
'# [ �g���� ]
'#  ���̃t�@�C���𓯊K�w��list.csv�Ƃ����e�L�X�g�t�@�C����u���Ă��������B
'#  list.csv�ɂ̓T�[�`�������ꏊ���J���}��؂�ňܓx,�o�x,���O�������܂��B
'#  ��s��ӏ��ɂȂ�܂��̂ŁA��ӏ��ڂ͓�s�ڂɏ����܂��B
'#  list.csv���ł�����A���̃t�@�C�����_�u���N���b�N���Ď��s���Ă��������B
'#  �h�d�͔�\����Ԃŗ��œ����Ă��܂��̂ŁA�ʂ̃E�B���h�E�Ńs�S�T�̏󋵂��m�F���Ă��������B
'# 
'# [ �T�[�`���� ]
'#  �s�S�T�Ō��݁i2017/03/24�j����IP�A�h���X�ň���ŃT�[�`�ł���񐔂�840����x�炵���ł��B
'#  �񐔐����ɒB����ƃT�[�`�����΂炭�ł��Ȃ��Ȃ�܂��B
'#  840���24���ԂŊ����2�����ɃT�[�`����v�Z�ɂȂ�܂��B
'#  10�����ɓ���̍��W���T�[�`�������ꍇ��5������list.csv�ɋL�ڂ��Ă��������B
'#  30�����ɓ���̍��W���T�[�`�������ꍇ��15������list.csv�ɋL�ڂ��Ă��������B
'#  �Ȃ��s�S�T�ɐ��n�E�����Ƃ��ēo�^����Ă���ꏊ�ł͂ӂ����Ȃ�����ɂ����
'#  24���Ԏ����T�[�`������Ă��܂��̂ŁA���̏ꏊ���T�[�`����K�v�͂���܂���B
'# 
'# [ IE�̊m�F ]
'#  ���iIE���g���Ă��Ȃ����́A��xIE���蓮�ŋN�����āA�������b�Z�[�W�Ȃǂ��\������Ȃ����m�F�����Ă��������B
'#  �܂���x�A�s�S�T�ɃA�N�Z�X���āA�|�P������|�P�X�g�b�v�̕\����S����\���ɂ��܂��B
'#  �����T�[�`�ł̓T�[�`�{�^����������΂��������Ȃ̂ŁA�ɗ͏����𑁂����邽�߂ł��B
'# 
'# [ ���Ԃ̎w�� ]
'#  ���Ԃ̎w�肪�������ꍇ�͊�{�ݒ��ύX���Ă��������B
'#
'# [ ���G�Ȏw�� ]
'#  pgo.vbs���R�s�[���ă��l�[������΁A���Ԏw��Ƒg�ݍ��킹�邱�ƂŁA���G�Ȏw�肪�\�ł��B
'#  �Ⴆ�΁A���ԑтŃT�[�`����ꏊ��ύX������A�x���̃T�[�`����ꏊ��ύX������ł��܂��B
'#  �R�s�[����pgo.vbs�����ׂċN�����Ă��������B�w�莞�Ԃł͂Ȃ����͎̂��ԂɂȂ�܂őҋ@���[�h�ɂȂ�܂��B
'#
'#==============================================================================
'# �ݒ�l�������ɋL�ڂ���B
Dim Conf
Set Conf = CreateObject( "Scripting.Dictionary" )

'#-- [ ��{�ݒ� ] --------------------------------------------------------------

Conf.Add "URL",  "https://pmap.kuku.lu/#" ' P-GO SEARCH URL
Conf.Add "BTN",  "area_buttonresearch" ' �T�[�`�{�^��ID
Conf.Add "WAIT", 120*1000 ' �T�[�`��ҋ@�b
Conf.Add "READ", 5*1000 ' �Ǎ��ҋ@�b
Conf.Add "IE",   True ' IE��\�����邩�A�\���FTrue, ��\���FFalse
Conf.Add "LIST", "list.csv" ' �T�[�`������W�������ꂽ�e�L�X�g�t�@�C��

Conf.Add "START", "0:00" ' �T�[�`�J�n����
Conf.Add "END",   "24:00" ' �T�[�`�I������

Conf.Add "WD1", True ' ���j���Ɏ��s���邩�A���s�FTrue, ����s�FFalse
Conf.Add "WD2", True ' ���j���Ɏ��s���邩�A���s�FTrue, ����s�FFalse
Conf.Add "WD3", True ' �Ηj���Ɏ��s���邩�A���s�FTrue, ����s�FFalse
Conf.Add "WD4", True ' ���j���Ɏ��s���邩�A���s�FTrue, ����s�FFalse
Conf.Add "WD5", True ' �ؗj���Ɏ��s���邩�A���s�FTrue, ����s�FFalse
Conf.Add "WD6", True ' ���j���Ɏ��s���邩�A���s�FTrue, ����s�FFalse
Conf.Add "WD7", True ' �y�j���Ɏ��s���邩�A���s�FTrue, ����s�FFalse

'#==============================================================================
'# �O���[�o���ϐ��錾, �ݒ�l����, ���C�u�����Ǎ�

'#------------------------------------------------------------------------------
Main()

'#=============================================================================#
'#                               [ ���C���֐� ]                                #
'#=============================================================================#
Function Main()
	Dim wsh, fso
	Set wsh = WScript.CreateObject("WScript.Shell")
	Set fso = CreateObject("Scripting.FileSystemObject")
	'---------------------------------------------------------------------------
	' CSCRIPT.EXE �Ŗ����ꍇ�ɐ؂�ւ���

	Dim strPath, strTarget, strMyPath, strParam, j
	strPath   = WScript.FullName ' �t���p�X
	strTarget = Right( strPath, 11 )
	strTarget = Ucase( strTarget )

	if strTarget <> "CSCRIPT.EXE" then

		' �������g�̂̃t���p�X
		strMyPath = WScript.ScriptFullName
		strParam  = " "

		' ������S�ēǂ݂����āA�X�y�[�X���܂܂�Ă����ꍇ�͍ēx�����n���ׂ� " �ň͂ݒ����Ă��܂�
		For j = 0 to Wscript.Arguments.Count - 1
			If instr(Wscript.Arguments(j), " ") < 1 Then
				strParam = strParam & Wscript.Arguments(j) & " "
			Else
				strParam = strParam & """" & Wscript.Arguments(j) & """ "
			End If
		Next

		' CSCRIPT.EXE �Ɉ����p���ŏI��
		Call wsh.Run( "cscript.exe //nologo """ & strMyPath & """" & strParam )
		WScript.Quit

	end if
	'---------------------------------------------------------------------------

	' CSV�ǂݍ���
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
	WScript.Echo "�T�[�`�J�n�B���̃E�B���h�E�����ΏI�����܂��B"
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
			WScript.Echo "�T�[�`�F" & CStr(n) & "���(" & Time & ") " & nameArray(i) & "(" & CStr(i+1) & "�s��) "
			WScript.Sleep Conf("WAIT") - Conf("READ")
			ie.Quit
			Set ie = Nothing

			If Err.Number <> 0 Then
				WScript.Echo "�G���[�ԍ� " & CStr(Err.Number) & " " & Err.Description
				Err.Clear
				WScript.Sleep 5000
			Else
				i = i + 1
				n = n + 1
				If i = r Then: i = 0: End If
			End If
		Else
			If Not offFlg Then
				WScript.Echo "�T�[�`���s�����ɂȂ�܂őҋ@���ł��B"
				offFlg = True
			End If
		End If

	Loop

	Main = True
'#------------------------------------------------------------------------------
End Function

' �b�ɕϊ�
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