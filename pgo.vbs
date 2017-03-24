Option Explicit
On Error Resume Next
'#������������������������������������������������������������������������������
'#��  �����T�[�` v0.0.2 (2017/03/24)
'#��  pgo.vbs 
'#������������������������������������������������������������������������������
'#
'# [ �X�V���� ]
'# 2017/03/24 -> v0.0.2
'#  IE���t���[�Y���̏����ǉ�
'# 
'# 2017/03/14 -> v0.0.1
'#  �����̊���
'# 
'# [ �g���� ]
'#  ���̃t�@�C���𓯊K�w��list.csv�Ƃ����e�L�X�g�t�@�C����u���Ă��������B
'#  list.csv�ɂ̓T�[�`�������ꏊ���J���}��؂�ňܓx,�o�x�������܂��B
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
'#
'#==============================================================================
'# �ݒ�l�������ɋL�ڂ���B
Dim CMN
Set CMN = CreateObject( "Scripting.Dictionary" )

'#-- [ ��{�ݒ� ] --------------------------------------------------------------

Call CMN.Add("URL",  "https://pmap.kuku.lu/#") ' P-GO SEARCH URL
Call CMN.Add("BTN",  "area_buttonsearch") ' �T�[�`�{�^��ID
Call CMN.Add("WAIT",120*1000) ' �T�[�`��ҋ@�b
Call CMN.Add("IE",   True) ' IE��\�����邩�A�\���FTrue, ��\���FFalse
Call CMN.Add("LIST", "list.csv") ' �T�[�`������W�������ꂽ�e�L�X�g�t�@�C��

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
	WScript.Echo "�T�[�`�J�n�B���̃E�B���h�E�����ΏI�����܂��B"
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
		WScript.Echo "�T�[�`�F" & CStr(n) & "���(" & Time & ") " & CStr(i+1) & "�s��(" & latArray(i) & "," & lngArray(i) & ") "
		WScript.Sleep CMN("WAIT")
		ie.Quit

		If Err.Number <> 0 Then
			WScript.Echo "�G���[�ԍ� " & CStr(Err.Number) & " " & Err.Description
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

