Option Explicit

' *************************************************************************************************
' �y�R�}���h�������s�z
' 
' ���T�v
'   �R�}���h��L�[����𑀍��`�t�@�C���ɋL�q����������e�ɏ]���Ď������s�ł���B
'   �����`�t�@�C���͑������Ƃ��ė^����K�v������B
'   �������Ȃ��ꍇ�͂��̃t�@�C���Ɠ���K�w�́uOperationList.dat�v�𑀍��`�t�@�C���Ƃ��ė��p����B
'   ���Ȃ݂ɑS�R�G���[�`�F�b�N�����Ă��Ȃ��̂ŗ��p�ɂ͒��ӁB
' 
' 
' ���g����R�}���h
'   AppRun
'     �����Ɏw�肵���A�v���P�[�V�����inotepad.exe���j�����s����B
' 
'   SendKey
'     �L�[�{�[�h���삪���̂܂܎��s�ł���B
'     �����̎w���WSH�̋K��ɏ]���B�i�Q�l�Fhttp://jscript.zouri.jp/Source/KeybordCtrl.html�j
' 
'   Input
'     ���݂̃J�[�\���ʒu�Ɉ����̕��������͂���B
'     (���Ԃ͂����̃R�s�y)
' 
'   Msg
'     �����̓��e�����b�Z�[�W�_�C�A���O�Ƃ��ĕ\������B
' 
'   Wait
'     �����Ɏw�肵�����ԁims�j������������Ȃ��B
'     �A�v���P�[�V�����̎��s�҂��̎��Ԓ����Ȃǂɂ����B
' 
' 
' �������`�t�@�C���ɂ���
'   1�s�ɂ�1�R�}���h�L�q����B
'   �C���f���g����Ȃ�^�u�͎g�킸���p�X�y�[�X���g�����ƁB
'   �擪��"'"������΃R�����g�s�ɂł���B
' 
'   ��F�ȉ��̂悤�ȊO���t�@�C�����쐬����B
'     Msg(�J�n)
'     AppRun(notepad.exe)
'     SendKey(All is well that ends well.)
'     Wait(1000)
'     SendKey({ENTER 1})
'     Input(��F�I���ǂ���ΑS�Ă悵)
'     Msg(�I��)
' 
' *************************************************************************************************

' �萔�錾
Const OPE_FILENAME = "OperationList.dat"			' ������e�̒�`�t�@�C��
Const WAIT_LOAD = 1000							' �A�v���P�[�V�������s���̑҂�����
Const WAIT_SENDKEY = 100						' �L�[����̑҂����ԁi�ꉞ�j

' �I�u�W�F�N�g�錾
Dim objShell
Dim objFSO
Dim objOpeFile

' �ϐ��錾
Dim file_name
Dim line, line_num
Dim command, params
Dim result
Dim err_num

If Wscript.Arguments.Count = 0 Then
	file_name = OPE_FILENAME
Else
	file_name = Wscript.Arguments(0)
End If

' �I�u�W�F�N�g�̃Z�b�g
Set objShell = CreateObject("WScript.Shell")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objOpeFile = objFSO.OpenTextFile(file_name)

' WScript.Echo("�J�n")

' �����`�t�@�C���̋L�q�G���[�擾�p�B
On Error Resume Next

' dat�t�@�C������͂��A�e��R�}���h�����s�B
line_num = 0
Do While objOpeFile.AtEndOfStream <> True
	line = Trim(objOpeFile.ReadLine)
	line_num = line_num + 1
	
	' �R�����g�s�Ƌ�s�ȊO�̏ꍇ�ɑ�������s����B
	If Left(line, 1) <> "'"  And line <> "" Then
		
		' dat�t�@�C���̒��g���擾
		SetCommand(line)
		
		' �R�}���h���s
		Select Case command
			
			' �A�v���P�[�V�����̎��s
			Case "AppRun"
				objShell.Run(params)
				WScript.Sleep(WAIT_LOAD)
				
			' �L�[�{�[�h����̎��s
			Case "SendKey"
				objShell.SendKeys(params)
				WScript.Sleep(WAIT_SENDKEY)
				
			' �w�蕶����̓��͑���
			Case "Input"
				' ��������N���b�v�{�[�h�ɃR�s�[
				result = objShell.Run("cmd /c ""set /P =""" & params & """ < NUL | clip""", 0, true)
				' �\��t���iCtrl + V�j�̎��s
				objShell.SendKeys("^" + "v")
				WScript.Sleep(WAIT_SENDKEY)
				
			' ���b�Z�[�W�\��
			Case "Msg"
				WScript.Echo(params)
				
			' �ҋ@�̎��{
			Case "Wait"
				WScript.Sleep(params)
		End Select
		
		' �G���[����
		If Err.Number <> 0 Then
			'�����`�t�@�C���̃G���[�s���o��
			err_num = Err.Number
			WScript.Echo("�����`�t�@�C���G���[�I" & vbCrLf & "�G���[�s : " & CStr(line_num) & vbCrLf & "�G���[�R�[�h : " & err_num)
			
			' �ʏ�̃G���[���O�̈׏o��
			On Error Goto 0
			Err.Raise(err_num)
			WScript.Quit()
		End If
	End If
Loop

'�uOn Error Resume Next�v������
On Error Goto 0

' WScript.Echo("�I��")

' �I�u�W�F�N�g�̔j��
Set objShell = Nothing
Set objFSO = Nothing
Set objOpeFile = Nothing


' *************************************************************************************************
' �ȉ��A�T�u�v���V�[�W��
' *************************************************************************************************

' ���������͂��ăR�}���h�ƈ������擾
' ��F�uAppRun(notepad.exe)�v�ł����command="AppRun",params="notepad.exe"�Ƃ��Ď擾����B 
Sub SetCommand(line)
	Dim pos_bracket_start
	Dim pos_bracket_end
	
	pos_bracket_start = InStr(line, "(")
	pos_bracket_end = InStrRev(line, ")")
	command = Mid(line, 1, pos_bracket_start - 1)
	params = Mid(line, pos_bracket_start + 1, pos_bracket_end - pos_bracket_start - 1)
	
End Sub

