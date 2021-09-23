Option Explicit

' *************************************************************************************************
' �y�R�}���h�������s�z
' 
' ���T�v
'   �R�}���h��L�[����𑀍��`�t�@�C���ɋL�q����������e�ɏ]���Ď������s�ł���B
'   �����`�t�@�C���͑������Ƃ��ė^����K�v������B
'   �������Ȃ��ꍇ�͂��̃t�@�C���Ɠ���K�w�́uOperationList.dat�v�𑀍��`�t�@�C���Ƃ��ė��p����B
'   ���Ȃ݂ɑS�R�G���[�`�F�b�N�����Ă��Ȃ��̂ŗ��p�ɂ͒��ӁB
'   �ڍׂ�Readme���Q�ƁB
' 
' *************************************************************************************************

' �萔�錾
Const OPE_FILENAME = "OperationList.dat"			' ������e�̒�`�t�@�C��
Const WAIT_LOAD = 1000								' �A�v���P�[�V�������s���̑҂�����
Const WAIT_SENDKEY = 100							' �L�[����̑҂����ԁi�ꉞ�j

' �I�u�W�F�N�g�錾
Dim objShell
Dim objFSO
Dim objOpeFile
Dim objArray										' List���g�����߂̃I�u�W�F�N�g

' �ϐ��錾
Dim file_name
Dim line, line_num
Dim command, params
Dim result
Dim err_num
Dim loop_flag, loop_num, loop_count

' ���s���������ݒ肳��Ă���ꍇ�A����𑀍��`�t�@�C�����Ƃ��ăZ�b�g����B
If Wscript.Arguments.Count = 0 Then
	file_name = OPE_FILENAME
Else
	file_name = Wscript.Arguments(0)
End If

' �I�u�W�F�N�g�̃Z�b�g
Set objShell = CreateObject("WScript.Shell")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objOpeFile = objFSO.OpenTextFile(file_name)
Set objArray = CreateObject("System.Collections.ArrayList")

' WScript.Echo("�J�n")

' �����`�t�@�C���̋L�q�G���[�擾�p�B
On Error Resume Next

' dat�t�@�C������͂��A�e��R�}���h�����s�B
line_num = 0
loop_flag = False
Do While objOpeFile.AtEndOfStream <> True
	line = Trim(objOpeFile.ReadLine)
	line_num = line_num + 1
	
	' �R�����g�s�Ƌ�s�ȊO�̏ꍇ�ɑ�������s����B
	If Left(line, 1) <> "'"  And line <> "" Then
		
		' Loop�R�}���h���L���ȊԁA���s����Ă���R�}���h��ArrayList�ɕۑ����Ă����B
		If loop_flag = True Then
			objArray.Add line
		End If
		
		' �R�}���h�����1�s���s
		DoCommand(line)
		
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
Set objArray = Nothing


' *************************************************************************************************
' �ȉ��A�T�u�v���V�[�W��
' *************************************************************************************************

' �Ăь�����󂯎�����R�}���h�����s�B�R�}���h���e�ɂ���ĕ���B
Sub DoCommand(text)
	
	' �R�}���h�̓��e���R�}���h���ƈ������ɕ����邽�߂ɉ�͂���B
	SetCommand(text)
	
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
			
		' ���[�v�̎��s
		' ����ȍ~�ALoop(end)�̃R�}���h�����s�����܂ł̃R�}���h�����ׂĕۑ����Ă����A�����̃R�}���h���w��̃��[�v�񐔕����s����B
		Case "Loop"
			' ���������l�̏ꍇ�̓��[�v�J�n���������s����B
			If IsNumeric(params) Then
				objArray.Clear
				loop_flag = True
				loop_num = CInt(params) - 1		' ���[�v��1��ڂ͉�͒��Ɏ��s����̂Łu-1�v����B
			' ������"end"�̏ꍇ�̓��[�v�I�����������s����B
			ElseIf params = "end" and loop_flag = True Then
				loop_flag = False
				' ���[�v�J�n���Ɏ擾�������[�v�񐔕��R�}���h�S�����s����B
				For loop_count = 1 To loop_num
					' �ۑ����Ă������S�ẴR�}���h�����s����B
					For Each line In objArray
						DoCommand(line)
					Next
				Next
			End IF
	End Select
End Sub

' ���������͂��ăR�}���h�ƈ������擾
' ��F�uAppRun(notepad.exe)�v�ł����command="AppRun",params="notepad.exe"�Ƃ��Ď擾����B 
Sub SetCommand(text)
	Dim pos_bracket_start
	Dim pos_bracket_end
	
	' �J�b�R�̑O�ƒ������ꂼ��R�}���h�ƈ����Ƃ��Ď擾����B
	pos_bracket_start = InStr(text, "(")
	pos_bracket_end = InStrRev(text, ")")
	command = Mid(text, 1, pos_bracket_start - 1)
	params = Mid(text, pos_bracket_start + 1, pos_bracket_end - pos_bracket_start - 1)
	
End Sub

