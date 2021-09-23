Option Explicit

' *************************************************************************************************
' 【コマンド自動実行】
' 
' ◆概要
'   コマンドやキー操作を操作定義ファイルに記述した操作内容に従って自動実行できる。
'   操作定義ファイルは第一引数として与える必要がある。
'   引数がない場合はこのファイルと同一階層の「OperationList.dat」を操作定義ファイルとして利用する。
'   ちなみに全然エラーチェックをしていないので利用には注意。
'   詳細はReadmeを参照。
' 
' *************************************************************************************************

' 定数宣言
Const OPE_FILENAME = "OperationList.dat"			' 操作内容の定義ファイル
Const WAIT_LOAD = 1000								' アプリケーション実行時の待ち時間
Const WAIT_SENDKEY = 100							' キー操作の待ち時間（一応）

' オブジェクト宣言
Dim objShell
Dim objFSO
Dim objOpeFile
Dim objArray										' Listを使うためのオブジェクト

' 変数宣言
Dim file_name
Dim line, line_num
Dim command, params
Dim result
Dim err_num
Dim loop_flag, loop_num, loop_count

' 実行時引数が設定されている場合、それを操作定義ファイル名としてセットする。
If Wscript.Arguments.Count = 0 Then
	file_name = OPE_FILENAME
Else
	file_name = Wscript.Arguments(0)
End If

' オブジェクトのセット
Set objShell = CreateObject("WScript.Shell")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objOpeFile = objFSO.OpenTextFile(file_name)
Set objArray = CreateObject("System.Collections.ArrayList")

' WScript.Echo("開始")

' 操作定義ファイルの記述エラー取得用。
On Error Resume Next

' datファイルを解析し、各種コマンドを実行。
line_num = 0
loop_flag = False
Do While objOpeFile.AtEndOfStream <> True
	line = Trim(objOpeFile.ReadLine)
	line_num = line_num + 1
	
	' コメント行と空行以外の場合に操作を実行する。
	If Left(line, 1) <> "'"  And line <> "" Then
		
		' Loopコマンドが有効な間、実行されているコマンドをArrayListに保存しておく。
		If loop_flag = True Then
			objArray.Add line
		End If
		
		' コマンド操作を1行実行
		DoCommand(line)
		
		' エラー処理
		If Err.Number <> 0 Then
			'操作定義ファイルのエラー行を出力
			err_num = Err.Number
			WScript.Echo("操作定義ファイルエラー！" & vbCrLf & "エラー行 : " & CStr(line_num) & vbCrLf & "エラーコード : " & err_num)
			
			' 通常のエラーも念の為出力
			On Error Goto 0
			Err.Raise(err_num)
			WScript.Quit()
		End If
	End If
Loop

'「On Error Resume Next」を解除
On Error Goto 0

' WScript.Echo("終了")

' オブジェクトの破棄
Set objShell = Nothing
Set objFSO = Nothing
Set objOpeFile = Nothing
Set objArray = Nothing


' *************************************************************************************************
' 以下、サブプロシージャ
' *************************************************************************************************

' 呼び元から受け取ったコマンドを実行。コマンド内容によって分岐。
Sub DoCommand(text)
	
	' コマンドの内容をコマンド部と引数部に分けるために解析する。
	SetCommand(text)
	
	' コマンド実行
	Select Case command
		
		' アプリケーションの実行
		Case "AppRun"
			objShell.Run(params)
			WScript.Sleep(WAIT_LOAD)
			
		' キーボード操作の実行
		Case "SendKey"
			objShell.SendKeys(params)
			WScript.Sleep(WAIT_SENDKEY)
			
		' 指定文字列の入力操作
		Case "Input"
			' 文字列をクリップボードにコピー
			result = objShell.Run("cmd /c ""set /P =""" & params & """ < NUL | clip""", 0, true)
			' 貼り付け（Ctrl + V）の実行
			objShell.SendKeys("^" + "v")
			WScript.Sleep(WAIT_SENDKEY)
			
		' メッセージ表示
		Case "Msg"
			WScript.Echo(params)
			
		' 待機の実施
		Case "Wait"
			WScript.Sleep(params)
			
		' ループの実行
		' これ以降、Loop(end)のコマンドが実行されるまでのコマンドをすべて保存しておき、それらのコマンドを指定のループ回数分実行する。
		Case "Loop"
			' 引数が数値の場合はループ開始処理を実行する。
			If IsNumeric(params) Then
				objArray.Clear
				loop_flag = True
				loop_num = CInt(params) - 1		' ループの1回目は解析中に実行するので「-1」する。
			' 引数が"end"の場合はループ終了処理を実行する。
			ElseIf params = "end" and loop_flag = True Then
				loop_flag = False
				' ループ開始時に取得したループ回数分コマンド郡を実行する。
				For loop_count = 1 To loop_num
					' 保存しておいた全てのコマンドを実行する。
					For Each line In objArray
						DoCommand(line)
					Next
				Next
			End IF
	End Select
End Sub

' 文字列を解析してコマンドと引数を取得
' 例：「AppRun(notepad.exe)」であればcommand="AppRun",params="notepad.exe"として取得する。 
Sub SetCommand(text)
	Dim pos_bracket_start
	Dim pos_bracket_end
	
	' カッコの前と中をそれぞれコマンドと引数として取得する。
	pos_bracket_start = InStr(text, "(")
	pos_bracket_end = InStrRev(text, ")")
	command = Mid(text, 1, pos_bracket_start - 1)
	params = Mid(text, pos_bracket_start + 1, pos_bracket_end - pos_bracket_start - 1)
	
End Sub

