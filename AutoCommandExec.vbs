Option Explicit

' *************************************************************************************************
' 【コマンド自動実行】
' 
' ◆概要
'   コマンドやキー操作を操作定義ファイルに記述した操作内容に従って自動実行できる。
'   操作定義ファイルは第一引数として与える必要がある。
'   引数がない場合はこのファイルと同一階層の「OperationList.dat」を操作定義ファイルとして利用する。
'   ちなみに全然エラーチェックをしていないので利用には注意。
' 
' 
' ◆使えるコマンド
'   AppRun
'     引数に指定したアプリケーション（notepad.exe等）を実行する。
' 
'   SendKey
'     キーボード操作がそのまま実行できる。
'     引数の指定はWSHの規定に従う。（参考：http://jscript.zouri.jp/Source/KeybordCtrl.html）
' 
'   Input
'     現在のカーソル位置に引数の文字列を入力する。
'     (実態はただのコピペ)
' 
'   Msg
'     引数の内容をメッセージダイアログとして表示する。
' 
'   Wait
'     引数に指定した時間（ms）何も操作をしない。
'     アプリケーションの実行待ちの時間調整などにつかう。
' 
' 
' ◆操作定義ファイルについて
'   1行につき1コマンド記述する。
'   インデントするならタブは使わず半角スペースを使うこと。
'   先頭に"'"があればコメント行にできる。
' 
'   例：以下のような外部ファイルを作成する。
'     Msg(開始)
'     AppRun(notepad.exe)
'     SendKey(All is well that ends well.)
'     Wait(1000)
'     SendKey({ENTER 1})
'     Input(訳：終わり良ければ全てよし)
'     Msg(終了)
' 
' *************************************************************************************************

' 定数宣言
Const OPE_FILENAME = "OperationList.dat"			' 操作内容の定義ファイル
Const WAIT_LOAD = 1000							' アプリケーション実行時の待ち時間
Const WAIT_SENDKEY = 100						' キー操作の待ち時間（一応）

' オブジェクト宣言
Dim objShell
Dim objFSO
Dim objOpeFile

' 変数宣言
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

' オブジェクトのセット
Set objShell = CreateObject("WScript.Shell")
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objOpeFile = objFSO.OpenTextFile(file_name)

' WScript.Echo("開始")

' 操作定義ファイルの記述エラー取得用。
On Error Resume Next

' datファイルを解析し、各種コマンドを実行。
line_num = 0
Do While objOpeFile.AtEndOfStream <> True
	line = Trim(objOpeFile.ReadLine)
	line_num = line_num + 1
	
	' コメント行と空行以外の場合に操作を実行する。
	If Left(line, 1) <> "'"  And line <> "" Then
		
		' datファイルの中身を取得
		SetCommand(line)
		
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
		End Select
		
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


' *************************************************************************************************
' 以下、サブプロシージャ
' *************************************************************************************************

' 文字列を解析してコマンドと引数を取得
' 例：「AppRun(notepad.exe)」であればcommand="AppRun",params="notepad.exe"として取得する。 
Sub SetCommand(line)
	Dim pos_bracket_start
	Dim pos_bracket_end
	
	pos_bracket_start = InStr(line, "(")
	pos_bracket_end = InStrRev(line, ")")
	command = Mid(line, 1, pos_bracket_start - 1)
	params = Mid(line, pos_bracket_start + 1, pos_bracket_end - pos_bracket_start - 1)
	
End Sub

