'スリープ時間
wait = 60000
' 送信キー
sendkey = "{F16}"
 
'起動チェック
query = "Select * FROM Win32_Process WHERE (Caption = 'wscript.exe' OR Caption = 'cscript.exe') AND " _
         & " CommandLine LIKE '%" & WScript.ScriptName & "%'"
Set wmiLocator = CreateObject("WbemScripting.SWbemLocator")
Set wmiService = wmiLocator.ConnectServer
Set objEnumerator = wmiService.ExecQuery(query)
 
Dim counter
counter = 0
 
If objEnumerator.Count > 1 Then
    '>> 実行中
    message = "ロック無効化が実行中です。" & vbCrLf _
            & "終了しますか？"
    If MsgBox(message , vbYesNo + vbQuestion, "実行中") = vbYes Then
        For Each objProcess In objEnumerator
            counter = counter + 1
            If counter <> objEnumerator.Count then
                '最後のプロセス（自分自身）以外を終了
                objProcess.Terminate
            End If
        Next
    End If
    WScript.Quit 0
Else
    '>> 新規
    message = "一定間隔でキーを送信し、画面ロックを無効化します。" & vbCrLf _
            & "実行しますか？"
    If MsgBox(message , vbYesNo + vbQuestion, "画面ロック無効化") = vbYes Then
        Call StopWindowsLock(sendkey, wait)
    End If
End If
 
WScript.Quit 0
 

'送信キーの定期送信
Sub StopWindowsLock(key, wait)
    Set WshShell = CreateObject("Wscript.Shell")
    Do
        WshShell.SendKeys(key)
        WScript.Sleep wait
    Loop
End Sub
