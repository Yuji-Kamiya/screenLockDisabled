'�X���[�v����
wait = 60000
' ���M�L�[
sendkey = "{F16}"
 
'�N���`�F�b�N
query = "Select * FROM Win32_Process WHERE (Caption = 'wscript.exe' OR Caption = 'cscript.exe') AND " _
         & " CommandLine LIKE '%" & WScript.ScriptName & "%'"
Set wmiLocator = CreateObject("WbemScripting.SWbemLocator")
Set wmiService = wmiLocator.ConnectServer
Set objEnumerator = wmiService.ExecQuery(query)
 
Dim counter
counter = 0
 
If objEnumerator.Count > 1 Then
    '>> ���s��
    message = "���b�N�����������s���ł��B" & vbCrLf _
            & "�I�����܂����H"
    If MsgBox(message , vbYesNo + vbQuestion, "���s��") = vbYes Then
        For Each objProcess In objEnumerator
            counter = counter + 1
            If counter <> objEnumerator.Count then
                '�Ō�̃v���Z�X�i�������g�j�ȊO���I��
                objProcess.Terminate
            End If
        Next
    End If
    WScript.Quit 0
Else
    '>> �V�K
    message = "���Ԋu�ŃL�[�𑗐M���A��ʃ��b�N�𖳌������܂��B" & vbCrLf _
            & "���s���܂����H"
    If MsgBox(message , vbYesNo + vbQuestion, "��ʃ��b�N������") = vbYes Then
        Call StopWindowsLock(sendkey, wait)
    End If
End If
 
WScript.Quit 0
 

'���M�L�[�̒�����M
Sub StopWindowsLock(key, wait)
    Set WshShell = CreateObject("Wscript.Shell")
    Do
        WshShell.SendKeys(key)
        WScript.Sleep wait
    Loop
End Sub
