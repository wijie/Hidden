Attribute VB_Name = "CmdExec"
Option Explicit

Sub Main()

    Dim strComspec As String
    Dim strCmdLine As String

    If Command <> "" Then
        strComspec = Environ("COMSPEC")
        strCmdLine = strComspec & " /C " & Command
        Call ShellWait(strCmdLine, vbHide)
    Else
        With App
            MsgBox "�����������Ă��J�����ɺݿ�ٱ��؂����s" & vbCrLf & _
                   vbCrLf & _
                   "Hidden.exe Ver." & .Major & "." & .Minor & "." & .Revision & vbCrLf & _
                   "Copyright (C) 2001 WATABE Eiji" & vbCrLf & _
                   vbCrLf & _
                   "usage: hidden ����ޖ�", _
                   vbDefaultButton1, _
                   "Hidden ���ް�ޮݏ��"
        End With
    End If

End Sub

