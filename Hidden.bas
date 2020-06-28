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
            MsgBox "∫œ›ƒﬁÃﬂ€›ÃﬂƒÇäJÇ©Ç∏Ç…∫›ø∞Ÿ±ÃﬂÿÇé¿çs" & vbCrLf & _
                   vbCrLf & _
                   "Hidden.exe Ver." & .Major & "." & .Minor & "." & .Revision & vbCrLf & _
                   "Copyright (C) 2001 WATABE Eiji" & vbCrLf & _
                   vbCrLf & _
                   "usage: hidden ∫œ›ƒﬁñº", _
                   vbDefaultButton1, _
                   "Hidden ÇÃ ﬁ∞ºﬁÆ›èÓïÒ"
        End With
    End If

End Sub

