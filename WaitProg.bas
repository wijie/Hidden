Attribute VB_Name = "WaitProg"
Option Explicit

'�n���h�����擾
Declare Function OpenProcess Lib "kernel32.dll" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

'�I���X�e�[�^�X���擾
Public Declare Function GetExitCodeProcess Lib "kernel32.dll" ( _
    ByVal hProcess As Long, _
    lpExitCode As Long) As Long

'�n���h�������
Declare Function CloseHandle Lib "kernel32.dll" ( _
    ByVal hObject As Long) As Long

Const PROCESS_QUERY_INFORMATION = &H400&
Const STILL_ACTIVE = &H103&

Public Sub ShellWait(ByVal strProg As String, _
                     ByVal bytStyle As Byte)

    Dim lnProcHandle As Long
    Dim lnExitCode As Long
    Dim lnReturnCode As Long
    Dim lnTaskID As Long

    lnTaskID = Shell(strProg, bytStyle)
    DoEvents

    '�n���h�����擾����
    lnProcHandle = OpenProcess(PROCESS_QUERY_INFORMATION, _
                               1, _
                               lnTaskID)
    '�I���܂ő҂�
    Do
        lnReturnCode = GetExitCodeProcess(lnProcHandle, _
                                          lnExitCode)
        DoEvents
    Loop While lnExitCode = STILL_ACTIVE

    '�n���h�������
    CloseHandle lnProcHandle

End Sub
