Attribute VB_Name = "modWaitForExecCmd"
Option Compare Database
Option Explicit

'modWaitForExecCmd
'use to wait for a shelled process to complete before continueing
'

Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
                                                             ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, _
                                                        ByVal lpCommandLine As String, _
                                                        ByVal lpProcessAttributes As Long, _
                                                        ByVal lpThreadAttributes As Long, _
                                                        ByVal bInheritHandles As Long, _
                                                        ByVal dwCreationFlags As Long, _
                                                        ByVal lpEnvironment As Long, _
                                                        ByVal lpCurrentDirectory As Long, _
                                                        lpStartupInfo As STARTUPINFO, _
                                                        lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, _
                                                            lpExitCode As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Public Function WaitForExecCmd(cmdline$)
    Dim proc As PROCESS_INFORMATION
    Dim Start As STARTUPINFO
    Dim ret As Variant
    ' Initialize the STARTUPINFO structure:
    Start.cb = Len(Start)
    'Start.wShowWindow = 6 'SW_MINIMIZE
    ' Start the shelled application:
    ret = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
    NORMAL_PRIORITY_CLASS, 0&, 0&, Start, proc)

    ' Wait for the shelled application to finish:
    ret = WaitForSingleObject(proc.hProcess, INFINITE)
    Call GetExitCodeProcess(proc.hProcess, ret)
    Call CloseHandle(proc.hThread)
    Call CloseHandle(proc.hProcess)
    WaitForExecCmd = ret
End Function
