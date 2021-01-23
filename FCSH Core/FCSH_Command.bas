Attribute VB_Name = "FCSH_Command"
'FCSH Script Runtime Core
'By WMProject1217

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Public FCSH_VAL_ScriptPath
Public FCSH_VAL_ScriptLine
Public FCSH_VAL_Sysdisk
Public FCSH_VAL_Sysroot
Public FCSH_VAL_Pathnow
Public FCSH_VAL_CMD
Public Retval

Public Function ErrorReport(ByVal SubName As String, ByVal ErrorCode As String, ByVal ErrorIndex As String)
Retval = MsgBox("发生错误" & vbCrLf & "在 " & FCSH_VAL_ScriptPath & " [Line " & FCSH_VAL_ScriptLine & "]" & vbCrLf & "在 " & SubName & vbCrLf & "错误 " & ErrorCode & vbCrLf & ErrorIndex, 16, "Fatal Error - FCSH Script Lanucher")
End
End Function

Public Sub Timeout(ByVal Milliseconds As Long)
Dim lngTime As Long
lngTime = timeGetTime
While timeGetTime < lngTime + Milliseconds
DoEvents
Wend
End Sub

Public Sub FCSHCommand()
On Error GoTo FCSHCOMMANDEE
Dim TextLine
FCSH_VAL_CMD = Command
FCSH_VAL_CMD = Replace(FCSH_VAL_CMD, Chr(34), "")
FCSH_VAL_CMD = Replace(FCSH_VAL_CMD, Chr(39), "")

FCSH_VAL_Sysdisk = Environ("SystemDrive")
FCSH_VAL_Sysroot = Environ("SystemRoot")
FCSH_VAL_Paathnow = App.Path
FCSH_VAL_ScriptPath = FCSH_VAL_CMD
If FCSH_VAL_CMD = "help" Or FCSH_VAL_CMD = "HELP" Or FCSH_VAL_CMD = "Help" Then
FCSH_Help.Show
Exit Sub
End
End If
Open FCSH_VAL_CMD For Input As #1
FCSH_VAL_ScriptLine = 0
Dim CMDSplit() As String
Do While Not EOF(1)
Line Input #1, TextLine
TextLine = Replace(TextLine, "#AppPath#", App.Path)
TextLine = Replace(TextLine, "#ScriptPath#", FCSH_VAL_ScriptPath)
TextLine = Replace(TextLine, "#SystemDrive#", FCSH_VAL_Sysdisk)
TextLine = Replace(TextLine, "#SystemRoot#", FCSH_VAL_Sysroot)
TextLine = Replace(TextLine, "#Date#", Date)
TextLine = Replace(TextLine, "#Time#", Time)
CMDSplit = Split(TextLine, "=")

If CMDSplit(0) = "playsound" Or CMDSplit(0) = "PLAYSOUND" Or CMDSplit(0) = "Playsound" Then
On Error GoTo PLAYSOUNDEE
Rteval = Shell(App.Path & "\BGM.exe " & CMDSplit(1), vbHide)
GoTo PLAYSOUNDED
PLAYSOUNDEE:
Retval = ErrorReport("Playsound()", "0x00000053", "不能启动指定的文件")
PLAYSOUNDED:
End If

If CMDSplit(0) = "exec" Or CMDSplit(0) = "EXEC" Or CMDSplit(0) = "Exec" Then
On Error GoTo EXECEE
Retval = Shell(FCSH_VAL_Pathnow & CMDSplit(1), vbNormalFocus)
GoTo EXECED
EXECEE:
Retval = ErrorReport("EXEC()", "0x00000053", "不能启动指定的文件")
EXECED:
End If

If CMDSplit(0) = "exechc" Or CMDSplit(0) = "EXECHC" Or CMDSplit(0) = "Exechc" Then
On Error GoTo EXECHCEE
Retval = Shell(FCSH_VAL_Pathnow & CMDSplit(1), vbHide)
GoTo EXECHCED
EXECHCEE:
Retval = ErrorReport("EXECHC()", "0x00000053", "不能启动指定的文件")
EXECHCED:
End If

If CMDSplit(0) = "timeout" Or CMDSplit(0) = "TIMEOUT" Or CMDSplit(0) = "Timeout" Then
Timeout (CMDSplit(1))
End If

If CMDSplit(0) = "cd" Or CMDSplit(0) = "CD" Or CMDSplit(0) = "ChooseDir" Then
FCSH_VAL_Paathnow = CMDSplit(1)
End If

If CMDSplit(0) = "kill" Or CMDSplit(0) = "KILL" Or CMDSplit(0) = "Kill" Then
On Error GoTo KILLEE
Retval = Shell(FCSH_VAL_Sysroot & "\System32\taskkill.exe -f -im " & CMDSplit(1), vbHide)
GoTo KILLED
KILLEE:
Retval = ErrorReport("Kill()", "0x00000053", "不能启动指定的文件")
KILLED:
End If

If CMDSplit(0) = "tipmessage" Or CMDSplit(0) = "TIPMESSAGE" Or CMDSplit(0) = "TipMessage" Then
On Error GoTo TIPMESSAGEEE
Rteval = Shell(App.Path & "\TipMessage.exe " & CMDSplit(1) & " " & CMDSplit(2), vbHide)
GoTo TIPMESSAGEED
TIPMESSAGEEE:
Retval = ErrorReport("TipMessage()", "0x00000053", "不能启动指定的文件")
TIPMESSAGEED:
End If

FCSH_VAL_ScriptLine = FCSH_VAL_ScriptLine + 1
Loop
Close #1
GoTo FCSHCOMMANDED
FCSHCOMMANDEE:
Retval = ErrorReport("FCSHCommand()", "0x00000001", "指令 : " & FCSH_VAL_CMD)
FCSHCOMMANDED:
End Sub

Sub Main()
FCSHCommand
End Sub
