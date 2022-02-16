Attribute VB_Name = "modMain"
Option Explicit
' Main exec module for CFLDC

#Const ReleaseBuild = 1

Private Sub Main()
    Dim success As Boolean
    Dim file As String, outFile As String
    Dim DecompileTime As Single
    
    outFile = CStr(InStr(1, Command$(), " -o:" & Chr$(34)))
    
    If InStr(1, Command$(), " -o:" & Chr$(34)) > 0 Then
        file = LTrim$(RTrim$(Left$(Command$(), CLng(outFile))))
        file = UnquoteString(file)
        
        outFile = LTrim$(Right$(Command$(), Len(Command$()) - CLng(outFile)))
        outFile = Right$(outFile, Len(outFile) - 3)
        outFile = UnquoteString(outFile)
        
        GoTo DecompileLUA
    Else    '   End If  '   If InStr(1, Command$(), " -o:" & Chr$(34)) > 0 Then
        file = UnquoteString(Command$())
    End If  '   If InStr(1, Command$(), " -o:" & Chr$(34)) > 0 Then
    
    If Len(file) > 0 Then
GetFileName:
        outFile = Left$(file, InStrRev(file, ".") - 1) & Replace$(file, ".", "_DC.", InStrRev(file, "."))
        
DecompileLUA:
        DecompileTime = Timer
        
        If Not LUA_Decompile(file, outFile) Then _
            MsgBox "The LUA was not successfully decompiled!", _
                vbApplicationModal + vbCritical + vbDefaultButton1 + vbOKOnly, _
                "Cold Fusion LUA Decompiler"
                
        DecompileTime = Timer - DecompileTime
        Debug.Print DecompileTime
    Else    '   If Len(file) > 0 Then
#If ReleaseBuild Then
        Load frmMain
        frmMain.Show vbModal
#Else
        file = App.path & "\Test.lua"
        GoTo GetFileName
#End If
    End If  '   If Len(file) > 0 Then
    
End Sub
