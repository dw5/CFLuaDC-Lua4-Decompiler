Attribute VB_Name = "modLUA"
Option Explicit
' Module containing LUA file format, for opening LUA files.

#Const HaltOnErrors = 1                     ' Halt on errors like unknown instructions?

Const LUA_BinaryID = vbKeyEscape            ' binary files start with ESC...
Const LUA_Signature = "Lua"                 ' ...followed by this signature
Const LUA_Version = &H40                    ' ie 64 or "@"
Const LUA_Endianess = 1                     ' Must have one in the endianess.

' Test size values for LUA virtual machine...
' in bytes
Const LUA_TS_int = 4                        '
Const LUA_TS_size_t = 4                     '
Const LUA_TS_Instruction = 4                '

' In bits
Const LUA_TS_SIZE_INSTRUCTION = 32          '
Const LUA_TS_SIZE_SIZE_OP = 6               '
Const LUA_TS_SIZE_B = 9                     '

' In bytes again
Const LUA_TS_Number = 8                     ' A double

Const LUA_TestNumber = 314159265.358979     ' a multiple of PI for testing native format

' All strings that denote a token. Seperated by Null.
Public Const LUA_Tokens = _
    "~=" & " " & _
    "<=" & " " & _
    ">=" & " " & _
    "<" & " " & _
    ">" & " " & _
    "==" & " " & _
    "=" & " " & _
    "+" & " " & _
    "-" & " " & _
    "*" & " " & _
    "/" & " " & _
    "%" & " " & _
    "(" & " " & _
    ")" & " " & _
    "{" & " " & _
    "}" & " " & _
    "[" & " " & _
    "]" & " " & _
    ";" & " " & _
    "," & " " & _
    "." & " " & _
    ".." & " " & _
    "..."

' (INTERNAL) FILE FORMAT RELATED STUFF
Private Const MAX_FUNCTIONS = 256           ' number of functions we'll keep (this is to prevent direct
                                            ' recrusion)

Type LUA_Header
    binID As Byte                           ' = LUA_BinaryID
    sign As String * 3                      ' = LUA_Signature
    vers As Byte                            ' = LUA_Version
    byteOrder As Byte                       ' = LUA_Endianess
    
    ts_int As Byte                          ' = LUA_TS_int
    ts_size_t As Byte                       ' = LUA_TS_size_t
    ts_instruction As Byte                  ' = LUA_TS_Instruction
    
    ts_size_instruction  As Byte            ' = LUA_TS_SIZE_INSTRUCTION
    ts_size_size_op As Byte                 ' = LUA_TS_SIZE_SIZE_OP
    ts_size_b As Byte                       ' = LUA_TS_SIZE_B
    
    ts_number As Byte                       ' = LUA_TS_Number
    
    testNumber As Double                    ' = LUA_TestNumber
End Type

' Local variable.
Type LUA_LocalVariable
    lenName As Long                         ' name of the local var.
    name As String
    
    startpc As Long                         ' first point where variable is active
    endpc As Long                           ' first point where variable is dead
End Type

' This 'dummy' type sits here, just because I don't want lenData and data to be
' manually redimmed one after another!
Type LUA_String
    lenData As Long                         ' data, usually name.
    data As String
End Type

Type LUA_Chunk                              ' A chunk or a function; it's the same data
    ' Main header of this function.
    lenSource As Long                       ' source path of the LUA in question
    Source As String
    
    lineDefined As Long                     ' the line where this function is defined
    numParams As Long                       ' nr. of parameters for this function
    isVarArg As Byte                        ' does this function have variable nr. of arguments?
    maxStackSize As Long                    ' maximum stack size for this function
    
    ' data for local variables.
    numLocals As Long                       ' local variables
    locals() As LUA_LocalVariable
    
    ' This is related to debugging; not needed to decompile LUA (since it's possible to decompile
    ' "stripped" luas; these line info don't play a role in functioning of LUA).
    numLineInfo As Long                     ' line info (map from opcodes to source lines) ???
    lineInfo() As Long
    
    ' Constants section, strings, numbers, and functions(???) in functions, which is
    ' probably not used, atleast for HW2...
    numStrings As Long                      ' Constants->Strings
    strings() As LUA_String
    
    numNumbers As Long                      ' Constants->Numbers
    numbers() As Double
    
    numFunctions As Long                    ' Constants->Functions
    functions() As Long                     ' Pointers to the function in this
    
    ' and FINALLY... the code section.
    numInstructions As Long                 ' instructions
    instructions() As Long
    
    funcPtr As Long                         ' Pointers to the function and it's parent (not
    parentPtr As Long                       ' actually present in LUA! Just for convinience...)
End Type

Type LUA_File
    hdr As LUA_Header                       ' header fist...
    chunks() As LUA_Chunk                   ' ...and then chunks\functions whatever it may be.
    funcs(1 To MAX_FUNCTIONS) As LUA_Chunk  ' these are outside, since recrusion is maybe not possible here.
End Type

' Opens the given file and reads the given LUA, just loads into the memory, doesn't "undump" it.
Function ReadLUA(ByRef fileName As String, ByRef outLUA As LUA_File) As Boolean
    Dim fileNum As Integer
    Dim I As Long
    
    Dim test As Boolean
    
    fileNum = FreeFile
    Open fileName For Binary As #fileNum
        With outLUA
            Get #fileNum, , .hdr
            
            ' Verify the LUA is OK
            test = _
                (.hdr.binID = LUA_BinaryID) And _
                (.hdr.sign = LUA_Signature) And _
                (.hdr.vers = LUA_Version) And _
                (.hdr.byteOrder = LUA_Endianess) And _
                (.hdr.ts_int = LUA_TS_int) And _
                (.hdr.ts_size_t = LUA_TS_size_t) And _
                (.hdr.ts_instruction = LUA_TS_Instruction) And _
                (.hdr.ts_size_instruction = LUA_TS_SIZE_INSTRUCTION) And _
                (.hdr.ts_size_size_op = LUA_TS_SIZE_SIZE_OP) And _
                (.hdr.ts_size_b = LUA_TS_SIZE_B) And _
                (.hdr.ts_number = LUA_TS_Number) And _
                (Fix(.hdr.testNumber) = Fix(LUA_TestNumber))
                
            If LOF(fileNum) = 0 Then
                MsgBox "Size of given file = 0", _
                    vbApplicationModal + vbExclamation + vbDefaultButton1 + vbOKOnly, _
                    "Cold Fusion LUA Decompiler"
                
                Close #fileNum
                Kill fileName
                
                GoTo Failure
            End If  '   If LOF(fileNum) = 0 Then
            
            ' Verify that it's a LUA bin, not LUA text
            If Not ((.hdr.binID = LUA_BinaryID) And (.hdr.sign = LUA_Signature)) Then
                MsgBox "Need compiled binary files to be decompiled!" & vbCrLf & _
                    "This does not seem to be a compiled LUA binary!", _
                    vbApplicationModal + vbCritical + vbDefaultButton1 + vbOKOnly, _
                    "Cold Fusion LUA Decompiler"
                
                Close #fileNum
                
                GoTo Failure
            End If  '   If Not ((.hdr.binID = LUA_BinaryID) And (.hdr.sign = LUA_Signature)) Then
            
            If Not test Then
                MsgBox "This LUA is probably corrupted. Please get this LUA" & vbCrLf & _
                    "recompiled through LuaC to ensure that the header is OK.", _
                    vbApplicationModal + vbCritical + vbDefaultButton1 + vbOKOnly, _
                    "Cold Fusion LUA Decompiler"
                
                Close #fileNum
                
                GoTo Failure
            End If  '   If Not test Then
            
            ReDim .chunks(0)
            
            ' Read all chunks.
            Do While (Loc(fileNum) < LOF(fileNum)) And (Not EOF(fileNum))
                ReDim Preserve .chunks(UBound(.chunks()) + 1)
                
                ' Give pointers before, this is to avoid change in address (these just reflect hierarchy)
                .chunks(UBound(.chunks)).funcPtr = VarPtr(.chunks(UBound(.chunks)).funcPtr)
                .chunks(UBound(.chunks)).parentPtr = 0  ' apparently no parent.
                
                ReadLUAFunction fileNum, .chunks(UBound(.chunks())), .chunks(UBound(.chunks())), outLUA
            Loop    '   Do While (Loc(fileNum) < LOF(fileNum)) And (Not EOF(fileNum))
        End With
    Close #fileNum
    
    ' Done!
    ReadLUA = True
    Exit Function
Failure:
    ReadLUA = False
    Exit Function
End Function

' Reads the given LUA function\chunk.
Function ReadLUAFunction(ByVal fileNum As Integer, ByRef outFunc As LUA_Chunk, ByRef parentChunk As LUA_Chunk, ByRef theLUA As LUA_File) As Boolean
Attribute ReadLUAFunction.VB_Description = "Reads the given LUA function\\chunk. Do NOT call from a function! (OK to call from a chunk). A ""This array is temporary locked"" error will occur."
    Dim test As Boolean
    Dim I As Long, funcIndex As Long
    
    Static funcCallNr As Long
    
    funcCallNr = funcCallNr + 1
    
    With outFunc
        ReDim .locals(0)
        ReDim .lineInfo(0)
        
        ReDim .strings(0)       ' \
        ReDim .numbers(0)       '  > Togetherly known as constants...
        'ReDim .functions(0)    ' /
        
        ReDim .instructions(0)
        
        ' -- Chunk header -- '
        Get #fileNum, , .lenSource
        .Source = Space$(.lenSource)
        
        Get #fileNum, , .Source
        .Source = ChopTerminatingNull(.Source)
        
        Get #fileNum, , .lineDefined
        Get #fileNum, , .numParams
        Get #fileNum, , .isVarArg
        Get #fileNum, , .maxStackSize
        
        ' -- Local Vars -- '
        Get #fileNum, , .numLocals
        If .numLocals > 0 Then ReDim _
            .locals(1 To .numLocals)
        
        For I = 1 To .numLocals
            Get #fileNum, , .locals(I).lenName
            .locals(I).name = Space$(.locals(I).lenName)
            
            Get #fileNum, , .locals(I).name
            .locals(I).name = ChopTerminatingNull(.locals(I).name)
            
            Get #fileNum, , .locals(I).startpc
            Get #fileNum, , .locals(I).endpc
        Next I  '   For I = 1 To .numLocals
        
        ' -- Line Infos -- '
        Get #fileNum, , .numLineInfo
        
        If .numLineInfo > 0 Then _
            ReDim .lineInfo(1 To .numLineInfo): _
            Get #fileNum, , .lineInfo
        
        ' -- Constants -- '
            ' -- Strings -- '
            Get #fileNum, , .numStrings
            If .numStrings > 0 Then ReDim _
                .strings(1 To .numStrings)
            
            For I = 1 To .numStrings
                Get #fileNum, , .strings(I).lenData
                .strings(I).data = Space$(.strings(I).lenData)
                
                Get #fileNum, , .strings(I).data
                .strings(I).data = ChopTerminatingNull(.strings(I).data)
            Next I  '   For I = 1 To .numStrings
            
            ' -- Numbers -- '
            Get #fileNum, , .numNumbers
            
            If .numNumbers > 0 Then _
                ReDim .numbers(1 To .numNumbers): _
                Get #fileNum, , .numbers
            
            ' -- Functions -- '
            Get #fileNum, , .numFunctions
            
            If .numFunctions > 0 Then _
                ReDim .functions(1 To .numFunctions)
            
            For I = 1 To .numFunctions
                ' Save this, incase we have more functions inside this
                funcIndex = funcCallNr
                
                ' Give pointers before, this is to avoid change in address (these just reflect hierarchy)
                theLUA.funcs(funcIndex).funcPtr = VarPtr(theLUA.funcs(funcIndex))
                theLUA.funcs(funcIndex).parentPtr = outFunc.funcPtr
                
                ' and note the function index here, too.
                .functions(I) = theLUA.funcs(funcIndex).funcPtr
                
                If theLUA.funcs(funcIndex).parentPtr = theLUA.funcs(funcIndex).funcPtr Then _
                    theLUA.funcs(funcIndex).parentPtr = 0
                
                ' Finally read the function.
                test = ReadLUAFunction(fileNum, theLUA.funcs(funcIndex), outFunc, theLUA)
                Debug.Assert test
            Next I  '   For I = 1 To .numFunctions
            
        ' -- Instructions -- '
        Get #fileNum, , .numInstructions
        If .numInstructions > 0 Then _
            ReDim .instructions(1 To .numInstructions): _
            Get #fileNum, , .instructions
        
        ' HOPE that last is OP_END
        Debug.Assert .instructions(.numInstructions) = 0
    End With
    
    ReadLUAFunction = True  ' Success!
End Function
