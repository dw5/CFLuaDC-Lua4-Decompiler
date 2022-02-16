Attribute VB_Name = "modLUAInstruction"
Option Explicit
' Module containing LUA instruction-specific code, for decoding instructions.

#Const PrintStack = 0           ' To print debug and stack info (causes an information overload!).
#Const HaltOnErrors = 1         ' Halt on errors like unknown instructions?
#Const UseLineInfo = 1          ' Use line info (ex. for multiple conditions?)

' K = U argument used as index to `kstr'
' J = S argument used as jump offset (relative to pc of next instruction)
' L = unsigned argument used as index of local variable
' N = U argument used as index to `knum'

Private Enum LUA_OPCodes
    ' ------------------------------------------------------------------------'
    ' name            args    stack before    stack after       side effects  '
    ' ------------------------------------------------------------------------'
    OP_END = 0 '    -       -                 (return)        no results      '
    OP_RETURN '     U       v_n-v_x(at u)     (return)        returns v_x-v_n '
    
    OP_CALL '       A B     v_n-v_1 f(at a)   r_b-r_1         f(v1,...,v_n)   '
    OP_TAILCALL '   A B     v_n-v_1 f(at a)   (return)        f(v1,...,v_n)   '
    
    OP_PUSHNIL '    U       -                 IDT_Nil_1-IDT_Nil_u             '
    OP_POP '        U       a_u-a_1           -                               '
    
    OP_PUSHINT '    S       -                 (Number)s                       '
    OP_PUSHSTRING ' K       -                 KSTR[k]                         '
    OP_PUSHNUM '    N       -                 KNUM[n]                         '
    OP_PUSHNEGNUM ' N       -                 -KNUM[n]                        '
    
    OP_PUSHUPVALUE 'U       -                 IDT_Closure[u]                  '
    
    OP_GETLOCAL '   L       -                 LOC[l]                          '
    OP_GETGLOBAL '  K       -                 VAR[KSTR[k]]                    '
    
    OP_GETTABLE '   -       i t               t[i]                            '
    OP_GETDOTTED '  K       t                 t[KSTR[k]]                      '
    OP_GETINDEXED ' L       t                 t[LOC[l]]                       '
    OP_PUSHSELF '   K       t                 t t[KSTR[k]]                    '
    
    OP_CREATETABLE 'U       -                 newarray(size = u)              '
    
    OP_SETLOCAL '   L       x                 -               LOC[l]=x        '
    OP_SETGLOBAL '  K       x                 -               VAR[KSTR[k]]=x  '
    OP_SETTABLE '   A B     v a_a-a_1 i t     (pops b values) t[i]=v          '
    
    OP_SETLIST '    A B     v_b-v_1 t         t               t[i+a*FPF]=v_i  '
    OP_SETMAP '     U       v_u k_u - v_1 k_1 t     t         t[k_i]=v_i      '
    
    OP_ADD '        -       y x               x+y                             '
    OP_ADDI '       S       x                 x+s                             '
    OP_SUB '        -       y x               x-y                             '
    OP_MULT '       -       y x               x*y                             '
    OP_DIV '        -       y x               x/y                             '
    OP_POW '        -       y x               x^y                             '
    OP_CONCAT '     U       v_u-v_1           v1..-..v_u                      '
    OP_MINUS '      -       x                 -x                              '
    OP_NOT '        -       x                 (x==nil)? 1 : IDT_Nil           '
    
    OP_JMPNE '      J       y x               -               (x~=y)? PC+=s   '
    OP_JMPEQ '      J       y x               -               (x==y)? PC+=s   '
    OP_JMPLT '      J       y x               -               (x<y)? PC+=s    '
    OP_JMPLE '      J       y x               -               (x<y)? PC+=s    '
    OP_JMPGT '      J       y x               -               (x>y)? PC+=s    '
    OP_JMPGE '      J       y x               -               (x>=y)? PC+=s   '
    
    OP_JMPT '       J       x                 -               (x~=nil)? PC+=s '
    OP_JMPF '       J       x                 -               (x==nil)? PC+=s '
    OP_JMPONT '     J       x                 (x~=nil)? x : - (x~=nil)? PC+=s '
    OP_JMPONF '     J       x                 (x==nil)? x : - (x==nil)? PC+=s '
    OP_JMP '        J       -                 -               PC+=s           '
    
    OP_PUSHNILJMP ' -       -                 IDT_Nil             PC++;       '
    
    OP_FORPREP '    J                                                         '
    OP_FORLOOP '    J                                                         '
    
    OP_LFORPREP '   J                                                         '
    OP_LFORLOOP '   J                                                         '
    
    OP_IDT_Closure '    A B      v_b-v_1           IDT_Closure(KPROTO[a], v_1-v_b)    '
End Enum

' first six bytes
Private Const LUA_Instruction_Mask As Long = 63     ' ie [bits:00011111 00000000 00000000 00000000]
Private Const LUA_BArg_Mask = 511                   ' ie [bits:11111111 00000001 00000000 00000000]
Private Const LUA_Intruction_SArg_Zero = 33554431   ' ie 2^25 - 1 for -33554431 to 33554431 (which is ~2^26 values)

Private Const LUA_Size_OP = 6                       ' ref: LUA Source Code (lopcodes.h and llimits.h)
Private Const LUA_Size_A = 17                       '
Private Const LUA_Size_B = 9                        '

Private Const LUA_Pos_U = LUA_Size_OP               '
Private Const LUA_Pos_S = LUA_Size_OP               '
Private Const LUA_Pos_A = LUA_Size_OP + LUA_Size_B  '
Private Const LUA_Pos_B = LUA_Size_OP               '

Private Const LUA_MaxStackSize = 1024               ' Max. (internal) stack size.
Private Const LUA_NumJumpsStored = 128              ' NR. of jumps stored (may be nested)
Private Const LUA_FieldsPerFlush = 64               ' Fields per flush. Looks like max. stack size...
Private Const LUA_ExtraFields = 8                   ' Some extra no. of fields (for calculating the stack size
                                                    ' while DC'ing LUA).

Private Const LUA_Null = "IDT_Nil"                  ' not "null" but "IDT_Nil"

' Function statements and their search strings (the sstr will be chopped off).
Private Const FuncStatement_SStr = "--<< Position Reserved for Function "
Private Const FuncStatement_SStr_R = " >>--"

' Main body of function statement, this stores name (%n), function pointer (%f), parent's pointer (%p)
' [all references to pointers internal and may not neccessarily be accurate...]
Private Const FuncStatement = FuncStatement_SStr & "%n, &f, &p" & FuncStatement_SStr_R

Private Enum interpretationDataType
    IDT_Nil         ' null has it's own type.
    IDT_Integral    ' \ This would help to
    IDT_Float       ' / trace origin of value.
    IDT_Char
    IDT_Table       ' a IDT_Table.
    IDT_LocalVar    ' local var.
    IDT_Closure     ' function.
End Enum

'Private Enum jumpTypes
'    JT_Invalid                                  ' uninitialized...
'
'    JT_Conditional                              ' OP_JMPNE <= JT_Conditional <= OP_JMPONF
'    JT_Unconditional                            ' OP_JMP = JT_Unconditional
'    JT_While                                    ' OP_JMPNE <= JT_Conditional <= OP_JMPONF
'    JT_For                                      ' OP_FORPREP <= JT_For <= OP_FORLOOP
'End Enum

Private Enum interpretationDataFlags
    IDF_IsALocalValue = 1
    
    IDF_FunctionReturn = 2
    IDF_FunctionReturnWithEQ = 4
End Enum

Private Type interpretationStack                ' stack Type
    value As String
    type As interpretationDataType
    flags As interpretationDataFlags
    
    extraValue As Long
    extraString As String
End Type

'Private Type jumpRegister                       ' for storing jump infos
'    type As jumpTypes                           ' type.
'    condition As String                         ' actual string condition
'
'    jmpT As Long                                ' where to jump if true. Obviously pos+1
'    jmpF As Long                                ' where to jump if false.
'    jmpE As Long                                ' where to jump after this jump is proved true
'End Type

Private localProcessed() As Boolean             ' has this local been assigned value?

Private stack() As interpretationStack          ' stack data
'Private jumpR(1 To LUA_NumJumpsStored) As jumpRegister

Private stackP As Long
'Private jumpRP As Long

Private currIns As Long                         ' current instruction
Private level As Long                           ' tab level

' Returns the OP Code of a given LUA 32-bit instruction.
Private Function Instruction_GetOPCode(ByVal instruction As Long) As LUA_OPCodes
    Instruction_GetOPCode = (instruction And LUA_Instruction_Mask)
End Function

' Returns U (unsigned) from a instruction.
Private Function Instruction_GetUArg(ByVal instruction As Long)
    Instruction_GetUArg = ShiftRight(instruction, LUA_Pos_U)
End Function

' Returns S (signed) from a instruction.
Private Function Instruction_GetSArg(ByVal instruction As Long)
    Instruction_GetSArg = ShiftRight(instruction, LUA_Pos_S) - LUA_Intruction_SArg_Zero
End Function

' Returns A (1st argument of 17 bits in the upper bits) from a instruction.
Private Function Instruction_GetAArg(ByVal instruction As Long)
    Instruction_GetAArg = ShiftRight(instruction, LUA_Pos_A)
End Function

' Returns B (2nd argument of 9 bits in the middle bits) from a instruction.
Private Function Instruction_GetBArg(ByVal instruction As Long)
    Instruction_GetBArg = ShiftRight(instruction, LUA_Pos_B) And LUA_BArg_Mask
End Function

' Reverses the given condition (ie >= is chaned to <, and please, conditional jump opcodes only)
Private Function ReverseCondition_OP(ByVal condition As LUA_OPCodes) As LUA_OPCodes
    ReverseCondition_OP = condition
    
    Select Case ReverseCondition_OP
        Case LUA_OPCodes.OP_JMPEQ
            ReverseCondition_OP = OP_JMPNE
        Case LUA_OPCodes.OP_JMPNE
            ReverseCondition_OP = OP_JMPEQ
        Case LUA_OPCodes.OP_JMPLT
            ReverseCondition_OP = OP_JMPGE
        Case LUA_OPCodes.OP_JMPGT
            ReverseCondition_OP = OP_JMPLE
        Case LUA_OPCodes.OP_JMPLE
            ReverseCondition_OP = OP_JMPGT
        Case LUA_OPCodes.OP_JMPGE
            ReverseCondition_OP = OP_JMPLT
    End Select  '   Select Case ReverseCondition_OP
End Function

' Returns the index for the FIRST local with matching intruction number (startpc).
Private Function FindMatchingStartPCForLocal(ByRef luaChunk As LUA_Chunk, ByRef find As Long) As Long
    Dim I As Long
    
    FindMatchingStartPCForLocal = 0
    
    For I = 1 To luaChunk.numLocals
        If luaChunk.locals(I).startpc = find Then
            FindMatchingStartPCForLocal = I
            Exit For
        End If  '   If luaChunk.locals(I).startpc = find Then
    Next I  '   For I = 1 To luaChunk.numLocals
End Function

' Returns the index for the LAST local with matching intruction number (startpc)
Private Function FindMatchingStartPCForLastLocal(ByRef luaChunk As LUA_Chunk, ByRef find As Long) As Long
    Dim I As Long
    
    FindMatchingStartPCForLastLocal = -1    ' So that a loop out of this is never executed...
    
    For I = 1 To luaChunk.numLocals
        If luaChunk.locals(I).startpc = find Then
            FindMatchingStartPCForLastLocal = I
        End If  '   If luaChunk.locals(I).startpc = find Then
    Next I  '   For I = 1 To luaChunk.numLocals
End Function

' Returns the nr. of locals which end at the given instruction.
Private Function FindNumLocalsWithEndPC(ByRef luaChunk As LUA_Chunk, ByRef find As Long) As Long
    Dim I As Long
    
    For I = 1 To luaChunk.numLocals
        If luaChunk.locals(I).endpc = find Then _
            FindNumLocalsWithEndPC = FindNumLocalsWithEndPC + 1
    Next I  '   For I = 1 To luaChunk.numLocals
End Function

' Returns the nr. of locals which start at the given instruction.
Private Function FindNumLocalsWithStartPC(ByRef luaChunk As LUA_Chunk, ByRef find As Long) As Long
    Dim I As Long
    
    For I = 1 To luaChunk.numLocals
        If luaChunk.locals(I).startpc = find Then _
            FindNumLocalsWithStartPC = FindNumLocalsWithStartPC + 1
    Next I  '   For I = 1 To luaChunk.numLocals
End Function

' Finds the line info for given instruction number and lua chunk.
Private Function FindLineInfo(ByRef luaChunk As LUA_Chunk, ByVal ins As Long) As Long
#If UseLineInfo Then
    ' Function converted from ANSI C to BASIC (ref: "\luac\print.c" or "\ldebug.c")
    Dim refLine As Long, refi As Long
    Dim nextLine As Long, nextRef As Long
    Dim lineInfo
    
    refLine = 1
    refi = 1
    
    If (ins < 1) Or (ins > luaChunk.numInstructions) Or (luaChunk.numLineInfo = 0) Then _
        FindLineInfo = -1: _
        Exit Function _
    Else _
        lineInfo = luaChunk.lineInfo
    
    ins = ins - 1
    
    If lineInfo(refi) < 0 Then _
        refLine = refLine - lineInfo(refi):
        refi = refi + 1
    
    Debug.Assert lineInfo(refi) >= 0
    
    Do While lineInfo(refi) > ins
        refLine = refLine - 1
        refi = refi - 1
        
        If lineInfo(refi) < 0 Then _
            refLine = refLine + lineInfo(refi): _
            refi = refi - 1
        
        Debug.Assert lineInfo(refi) >= 0
    Loop    '   Do While lineInfo(refi) > ins
    
    Debug.Assert lineInfo(refi) >= 0
    
    Do
        nextLine = refLine + 1
        nextRef = refi + 1
        
        If lineInfo(nextRef) < 0 Then _
            nextLine = nextLine - lineInfo(nextRef): _
            nextRef = nextRef + 1
        
        Debug.Assert lineInfo(nextRef) >= 0
        
        If lineInfo(nextRef) > ins Then _
            Exit Do
        
        refLine = nextLine
        refi = nextRef
    Loop
    
    FindLineInfo = refLine
#Else
    FindLineInfo = -1
#End If
End Function

' Finds the previous condition for the 'if' line (uses line info if enabled)
Private Function FindPreviousCondition(ByVal currIns As Long, ByVal ins As Long, ByRef luaChunk As LUA_Chunk, Optional ByVal includeJMP As Boolean = True, Optional ByVal ingnoreLineInfo As Boolean = False, Optional ByVal bruteForce As Boolean = False) As Long
    Dim opCode As LUA_OPCodes
    Dim I As Long
    
    FindPreviousCondition = 0
    
    ' Try to get the previous IDT_Nil-comparision jump... if possible.
    If ins - 2 > 0 Then
        opCode = Instruction_GetOPCode(luaChunk.instructions(ins - 2))
        
        ' Determine whether the instruction is a jump
        If (OP_JMPNE <= opCode) And (opCode <= OP_JMP + IIf(includeJMP, 0, -1)) Then
            If (FindLineInfo(luaChunk, ins - 2) = FindLineInfo(luaChunk, ins)) Or ingnoreLineInfo Then
                FindPreviousCondition = ins - 2
                Exit Function
            End If  '   If (FindLineInfo(luaChunk, ins - 2) = FindLineInfo(luaChunk, ins)) Or ingnoreLineInfo Then
        End If  '   If (OP_JMPNE <= opCode) And (opCode <= OP_JMP + IIf(includeJMP, 0, -1)) Then
    End If  '   If ins - 2 > 0 Then
    
    ' Try to get the previous two-comparision jump... if possible.
    If ins - 3 > 0 Then
        opCode = Instruction_GetOPCode(luaChunk.instructions(ins - 3))
        
        ' Determine whether the previous instruction is a jump
        If (OP_JMPNE <= opCode) And (opCode <= OP_JMP + IIf(includeJMP, 0, -1)) Then
            If (FindLineInfo(luaChunk, ins - 3) = FindLineInfo(luaChunk, ins)) Or ingnoreLineInfo Then
                FindPreviousCondition = ins - 3
                Exit Function
            End If  '   If (FindLineInfo(luaChunk, ins - 3) = FindLineInfo(luaChunk, ins)) Or ingnoreLineInfo Then
        End If  '   If (OP_JMPNE <= opCode) And (opCode <= OP_JMP + IIf(includeJMP, 0, -1)) Then
    End If  '   If ins - 3 > 0 Then
    
    If bruteForce Then
        For I = ins - 1 To 1 Step -1
            opCode = Instruction_GetOPCode(luaChunk.instructions(I))
            
            ' Determine whether the previous instruction is a jump
            If (OP_JMPNE <= opCode) And (opCode <= OP_JMP + IIf(includeJMP, 0, -1)) Then
                If (FindLineInfo(luaChunk, I) = FindLineInfo(luaChunk, ins)) Or ingnoreLineInfo Then
                    FindPreviousCondition = I
                    Exit Function
                End If  '   If (FindLineInfo(luaChunk, I) = FindLineInfo(luaChunk, ins)) Or ingnoreLineInfo Then
            End If  '   If (OP_JMPNE <= opCode) And (opCode <= OP_JMP + IIf(includeJMP, 0, -1)) Then
        Next I  '   For I = ins - 1 To 1 Step -1
    End If  '   If bruteForce Then
End Function

' Finds the next condition for the 'if' line (uses line info if enabled)
Private Function FindNextCondition(ByVal currIns As Long, ByVal ins As Long, ByRef luaChunk As LUA_Chunk, Optional ByVal includeJMP As Boolean = True, Optional ByVal overrideLineInfo As Boolean = False) As Long
    Dim opCode As LUA_OPCodes
    
    FindNextCondition = 0
    
    ' Try to get the next IDT_Nil-comparision jump... if possible.
    If luaChunk.numInstructions > 2 + ins Then
        opCode = Instruction_GetOPCode(luaChunk.instructions(2 + ins))
        
        ' Determine whether the next instruction is a jump
        If (OP_JMPNE <= opCode) And (opCode <= OP_JMP + IIf(includeJMP, 0, -1)) Then
            If (FindLineInfo(luaChunk, ins + 2) = FindLineInfo(luaChunk, ins)) Or overrideLineInfo Then
                FindNextCondition = 2 + ins
                Exit Function
            End If  '   If (FindLineInfo(luaChunk, ins + 2) = FindLineInfo(luaChunk, ins)) Or overrideLineInfo Then
        End If  '   If (OP_JMPNE <= opCode) And (opCode <= OP_JMP + IIf(includeJMP, 0, -1)) Then
    End If  '   If luaChunk.numInstructions > 2 + ins Then
    
    ' Try to get the next two-comparision jump... if possible.
    If luaChunk.numInstructions > 3 + ins Then
        opCode = Instruction_GetOPCode(luaChunk.instructions(3 + ins))
        
        ' Determine whether the instruction is a jump
        If (OP_JMPNE <= opCode) And (opCode <= OP_JMP + IIf(includeJMP, 0, -1)) Then
            If (FindLineInfo(luaChunk, ins + 3) = FindLineInfo(luaChunk, ins)) Or overrideLineInfo Then
                FindNextCondition = 3 + ins
                Exit Function
            End If  '   If (FindLineInfo(luaChunk, ins + 3) = FindLineInfo(luaChunk, ins)) Or overrideLineInfo Then
        End If  '   If (OP_JMPNE <= opCode) And (opCode <= OP_JMP + IIf(includeJMP, 0, -1)) Then
    End If  '   If luaChunk.numInstructions > 3 + ins Then
    
    opCode = Instruction_GetOPCode(luaChunk.instructions(ins + Instruction_GetSArg(currIns)))
    
    ' Determine whether the instruction is a jump
    If (OP_JMPNE <= opCode) And (opCode <= OP_JMP + IIf(includeJMP, 0, -1)) Then
        If (FindLineInfo(luaChunk, ins + Instruction_GetSArg(currIns)) = FindLineInfo(luaChunk, ins)) Or overrideLineInfo Then
            FindNextCondition = ins + Instruction_GetSArg(currIns)
        End If  '   If (FindLineInfo(luaChunk, ins + Instruction_GetSArg(currIns)) = FindLineInfo(luaChunk, ins)) Or overrideLineInfo Then
    End If  '   If (OP_JMPNE <= opCode) And (opCode <= OP_JMP + IIf(includeJMP, 0, -1)) Then
End Function

' Finds the next condition for the 'if' line (uses line info if enabled) (preference to JMP value)
Private Function FindNextCondition_PrefJMP(ByVal currIns As Long, ByVal ins As Long, ByRef luaChunk As LUA_Chunk, Optional ByVal includeJMP As Boolean = True) As Long
    Dim opCode As LUA_OPCodes
    
    FindNextCondition_PrefJMP = 0
    
    opCode = Instruction_GetOPCode(luaChunk.instructions(ins + Instruction_GetSArg(currIns)))
    
    ' Determine whether the instruction is a jump
    If (OP_JMPNE <= opCode) And (opCode <= OP_JMP + IIf(includeJMP, 0, -1)) Then
        If FindLineInfo(luaChunk, Instruction_GetSArg(currIns) + ins) = FindLineInfo(luaChunk, ins) Then
            FindNextCondition_PrefJMP = ins + Instruction_GetSArg(currIns)
            Exit Function
        End If  '   If FindLineInfo(luaChunk, Instruction_GetSArg(currIns) + ins) = FindLineInfo(luaChunk, ins) Then
    End If  '   If (OP_JMPNE <= opCode) And (opCode <= OP_JMP + IIf(includeJMP, 0, -1)) Then
    
    ' Try to get the next IDT_Nil-comparision jump... if possible.
    If luaChunk.numInstructions > 2 + ins Then
        opCode = Instruction_GetOPCode(luaChunk.instructions(2 + ins))
        
        ' Determine whether the next instruction is a jump
        If (OP_JMPNE <= opCode) And (opCode <= OP_JMP + IIf(includeJMP, 0, -1)) Then
            If FindLineInfo(luaChunk, ins + 2) = FindLineInfo(luaChunk, ins) Then
                FindNextCondition_PrefJMP = 2 + ins
                Exit Function
            End If  '   If FindLineInfo(luaChunk, ins + 2) = FindLineInfo(luaChunk, ins) Then
        End If  '   If (OP_JMPNE <= opCode) And (opCode <= OP_JMP + IIf(includeJMP, 0, -1)) Then
    End If  '   If luaChunk.numInstructions > 2 + ins Then
    
    ' Try to get the next two-comparision jump... if possible.
    If luaChunk.numInstructions > 3 + ins Then
        opCode = Instruction_GetOPCode(luaChunk.instructions(3 + ins))
        
        ' Determine whether the instruction is a jump
        If (OP_JMPNE <= opCode) And (opCode <= OP_JMP + IIf(includeJMP, 0, -1)) Then
            If FindLineInfo(luaChunk, ins + Instruction_GetSArg(currIns)) = FindLineInfo(luaChunk, ins) Then
                FindNextCondition_PrefJMP = 3 + ins
            End If  '   If FindLineInfo(luaChunk, ins + Instruction_GetSArg(currIns)) = FindLineInfo(luaChunk, ins) Then
        End If  '   If (OP_JMPNE <= opCode) And (opCode <= OP_JMP + IIf(includeJMP, 0, -1)) Then
    End If  '   If luaChunk.numInstructions > 3 + ins Then
End Function

' Finds the end of this 'if' line (uses line info if enabled)
Private Function FindEndOfJump(ByVal currIns As Long, ByVal ins As Long, ByRef luaChunk As LUA_Chunk) As Long
    Dim lastValue As Long
    
    lastValue = ins ' Don't forget to initialize me!
    FindEndOfJump = FindNextCondition_PrefJMP(currIns, ins, luaChunk)
    
    Do While Not FindEndOfJump = 0
        lastValue = FindEndOfJump
        FindEndOfJump = FindNextCondition_PrefJMP(luaChunk.instructions(lastValue), lastValue, luaChunk)
        
        ' Safety catch: Jumps with 0 pc offset!
        If lastValue = FindEndOfJump Then Exit Do
    Loop    '   Do While Not FindEndOfJump = 0
    
    FindEndOfJump = Instruction_GetSArg(luaChunk.instructions(lastValue)) + lastValue + 1
End Function

' Tries to findout whether the given block (use last condition only!) is a 'elseif' or an 'if' block.
Private Function FindELSEIFPresence(ByVal ins As Long, ByRef luaChunk As LUA_Chunk) As Boolean
    Dim I As Long, J As Long
    Dim jmpPrev As Long, jmpNext As Long, jmpCurr As Long
    
    ' For an 'elseif', either
    '      i) we have the previous unconditional jump pointing to instruction which this instruction does,
    ' or, ii) we have two unconditional jumps, both pointing to same instruction.
    ' Note: The previous jump will always be 1 'line' behind. Not more, not less (it's true, think on it).
    
    ' First try using the jump register.
    'If jumpRP < 2 Then Exit Function
    '
    'For I = 2 To jumpRP
    '    If jumpR(I).jmpE = jumpR(I - 1).jmpE Then
    '        FindELSEIFPresence = True
    '        Exit Function
    '    End If  '   If jumpR(I).jmpE = jumpR(I - 1).jmpE Then
    'Next I  '   For I = 1 To jumpRP
    '
    ' Then manually check, if jump register doesn't reference it...
    'Debug.Assert 0      ' should never come here in the first place (if we have a last jump, ofcourse).
    
GetCurrJmp:
    jmpCurr = 1 + ins + Instruction_GetSArg(luaChunk.instructions(ins))
    
GetPrevJump:
    ' Try to find the previous condition on the same line.
    I = FindPreviousCondition(luaChunk.instructions(ins), ins, luaChunk, , True, True)
    If I = 0 Then _
        FindELSEIFPresence = False: _
        Exit Function
    
    ' Get the unconditional jump!
    Do While Not I = 0
        If Instruction_GetOPCode(luaChunk.instructions(I)) = OP_JMP Then
            ' Compare it's lineinfo.
            If FindLineInfo(luaChunk, I) = FindLineInfo(luaChunk, ins) - 1 Then
                jmpPrev = 1 + I + Instruction_GetSArg(luaChunk.instructions(I))
                Exit Do
            End If  '   If FindLineInfo(luaChunk, I) = FindLineInfo(luaChunk, ins) - 1 Then
        End If  '   If Instruction_GetOPCode(luaChunk.instructions(I)) = OP_JMP Then
        
        J = I
        I = FindPreviousCondition(luaChunk.instructions(J), J, luaChunk, , True)
        
        If I = J Then _
            Exit Do
    Loop    '   Do While Not I = 0
    I = J
    
    ' If jmpPrev = 0, then this is no 'elseif' block...
    If jmpPrev = 0 Then _
        FindELSEIFPresence = False: _
        Exit Function
    
    If jmpCurr = jmpPrev Then _
        FindELSEIFPresence = True: _
        Exit Function
    
    ' Try to get the next condition (no lineinfos).
    I = FindNextCondition(luaChunk.instructions(ins), ins, luaChunk, , True)
    
GetNextJump:
    If I = 0 Then
        I = FindNextCondition(luaChunk.instructions(I), I, luaChunk, , True)
        
        If Instruction_GetOPCode(I) = OP_JMP Then
            J = 1 + I + Instruction_GetSArg(luaChunk.instructions(I))
            
            If jmpCurr = J Then _
                jmpNext = J: _
                GoTo Ending
        Else    '   If Instruction_GetOPCode(I) = OP_JMP Then
            GoTo GetNextJump
        End If  '   If Instruction_GetOPCode(I) = OP_JMP Then
    Else    '   If I = 0 Then
        jmpNext = 1 + I + Instruction_GetSArg(luaChunk.instructions(I))
    End If  '   If I = 0 Then
    
Ending:
    If (jmpCurr = jmpNext) Or (jmpPrev = jmpNext) Then _
        FindELSEIFPresence = True: _
        Exit Function _
    Else _
        FindELSEIFPresence = False: _
        Exit Function
End Function

' Tries to find out whether 'else' should be written at this JMP or not.
Private Function FindELSERequirement(ByVal ins As Long, ByRef luaChunk As LUA_Chunk) As Boolean
    Dim I As Long, J As Long, K As Long
    
    ' There is a need of writing 'else', when the next jump instruction on the next line is present.
    ' (and we are NOT in a 'while' loop)
    FindELSERequirement = True
    
    ' If this is part of a while loop, then no 'else' is needed.
    If Instruction_GetSArg(luaChunk.instructions(ins)) < 0 Then FindELSERequirement = False: Exit Function
    
    For I = ins + 1 To luaChunk.numInstructions
        J = Instruction_GetOPCode(luaChunk.instructions(I))
        
        If (OP_JMPNE <= J) And (J <= OP_JMPONF) Then _
            If FindLineInfo(luaChunk, I) - 1 = FindLineInfo(luaChunk, ins) Then _
                Exit For
    Next I  '   For I = ins + 1 To luaChunk.numInstructions
    
    If I = luaChunk.numInstructions + 1 Then _
        Exit Function
    
    FindELSERequirement = False
End Function

' Tries to find out whether we need an 'end' at this instruction or not.
Private Function FindENDRequirement(ByVal ins As Long, ByRef luaChunk As LUA_Chunk) As Long
    Dim I As Long
    
    ' We always need an 'end' on the last instruction, whenever applicable.
    If ins = luaChunk.numInstructions Then _
        FindENDRequirement = ins: _
        Exit Function
    
    ' There is a need for 'end', when two consecutive jumps (no relation betw. line infos) send jump to
    ' the same pos (this happens in case of nested 'if' blocks ending at same instruction).
    For I = ins + 1 To luaChunk.numInstructions
        If (OP_JMPNE <= Instruction_GetOPCode(luaChunk.instructions(I))) And _
            (Instruction_GetOPCode(luaChunk.instructions(I)) <= OP_JMPONF) And _
                Not Instruction_GetOPCode(luaChunk.instructions(I)) = OP_JMP Then _
                    I = luaChunk.numInstructions + 1: _
                    Exit For
        
        If Instruction_GetOPCode(luaChunk.instructions(I)) = OP_JMP Then _
            If Instruction_GetSArg(luaChunk.instructions(I)) + I = Instruction_GetSArg(luaChunk.instructions(ins)) + ins Then _
                Exit For
    Next I  '   For I = ins + 1 To luaChunk.numInstructions
    
    If I = luaChunk.numInstructions + 1 Then _
        Exit Function
        
    FindENDRequirement = I
End Function

' Determines presence of 'or' or 'and' between this jump instruction and next jump instruction.
Private Function FindORPresence(ByVal instructionNr As Long, ByRef luaChunk As LUA_Chunk) As Boolean
    Dim I As Long, J As Long
    Dim ins() As Long, nIns() As Long, jmpT() As Long, jmpF() As Long, swaped() As Boolean
    Dim tIns As Long, tNIns As Long, tJmpT As Long, tJmpF As Long
    Dim cIPos As Long
    
    ' For an 'or', the only known condition is (due to their nature),
    ' that: in the IDT_Table of conditions, if T and F jump instructions are in a descending order, and
    ' if one set doesn't match with the instruction, then an 'and' is present for that instruction.
    
    ' FIRST CACHE CURRENT STATE OF INSTRUCTIONS...
    ReDim ins(0), nIns(0), jmpT(0), jmpF(0), swaped(0)
    
    tNIns = instructionNr
    tIns = luaChunk.instructions(tNIns)
    tJmpT = Instruction_GetSArg(tIns) + tNIns + 1
    tJmpF = tNIns + 1
    
    ' Incase this is the very first instruction...
    nIns(0) = tNIns
    ins(0) = tIns
    jmpT(0) = tJmpT
    jmpF(0) = tJmpF
    
    ' Get the number of instructions, for 'if' before this...
    J = FindPreviousCondition(tIns, tNIns, luaChunk, False)
    
    Do While Not J = 0
        nIns(I) = J
        ins(I) = luaChunk.instructions(J)
        jmpT(I) = Instruction_GetSArg(ins(I)) + J + 1
        jmpF(I) = J + 1
        
        I = I + 1
        J = FindPreviousCondition(ins(I - 1), nIns(I - 1), luaChunk, False)
        
        ReDim Preserve ins(I), nIns(I), jmpT(I), jmpF(I), swaped(I)
    Loop    '   Do While Not J = 0
    
    ' But the instructions we have are inverted!
    For J = 0 To (I - 1) \ 2
        If (Not I - J - 1 = J) And (Not I - J - 1 < 0) Then
            ' So invert them, we'll instructions in the correct order.
            swap nIns(J), nIns(I - J - 1)
            swap ins(J), ins(I - J - 1)
            swap jmpT(J), jmpT(I - J - 1)
            swap jmpF(J), jmpF(I - J - 1)
        End If  '   If (Not I - J - 1 = J) And (Not I - J - 1 < 0) Then
    Next J  '   For J = 0 To (I - 1) \ 2
    
    ' This instruction...
    cIPos = I
    
    nIns(I) = tNIns
    ins(I) = tIns
    jmpT(I) = tJmpT
    jmpF(I) = tJmpF
    
    '  ... and the instructions after this...
    J = FindNextCondition(ins(I), nIns(I), luaChunk, False)
    If J = instructionNr Then J = 0
    
    ' If J = 0 (ie NO conditions after this) then this is an 'and' (and an end too, no doubt!),
    ' else, go ahead, caching...
    If J Then _
        I = I + 1: _
        ReDim Preserve ins(I), nIns(I), jmpT(I), jmpF(I), swaped(I) _
    Else _
        FindORPresence = False: _
        Exit Function
    
    Do While Not J = 0
        nIns(I) = J
        ins(I) = luaChunk.instructions(J)
        jmpT(I) = Instruction_GetSArg(ins(I)) + J + 1
        jmpF(I) = J + 1
        
        J = FindNextCondition(ins(I), nIns(I), luaChunk, False)
        
        If J Then _
            I = I + 1: _
            ReDim Preserve ins(I), nIns(I), jmpT(I), jmpF(I), swaped(I) _
        Else _
            Exit Do
    Loop    '   Do While Not J = 0
    
#If PrintStack Then
    For J = 0 To UBound(ins())
        Debug.Print nIns(J), Instruction_GetOPCode(ins(J)), jmpT(J), jmpF(J), IIf(J = cIPos, "<<", "")
    Next J
#End If
    
    ' THEN, SWAP AND STORE AS NEEDED...
    ' first swap the last one... if we have only two conditions, then this could get messy if we dont...
    swap jmpT(I), jmpF(I)
    swaped(I) = True
    
    ' swap the first one, if needed.
    If (jmpT(0) = jmpF(1)) Or (jmpT(1) = jmpF(0)) Then _
        swap jmpT(0), jmpF(0): _
        swaped(0) = True
    
    ' swap others...
    For J = 1 To I - 1
        If (jmpT(J) = jmpF(J - 1)) Or (jmpT(J - 1) = jmpF(J)) Then _
            swap jmpT(J), jmpF(J): _
            swaped(J) = True
    Next J  '   For J = 1 To I - 1
    
    ' finally swap the last one (last is always and, like, 'x and true' where x = <condition>.
    swap jmpT(J), jmpF(J)
    swaped(J) = True
    
    ' FINALLY, IF THE CONCERNED VARIABLE WAS SWAPPED, THEN 'or' IS NOT PRESENT AFTER IT,
    ' ELSE 'or' IS PRESENT.
    FindORPresence = Not swaped(cIPos)
End Function

' Gets the LUA function with this pointer.
Private Function FindLUAFunction(ByVal funcPtr As Long, ByRef theLUA As LUA_File) As LUA_Chunk
    Dim I As Long
    
    For I = 1 To UBound(theLUA.chunks())
        If theLUA.chunks(I).funcPtr = funcPtr Then _
            FindLUAFunction = theLUA.chunks(I): _
            Exit Function
    Next I  '   For I = 1 To UBound(theLUA.chunks())
    
    For I = 1 To UBound(theLUA.funcs())
        If theLUA.funcs(I).funcPtr = 0 Then _
            Exit For
            
        If theLUA.funcs(I).funcPtr = funcPtr Then _
            FindLUAFunction = theLUA.funcs(I): _
            Exit Function
    Next I  '   For I = 1 To UBound(theLUA.funcs())
    
    ' We don't 'know' this function.
    If Not funcPtr = 0 Then _
        Debug.Assert 0
End Function

' Gets the tab for the function with the given pointer.
Private Function FindLUAFunctionTab(ByVal funcPtr As Long, ByRef theLUA As LUA_File) As Long
    Dim theFunc As LUA_Chunk
    
    ' First get this function
    theFunc = FindLUAFunction(funcPtr, theLUA)
    
    Do While Not theFunc.parentPtr = 0
        ' For each valid parent, we have an extra tab...
        FindLUAFunctionTab = FindLUAFunctionTab + 1
        
        ' Get this function's parent.
        If Not theFunc.parentPtr = 0 Then _
            theFunc = FindLUAFunction(theFunc.parentPtr, theLUA)
    Loop    '   Do While Not theFunc.parentPtr = 0
End Function

' Processes the given value for output, for example:
' - In a IDT_Table, if it's raw (no data, just value indicating number of entries), then IDT_Nil's are written.
' - For a stack type 'IDT_Nil', 'IDT_Nil' is written out, even if it's not so in the stack value (for safety).
Private Function ProcessValueForOutput(Optional ByVal offset As Long = 0) As String
    Dim I As Long
    
    Select Case stack(stackP + offset).type
        Case interpretationDataType.IDT_Table
            If IsNumeric(stack(stackP + offset).value) Then
                ' Empty IDT_Table.
                If CLng(stack(stackP + offset).value) = 0 Then
                    ' No data
                    ProcessValueForOutput = "{}"
                Else    '   If stack(stackP + offset).value = 0 Then
                    ' Data, but not defined yet...
                    ' fill with Null's (IDT_Nil's)
                    ProcessValueForOutput = "{"
                    
                    For I = 1 To stack(stackP + offset).value
                        ProcessValueForOutput = ProcessValueForOutput & _
                            LUA_Null & _
                            IIf(Not I = stack(stackP + offset).value, ", ", "")
                    Next I  '   For I = 1 To stack(stackP + offset).value
                    
                    ProcessValueForOutput = ProcessValueForOutput & _
                        "}"
                End If  '   If stack(stackP + offset).value = 0 Then
            Else    '   If IsNumeric(stack(stackP + offset).value) Then
                ' Filled IDT_Table.
                ProcessValueForOutput = stack(stackP + offset).value
                ProcessValueForOutput = EscapeBacklashesInString(ProcessValueForOutput)
            End If  '   If IsNumeric(stack(stackP + offset).value) Then
        Case interpretationDataType.IDT_Nil
            ProcessValueForOutput = LUA_Null
            Debug.Assert stack(stackP + offset).value = LUA_Null
        Case Else
            ProcessValueForOutput = IIf(InStr(1, stack(stackP + offset).value, "[[") > 0, _
                    " ", _
                    "" _
                ) & _
                stack(stackP + offset).value
            
            ProcessValueForOutput = EscapeBacklashesInString(ProcessValueForOutput)
    End Select  '   Select Case stack(stackP + offset).type
End Function

Private Function ProcessTableValue(ByRef tbl As String, ByRef key As String) As String
    Dim dotted As Boolean
    
#If 0 Then
    dotted = (NumTokens(key, LUA_Tokens, " ") = 1)
    dotted = dotted And (Not IsNumeric(key))
    dotted = dotted And ((InStr(1, key, Chr$(34)) = 1) Or (InStr(1, key, "[[") = 1))
    
    If ((InStr(1, key, Chr$(34)) = 1) Or (InStr(1, key, "[[") = 1)) Then _
        dotted = dotted And (NumTokens(UnquoteString(key), LUA_Tokens, " ") = 1)
#Else
    dotted = False
#End If

    If dotted Then
        ProcessTableValue = tbl & _
            "." & _
            (key)
    Else    '   If dotted Then
        ProcessTableValue = tbl & _
            "[" & _
            key & _
            "]"
    End If  '   If dotted Then
End Function


' Processes current line for "output", ie for ";" and CR-LF.
Private Sub PushStatement(ByRef out As String, ByRef statement As String, Optional ByVal semiColon As Boolean = True, Optional crlf As Boolean = True)
    Dim tmpStr As String
    tmpStr = statement
    
    ' Remove multiple spaces.
    Do While InStr(1, tmpStr, "  ")
        tmpStr = Replace(tmpStr, "  ", " ")
    Loop

    ' Dump statement here.
    out = out & _
        String(max(0, level), Chr$(9)) & _
        LTrim( _
            RTrim(tmpStr) _
        ) & _
        IIf(Left(statement, 2) = "--", "", IIf(semiColon, IIf(Len(statement) > 0, ";", ""), "")) & _
        IIf(crlf, vbCrLf, "")
End Sub

' Pushes the given jump into the register
Private Function GetLocalInStack(Optional ByRef name As String) As Long
    GetLocalInStack = 1
    
    Do While BitsPresent(stack(GetLocalInStack).flags, IDF_IsALocalValue) Or (stack(GetLocalInStack).type = IDT_LocalVar)
        ' Is this the one?
        If Not Len(name) = 0 Then _
            If stack(GetLocalInStack).value = name Or stack(GetLocalInStack).extraString = name Then _
                Exit Function
        
        ' Next!
        GetLocalInStack = GetLocalInStack + 1
    Loop    '   Do While BitsPresent(stack(GetLocalInStack).flags, IDF_IsALocalValue) Or (stack(GetLocalInStack).type = IDT_LocalVar)
    
    GetLocalInStack = GetLocalInStack - 1
End Function

' Method for "Push"-ing values into the stack.
' First search for a viable local, and
'  i) if found, assign value to it (change to stack, 1st value is popped, and this is pushed).
' ii) if not found, push value in stack.
' If we're return value for a function, then set this IN for return.
' Otherwise, if FORCED to push value, then do so.
Private Sub PushValueInStack(ByRef value As String, ByVal dType As interpretationDataType, ByVal instructionNr As Long, ByRef luaChunk As LUA_Chunk, ByRef outLUA As String, Optional ByVal searchForLocal As Boolean = True)
    Dim I As Long, J As Long
    Dim tmpStr As String
    
    If (searchForLocal) Then
        ' Try to search for a feasible (ie unused, here) local with 'active' state from here.
        For J = FindMatchingStartPCForLocal(luaChunk, instructionNr) To FindMatchingStartPCForLastLocal(luaChunk, instructionNr)
            If Not localProcessed(J) Then
                PopLocalFromStack luaChunk.locals(J).name
                I = PushLocalInStack(value, dType)
                
                stack(I).extraString = luaChunk.locals(J).name
                stack(I).flags = IDF_IsALocalValue
                
                tmpStr = _
                    "local " & _
                    luaChunk.locals(J).name & _
                    "=" & _
                    ProcessValueForOutput
                
                ' mark this local done.
                localProcessed(J) = True
                
                ' Don't forget this!
                PushStatement outLUA, tmpStr
                
                Exit Sub
            End If  '   If Not localProcessed(J) Then
        Next J  '   For J = FindMatchingStartPCForLocal(luaChunk, instructionNr) To FindMatchingStartPCForLastLocal(luaChunk, instructionNr)
    End If  '   If (searchForLocal) Then
    
    stackP = stackP + 1
    
    stack(stackP).value = value
    stack(stackP).type = dType
End Sub

' Method for "Push"-ing values into the stack on the TOP.
Private Function PushValueInStackTop(ByRef value As String, ByVal dType As interpretationDataType) As Long
    Dim I As Long
    
    stackP = stackP + 1
    
    For I = stackP - 1 To 1 Step -1
        If Not (stack(I).type = IDT_LocalVar) And Not BitsPresent(stack(I).flags, IDF_IsALocalValue) Then
            stack(I + 1).type = stack(I).type
            stack(I + 1).value = stack(I).value
            
            stack(I + 1).flags = stack(I).flags
            stack(I + 1).extraString = stack(I).extraString
            stack(I + 1).extraValue = stack(I).extraValue
        Else    '   If Not (stack(I).type = IDT_LocalVar) And Not BitsPresent(stack(I).flags, IDF_IsALocalValue) Then
            ' don't move a local var!
            stack(I + 1).type = dType
            stack(I + 1).value = value
            
            stack(I + 1).flags = 0
            stack(I + 1).extraString = ""
            stack(I + 1).extraValue = 0
            
            PushValueInStackTop = I + 1
            
            Exit Function
        End If  '   If Not (stack(I).type = IDT_LocalVar) And Not BitsPresent(stack(I).flags, IDF_IsALocalValue) Then
    Next I  '   For I = stackP - 1 To 1 Step -1
    
    stack(1).type = dType
    stack(1).value = value
    
    stack(1).flags = 0
    stack(1).extraString = ""
    stack(1).extraValue = 0
    
    PushValueInStackTop = 1
End Function

' Pushes a local in the stack.
Private Function PushLocalInStack(ByRef value As String, ByVal dType As interpretationDataFlags) As Long
    PushLocalInStack = PushValueInStackTop(value, dType)
    stack(PushLocalInStack).flags = IDF_IsALocalValue
End Function

' Pushes the given jump into the jump register.
'Private Sub PushJumpInRegister(ByRef luaChunk As LUA_Chunk, ByVal pos As Long, ByVal jType As jumpTypes, Optional ByRef condition As String = "")
'    Dim tmpLong
'
'    jumpRP = jumpRP + 1
'
'    jumpR(jumpRP).type = jType
'    jumpR(jumpRP).condition = condition
'
'    Select Case jType
'        Case JT_Unconditional
'            jumpR(jumpRP).jmpT = Instruction_GetSArg(luaChunk.instructions(pos)) + pos + 1
'            jumpR(jumpRP).jmpF = jumpR(jumpRP).jmpT
'            jumpR(jumpRP).jmpE = jumpR(jumpRP).jmpT
'        Case Else   ' JT_Unconditional, JT_While
'            jumpR(jumpRP).jmpT = pos + 1
'            jumpR(jumpRP).jmpF = Instruction_GetSArg(luaChunk.instructions(pos)) + pos + 1
'
'            tmpLong = Instruction_GetOPCode(luaChunk.instructions(jumpR(jumpRP).jmpF - 1))
'
'            If (OP_JMPNE <= tmpLong) And (tmpLong <= OP_JMPONF) Then
'                jumpR(jumpRP).jmpE = FindEndOfJump(luaChunk.instructions(jumpR(jumpRP).jmpF - 1), jumpR(jumpRP).jmpF - 1, luaChunk)
'            ElseIf OP_JMP = tmpLong Then
'                jumpR(jumpRP).jmpE = Instruction_GetSArg(luaChunk.instructions(jumpR(jumpRP).jmpF - 1)) + _
'                    (jumpR(jumpRP).jmpF - 1) + _
'                    1
'            Else
'                jumpR(jumpRP).jmpE = 0
'            End If
'    End Select  '   Select Case jType
'End Sub

' Method for "Pop"-ing values from the stack.
Private Sub PopValueFromStack(Optional ByVal numValues As Long = 1)
    Dim I As Long
    
    Debug.Assert numValues <= stackP
    
    For I = 1 To numValues
        If stackP = 0 Then Exit For
        
        stack(stackP).value = ""
        stack(stackP).type = IDT_Nil
        
        stack(stackP).flags = 0
        stack(stackP).extraValue = 0
        stack(stackP).extraString = 0
        
        stackP = stackP - 1
    Next I  '   For I = 1 To numValues
End Sub

' Method for "Pop"-ing values from the stack (the TOP ones).
Private Sub PopValueFromStackTop(Optional ByVal numValues As Long = 1)
    Dim I As Long
    
    For I = numValues + 1 To stackP
        stack(I - numValues).type = stack(I).type
        stack(I - numValues).value = stack(I).value
        
        stack(I - numValues).flags = stack(I).flags
        stack(I - numValues).extraString = stack(I).extraString
        stack(I - numValues).extraValue = stack(I).extraValue
    Next I  '   For I = 1 To numValues
    
    PopValueFromStack numValues
    
End Sub

' Pops a local from the stack (bottom-most local, if no name, otherwise the specified local is popped)
Private Sub PopLocalFromStack(Optional ByRef name As String = "")
    Dim I As Long, move As Boolean
    
    For I = 1 To stackP
        If name = "" Then
            If I = stackP Then _
                move = True: _
                Exit For            ' This part is not meant to...
            
            If Not (stack(I + 1).type = IDT_LocalVar Or BitsPresent(stack(I + 1).flags, IDF_IsALocalValue)) Then _
                move = True
        Else    '   If name = "" Then
            If (stack(I).value = name And stack(I).type = IDT_LocalVar) Or (stack(I).extraString = name And BitsPresent(stack(I).flags, IDF_IsALocalValue)) Then _
                move = True
        End If  '   If name = "" Then
        
        If move Then
            If I = stackP Then Exit For
            
            stack(I).type = stack(I + 1).type
            stack(I).value = stack(I + 1).value
            
            stack(I).flags = stack(I + 1).flags
            stack(I).extraString = stack(I + 1).extraString
            stack(I).extraValue = stack(I + 1).extraValue
        End If  '   If move Then
    Next I  '   For I = 1 To stackP - 1
    
    ' Finally pop a value. Use default method (if we moved, or if the only thing present is a local)
    If (move) Or (I = 1) Then _
        PopValueFromStack
End Sub

' Removes the top-most jump (with max. indice) from the register.
'Private Sub PopJumpFromRegister()
'    jumpR(jumpRP).type = 0
'    jumpR(jumpRP).condition = ""
'
'    jumpR(jumpRP).jmpE = 0
'    jumpR(jumpRP).jmpF = 0
'    jumpR(jumpRP).jmpT = 0
'
'    jumpRP = jumpRP - 1
'End Sub

' Prints all stack info.
Private Sub PrintStack()
    Dim I As Byte, typeStr As String
    
    Debug.Print "Instruction OPCode:", Instruction_GetOPCode(currIns)
    Debug.Print " ---- "
    
    If stackP = 0 Then _
        Debug.Print "No stack" _
    Else _
        Debug.Print "Index", "Value", "Type", "Flags", "Extra Value", "Extra String"
    
    For I = 1 To stackP
        Select Case stack(I).type
            Case IDT_Integral
                typeStr = "Integer"
            Case IDT_Float
                typeStr = "Number"
            Case IDT_Char
                typeStr = "String"
            Case IDT_Table
                typeStr = "Table"
            Case IDT_Nil
                typeStr = LUA_Null
            Case IDT_Closure
                typeStr = "Function"
            Case IDT_LocalVar
                typeStr = "Local"
            Case Else
                typeStr = "UNKNWON: " & stack(I).type
        End Select  '   Select Case stack(I).type
        
        Debug.Print I, stack(I).value, typeStr, stack(I).flags, stack(I).extraValue, stack(I).extraString
    Next I  '   For I = 1 To stackP
    
    Debug.Print " ---- "
End Sub

' Prints information in the jump register.
'Private Sub PrintRegister()
'    Dim I As Long
'    Dim jmpType As String, condition As String
'
'    Debug.Print " ---- "
'
'    If jumpRP = 0 Then _
'        Debug.Print "No stack" _
'    Else _
'        Debug.Print "Index", "Type", "Condition", , "Jump (True)", "Jump (False)", "Jump (End)"
'
'    For I = 1 To jumpRP
'        ' if a then
'        condition = Replace$(jumpR(I).condition, "if ", "")
'        condition = Replace$(condition, " then", "")
'
'        ' elseif a then
'        condition = Replace$(condition, "else", "")
'        'condition = Replace$(condition, " then", "")
'
'        ' for a=x,y,s do
'        condition = Replace$(condition, "for ", "")
'        condition = Replace$(condition, " do", "")
'
'        ' for a,b in tbl do
'        'condition = Replace$(condition, "for ", "")
'        condition = Replace$(condition, " in", "")
'        'condition = Replace$(condition, " do", "")
'
'        ' while a do
'        condition = Replace$(condition, "while ", "")
'        'condition = Replace$(condition, " do", "")
'
'        condition = KillSpaces(condition)
'        condition = condition & Space$(14 * 2)
'
'        condition = Left$(condition, 14 * 2 - 1)
'
'        Select Case jumpR(I).type
'            Case JT_Conditional
'                jmpType = "Conditional"
'            Case JT_Unconditional
'                jmpType = "Unconditional"
'            Case JT_For
'                jmpType = "For"
'            Case JT_While
'                jmpType = "While"
'            Case Else
'                jmpType = "UNKNOWN: " & jumpR(I).type
'        End Select  '   Select Case jumpR(I).type
'
'        Debug.Print I, jmpType, condition, jumpR(I).jmpT, jumpR(I).jmpF, jumpR(I).jmpE
'    Next I  '   For I = 1 To jumpRP
'
'    Debug.Print " ---- "
'End Sub

' Prints all info about locals.
Private Sub PrintLocalInfo(ByRef theLuaChunk As LUA_Chunk)
    Dim I As Integer
    
    If theLuaChunk.numLocals = 0 Then _
        Debug.Print "-- ------": _
        Debug.Print "No Locals" _
    Else _
        Debug.Print "-----", "----", "-------", "-----": _
        Debug.Print "Index", "Name", "StartPC", "EndPC"
    
    With theLuaChunk
        For I = 1 To .numLocals
            Debug.Print I, .locals(I).name, .locals(I).startpc, .locals(I).endpc
        Next I  '   For I = 1 To .numLocals
    End With
    
    Debug.Print "-----", "----", "-------", "-----"
End Sub

' Prints all info about functions.
Private Sub PrintFunctionInfo(ByRef theLUA As LUA_File)
    Dim I As Long, J As Long
    
    Debug.Print "LUA Chunks"
    Debug.Print "----------"
    
    Debug.Print "Index", "Pointer", "Parent Pointer"
    For I = 1 To UBound(theLUA.chunks())
        Debug.Print I, theLUA.chunks(I).funcPtr, theLUA.chunks(I).parentPtr
    Next I  '   For I = 1 To UBound(theLUA.chunks())
    Debug.Print "-----", "-------", "--------------"
    
    Debug.Print "LUA Functions"
    Debug.Print "-------------"
    
    Debug.Print "Index", "Pointer", "Parent Pointer"
    For I = 1 To UBound(theLUA.funcs)
        If theLUA.funcs(I).funcPtr = 0 Then Exit For
        
        Debug.Print I, theLUA.funcs(I).funcPtr, theLUA.funcs(I).parentPtr
        Debug.Print , _
            IIf(theLUA.funcs(I).numFunctions = 0, _
                "No", _
                CStr(theLUA.funcs(I).numFunctions) _
            ); _
            " function(s) defined in this function"
        
        If theLUA.funcs(I).numFunctions > 0 Then _
            Debug.Print , "Index", "Function Pointer"
        
        For J = 1 To theLUA.funcs(I).numFunctions
            Debug.Print , I, theLUA.funcs(I).functions(J)
        Next J  '   For J = 1 To theLUA.funcs(I).numFunctions
        
        If theLUA.funcs(I).numFunctions > 0 Then _
            Debug.Print , "-----", "----------------"
    Next I  '   For I = 1 To UBound(theLUA.funcs)
    Debug.Print "-----", "-------", "--------------"
End Sub

' Cleans up stuff for initialization.
Private Function Init(ByRef luaChunk As LUA_Chunk, ByVal tabLevel As Long) As Boolean
    stackP = 0
    level = tabLevel
    If luaChunk.numInstructions > 0 Then _
        currIns = luaChunk.instructions(1) _
    Else _
        currIns = 0
    
    Erase stack
    Erase localProcessed
    
    ReDim stack(1 To _
        min( _
            NearestPowerOfBase(luaChunk.maxStackSize + LUA_ExtraFields, 2), _
            LUA_MaxStackSize _
        ) _
    )
    
    If luaChunk.numLocals > 0 Then _
        ReDim localProcessed(1 To luaChunk.numLocals) _
    Else _
        ReDim localProcessed(0)
    
    Init = True
End Function

' Ripped from LUA_Decompile_Chunk
Private Sub ProcessLocals(ByRef luaChunk As LUA_Chunk, ByVal I As Long)
    Dim J As Long
    Static lastLocal As Long
    
    ' Init, if we reset-ed at some while (this sub is supposed to be called on all instructions)
    If (I = 0) Then _
        lastLocal = 1
    
    ' Initialize each local (the 'dummy' way!)
    For J = lastLocal To luaChunk.numLocals
        If luaChunk.locals(J).startpc = I Then _
            PushLocalInStack luaChunk.locals(J).name, IDT_LocalVar: _
            lastLocal = J + 1
        
        If luaChunk.locals(J).startpc > I Then _
            Exit For
    Next J  '   For J = lastLocal To luaChunk.numLocals
End Sub

' Ripped from LUA_Decompile_Chunk
Private Sub ProcessDOENDChunks(ByRef luaChunk As LUA_Chunk, ByRef strForDO As String, ByRef strForEND As String)
    Dim I As Long
    Dim tmpLong As Long, tmpLong2 As Long
    Dim A As Boolean, B As Boolean, C As Boolean    ' conditions
    
    For I = 1 To luaChunk.numLocals - 1
        ' whether this does not ends at the last instruction
        A = Not (luaChunk.locals(I).endpc = luaChunk.numInstructions)
        
        ' whether this 'do...end' chunk hasn been added (yet)
        B = Not (luaChunk.locals(I).endpc = tmpLong2)
        
        ' Whether this is not a local var of 'for'
        If luaChunk.locals(I).startpc > 0 Then _
            C = Not (Instruction_GetOPCode(luaChunk.instructions(luaChunk.locals(I).endpc)) = OP_POP) And _
                Not (OP_FORPREP <= Instruction_GetOPCode(luaChunk.instructions(luaChunk.locals(I).startpc)) And _
                    Instruction_GetOPCode(luaChunk.instructions(luaChunk.locals(I).startpc)) <= OP_LFORLOOP) _
        Else _
            C = False   ' startPC = 0; of a function arg; 'do...end' not needed
        
        If A And B And C Then
            tmpLong = luaChunk.locals(I).startpc
            tmpLong2 = luaChunk.locals(I).endpc
            
            GoSub AddDOChunk
        End If  '   If A And B And C Then
    Next I  '   For I = 1 To luaChunk.numLocals - 1
    
    Exit Sub
AddDOChunk:
    If tmpLong > 0 Then
        If tmpLong2 > 0 Then
            strForDO = "," & _
                CStr(tmpLong) & _
                strForDO
            
            strForEND = "," & _
                CStr(tmpLong2) & _
                strForEND
        End If  '   If tmpLong2 > 0 Then
    End If  '   If tmpLong > 0 Then
    
    Return
End Sub

' Ripped from LUA_Decompile_Chunk
Private Sub ProcessDOChunk(ByRef luaChunk As LUA_Chunk, ByVal I As Long, ByRef strForDO As String, ByRef outLUA As String)
    Dim tmpLong As Long
    Dim tmpStr As String
    
Starting:
    tmpLong = InStrRev(strForDO, ",")
    If tmpLong < Len(strForDO) Then _
        tmpStr = Right(strForDO, Len(strForDO) - tmpLong)
    
    If Not (Len(tmpStr) = 0) And tmpLong > 0 Then
        ' Is this the one?
        If CLng(tmpStr) = I Then
            PushStatement outLUA, "do", False
            level = level + 1
            
            ' Chop off this part.
            strForDO = Left$(strForDO, tmpLong - 1)
            
            GoTo Starting   ' because we need to work on all
        End If  '   If CLng(tmpStr) = I Then
    End If  '   If Not (Len(tmpStr) = 0) And tmpLong > 0 Then
End Sub

' Ripped from LUA_Decompile_Chunk
Private Sub ProcessENDStatement(ByRef luaChunk As LUA_Chunk, ByVal I As Long, ByRef strForEND As String, ByRef strForJumpType As String, ByRef outLUA As String)
    Dim tmpLong As Long
    Dim tmpStr As String
    
Starting:
    tmpLong = InStrRev(strForEND, ",")
    
    If tmpLong < Len(strForEND) Then _
        tmpStr = Right(strForEND, Len(strForEND) - tmpLong)
    
    If Not (Len(tmpStr) = 0) And tmpLong > 0 Then
        ' Is this the one?
        If CLng(tmpStr) = I Then
            level = level - 1
            PushStatement outLUA, "end", False
            PushStatement outLUA, ""
            
            ' Chop off this part.
            strForEND = Left$(strForEND, tmpLong - 1)
            strForJumpType = Left$(strForJumpType, Len(strForJumpType) - 1)
            
            GoTo Starting
        End If  '   If CLng(tmpStr) = I Then
    End If  '   If Not (Len(tmpStr) = 0) And tmpLong > 0 Then
End Sub

' Function for decompiling LUA chunk (or function).
Private Function LUA_Decompile_Chunk(ByRef luaChunk As LUA_Chunk, ByRef parentChunk As LUA_Chunk, ByRef outLUA As String, ByVal tabLevel As Long) As Boolean
Attribute LUA_Decompile_Chunk.VB_Description = "Description: Main function for decompiling LUA chunk (or function).\r\n    Inputs: Compiled LUA chunk; it's parent chunk (for upvalues...), tabbing level, and string where LUA is to be outputted.\r\n    Outputs: Whether the LUA was sucessfully decompiled or "
    Dim I As Long, J As Long
    Dim tmpLong As Long, tmpLong2 As Long
    
    Dim tmpStr As String, tmpStr2 As String, tmpStr3 As String
    Dim strForCondition As String, strForDO As String, strForEND As String, strForJumpType As String
    
    Dim tmpIDT As interpretationDataType
    Dim insOPCode As LUA_OPCodes
    
    LUA_Decompile_Chunk = Init(luaChunk, tabLevel)
    
    ProcessDOENDChunks luaChunk, strForDO, strForEND
    ProcessLocals luaChunk, 0               ' <- observed in functions (!!!)
    
    For I = 1 To luaChunk.numInstructions
        Debug.Assert FindLineInfo(luaChunk, I) > 0
    
        currIns = luaChunk.instructions(I)  ' Next instruction
        insOPCode = Instruction_GetOPCode(currIns)
        
        ProcessLocals luaChunk, I
        ProcessDOChunk luaChunk, I, strForDO, outLUA
        ProcessENDStatement luaChunk, I, strForEND, strForJumpType, outLUA
        
        Select Case insOPCode
            Case LUA_OPCodes.OP_END
                ' pop all vars ending here
                For J = 1 To FindNumLocalsWithEndPC(luaChunk, I)
                    PopLocalFromStack
                Next J  '   For J = 1 To FindNumLocalsWithEndPC(luaChunk, I)
                
                If Not stackP = 0 Then
                    Debug.Print "Stack not empty!"
                    PrintStack
                    
                    MsgBox "Warning! Stack is not empty! The written LUA" & vbCrLf & _
                        "may have been incorrectly decompiled!", _
                        vbApplicationModal + vbExclamation + vbDefaultButton1 + vbOKOnly, _
                        "Cold Fusion LUA Decompiler"
                    
                    LUA_Decompile_Chunk = True
                End If  '   If Not stackP = 0 Then
                                
                Debug.Assert stackP = 0
                
                Exit For
            Case LUA_OPCodes.OP_RETURN
                ' If this is not the second last instruction, then dunk this in another block.
                tmpLong = Instruction_GetUArg(currIns) + 1
                tmpLong2 = (Instruction_GetOPCode(luaChunk.instructions(I + 1)) = OP_JMP) Or _
                    (Instruction_GetOPCode(luaChunk.instructions(I + 1)) = OP_END)
                
                tmpStr = IIf(Not tmpLong2, _
                            "do ", _
                            "" _
                        ) & _
                    "return "
                
                For J = tmpLong To stackP
                    If Not J = 0 Then _
                        tmpStr = tmpStr & stack(J).value & _
                            IIf(Not J = stackP, _
                                ", ", _
                                " " _
                            )
                Next J  '   For J = tmpLong  To stackP
                
                tmpStr = tmpStr & _
                    IIf(Not tmpLong2, _
                        "end", _
                        "" _
                    )
                
                PushStatement outLUA, tmpStr
                PushStatement outLUA, ""
                
                PopValueFromStack stackP - tmpLong + 1
            Case LUA_OPCodes.OP_CALL, LUA_OPCodes.OP_TAILCALL
                ' Pop the locals, which will be awarded this function.
                tmpLong = FindMatchingStartPCForLocal(luaChunk, I)
                For J = 1 To FindNumLocalsWithStartPC(luaChunk, I)
                    PopLocalFromStack luaChunk.locals(tmpLong + J - 1).name
                Next J  '   For J = 1 To FindNumLocalsWithStartPC(luaChunk, I)
                
                ' Then get stack pos for function and num returns.
                tmpLong = Instruction_GetAArg(currIns) + 1
                tmpLong2 = IIf(insOPCode = OP_CALL, _
                    Instruction_GetBArg(currIns), _
                    0 _
                )
                
                tmpStr = stack(tmpLong).value & "("
                
                For J = tmpLong + 1 To stackP
                    tmpStr = tmpStr & stack(J).value & IIf(Not J = stackP, ", ", "")
                Next J  '   For J = tmpLong + 1 To stackP
                
                tmpStr = tmpStr & ")"
                
                ' FUNCTION ASSIGNMENT TO INITIALIZED LOCAL: Different handling...
                If (FindMatchingStartPCForLocal(luaChunk, I) > 0) Then
                    ' pop function and it's vars; and pump in this function (don't assign to local)
                    PopValueFromStack stackP - tmpLong + 1
                    PushValueInStack tmpStr, IDT_Char, I, luaChunk, outLUA, False
                    
                    ' retrieve local names (all will be consecutive, hehe...)
                    For J = 1 To tmpLong2
                        tmpStr2 = tmpStr2 & _
                            IIf(Not J = 1, ", ", "") & _
                           luaChunk.locals(FindMatchingStartPCForLocal(luaChunk, I) + J - 1).name
                    Next J  '   For J = 1 To tmpLong2
                    
                    PushStatement outLUA, "local " & _
                        tmpStr2 & _
                        "=" & _
                        tmpStr
                    
                    ' don't pop the function, since it's assigned to a local (we popped the locals earlier, remember?)
                    PopValueFromStack stackP - tmpLong
                    
                    stack(stackP).extraString = KillSpaces(tmpStr2)
                    stack(stackP).flags = IDF_IsALocalValue
                Else    '   If (FindMatchingStartPCForLocal(luaChunk, I) > 0) Then
                    ' pop function and it's vars and pump in only one return... the function itself
                    ' with an extra flag...
                    PopValueFromStack stackP - tmpLong + 1
                    PushValueInStack tmpStr, IDT_Char, I, luaChunk, outLUA
                    
                    stack(stackP).flags = IDF_FunctionReturn
                    stack(stackP).extraValue = tmpLong2
                    
                    ' If no returns then print immediately
                    If tmpLong2 = 0 Then _
                        PushStatement outLUA, stack(stackP).value: _
                        PopValueFromStack
                End If  '   If (FindMatchingStartPCForLocal(luaChunk, I) > 0) Then
            Case LUA_OPCodes.OP_PUSHNIL
                For J = 1 To Instruction_GetUArg(currIns)
                    PushValueInStack LUA_Null, IDT_Nil, I, luaChunk, outLUA
                Next J  '   For J = 1 To Instruction_GetUArg(currIns)
            Case LUA_OPCodes.OP_POP
                PopValueFromStack Instruction_GetUArg(currIns)
            Case LUA_OPCodes.OP_PUSHINT, LUA_OPCodes.OP_PUSHNUM, LUA_OPCodes.OP_PUSHNEGNUM, LUA_OPCodes.OP_PUSHSTRING
                ' Extract the appropriate thing and appropriately decide whether it's to be converted or not...
                If insOPCode = OP_PUSHINT Then
                    tmpStr = CStr(Instruction_GetSArg(currIns))
                ElseIf insOPCode = OP_PUSHSTRING Then
                    tmpStr = luaChunk.strings(Instruction_GetUArg(currIns) + 1).data
                    tmpLong = InStr(1, tmpStr, Chr$(34)) Or InStr(1, tmpStr, vbCr) Or InStr(1, tmpStr, vbLf)
                    tmpStr = IIf(tmpLong > 0, _
                            "[[", _
                            Chr$(34) _
                        ) & _
                        tmpStr & _
                        IIf(tmpLong > 0, _
                            "]]", _
                            Chr$(34) _
                        )
                Else
                    ' +ve or -ve number.
                    tmpStr = IIf(insOPCode = OP_PUSHNUM, _
                        CStr(luaChunk.numbers(Instruction_GetUArg(currIns) + 1)), _
                        CStr(-luaChunk.numbers(Instruction_GetUArg(currIns) + 1)) _
                    )
                End If
                
                ' Get the appropriate data-type for this.
                tmpIDT = IIf(insOPCode = OP_PUSHINT, IDT_Integral, 0) + _
                        IIf(insOPCode = OP_PUSHNUM, IDT_Float, 0) + _
                        IIf(insOPCode = OP_PUSHNEGNUM, IDT_Float, 0) + _
                        IIf(insOPCode = OP_PUSHSTRING, IDT_Char, 0)
                
                ' Push in stack.
                PushValueInStack tmpStr, tmpIDT, I, luaChunk, outLUA
            Case LUA_OPCodes.OP_PUSHUPVALUE
                ' OK, lets hope we have a valid parent chunk.
                Debug.Assert Not VarPtr(parentChunk) = VarPtr(luaChunk)
                
                If Not VarPtr(parentChunk) = VarPtr(luaChunk) Then
                    tmpStr = "%" & _
                        parentChunk.strings(Instruction_GetUArg(currIns) + 1).data
                    
                    PushValueInStack tmpStr, IDT_Char, I, luaChunk, outLUA
                Else    '   If Not VarPtr(parentChunk) = VarPtr(luaChunk) Then
                    MsgBox "Error: Trying to access an upvalue from a non-existant parent chunk." & _
                        vbCrLf & "This value will be taken as 'IDT_Nil'", _
                        vbApplicationModal + vbDefaultButton1 + vbExclamation + vbOKOnly, _
                        "Cold Fusion LUA Decompiler"
                    
                    PushValueInStack LUA_Null, IDT_Nil, I, luaChunk, outLUA
                End If  '   If Not VarPtr(parentChunk) = VarPtr(luaChunk) Then
            Case LUA_OPCodes.OP_GETLOCAL, LUA_OPCodes.OP_GETGLOBAL
                ' Extract the appropriate name.
                If insOPCode = OP_GETLOCAL Then _
                    tmpStr = luaChunk.locals(Instruction_GetUArg(currIns) + 1).name _
                Else _
                    tmpStr = luaChunk.strings(Instruction_GetUArg(currIns) + 1).data
                
                PushValueInStack tmpStr, IDT_Char, I, luaChunk, outLUA
            Case LUA_OPCodes.OP_GETTABLE, LUA_OPCodes.OP_GETDOTTED, LUA_OPCodes.OP_GETINDEXED
                If Not insOPCode = OP_GETTABLE Then _
                    If insOPCode = OP_GETDOTTED Then _
                        tmpStr = luaChunk.strings(Instruction_GetUArg(currIns) + 1).data: _
                        tmpStr = Chr$(34) & tmpStr & Chr$(34) _
                    Else _
                        tmpStr = luaChunk.locals(Instruction_GetUArg(currIns) + 1).name _
                Else _
                    tmpStr = ProcessValueForOutput
                
                ' Extract the name and IDT_Table entry. (IDT_Table is at -1, and value is at the pointer, ie 0)
                tmpStr = ProcessTableValue(stack(stackP - IIf(insOPCode = OP_GETTABLE, 1, 0)).value, tmpStr)
                    
                ' Don't forget to pop the value(s) first! Then push our result ("t[i]")
                PopValueFromStack 1 + IIf(insOPCode = OP_GETTABLE, 1, 0)
                PushValueInStack tmpStr, IDT_Char, I, luaChunk, outLUA
            Case LUA_OPCodes.OP_PUSHSELF
                tmpStr = stack(stackP).value & _
                    "[" & _
                        Chr$(34) & _
                            luaChunk.strings(Instruction_GetUArg(currIns) + 1).data & _
                        Chr$(34) & _
                    "]"
                
                PushValueInStack tmpStr, IDT_Char, I, luaChunk, outLUA
            Case LUA_OPCodes.OP_CREATETABLE
                tmpLong = Instruction_GetUArg(currIns)
                
                ' just create the raw IDT_Table, and yes, DON'T SEARCH FOR A LOCAL IF NO DATA!!!
                ' (that's done for OP_SETLIST, if data is present, else, create empty IDT_Table for local)
                PushValueInStack CStr(tmpLong), IDT_Table, I, luaChunk, outLUA, IIf(tmpLong = 0, True, False)
            Case LUA_OPCodes.OP_SETLOCAL, LUA_OPCodes.OP_SETGLOBAL
                ' Extract the appropriate name (for local or global var).
                tmpLong = BitsPresent(stack(stackP).flags, IDF_FunctionReturnWithEQ)
                
                If Instruction_GetOPCode(currIns) = LUA_OPCodes.OP_SETLOCAL Then _
                    tmpStr = luaChunk.locals(Instruction_GetUArg(currIns) + 1).name _
                Else _
                    tmpStr = luaChunk.strings(Instruction_GetUArg(currIns) + 1).data
                
                ' If tis a function...
                If Not stack(stackP).type = IDT_Closure Then
                    tmpStr = tmpStr & _
                            IIf(BitsPresent(stack(stackP).flags, IDF_FunctionReturnWithEQ), _
                                ", ", _
                                "=" _
                            ) & _
                            ProcessValueForOutput
                    
                    stack(stackP).extraValue = stack(stackP).extraValue + _
                        BitsPresent(stack(stackP).flags, IDF_FunctionReturnWithEQ)
                    
                    If BitsPresent(stack(stackP).flags, IDF_FunctionReturn) Then _
                        stack(stackP).flags = IDF_FunctionReturnWithEQ: _
                        stack(stackP).extraValue = stack(stackP).extraValue - 1
                Else    '   If Not stack(stackP).type = interpretationDataType.IDT_Closure Then
                    tmpStr = Replace(stack(stackP).value, "%n", tmpStr) & vbCrLf
                End If  '   If Not stack(stackP).type = interpretationDataType.IDT_Closure Then
                
                ' Remove from stack only if it's not of a function return
                If Not BitsPresent(stack(stackP).flags, IDF_FunctionReturnWithEQ) Then _
                    PushStatement outLUA, tmpStr: _
                    PopValueFromStack _
                Else _
                    stack(stackP).value = tmpStr: _
                    If stack(stackP).extraValue = 0 Then _
                        PushStatement outLUA, tmpStr: _
                        PopValueFromStack
            Case LUA_OPCodes.OP_SETTABLE
                tmpLong = Instruction_GetAArg(currIns)
                tmpStr = ProcessTableValue(stack(stackP - tmpLong + 1).value, ProcessValueForOutput(-tmpLong + 2))
                
                If Not stack(stackP).type = interpretationDataType.IDT_Closure Then
                    tmpStr = tmpStr & _
                        IIf(BitsPresent(stack(stackP).flags, IDF_FunctionReturnWithEQ), _
                            ", ", _
                            "=" _
                        ) & _
                        ProcessValueForOutput
                        
                    stack(stackP).extraValue = stack(stackP).extraValue + BitsPresent(stack(stackP).flags, IDF_FunctionReturnWithEQ)
                    
                    If BitsPresent(stack(stackP).flags, IDF_FunctionReturn) Then _
                        stack(stackP).flags = IDF_FunctionReturnWithEQ: _
                        stack(stackP).extraValue = stack(stackP).extraValue - 1
                Else    '   If Not stack(stackP).type = interpretationDataType.IDT_Closure Then
                    tmpStr = Replace(stack(stackP).value, "%n", tmpStr) & vbCrLf
                End If  '   If Not stack(stackP).type = interpretationDataType.IDT_Closure Then
                
                If Not BitsPresent(stack(stackP).flags, IDF_FunctionReturnWithEQ) Then _
                    PushStatement outLUA, tmpStr: _
                    PopValueFromStack Instruction_GetBArg(currIns) _
                Else _
                    stack(stackP).value = tmpStr: _
                    If stack(stackP).extraValue = 0 Then _
                        PushStatement outLUA, tmpStr: _
                        PopValueFromStack Instruction_GetBArg(currIns)
            Case LUA_OPCodes.OP_SETLIST
                tmpLong = Instruction_GetBArg(currIns)
                
                If IsNumeric(stack(stackP - tmpLong).value) Then
                    ' Stock IDT_Table.
                    tmpStr = "{"
                Else    '   If IsNumeric(stack(stackP - tmpLong).value) Then
                    ' IDT_Table has more than 62 entries. Not stock. Chop off the right brace "}" and put a ", "
                    tmpStr = stack(stackP - tmpLong).value
                    tmpStr = Left(tmpStr, Len(tmpStr) - 1) & "," & _
                        IIf(stack(stackP - tmpLong).extraString = "OP_SETLIST", "", ";") & _
                        " "
                End If  '   If IsNumeric(stack(stackP - tmpLong).value) Then
                    
                For J = tmpLong To 1 Step -1
                    tmpStr = tmpStr & _
                        ProcessValueForOutput(-J + 1) & _
                            IIf(Not J = 1, "," & vbCrLf & String$(level + 1, Chr$(9)), "")
                Next J  '   For J = tmpLong To 1 Step -1
                
                tmpStr = tmpStr & "}"
                tmpLong2 = stack(stackP).extraValue - tmpLong
                
                ' Pop the IDT_Table, too. We'll add it manually...
                PopValueFromStack tmpLong + 1
                PushValueInStack tmpStr, IIf(tmpLong2, IDT_Table, IDT_Char), I, luaChunk, outLUA
                
                stack(stackP).extraString = "OP_SETLIST"
                stack(stackP).extraValue = tmpLong2
            Case LUA_OPCodes.OP_SETMAP
                tmpLong = Instruction_GetUArg(currIns) * 2
                tmpStr2 = ""
                
                If IsNumeric(stack(stackP - tmpLong).value) Then
                    ' Stock IDT_Table.
                    tmpStr = "{"
                Else    '   If IsNumeric(stack(stackP - tmpLong).value) Then
                    ' IDT_Table has more than 62 entries. Not stock. Or a hybid (well, could be...)
                    ' Chop off the right brace ("}") and put a ", "
                    tmpStr = stack(stackP - tmpLong).value
                    tmpStr = Left(tmpStr, Len(tmpStr) - 1) & "," & _
                        IIf(stack(stackP - tmpLong).extraString = "OP_SETMAP", "", ";") & _
                        " "
                End If  '   If IsNumeric(stack(stackP - tmpLong).value) Then
                
                For J = tmpLong To 1 Step -2
                    tmpStr3 = ProcessValueForOutput(-J + 1)
                    tmpStr3 = IIf(Len(UnquoteString(tmpStr3)) - Len(tmpStr3), _
                        UnquoteString(tmpStr3), _
                        "[" & tmpStr3 & "]" _
                    ) ' if quotes are present, remove them, else if not, then put in '[]' (they're vars)
                    
                    tmpStr3 = tmpStr3 & _
                        "=" & _
                        ProcessValueForOutput(-J + 2)
                    
                    tmpStr2 = tmpStr2 & tmpStr3 & IIf(Not J = 2, "," & vbCrLf & String$(level, Chr$(9)), "")
                Next J  '   For J = tmpLong To 1 Step -1
                
                tmpStr = tmpStr & tmpStr2 & "}"
                tmpLong2 = stack(stackP).extraValue - tmpLong
                
                ' Pop all values AND the IDT_Table...
                PopValueFromStack tmpLong + 1
                
                ' Push this IDT_Table in manually.
                PushValueInStack tmpStr, IIf(tmpLong2, IDT_Table, IDT_Char), I, luaChunk, outLUA
                
                stack(stackP).extraString = "OP_SETMAP"
                stack(stackP).extraValue = tmpLong2
            Case LUA_OPCodes.OP_ADD, LUA_OPCodes.OP_ADDI, LUA_OPCodes.OP_SUB, LUA_OPCodes.OP_MULT, LUA_OPCodes.OP_DIV, LUA_OPCodes.OP_POW
                ' Retrieve the value, pop the stack, and then work on appropriate OP.
                If Not insOPCode = OP_ADDI Then
                    tmpStr = stack(stackP).value
                    PopValueFromStack
                Else    '   If Not INtruction_OPCode(currIns) = OP_ADDI Then
                    tmpStr = Instruction_GetSArg(currIns)
                End If  '   If Not INtruction_OPCode(currIns) = OP_ADDI Then
                
                tmpStr = "(" & _
                    stack(stackP).value & _
                    " " & _
                        IIf(insOPCode = OP_ADD, "+", "") & _
                        IIf(insOPCode = OP_ADDI, "+", "") & _
                        IIf(insOPCode = OP_SUB, "-", "") & _
                        IIf(insOPCode = OP_MULT, "*", "") & _
                        IIf(insOPCode = OP_DIV, "/", "") & _
                        IIf(insOPCode = OP_POW, "^", "") & _
                    " " & _
                    tmpStr & _
                    ")"
                    
                ' if we ADDI with -x, then change 'y + -x' to 'y - x'
                tmpStr = Replace(tmpStr, "+ -", "- ")
                tmpIDT = stack(stackP).type
                
                PopValueFromStack
                PushValueInStack tmpStr, tmpIDT, I, luaChunk, outLUA
            Case LUA_OPCodes.OP_CONCAT
                ' Get nr. of strings.
                tmpLong = Instruction_GetUArg(currIns)
                tmpStr = ""
                
                ' Concat reverse.
                For J = tmpLong To 1 Step -1
                    tmpStr = tmpStr & _
                        stack(stackP - J + 1).value & _
                        IIf(Not J = 1, "..", "")
                Next J  '   For J = tmpLong To 1 Step -1
                
                ' Pop all strings, and manually push our string
                PopValueFromStack tmpLong
                PushValueInStack tmpStr, IDT_Char, I, luaChunk, outLUA
            Case LUA_OPCodes.OP_MINUS, LUA_OPCodes.OP_NOT
                tmpStr = stack(stackP).value
                tmpIDT = stack(stackP).type
                
                PopValueFromStack
                
                tmpStr = IIf(insOPCode = OP_MINUS, _
                        "-", _
                        "(not " _
                    ) & _
                    tmpStr & _
                    IIf(insOPCode = OP_MINUS, _
                        "", _
                        ")" _
                    )
                
                PushValueInStack tmpStr, tmpIDT, I, luaChunk, outLUA
            Case LUA_OPCodes.OP_JMPNE, LUA_OPCodes.OP_JMPEQ, LUA_OPCodes.OP_JMPLT, LUA_OPCodes.OP_JMPLE, LUA_OPCodes.OP_JMPGT, LUA_OPCodes.OP_JMPGE
                ' decide whether
                '      i) we encountered a while (...) do ... end statement.
                '     ii) we encountered a if (...) then ... else ... end condition.
                ' or iii) we encountered a <var> = <cond> condition.
                tmpLong = luaChunk.instructions(FindEndOfJump(currIns, I, luaChunk) - 1)
                
                ' since jump offset cannot be 0, -ve jump offset indicates that it's a (i) and NOT (ii)
                If (Instruction_GetOPCode(tmpLong) = OP_JMP) And (Instruction_GetSArg(tmpLong) < 0) Then _
                    strForJumpType = strForJumpType & "w" _
                Else _
                    strForJumpType = strForJumpType & "c"
                
                ' find the correct string for this OP Code.
                tmpLong = insOPCode
                tmpStr2 = IIf(tmpLong = OP_JMPNE, "~=", "") & _
                    IIf(tmpLong = OP_JMPEQ, "==", "") & _
                    IIf(tmpLong = OP_JMPLT, "<", "") & _
                    IIf(tmpLong = OP_JMPLE, "<=", "") & _
                    IIf(tmpLong = OP_JMPGT, ">", "") & _
                    IIf(tmpLong = OP_JMPGE, ">=", "")
                
                ' and get the next condition, too.
                tmpLong = FindNextCondition(currIns, I, luaChunk, False)
                
                ' For forms like A=(<condition>)
                If Instruction_GetOPCode(luaChunk.instructions(I + 1)) = OP_PUSHNILJMP Then
                    tmpStr = strForCondition & _
                            "(" & _
                            ProcessValueForOutput(-1) & _
                            tmpStr2 & _
                            ProcessValueForOutput & _
                            ")"
                    
                    If Not tmpLong = 0 Then
                        ' negate all inequalities.
                        tmpStr = Replace(tmpStr, "==", CStr(OP_JMPEQ))
                        tmpStr = Replace(tmpStr, "~=", CStr(OP_JMPNE))
                        tmpStr = Replace(tmpStr, ">=", CStr(OP_JMPGE))
                        tmpStr = Replace(tmpStr, "<=", CStr(OP_JMPLE))
                        tmpStr = Replace(tmpStr, ">", CStr(OP_JMPGT))
                        tmpStr = Replace(tmpStr, "<", CStr(OP_JMPLT))
                        
                        tmpStr = Replace(tmpStr, CStr(OP_JMPEQ), "~=")
                        tmpStr = Replace(tmpStr, CStr(OP_JMPNE), "==")
                        tmpStr = Replace(tmpStr, CStr(OP_JMPGT), "<=")
                        tmpStr = Replace(tmpStr, CStr(OP_JMPLT), ">=")
                        tmpStr = Replace(tmpStr, CStr(OP_JMPGE), "<")
                        tmpStr = Replace(tmpStr, CStr(OP_JMPLE), ">")
                        
                        ' replace or with and and vice-versa
                        tmpStr = Replace(tmpStr, "and", "1")
                        tmpStr = Replace(tmpStr, "or", "and")
                        tmpStr = Replace(tmpStr, "1", "or")
                    End If  '   If Not tmpLong = 0 Then
                    
                    PopValueFromStack 2 ' the two conditions, ie x and y
                    PushValueInStack tmpStr, IDT_Char, I, luaChunk, outLUA
                    
                    I = I + 2   ' jump away from PUSHNILJMP and PUSHINT
                    
                    GoTo ExitSelect
                End If  '   If Instruction_GetOPCode(luaChunk.instructions(I + 1)) = OP_PUSHNILJMP Then
                
                ' For an 'or' don't reverse condition, else... rev!
                ' (negate in case of "w"hile loops).
                If Right$(strForJumpType, 1) = "c" Then
                    If FindORPresence(I, luaChunk) Then _
                        tmpStr = "(" & _
                            ProcessValueForOutput(-1) & _
                            tmpStr2 & _
                            ProcessValueForOutput & _
                            ") or " _
                    Else _
                        tmpStr = "(" & _
                            ProcessValueForOutput(-1) & _
                            ReverseCondition(tmpStr2) & _
                            ProcessValueForOutput & _
                            ") and "
                Else    '   If Right$(strForJumpType, 1) = "c" Then
                    ' Mimic 'and' for 'or' and 'or' for 'and'; for last, mimic 'and' not 'or'
                    If FindORPresence(I, luaChunk) Or (tmpLong = 0) Then _
                        tmpStr = "(" & _
                            ProcessValueForOutput(-1) & _
                            ReverseCondition(tmpStr2) & _
                            ProcessValueForOutput & _
                            ") and " _
                    Else _
                        tmpStr = "(" & _
                            ProcessValueForOutput(-1) & _
                            tmpStr2 & _
                            ProcessValueForOutput & _
                            ") or "
                End If  '   If Right$(strForJumpType, 1) = "c" Then
                
                strForCondition = strForCondition & _
                    tmpStr
                
                If tmpLong > 0 Then
                    ' Don't consider this, we'll consider the next one.
                    strForJumpType = Left$(strForJumpType, Len(strForJumpType) - 1)
                Else    '   If tmpLong > 0 Then
                    tmpStr2 = Right$(strForJumpType, 1)
                    
                    ' If we don't have any more conditions, then first trim last 4 characters
                    ' (they're an 'and ' sureshot!) and put an 'if' and 'then' (or 'while' and
                    ' 'do' respectively, whatever is applicable.)
                    tmpStr = IIf(FindELSEIFPresence(I, luaChunk), "else", "") & _
                        IIf(tmpStr2 = "c", "if ", "") & _
                        IIf(tmpStr2 = "w", "while  ", "") & _
                        LTrim(RTrim(Left(strForCondition, Len(strForCondition) - 4))) & _
                        IIf(tmpStr2 = "c", " then", "") & _
                        IIf(tmpStr2 = "w", " do", "")
                    
                    PushStatement outLUA, tmpStr, False
                    
                    level = level + 1           ' from now on, one tab extra!
                    strForCondition = ""        ' not to forget to purge this!
                    
                    'PushJumpInRegister luaChunk, I, JT_Conditional, tmpStr
                    
                    ' determine if there is a 'else' block (in this case, the jump will decide
                    ' the position of 'end') or not. If 'else' is absent, make a personal note
                    ' where 'end' is to be added. (don't do anything for 'while', OP_JMP will
                    ' automatically add 'end'.
                    tmpLong = luaChunk.instructions(Instruction_GetSArg(currIns) + I)
                    If Not (Instruction_GetOPCode(tmpLong) = OP_JMP) Then _
                        strForEND = strForEND & _
                            "," & _
                            CStr(Instruction_GetSArg(currIns) + I + 1)
                End If  '   If tmpLong > 0 Then
                
                PopValueFromStack 2 ' don't forget me!
            Case LUA_OPCodes.OP_JMPT, LUA_OPCodes.OP_JMPF, LUA_OPCodes.OP_JMPONT, LUA_OPCodes.OP_JMPONF
                ' decide whether
                '     i) we encountered a while (...) do ... end statement.
                ' or ii) we encountered a if (...) then ... else ... end condition.
                tmpLong = luaChunk.instructions(FindEndOfJump(currIns, I, luaChunk) - 1)
                tmpLong2 = insOPCode
                
                ' since jump offset cannot be 0, -ve jump offset indicates that it's a (i) and NOT (ii)
                If (Instruction_GetOPCode(tmpLong) = OP_JMP) And (Instruction_GetSArg(tmpLong) < 0) Then _
                    strForJumpType = strForJumpType & "w" _
                Else _
                    strForJumpType = strForJumpType & "c"
                
                ' find the correct string for this OP Code.
                tmpStr = IIf(tmpLong2 = OP_JMPT, " not (%1)", "") & _
                    IIf(tmpLong2 = OP_JMPF, " (%1)", "") & _
                    IIf(tmpLong2 = OP_JMPONT, " not (%1)", "") & _
                    IIf(tmpLong2 = OP_JMPONF, " (%1)", "")
                
                tmpStr = Replace(tmpStr, "%1", ProcessValueForOutput)
                tmpLong = FindNextCondition(currIns, I, luaChunk, False)
                
                If tmpLong > 0 Then
                    If Right$(strForJumpType, 1) = "c" Then
                        If FindORPresence(I, luaChunk) Then _
                            strForCondition = strForCondition & _
                                tmpStr & _
                                " and " _
                        Else _
                            strForCondition = strForCondition & _
                                " not " & _
                                tmpStr & _
                                " or "
                    Else    '   If Right$(strForJumpType, 1) = "c" Then
                        If FindORPresence(I, luaChunk) Then _
                            strForCondition = strForCondition & _
                                tmpStr & _
                                " and " _
                        Else _
                            strForCondition = strForCondition & _
                                " not " & _
                                tmpStr & _
                                " or "
                    End If  '   If Right$(strForJumpType, 1) = "c" Then
                    
                    ' Don't consider this, we'll consider the next one.
                    strForJumpType = Left$(strForJumpType, Len(strForJumpType) - 1)
                Else    '   If tmpLong > 0 Then
                    strForCondition = strForCondition & _
                        tmpStr
                    
                    tmpStr2 = Right$(strForJumpType, 1)
                    tmpStr = IIf(FindELSEIFPresence(I, luaChunk), "else", "") & _
                        IIf(tmpStr2 = "c", "if ", "") & _
                        IIf(tmpStr2 = "w", "while  ", "") & _
                        LTrim(RTrim(strForCondition)) & _
                        IIf(tmpStr2 = "c", " then", "") & _
                        IIf(tmpStr2 = "w", " do", "")
                    
                    Do While InStr(1, tmpStr, "  ") > 0
                        tmpStr = Replace(tmpStr, "  ", " ")
                    Loop    '   Do While InStr(1, tmpStr, "  ") > 0
                    
                    Do While InStr(1, tmpStr, "not not") > 0
                        tmpStr = Replace(tmpStr, "not not", "")
                    Loop    '   Do While InStr(1, tmpStr, "not not") > 0
                    
                    PushStatement outLUA, tmpStr, False
                    
                    level = level + 1           ' from now on, one tab extra!
                    strForCondition = ""        ' not to forget to purge this!
                    
                    'PushJumpInRegister luaChunk, I, JT_Conditional, tmpStr
                    
                    ' determine if there is a 'else' block (in this case, the jump will decide
                    ' the position of 'end') or not. If 'else' is absent, make a personal note
                    ' where 'end' is to be added. (don't do anything for 'while', OP_JMP will
                    ' automatically add 'end'.
                    tmpLong = luaChunk.instructions(Instruction_GetSArg(currIns) + I)
                    
                    If Not (Instruction_GetOPCode(tmpLong) = OP_JMP) Then _
                        strForEND = strForEND & _
                            "," & _
                            CStr(Instruction_GetSArg(currIns) + I + 1)
                End If  '   If tmpLong > 0 Then
                
                ' decide if we have to pop the value or not (this is gonna be crappy...)
                If (insOPCode = OP_JMPONT) Then
                    If Not tmpStr = LUA_Null Then _
                        PopValueFromStack
                ElseIf (insOPCode = OP_JMPONF) Then
                    If tmpStr = LUA_Null Then _
                        PopValueFromStack
                Else
                    PopValueFromStack
                End If
            Case LUA_OPCodes.OP_JMP
                level = level - 1
                tmpStr = Right$(strForJumpType, 1)
                
                'PushJumpInRegister luaChunk, I, JT_Unconditional, "-"
                
                ' Was the last logic a condition or a loop? Add the appropriate command for either case
                ' and set the appropriate level.
                Select Case tmpStr
                    Case "c"    ' if-...-then condition
                        tmpLong = FindENDRequirement(I, luaChunk)
                        tmpLong2 = FindELSERequirement(I, luaChunk)
                        
                        If tmpLong Then
                            PushStatement outLUA, "else", False
                            level = level + 1
                            
                            strForEND = strForEND & _
                                "," & _
                                CStr(FindENDRequirement(I, luaChunk))
                        ElseIf tmpLong2 Then
                            PushStatement outLUA, "else", False
                            level = level + 1
                            
                            strForEND = strForEND & _
                                "," & _
                                CStr(Instruction_GetSArg(currIns) + I + 1)
                        End If
                    Case "f"    ' for-loop
                        level = level + 1   ' Reset the level
                    
                        tmpLong = (Instruction_GetOPCode(luaChunk.instructions(I + 1)) = OP_FORLOOP) Or (Instruction_GetOPCode(luaChunk.instructions(I + 1)) = OP_LFORLOOP)
                        tmpStr = IIf(Not tmpLong, "do ", "") & _
                            "break" & _
                            IIf(Not tmpLong, " end", "") & _
                            vbCrLf
                        
                        PushStatement outLUA, tmpStr, False
                    Case "w"    ' while loop
                        ' put an 'end'
                        PushStatement outLUA, "end", False
                        PushStatement outLUA, ""
                        
                        strForJumpType = Left$(strForJumpType, Len(strForJumpType) - 1)
                End Select  '   Select Case tmpStr
            Case LUA_OPCodes.OP_PUSHNILJMP
                MsgBox "Instruction " & I & " references instruction 'OP_PUSHNILJMP', which was unexpected" & _
                    vbCrLf & " at this time. This LUA may not get decompiled properly...", _
                    vbApplicationModal + vbCritical + vbDefaultButton1 + vbOKOnly, _
                    "Cold Fusion LUA Decompiler"
                
                PushValueInStack LUA_Null, IDT_Nil, I, luaChunk, outLUA
            Case LUA_OPCodes.OP_FORPREP, LUA_OPCodes.OP_LFORPREP
                If insOPCode = OP_FORPREP Then
                    tmpStr = "for " & _
                       luaChunk.locals(FindMatchingStartPCForLocal(luaChunk, I)).name & _
                        "=" & _
                        ProcessValueForOutput(-2) & _
                        ", " & _
                        ProcessValueForOutput(-1) & _
                        IIf(Val(ProcessValueForOutput) = 1, _
                            "", _
                            ", " & _
                            ProcessValueForOutput _
                        ) & _
                        " do"
                Else    '   If insOPCode = OP_FORPREP Then
                    tmpStr = "for " & _
                       luaChunk.locals(FindMatchingStartPCForLocal(luaChunk, I) + 1).name & _
                        ", " & _
                       luaChunk.locals(FindMatchingStartPCForLocal(luaChunk, I) + 2).name & _
                        " in " & _
                        ProcessValueForOutput & _
                        " do"
                End If  '   If insOPCode = OP_FORPREP Then
                
                PushStatement outLUA, tmpStr, False
                PopValueFromStack IIf(insOPCode = OP_FORPREP, _
                    3, _
                    1 _
                )
                
                strForJumpType = strForJumpType & "f"
                
                level = level + 1
            Case LUA_OPCodes.OP_FORLOOP, LUA_OPCodes.OP_LFORLOOP
                level = level - 1
                
                strForJumpType = Left$(strForJumpType, Len(strForJumpType) - 1)
                
                PushStatement outLUA, "end", False
                PushStatement outLUA, ""
                
                ' 'Kill' our counter vars
                PopLocalFromStack
                PopLocalFromStack
                PopLocalFromStack
            Case LUA_OPCodes.OP_IDT_Closure
                tmpStr = FuncStatement  ' assume form of base function statement.
                tmpLong = Instruction_GetAArg(currIns) + 1
                
                tmpStr = Replace(tmpStr, "&f", CStr(luaChunk.functions(tmpLong)))
                tmpStr = Replace(tmpStr, "&p", CStr(luaChunk.funcPtr))
                
                PushValueInStack tmpStr, IDT_Closure, I, luaChunk, outLUA
            Case Else
                LUA_Decompile_Chunk = False
                
                MsgBox "Unknwon Instruction: " & insOPCode, _
                    vbApplicationModal + vbCritical + vbDefaultButton1 + vbOKOnly, _
                    "Cold Fusion LUA Decompiler"
                
#If HaltOnErrors Then
                Debug.Assert 0
#End If
        End Select  '  Select Case insOPCode
        
ExitSelect:
        ' print stack if told so.
#If PrintStack Then
        PrintStack
#End If
        
        ' Flush vars
        J = 0: tmpLong = 0: tmpLong2 = 0
        tmpStr = "": tmpStr2 = "": tmpStr3 = ""
        tmpIDT = IDT_Nil: insOPCode = OP_END
    Next I  '   For I = 1 To luaChunk.numInstructions
    
    ' Done.
    'PrintRegister
End Function

' Main function for decompiling a LUA.
Function LUA_Decompile(ByRef inLUAPath As String, ByRef outLUAPath As String) As Boolean
Attribute LUA_Decompile.VB_Description = "Main function for decompiling a LUA."
    Dim I As Long, J As Long, K As Long
    Dim tmpLong As Long, tmpLong2 As Long, tmpLong3 As Long, tmpLong4 As Long
    
    Dim outLUA As String, chunk As String
    Dim tmpStr As String, tmpStr2 As String
    Dim tmpStrA() As String, tmpStrA2() As String
    
    Dim theLUA As LUA_File, theFunc As LUA_Chunk
    
    LUA_Decompile = True
    
    If Not ReadLUA(inLUAPath, theLUA) Then _
        GoTo Failure
    
    Debug.Assert UBound(theLUA.chunks()) >= 1
    
    ' First decompile chunk, then functions in it.
    For I = 1 To UBound(theLUA.chunks())
        LUA_Decompile = LUA_Decompile And LUA_Decompile_Chunk(theLUA.chunks(I), theLUA.chunks(I), chunk, 0)
        
        outLUA = outLUA & chunk
        chunk = ""
        
ProcessFunctions:
        ' Then decompile all functions.
        For J = 1 To UBound(theLUA.funcs())
            If theLUA.funcs(J).funcPtr = 0 Then Exit For
            
            LUA_Decompile = LUA_Decompile And LUA_Decompile_Chunk( _
                theLUA.funcs(J), _
                theLUA.chunks(I), _
                chunk, _
                FindLUAFunctionTab(theLUA.funcs(J).funcPtr, theLUA) _
            )
            tmpStrA = Split(outLUA, vbCrLf)
            
            For K = 0 To UBound(tmpStrA())
                If InStr(1, tmpStrA(K), FuncStatement_SStr) > 0 Then
                    tmpStr = Right$(tmpStrA(K), Len(tmpStrA(K)) - InStr(1, tmpStrA(K), FuncStatement_SStr) + 1)
                    tmpStr = Replace$(tmpStr, FuncStatement_SStr, "")
                    tmpStr = Left$(tmpStr, InStrRev(tmpStr, " ") - 1)
                    
                    ' first is func #, second is ins # (to IDT_Closure instruction), third is name of func
                    tmpStrA2 = Split(tmpStr, ", ")
                    
                    tmpLong2 = FindLUAFunctionTab(theLUA.funcs(J).funcPtr, theLUA)
                    tmpLong3 = CLng(tmpStrA2(1))    ' Pointer to function
                    tmpLong4 = CLng(tmpStrA2(2))    ' Pointer to function's parent
                    
                    If tmpLong3 = theLUA.funcs(J).funcPtr Then
                        tmpStr2 = String(tmpLong2 - 1, Chr$(9)) & _
                            "function " & _
                            KillSpaces(IIf(tmpStrA2(0) = "%n", "", tmpStrA2(0))) & _
                            "("
                        
                        For tmpLong = 1 To theLUA.funcs(J).numParams
                            tmpStr2 = tmpStr2 & _
                                theLUA.funcs(J).locals(tmpLong).name & _
                                IIf(Not tmpLong = theLUA.funcs(J).numParams, ", ", "")
                        Next tmpLong    '   For tmpLong = 1 To theLUA.funcs(J).numParams
                        
                        If theLUA.funcs(J).isVarArg Then _
                            tmpStr2 = tmpStr2 & IIf(theLUA.funcs(J).numParams > 0, ", ", "") & "..."
                            
                        tmpStr2 = tmpStr2 & _
                            ")" & _
                            IIf(Len(chunk) > 0, _
                                vbCrLf & _
                                chunk & _
                                vbCrLf, _
                                " " _
                            ) & _
                            String(tmpLong2 - 1, Chr$(9)) & _
                            "end"
                        
                        tmpStr2 = vbCrLf & tmpStr2
                        
                        tmpStr = FuncStatement_SStr & tmpStr & FuncStatement_SStr_R
                        tmpStrA(K) = Replace$(tmpStrA(K), tmpStr, tmpStr2)
                        
                        outLUA = Join(tmpStrA, vbCrLf)
                        chunk = ""
                        
                        Exit For
                    End If  '   If tmpLong3 = theLUA.funcs(J).funcPtr Then
                End If  '   If InStr(1, tmpStrA(K), FuncStatement_SStr) > 0 Then
            Next K  '   For K = 0 To UBound(tmpStrA())
        Next J  '   For J = 1 To UBound(theLUA.funcs())
    Next I  '   For I = 1 To UBound(theLUA.chunks())
    
    I = FreeFile
    
    Open outLUAPath For Output As #I
        ' Stupid branding!
        Print #I, "-- " & "Cold Fusion LUA Decompiler" & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
            "-- By 4E534B" & vbCrLf & _
            "-- Date: " & Date$ & " Time: " & Time$() & vbCrLf & _
            "-- On error(s), send source (compiled) file to 4E534B@gmail.com" & vbCrLf
        
        ' Ah, finally the LUA!
        Print #I, outLUA
    Close #I
    
    ' clean up:
    Erase tmpStrA
    Erase tmpStrA2
    
    Exit Function
Failure:
    LUA_Decompile = False
    
End Function
