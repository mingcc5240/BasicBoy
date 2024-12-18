Attribute VB_Name = "z80cmd"
'This is a part of tha BasicBoy emulator
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel).
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'(I know the emulator is NOT OPTIMIZED AT ALL)

'This module contains the emulation of some z80 instructions
'Coments will be added with the next release

'Sory for my bad english ...

Option Explicit
Dim memptr As Long, memptr2 As Long
Public Sub rl(reg8 As Long)
    memptr = (reg8 And 128) = 128
    reg8 = ((reg8 * 2) Or -getC) And 255
    setZ reg8 = 0
    setH False
    setN False
    setC (memptr)
End Sub
Public Sub rla()
    memptr = (A And 128) = 128
    A = ((A * 2) Or -getC) And 255
    setZ A = 0
    setN False
    setH False
    setC (memptr)
End Sub

Public Sub rlc(reg8 As Long)
    setC ((reg8 And 128) = 128)
    reg8 = ((reg8 * 2) Or -(getC)) And 255
    setZ reg8 = 0
    setH False
    setN False
End Sub

Public Sub rlca()
    setC ((A And 128) = 128)
    A = ((A * 2) Or -getC) And 255
    setZ A = 0
    setN False
    setH False
End Sub

Public Sub rr(reg8 As Long)
    memptr = (reg8 And 1) = 1
    reg8 = (reg8 \ 2) Or (128 * -getC)
    setZ reg8 = 0
    setH False
    setN False
    setC (memptr)
End Sub

Public Sub rra()
    memptr = (A And 1) = 1
    A = (A \ 2) Or (128 * -getC)
    setZ A = 0
    setN False
    setH False
    setC (memptr)
End Sub

Public Sub rrc(reg8 As Long)
    setC ((reg8 And 1) = 1)
    reg8 = (reg8 \ 2) Or (128 * -getC)
    setZ reg8 = 0
    setH False
    setN False
End Sub
Public Sub rrca()
    setC ((A And 1) = 1)
    A = (A \ 2) Or (128 * -getC)
    setZ A = 0
    setN False
    setH False
End Sub
Public Sub sla(reg8 As Long)
    setC ((reg8 And 128) = 128)
    reg8 = (reg8 * 2) And 255
    setZ reg8 = 0
    setH False
    setN False
End Sub
Public Sub sra(reg8 As Long)
    setC ((reg8 And 1) = 1)
    reg8 = (reg8 \ 2) Or (reg8 And 128)
    setZ reg8 = 0
    setH False
    setN False
End Sub
Public Sub srl(reg8 As Long)
    setC ((reg8 And 1) = 1)
    reg8 = reg8 \ 2
    setZ reg8 = 0
    setH False
    setN False
End Sub
Public Sub zsub(reg8 As Long)
    memptr = A - reg8
    memptr2 = memptr And 255
    setZ memptr2 = 0
    setC A < memptr2
    setH ((A And 15) < (memptr2 And 15))
    setN True
    A = memptr2
End Sub
Public Sub sbc(reg8 As Long)
    memptr2 = -getC
    memptr = (A - reg8 - memptr2) And 255
    setZ memptr = 0
    setC A < memptr
    setH (A And &HF&) < (memptr And 15)
    setN True
    A = memptr
End Sub
Public Sub adc(reg8 As Long)
    memptr = -getC
    memptr2 = (A + reg8 + memptr) And 255
    setZ memptr2 = 0
    setC ((A + reg8 + memptr) And &H100) <> 0&
    setH (((A And &HF&) + (reg8 And &HF&) + memptr) And 16) <> 0&
    setN False
    A = memptr2
End Sub

Public Sub add(reg8 As Long)
    memptr = (A + reg8) And &HFF&
    setZ (memptr = 0&)
    setC ((A + reg8) And &H100) <> 0&
    setH ((((A And &HF&) + (reg8 And &HF&)) And 16) = 16)
    setN False
    A = memptr
End Sub

Sub push(reg8 As Long)
sp = sp - 1
WriteM sp, reg8
End Sub

Sub pop(reg8 As Long)
reg8 = readM(sp)
sp = sp + 1
End Sub

Sub zand(Val As Long)
A = A And Val
setZ A = 0
setN False
setC False
setH True
End Sub
Sub zor(Val As Long)
A = A Or Val
setZ A = 0
setN False
setH False
setC False
End Sub
Sub zxor(Val As Long)
A = A Xor Val
setZ A = 0
setN False
setH False
setC False
End Sub
Sub cp(Val As Long)
setC A < Val
setH (A And 15) < (Val And 15)
setZ A = Val
setN True
End Sub
Sub dec(reg8 As Long)
memptr = reg8
reg8 = reg8 - 1
reg8 = reg8 And 255
setH ((reg8 And 15) < (memptr And 15))
setZ reg8 = 0
setN True
End Sub
Sub dec16(reg81 As Long, reg82 As Long)
reg82 = reg82 - 1
reg81 = reg81 + (reg82 < 0)
reg81 = reg81 And 255
reg82 = reg82 And 255
End Sub
Sub inc(reg8 As Long)
reg8 = reg8 + 1
reg8 = reg8 And 255
setH ((reg8 And 15) = 0)
setZ reg8 = 0
setN False
End Sub
Sub inc16(reg81 As Long, reg82 As Long)
reg82 = reg82 + 1
reg81 = reg81 - (reg82 > 255)
reg81 = reg81 And 255: reg82 = reg82 And 255
End Sub
Sub addHL(r1h As Long, r1l As Long)
setC ((H * 256 + L) + (r1h * 256 + r1l)) > 65535
setH (((H * 256 + L) And 4095) + ((r1h * 256 + r1l) And 4095)) > 4095
L = L + r1l
H = H + r1h
If L > 255 Then H = H + 1
L = L And 255: H = H And 255
setN False
End Sub
Sub addSP(Value As Long)
If Value > 127 Then Value = Value - 256
memptr = sp + Value
memptr = memptr And 65535
        If Value > 0 Then
          setC sp > memptr
          setH ((sp Xor Value Xor memptr) And 4096) > 0
          sp = memptr
        Else
          setC sp < memptr
          setH ((sp Xor Value Xor memptr) And 4096) > 0
          sp = memptr
        End If
        setZ False
        setN False
End Sub
Sub swap(reg8 As Long)
reg8 = (reg8 \ 16) Or ((reg8 And 15) * 16)
setZ reg8 = 0
setN False
setH False
setC False
End Sub
Sub daa()
        If Not GetN Then
           If ((A And 15) >= 10 Or getH) Then A = A + 6
           If ((A And 240) >= 160 Or getC) Then A = A + 96: setC True
        Else
           If ((A And 15) >= 10 Or getH) Then A = A - 6
           If ((A And 240) >= 160 Or getC) Then A = A - 96: setC True
        End If
        A = A And 255
setZ A = 0
setH False
End Sub
Sub cpl()
A = 255 - A
setH True
setN True
End Sub
Sub halt()
   If IME = False Then Exit Sub
    memptr = RAM(65535, 0) And RAM(65295, 0)    ' AND IE, IF
    If memptr = 0 Then pc = pc - 1: Exit Sub                  'If no Interrupt occured exit
    'Process Interrput
    'Push pc
    sp = sp - 1
    WriteM sp, pc \ 256
    sp = sp - 1
    WriteM sp, pc And 255
    IME = False
    If (memptr And 1) = 1 Then        'V-Blank ?
        pc = 64
        RAM(65295, 0) = RAM(65295, 0) And 254
    ElseIf (memptr And 2) = 2 Then    'LCDC ?
        pc = 72
        RAM(65295, 0) = RAM(65295, 0) And 253
    ElseIf (memptr And 4) = 4 Then    'Timer ?
        pc = 80
        RAM(65295, 0) = RAM(65295, 0) And 251
    ElseIf (memptr And 8) = 8 Then    'Serial ?
        pc = 88
        RAM(65295, 0) = RAM(65295, 0) And 247
    ElseIf (memptr And 16) = 16 Then  'Joypad ?
        pc = 96
        RAM(65295, 0) = RAM(65295, 0) And 239
    End If
End Sub
Sub bit(reg8 As Long, b As Long)
setZ ((reg8 And b) = 0)
setN False
setH True
End Sub
Sub zset(reg8 As Long, b As Long)
reg8 = reg8 Or b
End Sub
Sub res(reg8 As Long, b As Long)
reg8 = reg8 And b
End Sub
Sub jp(Adr As Long, Optional cc As Boolean = True)
If cc Then pc = Adr
End Sub
Sub jr(Val As Long, Optional cc As Boolean = True)
If Val > 127 Then Val = Val - 256
If cc Then pc = pc + Val
End Sub
Sub zcall(Adr As Long, Optional cc As Boolean = True)
If cc Then
sp = sp - 1
WriteM sp, pc \ 256
sp = sp - 1
WriteM sp, pc And 255
pc = Adr
End If
End Sub
Sub rst(Value As Long)
sp = sp - 1
WriteM sp, pc \ 256
sp = sp - 1
WriteM sp, pc And 255
pc = Value
End Sub
Sub ret(Optional cc As Boolean = True)
If cc Then pc = readM(sp) + readM(sp + 1) * 256: sp = sp + 2
End Sub
Sub reti()
pc = readM(sp) + readM(sp + 1) * 256
sp = sp + 2
IME = True
End Sub
