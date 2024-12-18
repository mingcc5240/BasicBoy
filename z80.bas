Attribute VB_Name = "z80"
'This is a part of tha BasicBoy emulator
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel).
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'(I know the emulator is NOT OPTIMIZED AT ALL)

'This module is the main part of the z80 emulation
'Some Subs are also at the z80cmd module
'Coments will be added with the next release

'Sory for my bad english ...


Option Explicit
Public A As Long
Public b As Long
Public C As Long
Public D As Long
Public E As Long
Public H As Long
Public L As Long
Public F As Long
Public pc As Long
Public sp As Long
Public IME As Boolean
Private memval As Long
Private memptr As Long
Private temp As Long, temp2 As Long
Public brkAddr As Long
Private i As Long
Global Clcount As Long
Global lcdcline As Long
Global ime_stat As Long
Global SETT(0 To 56) As Byte, BITT(0 To 56) As Byte
Global CpuRun As Boolean, jv1 As Long, jv2 As Long, tval As Long, tvm As Long, cpc(&HFF) As Long, lw As Long
Global mm As Long, zm As Long, cm As Long, utu As Boolean
Global GBM As Long '0 = gb, 1 = gbc, 2 = gba
Global message As MSG, TGBC As Boolean, smp As Long, bCpuRun As Boolean
Global Clm0 As Long, clm3 As Long, cllc As Long, cldr As Long, drw As Long, cpuS As Long
Global lfp As Boolean, sta As Long, hline As Long, Mhz As Long


Global m_DividerVariable As Long
Global m_CurrentClockSpeed As Long
Global m_TimerVariable As Long


Sub setZ(Val As Boolean)
If Val Then F = F Or 128 Else F = F And 127
End Sub
Sub setN(Val As Boolean)
If Val Then F = F Or 64 Else F = F And 191
End Sub
Sub setH(Val As Boolean)
If Val Then F = F Or 32 Else F = F And 223
End Sub
Sub setC(Val As Boolean)
If Val Then F = F Or 16 Else F = F And 239
End Sub

Function getZ() As Boolean
getZ = F And 128
End Function
Function GetN() As Boolean
GetN = F And 64
End Function
Function getH() As Boolean
getH = F And 32
End Function
Function getC() As Boolean
getC = F And 16
End Function

Public Sub reset()
QueryPerformanceFrequency curFreq  'Get the timer frequency
curFreq = curFreq / 1000 ' in ms
hline = -1
smp = 0
IME = True
If GBM = 0 Then A = &H1 Else A = &H11
'- GB \ SGB , &hFF-GBP, &h11-GBC
F = &HB0
b = 0
C = &H13
D = 0
E = &HD8
H = 1
L = &H4D
pc = &H100
sp = 65534

cldr = 255
Clm0 = 251
clm3 = 79
cllc = 455
cpuS = 0
WriteM 65285, &H0   ' TIMA
WriteM 65286, &H0   ' TMA
WriteM 65287, &H0   ' TAC
WriteM 65296, &H80  ' NR10
WriteM 65297, &HBF  ' NR11
WriteM 65298, &HF3  ' NR12
WriteM 65300, &HBF  ' NR14
WriteM 65302, &H3F  ' NR21
WriteM 65303, &H0   ' NR22
WriteM 65305, &HBF  ' NR24
WriteM 65306, &H7F  ' NR30
WriteM 65307, &HFF  ' NR31
WriteM 65308, &H9F  ' NR32
WriteM 65310, &HBF  ' NR33
WriteM 65312, &HFF  ' NR41
WriteM 65313, &H0   ' NR42
WriteM 65314, &H0   ' NR43
WriteM 65315, &HBF  ' NR30
WriteM 65316, &H77  ' NR50
WriteM 65317, &HF3  ' NR51
WriteM 65318, &HF1  '- GB, &HF0 - SGB ' NR52
WriteM 65344, &H91  ' LCDC
WriteM 65346, &H0   ' SCY
WriteM 65347, &H0   ' SCX
WriteM 65349, &H0   ' LYC
WriteM 65351, &HE4  ' BGP
WriteM 65352, &HE4  ' OBP0
WriteM 65353, &HE4  ' OBP1
WriteM 65354, &H0   ' WY
WriteM 65355, &H0   ' WX
WriteM 65535, &H0   ' IE

 utu = True
End Sub


Public Sub saveS(filenum As String)
'Save Stage
End Sub
Public Sub loadS(filenum As String)
'Load Stage
End Sub

Sub IntReq()
    If IME = False Then Exit Sub
    temp = RAM(65535, 0) And RAM(65295, 0)    ' AND IE, IF
    If temp = 0 Then Exit Sub                       'If no Interrupt occured exit
    'Process Interrput
    'Push pc
    sp = sp - 1
    WriteM sp, pc \ 256
    sp = sp - 1
    WriteM sp, pc And 255
    IME = False
    If (temp And 1) = 1 Then        'V-Blank ?
        pc = 64
        RAM(65295, 0) = RAM(65295, 0) And 254
    ElseIf (temp And 2) = 2 Then    'LCDC ?
        pc = 72
        RAM(65295, 0) = RAM(65295, 0) And 253
    ElseIf (temp And 4) = 4 Then    'Timer ?
        pc = 80
        RAM(65295, 0) = RAM(65295, 0) And 251
    ElseIf (temp And 8) = 8 Then    'Serial ?
        pc = 88
        RAM(65295, 0) = RAM(65295, 0) And 247
    ElseIf (temp And 16) = 16 Then  'Joypad ?
        pc = 96
        RAM(65295, 0) = RAM(65295, 0) And 239
    End If
End Sub

Public Sub RunCpu2()
If lfp Then
QueryPerformanceCounter curStart 'Get the start time
End If
brkAddr = -1
'run @ 4mhz(4194304 hrz)
'60Fps(~59.7)(~70224 hrz per frame(~4194304 hrz))
cldr = 255
cllc = 455
Clm0 = cllc - 204
clm3 = cllc - 376
bCpuRun = True

While bCpuRun = True ' keep cpu running
RunCycles2

'Do interups
IntReq

If RAM(65344, 0) And 128 Then
If Clcount > clm3 And RAM(65348, 0) < 144 Then
'80
clm3 = clm3 + 456 + 456 * cpuS
'set stat mode 3
If Skipf = False Then If GBM Then Drawline Else Drawline4
RAM(65345, 0) = (RAM(65345, 0) And 252) Or 3
ElseIf Clcount > Clm0 And RAM(65348, 0) < 144 Then
'252
Clm0 = Clm0 + 456 + 456 * cpuS
'set h-blank

hline = RAM(65348, 0)
RAM(65345, 0) = RAM(65345, 0) And 252
If Hdma = True Then
For i = hdmaS To hdmaS + 15
        RAM(hdmaD, vRamB) = readM(i)
        hdmaD = hdmaD + 1
Next i
hdmaS = hdmaS + 16
Hdmal = Hdmal - 1
If Hdmal = -1 Then Hdma = False: RAM(65365, 0) = 255 Else RAM(65365, 0) = Hdmal
End If

'stat h-blank int
If RAM(65345, 0) And 8 Then RAM(65295, 0) = RAM(65295, 0) Or 2

End If


If Clcount > cllc Then
'456
cllc = cllc + 456 + 456 * cpuS
' Increment Line Counter
RAM(65348, 0) = (RAM(65348, 0) + 1) Mod 154

If RAM(65348, 0) = 144 Then
    If RAM(65344, 0) And 128 Then
    DrawScreen
    QueryPerformanceCounter curStart 'Get the start time
    End If
End If
'ly=lyc
If RAM(65348, 0) = RAM(65349, 0) Then
'ly=lyc int
If RAM(65345, 0) And 64 Then RAM(65295, 0) = RAM(65295, 0) Or 2
'set ly=lyc
RAM(65345, 0) = RAM(65345, 0) Or 4
Else
'reset ly=lyc
RAM(65345, 0) = RAM(65345, 0) And 251
End If

'check h-blank,v-blank ,ect
If RAM(65348, 0) < 144 Then
'set mode 2
RAM(65345, 0) = (RAM(65345, 0) And 252) Or 2
'stat mode 2 int
If RAM(65345, 0) And 32 Then RAM(65295, 0) = RAM(65295, 0) Or 2
ElseIf RAM(65348, 0) = 144 Then
'set v-blank (mode 1)
RAM(65345, 0) = (RAM(65345, 0) And 252) Or 1
'v-blank int
RAM(65295, 0) = RAM(65295, 0) Or 1
'stat mode 1 int too
If RAM(65345, 0) And 16 Then RAM(65295, 0) = RAM(65295, 0) Or 2
End If 'hck

End If 'mod456

Else
If Clcount > cllc Then cllc = cllc + 456 + 456 * cpuS
If Hdma = True Then
For i = hdmaS To hdmaS + 15
        RAM(hdmaD, vRamB) = readM(i)
        hdmaD = hdmaD + 1
Next i
hdmaS = hdmaS + 16
Hdmal = Hdmal - 1
If Hdmal = -1 Then Hdma = False: RAM(65365, 0) = 255 Else RAM(65365, 0) = Hdmal
End If

End If



If Clcount > cldr Then
    '256
    cldr = cldr + 456
    'Inc divreg
    RAM(65284, 0) = (RAM(65284, 0) + 1) And 255
End If


If Clcount > 70223 Then
Mhz = Mhz + 1
frmCheat.ChkCheats
Clcount = Clcount - 70224
'draw screen
cllc = cllc - 70224
Clm0 = cllc - 204 - 204 * cpuS
clm3 = cllc - 376 - 376 * cpuS

'Form1.di.Check_Keyboard
    If PeekMessage(message, 0&, 0&, 0&, PM_REMOVE) Then
        Call TranslateMessage(message)
        Call DispatchMessage(message)
    End If
End If


extif:
Wend
End Sub


Sub utimer(cycles As Long)
 Dim timerAtts As Long
 Dim overflow As Boolean
 Dim tem As Long
 
    timerAtts = readM(65287) '0xFF07
 
    m_DividerVariable = m_DividerVariable + cycles
    
    If (timerAtts And 4) Then
        m_TimerVariable = m_TimerVariable + cycles
        
        If (m_TimerVariable >= m_CurrentClockSpeed) Then
            m_TimerVariable = 0
            overflow = False
            tem = readM(65285) '0xff05
            If (tem = 255) Then overflow = True
            RAM(65285, 0) = RAM(65285, 0) + 1
             
            If (overflow) Then
                RAM(65285, 0) = RAM(65286, 0)
                tem = readM(65295) '0xff0f
                tem = tem Or 4   'interrupt routine
                WriteM 65295, tem
            End If
        End If
    
    End If
    
    If (m_DividerVariable >= 256) Then
        m_DividerVariable = 0
        RAM(65284, 0) = RAM(65284, 0) + 1
    End If
End Sub

Public Sub InitCPU()
    cpc(&H0) = 1:  cpc(&H1) = 3:  cpc(&H2) = 2:  cpc(&H3) = 2
    cpc(&H4) = 1:  cpc(&H5) = 1:  cpc(&H6) = 2:  cpc(&H7) = 1
    cpc(&H8) = 5:  cpc(&H9) = 2:  cpc(&HA) = 2:  cpc(&HB) = 2
    cpc(&HC) = 1:  cpc(&HD) = 1:  cpc(&HE) = 2:  cpc(&HF) = 1
    
    cpc(&H10) = 1: cpc(&H11) = 3: cpc(&H12) = 2: cpc(&H13) = 2
    cpc(&H14) = 1: cpc(&H15) = 1: cpc(&H16) = 2: cpc(&H17) = 1
    cpc(&H18) = 3: cpc(&H19) = 2: cpc(&H1A) = 2: cpc(&H1B) = 2
    cpc(&H1C) = 1: cpc(&H1D) = 1: cpc(&H1E) = 2: cpc(&H1F) = 1
    
    cpc(&H20) = 2: cpc(&H21) = 3: cpc(&H22) = 2: cpc(&H23) = 2
    cpc(&H24) = 1: cpc(&H25) = 1: cpc(&H26) = 2: cpc(&H27) = 1
    cpc(&H28) = 2: cpc(&H29) = 2: cpc(&H2A) = 2: cpc(&H2B) = 2
    cpc(&H2C) = 1: cpc(&H2D) = 1: cpc(&H2E) = 2: cpc(&H2F) = 1
    
    cpc(&H30) = 2: cpc(&H31) = 3: cpc(&H32) = 2: cpc(&H33) = 2
    cpc(&H34) = 3: cpc(&H35) = 3: cpc(&H36) = 3: cpc(&H37) = 1
    cpc(&H38) = 2: cpc(&H39) = 2: cpc(&H3A) = 2: cpc(&H3B) = 2
    cpc(&H3C) = 1: cpc(&H3D) = 1: cpc(&H3E) = 2: cpc(&H3F) = 1
    
    cpc(&H40) = 1: cpc(&H41) = 1: cpc(&H42) = 1: cpc(&H43) = 1
    cpc(&H44) = 1: cpc(&H45) = 1: cpc(&H46) = 2: cpc(&H47) = 1
    cpc(&H48) = 1: cpc(&H49) = 1: cpc(&H4A) = 1: cpc(&H4B) = 1
    cpc(&H4C) = 1: cpc(&H4D) = 1: cpc(&H4E) = 2: cpc(&H4F) = 1
    
    cpc(&H50) = 1: cpc(&H51) = 1: cpc(&H52) = 1: cpc(&H53) = 1
    cpc(&H54) = 1: cpc(&H55) = 1: cpc(&H56) = 2: cpc(&H57) = 1
    cpc(&H58) = 1: cpc(&H59) = 1: cpc(&H5A) = 1: cpc(&H5B) = 1
    cpc(&H5C) = 1: cpc(&H5D) = 1: cpc(&H5E) = 2: cpc(&H5F) = 1
    
    cpc(&H60) = 1: cpc(&H61) = 1: cpc(&H62) = 1: cpc(&H63) = 1
    cpc(&H64) = 1: cpc(&H65) = 1: cpc(&H66) = 2: cpc(&H67) = 1
    cpc(&H68) = 1: cpc(&H69) = 1: cpc(&H6A) = 1: cpc(&H6B) = 1
    cpc(&H6C) = 1: cpc(&H6D) = 1: cpc(&H6E) = 2: cpc(&H6F) = 1
    
    cpc(&H70) = 2: cpc(&H71) = 2: cpc(&H72) = 2: cpc(&H73) = 2
    cpc(&H74) = 2: cpc(&H75) = 2: cpc(&H76) = 1: cpc(&H77) = 2
    cpc(&H78) = 1: cpc(&H79) = 1: cpc(&H7A) = 1: cpc(&H7B) = 1
    cpc(&H7C) = 1: cpc(&H7D) = 1: cpc(&H7E) = 2: cpc(&H7F) = 1
    
    cpc(&H80) = 1: cpc(&H81) = 1: cpc(&H82) = 1: cpc(&H83) = 1
    cpc(&H84) = 1: cpc(&H85) = 1: cpc(&H86) = 2: cpc(&H87) = 1
    cpc(&H88) = 1: cpc(&H89) = 1: cpc(&H8A) = 1: cpc(&H8B) = 1
    cpc(&H8C) = 1: cpc(&H8D) = 1: cpc(&H8E) = 2: cpc(&H8F) = 1
    
    cpc(&H90) = 1: cpc(&H91) = 1: cpc(&H92) = 1: cpc(&H93) = 1
    cpc(&H94) = 1: cpc(&H95) = 1: cpc(&H96) = 2: cpc(&H97) = 1
    cpc(&H98) = 1: cpc(&H99) = 1: cpc(&H9A) = 1: cpc(&H9B) = 1
    cpc(&H9C) = 1: cpc(&H9D) = 1: cpc(&H9E) = 2: cpc(&H9F) = 1
    
    cpc(&HA0) = 1: cpc(&HA1) = 1: cpc(&HA2) = 1: cpc(&HA3) = 1
    cpc(&HA4) = 1: cpc(&HA5) = 1: cpc(&HA6) = 2: cpc(&HA7) = 1
    cpc(&HA8) = 1: cpc(&HA9) = 1: cpc(&HAA) = 1: cpc(&HAB) = 1
    cpc(&HAC) = 1: cpc(&HAD) = 1: cpc(&HAE) = 2: cpc(&HAF) = 1
    
    cpc(&HB0) = 1: cpc(&HB1) = 1: cpc(&HB2) = 1: cpc(&HB3) = 1
    cpc(&HB4) = 1: cpc(&HB5) = 1: cpc(&HB6) = 2: cpc(&HB7) = 1
    cpc(&HB8) = 1: cpc(&HB9) = 1: cpc(&HBA) = 1: cpc(&HBB) = 1
    cpc(&HBC) = 1: cpc(&HBD) = 1: cpc(&HBE) = 2: cpc(&HBF) = 1
    
    cpc(&HC0) = 2: cpc(&HC1) = 3: cpc(&HC2) = 3: cpc(&HC3) = 4
    cpc(&HC4) = 3: cpc(&HC5) = 4: cpc(&HC6) = 2: cpc(&HC7) = 4
    cpc(&HC8) = 2: cpc(&HC9) = 4: cpc(&HCA) = 3: cpc(&HCB) = 2
    cpc(&HCC) = 3: cpc(&HCD) = 6: cpc(&HCE) = 2: cpc(&HCF) = 4
    
    cpc(&HD0) = 2: cpc(&HD1) = 3: cpc(&HD2) = 3: cpc(&HD3) = 0
    cpc(&HD4) = 3: cpc(&HD5) = 4: cpc(&HD6) = 2: cpc(&HD7) = 4
    cpc(&HD8) = 2: cpc(&HD9) = 4: cpc(&HDA) = 3: cpc(&HDB) = 0
    cpc(&HDC) = 3: cpc(&HDD) = 0: cpc(&HDE) = 2: cpc(&HDF) = 4
    
    cpc(&HE0) = 3: cpc(&HE1) = 3: cpc(&HE2) = 2: cpc(&HE3) = 0
    cpc(&HE4) = 0: cpc(&HE5) = 4: cpc(&HE6) = 2: cpc(&HE7) = 4
    cpc(&HE8) = 4: cpc(&HE9) = 1: cpc(&HEA) = 4: cpc(&HEB) = 0
    cpc(&HEC) = 0: cpc(&HED) = 0: cpc(&HEE) = 2: cpc(&HEF) = 4
    
    cpc(&HF0) = 3: cpc(&HF1) = 3: cpc(&HF2) = 2: cpc(&HF3) = 1
    cpc(&HF4) = 0: cpc(&HF5) = 4: cpc(&HF6) = 2: cpc(&HF7) = 4
    cpc(&HF8) = 3: cpc(&HF9) = 2: cpc(&HFA) = 4: cpc(&HFB) = 1
    cpc(&HFC) = 0: cpc(&HFD) = 0: cpc(&HFE) = 2: cpc(&HFF) = 4
    For i = 0 To 255
    cpc(i) = cpc(i) * 4
    Next i
End Sub

Sub checkregs()
If A > 255 Or b > 255 Or C > 255 Or D > 255 Or E > 255 Or F > 255 Or H > 255 Or L > 255 Or sp > 65535 Or pc > 65535 Or _
        A < 0 Or b < 0 Or C < 0 Or D < 0 Or E < 0 Or F < 0 Or H < 0 Or L < 0 Or sp < 0 Or pc < 0 Then
Stop
End If

End Sub
Sub RunCycles2()

memval = pb
Clcount = Clcount + cpc(memval)
If utu Then utimer cpc(memval)

If ime_stat = 3 Then IME = True: ime_stat = 0
If ime_stat = 4 Then IME = False: ime_stat = 0
If ime_stat = 1 Then ime_stat = 3
If ime_stat = 2 Then ime_stat = 4

Select Case memval
Case 0 '     ' NOP

Case 1     '  'LD BC, nnnn
C = pb
b = pb
Case 2     '  LD (BC), a
WriteM b * 256 + C, A
Case 3     '  INC BC
inc16 b, C
Case 4     '  INC b
inc b
Case 5     '  DEC b
dec b
Case 6     '  LD b, nn
b = pb
Case 7     '  RLCA
rlca
Case 8      ' LD     '(nnnn),SP     ' ---- special (old ex af,af)
memptr = pw
WriteM memptr, sp And 255
WriteM memptr + 1, sp \ 256
Case 9     '  Add HL, BC
addHL b, C
Case &HA     ' LD     'A,(BC)
A = readM(b * 256 + C)
Case &HB     ' DEC  BC
dec16 b, C
Case &HC    ' INC  C
inc C
Case &HD      ' DEC  C
dec C
Case &HE      ' LD     'C,nn
C = pb
Case &HF  'RRCA
rrca
Case &H10 '00 STOP     '     '     '     '  ---- special ??? (old djnz disp)
If smp = 0 Then
    If mem.joyval1 = 0 And mem.joyval2 = 0 Then pc = pc - 1 Else pc = pc + 1
Else
    If cpuS = 0 Then cpuS = 1 Else cpuS = 0
    RAM(65357, 0) = cpuS * 128
    smp = 0
End If
Case &H11     ' LD DE, nnnn
E = pb
D = pb
Case &H12     ' LD (DE), a
WriteM D * 256 + E, A
Case &H13     ' INC DE
inc16 D, E
Case &H14     ' INC d
inc D
Case &H15     ' DEC d
dec D
Case &H16     ' LD d, nn
D = pb
Case &H17     ' RLA
rla
Case &H18     ' JR disp
jr pb
Case &H19     ' Add HL, DE
addHL D, E
Case &H1A    ' LD     'A,(DE)
A = readM(D * 256 + E)
Case &H1B    ' DEC  DE
dec16 D, E
Case &H1C    ' INC  E
inc E
Case &H1D     ' DEC  E
dec E
Case &H1E     ' LD     'E,nn
E = pb
Case &H1F     'RRA
rra
Case &H20     ' JR nz, disp
jr pb, Not getZ
Case &H21      ' LD HL, nnnn
L = pb
H = pb
Case &H22     ' LDI  (HL),A     '     ' ---- special (old ld (nnnn),hl)
WriteM H * 256 + L, A
inc16 H, L
Case &H23     ' INC HL
inc16 H, L
Case &H24     ' INC H
inc H
Case &H25     ' DEC H
dec H
Case &H26     ' LD H, nn
H = pb
Case &H27     ' DAA
daa
Case &H28     ' JR z, disp
jr pb, getZ
Case &H29     ' Add HL, HL
addHL H, L
Case &H2A    ' LDI  A,(HL)     '     ' ---- special (old ld hl,(nnnn))
A = readM(H * 256 + L)
inc16 H, L
Case &H2B    ' DEC  HL
dec16 H, L
Case &H2C    ' INC  L
inc L
Case &H2D     ' DEC  L
dec L
Case &H2E     ' LD     'L,nn
L = pb
Case &H2F 'CPL
cpl
Case &H30     ' JR NC, disp
jr pb, Not getC
Case &H31     ' LD sp, nnnn
sp = pw
Case &H32     ' LDD  (HL),A     '     ' ---- special (old remapped ld (nnnn),a)
WriteM H * 256 + L, A
dec16 H, L
Case &H33     ' INC sp
sp = sp + 1
sp = sp And 65535
Case &H34     ' INC (HL)
memptr = readM(H * 256 + L)
inc memptr
WriteM H * 256 + L, memptr
Case &H35     ' DEC (HL)
memptr = readM(H * 256 + L)
dec memptr
WriteM H * 256 + L, memptr
Case &H36     ' LD (HL), nn
WriteM H * 256 + L, pb
Case &H37     ' SCF
setC True
setN False
setH False
Case &H38     ' JR c, disp
jr pb, getC
Case &H39      ' Add HL, sp
addHL sp \ 256, sp And 256
Case &H3A    ' LDD  A,(HL)     '     ' ---- special (old remapped ld a,(nnnn))
A = readM(H * 256 + L)
dec16 H, L
Case &H3B    ' DEC  SP
sp = sp - 1
If sp = -1 Then sp = 65535
Case &H3C    ' INC  A
inc A
Case &H3D     ' DEC  A
dec A
Case &H3E     ' LD     'A,nn
A = pb
Case &H3F 'CCF
setC Not getC
setN False
setH False
Case &H40     ' LD     'B,B     '     '     '     '     '     '     '     '
'Stop 'nop
Case &H60     ' LD     'H,B
H = b
Case &H41     ' LD     'B,C     '     '     '     '     '     '     '     '
b = C
Case &H61     ' LD     'H,C
H = C
Case &H42     ' LD     'B,D     '     '     '     '     '     '     '     '
b = D
Case &H62     ' LD     'H,D
H = D
Case &H43     ' LD     'B,E     '     '     '     '     '     '     '     '
b = E
Case &H63     ' LD     'H,E
H = E
Case &H44     ' LD     'B,H     '     '     '     '     '     '     '     '
b = H
Case &H64     ' LD     'H,H
'Stop 'nop
Case &H45     ' LD     'B,L     '     '     '     '     '     '     '     '
b = L
Case &H65     ' LD     'H,L
H = L
Case &H46     ' LD     'B,(HL)     '     '     '     '     '     '     '
b = readM(H * 256 + L)
Case &H66     ' LD     'H,(HL)
H = readM(H * 256 + L)
Case &H47     ' LD     'B,A     '     '     '     '     '     '     '     '
b = A
Case &H67     ' LD     'H,A
H = A
Case &H48     ' LD     'C,B     '     '     '     '     '     '     '     '
C = b
Case &H68     ' LD     'L,B
L = b
Case &H49     ' LD     'C,C     '     '     '     '     '     '     '     '
'Stop 'nop
Case &H69     ' LD     'L,C
L = C
Case &H4A    ' LD     'C,D     '     '     '     '     '     '     '     '
C = D
Case &H6A     ' LD     'L,D
L = D
Case &H4B    ' LD     'C,E     '     '     '     '     '     '     '     '
C = E
Case &H6B     ' LD     'L,E
L = E
Case &H4C    ' LD     'C,H     '     '     '     '     '     '     '     '
C = H
Case &H6C     ' LD     'L,H
L = H
Case &H4D     ' LD     'C,L     '     '     '     '     '     '     '     '
C = L
Case &H6D     ' LD     'L,L
'Stop 'nop
Case &H4E     ' LD     'C,(HL)     '     '     '     '     '     '     '
C = readM(H * 256 + L)
Case &H6E     ' LD     'L,(HL)
L = readM(H * 256 + L)
Case &H4F    ' LD     'C,A     '     '     '     '     '     '     '     '
C = A
Case &H6F     ' LD     'L,A
L = A
Case &H50     ' LD     'D,B     '     '     '     '     '     '     '     '
D = b
Case &H70     ' LD     '(HL),B
WriteM H * 256 + L, b
Case &H51     ' LD     'D,C     '     '     '     '     '     '     '     '
D = C
Case &H71     ' LD     '(HL),C
WriteM H * 256 + L, C
Case &H52     ' LD     'D,D     '     '     '     '     '     '     '     '
'Stop 'nop
Case &H72     ' LD     '(HL),D
WriteM (H * 256 + L), D
Case &H53     ' LD     'D,E     '     '     '     '     '     '     '     '
D = E
Case &H73     ' LD     '(HL),E
WriteM (H * 256 + L), E
Case &H54     ' LD     'D,H     '     '     '     '     '     '     '     '
D = H
Case &H74     ' LD     '(HL),H
WriteM (H * 256 + L), H
Case &H55     ' LD     'D,L     '     '     '     '     '     '     '     '
D = L
Case &H75     ' LD     '(HL),L
WriteM (H * 256 + L), L
Case &H56     ' LD     'D,(HL)     '     '     '     '     '     '     '
D = readM(H * 256 + L)
Case &H76     ' HALT
halt
Case &H57     ' LD     'D,A     '     '     '     '     '     '     '     '
D = A
Case &H77     ' LD     '(HL),A
WriteM (H * 256 + L), A
Case &H58     ' LD     'E,B     '     '     '     '     '     '     '     '
E = b
Case &H78     ' LD     'A,B
A = b
Case &H59     ' LD     'E,C     '     '     '     '     '     '     '     '
E = C
Case &H79     ' LD     'A,C
A = C
Case &H5A    ' LD     'E,D     '     '     '     '     '     '     '     '
E = D
Case &H7A     ' LD     'A,D
A = D
Case &H5B    ' LD     'E,E     '     '     '     '     '     '     '     '
'Stop 'nop
Case &H7B     ' LD     'A,E
A = E
Case &H5C    ' LD     'E,H     '     '     '     '     '     '     '     '
E = H
Case &H7C     ' LD     'A,H
A = H
Case &H5D     ' LD     'E,L     '     '     '     '     '     '     '     '
E = L
Case &H7D     ' LD     'A,L
A = L
Case &H5E     ' LD     'E,(HL)     '     '     '     '     '     '     '
E = readM(H * 256 + L)
Case &H7E     ' LD     'A,(HL)
A = readM(H * 256 + L)
Case &H5F    ' LD     'E,A     '     '     '     '     '     '     '     '
E = A
Case &H7F     ' LD     'A,A
'Stop 'nop
Case &H80     ' ADD  A,B     '     '     '     '     '     '     '     '
add b
Case &HA0     ' AND  B
zand b
Case &H81     ' ADD  A,C     '     '     '     '     '     '     '     '
add C
Case &HA1     ' AND  C
zand C
Case &H82     ' ADD  A,D     '     '     '     '     '     '     '     '
add D
Case &HA2     ' AND  D
zand D
Case &H83     ' ADD  A,E     '     '     '     '     '     '     '     '
add E
Case &HA3     ' AND  E
zand E
Case &H84     ' ADD  A,H     '     '     '     '     '     '     '     '
add H
Case &HA4     ' AND  H
zand H
Case &H85     ' ADD  A,L     '     '     '     '     '     '     '     '
add L
Case &HA5     ' AND  L
zand L
Case &H86     ' ADD  A,(HL)     '     '     '     '     '     '     '
add readM(H * 256 + L)
Case &HA6     ' AND  (HL)
zand readM(H * 256 + L)
Case &H87     ' ADD  A,A     '     '     '     '     '     '     '     '
add A
Case &HA7     ' AND  A
zand A
Case &H88     ' ADC  A,B     '     '     '     '     '     '     '     '
adc b
Case &HA8     ' XOR  B
zxor b
Case &H89     ' ADC  A,C     '     '     '     '     '     '     '     '
adc C
Case &HA9     ' XOR  C
zxor C
Case &H8A    ' ADC  A,D     '     '     '     '     '     '     '     '
adc D
Case &HAA     ' XOR  D
zxor D
Case &H8B    ' ADC  A,E     '     '     '     '     '     '     '     '
adc E
Case &HAB     ' XOR  E
zxor E
Case &H8C    ' ADC  A,H     '     '     '     '     '     '     '     '
adc H
Case &HAC     ' XOR  H
zxor H
Case &H8D     ' ADC  A,L     '     '     '     '     '     '     '     '
adc L
Case &HAD     ' XOR  L
zxor L
Case &H8E     ' ADC  A,(HL)     '     '     '     '     '     '     '
adc readM(H * 256 + L)
Case &HAE     ' XOR  (HL)
zxor readM(H * 256 + L)
Case &H8F    ' ADC  A,A     '     '     '     '     '     '     '     '
adc A
Case &HAF     ' XOR  A
zxor A
Case &H90     ' SUB  B     '     '     '     '     '     '     '     '     '
zsub b
Case &HB0     ' OR     'B
zor b
Case &H91     ' SUB  C     '     '     '     '     '     '     '     '     '
zsub C
Case &HB1     ' OR     'C
zor C
Case &H92     ' SUB  D     '     '     '     '     '     '     '     '     '
zsub D
Case &HB2     ' OR     'D
zor D
Case &H93     ' SUB  E     '     '     '     '     '     '     '     '     '
zsub E
Case &HB3     ' OR     'E
zor E
Case &H94     ' SUB  H     '     '     '     '     '     '     '     '     '
zsub H
Case &HB4     ' OR     'H
zor H
Case &H95     ' SUB  L     '     '     '     '     '     '     '     '     '
zsub L
Case &HB5     ' OR     'L
zor L
Case &H96     ' SUB  (HL)     '     '     '     '     '     '     '     '
zsub readM(H * 256 + L)
Case &HB6     ' OR     '(HL)
zor readM(H * 256 + L)
Case &H97     ' SUB  A     '     '     '     '     '     '     '     '     '
zsub A
Case &HB7     ' OR     'A
zor A
Case &H98     ' SBC  A,B     '     '     '     '     '     '     '     '
sbc b
Case &HB8     ' CP     'B
cp b
Case &H99     ' SBC  A,C     '     '     '     '     '     '     '     '
sbc C
Case &HB9     ' CP     'C
cp C
Case &H9A    ' SBC  A,D     '     '     '     '     '     '     '     '
sbc D
Case &HBA     ' CP     'D
cp D
Case &H9B    ' SBC  A,E     '     '     '     '     '     '     '     '
sbc E
Case &HBB     ' CP     'E
cp E
Case &H9C    ' SBC  A,H     '     '     '     '     '     '     '     '
sbc H
Case &HBC     ' CP     'H
cp H
Case &H9D     ' SBC  A,L     '     '     '     '     '     '     '     '
sbc L
Case &HBD     ' CP     'L
cp L
Case &H9E     ' SBC  A,(HL)     '     '     '     '     '     '     '
sbc readM(H * 256 + L)
Case &HBE     ' CP     '(HL)
cp readM(H * 256 + L)
Case &H9F    ' SBC  A,A     '     '     '     '     '     '     '     '
sbc A
Case &HBF     ' CP     'A
cp A
Case &HC0     ' RET  NZ
ret Not getZ
Case &HC1     ' POP  BC
pop C
pop b
Case &HC2     ' JP     'NZ,nnnn
jp pw, Not getZ
Case &HC3     ' JP     'nnnn
pc = pw
Case &HC4     ' CALL NZ,nnnn
zcall pw, Not getZ
Case &HC5     ' PUSH BC
push b
push C
Case &HC6     ' ADD  A,nn
add pb
Case &HC7     ' RST  00H
rst 0
Case &HC8     ' RET  Z
ret getZ
Case &HC9 'RET
ret
Case &HCA     ' JP     'Z,nnnn
jp pw, getZ
Case &HCB 'nn ---(see beyond)---
memval = pb
'Put 99, , CByte(memval)
'If cbt(memval) = False Then: cbt(memval) = True: Stop
Select Case memval
    Case &H0   'RLC  B
    rlc b
    Case &H1   'RLC  C
    rlc C
    Case &H2   'RLC  D
    rlc D
    Case &H3   'RLC  E
    rlc E
    Case &H4   'RLC  H
    rlc H
    Case &H5   'RLC  L
    rlc L
    Case &H6   'RLC  (HL)
    memptr = readM(H * 256 + L)
    rlc memptr
    WriteM H * 256 + L, memptr
    Case &H7   'RLC  A
    rlc A
    Case &H8   'RRC  B
    rrc b
    Case &H9   'RRC  C
    rrc C
    Case &HA   'RRC  D
    rrc D
    Case &HB   'RRC  E
    rrc E
    Case &HC   'RRC  H
    rrc H
    Case &HD   'RRC  L
    rrc L
    Case &HE   'RRC  (HL)
    memptr = readM(H * 256 + L)
    rrc memptr
    WriteM H * 256 + L, memptr
    Case &HF   'RRC  A
    rrc A
    Case &H10  'RL     'B
    rl b
    Case &H11  'RL     'C
    rl C
    Case &H12  'RL     'D
    rl D
    Case &H13  'RL     'E
    rl E
    Case &H14  'RL     'H
    rl H
    Case &H15  'RL     'L
    rl L
    Case &H16  'RL     '(HL)
    memptr = readM(H * 256 + L)
    rl memptr
    WriteM H * 256 + L, memptr
    Case &H17  'RL     'A
    rl A
    Case &H18  'RR     'B
    rr b
    Case &H19  'RR     'C
    rr C
    Case &H1A  'RR     'D
    rr D
    Case &H1B  'RR     'E
    rr E
    Case &H1C  'RR     'H
    rr H
    Case &H1D  'RR     'L
    rr L
    Case &H1E  'RR     '(HL)
    memval = readM(H * 256 + L)
    rr memval
    WriteM H * 256 + L, memval
    Case &H1F  'RR     'A
    rr A
    Case &H20  'SLA  B
    sla b
    Case &H21  'SLA  C
    sla C
    Case &H22  'SLA  D
    sla D
    Case &H23  'SLA  E
    sla E
    Case &H24        'SLA  H
    sla H
    Case &H25  'SLA  L
    sla L
    Case &H26  'SLA  (HL)
    memval = readM(H * 256 + L)
    sla memval
    WriteM H * 256 + L, memval
    Case &H27  'SLA  A
    sla A
    Case &H28  'SRA  B
    sra b
    Case &H29  'SRA  C
    sra C
    Case &H2A  'SRA  D
    sra D
    Case &H2B  'SRA  E
    sra E
    Case &H2C  'SRA  H
    sra H
    Case &H2D  'SRA  L
    sra L
    Case &H2E  'SRA  (HL)
    memval = readM(H * 256 + L)
    sra memval
    WriteM H * 256 + L, memval
    Case &H2F  'SRA  A
    sra A
    Case &H30  'SWAP B     '     '     '  ---- special (old sll)
    swap b
    Case &H31  'SWAP C     '     '     '  ---- special ""
    swap C
    Case &H32  'SWAP D     '     '     '  ---- special ""
    swap D
    Case &H33  'SWAP E     '     '     '  ---- special ""
    swap E
    Case &H34  'SWAP H     '     '     '  ---- special ""
    swap H
    Case &H35  'SWAP L     '     '     '  ---- special ""
    swap L
    Case &H36  'SWAP (HL)     '     '  ---- special ""
    memval = readM(H * 256 + L)
    swap memval
    WriteM H * 256 + L, memval
    Case &H37  'SWAP A     '     '     '  ---- special ""
    swap A
    Case &H38  'SRL  B
    srl b
    Case &H39  'SRL  C
    srl C
    Case &H3A  'SRL  D
    srl D
    Case &H3B  'SRL  E
    srl E
    Case &H3C  'SRL  H
    srl H
    Case &H3D  'SRL  L
    srl L
    Case &H3E  'SRL  (HL)
    memval = readM(H * 256 + L)
    srl memval
    WriteM H * 256 + L, memval
    Case &H3F  'SRL  A
    srl A
    Case Else
    Select Case memval And 199
    Case &H40 '+n*38  BIT  n,B
    bit b, (BITT(memval And 56))
    Case &H41 '+n*38  BIT  n,C
    bit C, (BITT(memval And 56))
    Case &H42 '+n*38  BIT  n,D
    bit D, (BITT(memval And 56))
    Case &H43 '+n*38  BIT  n,E
    bit E, (BITT(memval And 56))
    Case &H44 '+n*38  BIT  n,H
    bit H, (BITT(memval And 56))
    Case &H45 '+n*38  BIT  n,L
    bit L, (BITT(memval And 56))
    Case &H46 '+n*38  BIT  n,(HL)
    memptr = readM(H * 256 + L)
    bit memptr, (BITT(memval And 56))
    Case &H47 '+n*38  BIT  n,A
    bit A, (BITT(memval And 56))
    Case &H80 '+ n * 38 'RES  n,B
    res b, (SETT(memval And 56))
    Case &H81 '+ n * 38 'RES  n,C
    res C, (SETT(memval And 56))
    Case &H82 '+ n * 38 'RES  n,D
    res D, (SETT(memval And 56))
    Case &H83 '+ n * 38 'RES  n,E
    res E, (SETT(memval And 56))
    Case &H84 '+ n * 38 'RES  n,H
    res H, (SETT(memval And 56))
    Case &H85 '+ n * 38 'RES  n,L
    res L, (SETT(memval And 56))
    Case &H86 '+ n * 38 'RES  n,(HL)
    memptr = readM(H * 256 + L)
    res memptr, (SETT(memval And 56))
    WriteM H * 256 + L, memptr
    Case &H87 '+ n * 38 'RES  n,A
    res A, (SETT(memval And 56))
    Case &HC0 '+ n * 38 'SET  n,B
    zset b, (BITT(memval And 56))
    Case &HC1 '+ n * 38 'SET  n,C
    zset C, (BITT(memval And 56))
    Case &HC2 '+ n * 38 'SET  n,D
    zset D, (BITT(memval And 56))
    Case &HC3 '+ n * 38 'SET  n,E
    zset E, (BITT(memval And 56))
    Case &HC4 '+ n * 38 'SET  n,H
    zset H, (BITT(memval And 56))
    Case &HC5 '+ n * 38 'SET  n,L
    zset L, (BITT(memval And 56))
    Case &HC6 '+ n * 38 'SET  n,(HL)
    memptr = readM(H * 256 + L)
    zset memptr, (BITT(memval And 56))
    WriteM H * 256 + L, memptr
    Case &HC7 '+ n * 38 'SET n,A
    zset A, (BITT(memval And 56))
    End Select
    
End Select
Case &HCC     ' CALL Z,nnnn
zcall pw, getZ
Case &HCD     ' CALL nnnn
zcall pw
Case &HCE     ' ADC  A,nn
adc pb
Case &HCF     ' RST  8
rst 8
Case &HD0     ' RET  NC
ret Not getC
Case &HD1     ' POP  DE
pop E
pop D
Case &HD2     ' JP     'NC,nnnn
jp pw, Not getC
Case &HD3     ' -     '     '     '     '     '  ---- ??? (old out (nn),a)
Stop
Case &HD4     ' CALL NC,nnnn
zcall pw, Not getC
Case &HD5     ' PUSH DE
push D
push E
Case &HD6     ' SUB  nn
zsub pb
Case &HD7     ' RST  10H
rst 16
Case &HD8     ' RET  C
ret getC
Case &HD9     ' RETI     '     '     '     '  ---- remapped (old exx)
reti
Case &HDA     ' JP     'C,nnnn
jp pw, getC
Case &HDB     ' -     '     '     '     '     '  ---- ??? (old in a,(nn))
Stop
Case &HDC     ' CALL C,nnnn
zcall pw, getC
Case &HDD     ' -     '     '     '     '     '  ---- ??? (old ix-commands)
Stop
Case &HDE     ' SBC  A,nn     '  (nocash added, this opcode does existed, e.g. used by kwirk)
sbc pb
Case &HDF     ' RST  18H
rst 24
Case &HE0     ' LD     '($FF00+nn),A ---- special (old ret po)
WriteM 65280 + pb, A
Case &HE1     ' POP  HL
pop L
pop H
Case &HE2     ' LD     '($FF00+C),A  ---- special (old jp po,nnnn)
WriteM 65280 + C, A
Case &HE3     ' -     '     '     '     '     '  ---- ??? (old ex (sp),hl)
Stop
Case &HE4     ' -     '     '     '     '     '  ---- ??? (old call po,nnnn)
Stop
Case &HE5     ' PUSH HL
push H
push L
Case &HE6     ' AND  nn
zand pb
Case &HE7     ' RST  20H
rst 32
Case &HE8     ' ADD  SP,dd     '     '  ---- special (old ret pe) (nocash extended as shortint)
addSP pb
Case &HE9 'JP(HL)
jp H * 256 + L
Case &HEA     ' LD     '(nnnn),A     '  ---- special (old jp pe,nnnn)
WriteM pw, A
Case &HEB     ' -     '     '     '     '     '  ---- ??? (old ex de,hl)
Stop
Case &HEC     ' -     '     '     '     '     '  ---- ??? (old call pe,nnnn)
Stop
Case &HED     ' -     '     '     '     '     '  ---- ??? (old ed-commands)
Stop
Case &HEE     ' XOR  nn
zxor pb
Case &HEF     ' RST  28H
rst 40
Case &HF0     ' LD     'A,($FF00+nn) ---- special (old ret p)
A = readM(65280 + pb)
Case &HF1     ' POP  AF
pop F
pop A
Case &HF2     ' LD     'A,(C)     '     '  ---- special (old jp p,nnnn)
A = readM(65280 + C)
Case &HF3 'DI
ime_stat = 2
Case &HF4     ' -     '     '     '     '     '  ---- ??? (old call p,nnnn)
Stop
Case &HF5     ' PUSH AF
push A
push F
Case &HF6     ' OR     'nn
zor pb
Case &HF7     ' RST  30H
rst 48
Case &HF8     ' LD     'HL,SP+dd     '  ---- special (old ret m) (nocash corrected)
memptr = pb
If memptr > 127 Then memptr = memptr - 256
memval = (sp + memptr) And 65535
        If memptr >= 0 Then
           setC sp > memval
           setH ((sp Xor memptr Xor memval) And 4096) > 0
           H = memval \ 256
           L = memval And 255
        Else
           setC sp > memval
           setH ((sp Xor memptr Xor memval) And 4096) > 0
           H = memval \ 256
           L = memval And 255
        End If
setZ False
setN False
Case &HF9     ' LD     'SP,HL
sp = H * 256 + L
Case &HFA     ' LD     'A,(nnnn)     '  ---- special (old jp m,nnnn)
A = readM(pw)
Case &HFB 'EI
ime_stat = 1
Case &HFC     ' -     '     '     '     '     '  ---- ??? (old call m,nnnn)
Stop
Case &HFD     ' -     '     '     '     '     '  ---- ??? (old iy-commands)
Stop
Case &HFE     ' CP     'nn
cp pb
Case &HFF     ' RST  38H
rst 56
End Select
If pc = brkAddr Then Stop

End Sub
