Attribute VB_Name = "mem"
'This is a part of tha BasicBoy emulator
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel).
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'(I know the emulator is NOT OPTIMIZED AT ALL)

'Ram I/O functions...
'Base taken from VisBoy (Uptaded,Corected,Recoded)
'Currenlty no optimizations
'Coments will be added with the next release

'Sory for my bad english ...

Option Explicit
Option Base 0

Public Type CartIinfo
    NGraph(47) As Byte
    title As String
    titleB(15) As Byte
    GBC As Byte
    lcode(1) As Byte
    isSGB As Byte
    Ctype As Byte
    romsize As Byte
    ramsize As Byte
    DestCode As Byte
    lcodeold As Byte
    MaskRomV As Byte
    CCheck As Byte
    Checksum(1) As Byte
End Type

Global rominfo As CartIinfo
Global Ct(255) As String
Global Ros(255) As String
Global Ras(255) As String
Global Rosn(255) As Long
Global Rasn(255) As Long
Global ROM(16383, 128) As Long
Global RAM(32768 To 65535, 7) As Long
Global wRamB As Long, vRamB As Long, objpi As Long, bgpi As Long, bgai As Boolean, objai As Boolean
Global hdmaS As Long, hdmaD As Long
Global colconv0(255) As Integer, colconv1(255) As Integer
Global bRam(8191, 15) As Byte
Global joyval1 As Long, joyval2 As Long
Global CurROMBank As Integer, CurRAMBank As Integer, tmpcolor As Byte, tmpcolor2 As Byte
Private i As Long, j As Long
Dim mbc1mode As Long
Dim memptr2 As Long
Global ro As String
Global Hdma As Boolean, Hdmal As Long, tHdmal As Long, nwr As Boolean

Public Function readM(memptr As Long) As Long
If GBM = 0 Then
    If memptr < 16384 Then
            readM = ROM(memptr, 0)      ' Read from ROM
    ElseIf memptr < 32768 Then
            readM = ROM(memptr - 16384, CurROMBank)      ' Read from ROM
    Else
            If memptr > 40959 And memptr < 49152 Then
            readM = bRam(memptr - 40960, CurRAMBank)    ' Read from sRAM
            Else
            readM = RAM(memptr, 0)      ' Read from RAM
            End If
    End If
Else
    If memptr < 16384 Then
            readM = ROM(memptr, 0)      ' Read from ROM
    ElseIf memptr < 32768 Then
            readM = ROM(memptr - 16384, CurROMBank)      ' Read from ROM
    ElseIf memptr < 40960 Then 'read Vram
        readM = RAM(memptr, vRamB)
    ElseIf memptr < 49152 Then 'read sRam
        readM = bRam(memptr - 40960, CurRAMBank)
    ElseIf memptr < 53248 Then 'read wRam(0)
        readM = RAM(memptr, 0)
    ElseIf memptr < 57344 Then 'read wRam(1-7)
        readM = RAM(memptr, wRamB)
    Else 'read ram
        readM = RAM(memptr, 0)      ' Read from RAM
    End If
End If
End Function
Public Sub WriteM(memptr As Long, ByVal Value As Long)
    If memptr > 32767 Then 'ram/mmio
    'ram
    If GBM = 0 Then
                If memptr > 40959 And memptr < 49152 Then    ' write to sRAM
                   bRam(memptr - 40960, CurRAMBank) = Value
                Else
                   RAM(memptr, 0) = Value    ' write to RAM
                   If memptr > &HE000 And memptr < &HFE00 Then ' echo
                      RAM(memptr - 8192, 0) = Value
                   ElseIf memptr > &HC000 And memptr < &HDE00 Then ' echo
                      RAM(memptr + 8192, 0) = Value
                   End If
                End If
    Else
    Select Case memptr
            Case Is < 40960 'write Vram
               RAM(memptr, vRamB) = Value
            'exit sub
            Case Is < 49152 'write sRam
               bRam(memptr - 40960, CurRAMBank) = Value
            'exit sub
            Case Is < 53248 'write wRam(0)
               RAM(memptr, 0) = Value
            'exit sub
            Case Is < 57344 'write wRam(1-7)
               RAM(memptr, wRamB) = Value
            'exit sub
            Case Else 'write ram
               RAM(memptr, 0) = Value    ' Read from RAM
      End Select
End If

    
    'mmio
    If memptr > 65279 Then
    Select Case memptr
        Case Is = 65280     ' Joypad
            If (Value And 32) = 32 Then         'Directional
                RAM(65280, 0) = 223 And (255 - joyval1)
            ElseIf (Value And 16) = 16 Then     ' Buttons
                RAM(65280, 0) = 239 And (255 - joyval2)
            Else
                RAM(65280, 0) = 255
            End If
        Case Is = 65350     ' DMA Xfer
            RAM(65350, 0) = Value
            j = Value * 256
            For i = 65024 To 65183
                RAM(i, 0) = readM(j)
                j = j + 1
            Next i
        Case 65351, 65352, 65353
            RAM(memptr, 0) = Value
            ccolid2 CByte(Value), memptr - 65351
        Case 65287 'timer
            RAM(memptr, 0) = Value
            Select Case Value And 3
            Case 0
                z80.tvm = 1024
            Case 1
                z80.tvm = 65536
            Case 2
                z80.tvm = 16384
            Case 3
                z80.tvm = 4096
            End Select
    End Select
    
    If GBM = 1 Then
        Select Case memptr
        Case 65357  'Speed SW
        smp = Value And 1
        RAM(65357, 0) = cpuS * 128 + smp
        Case 65359  'Vram Bank
        vRamB = Value And 1
        RAM(65359, 0) = vRamB
        Case 65361  'HDMA1 sh
        hdmaS = (hdmaS And 255) + Value * 256
        Case 65362 'HDMA2 sl
        hdmaS = (hdmaS And 65280) + Value
        Case 65363  'HDMA3 dh
        hdmaD = (hdmaD And 255) + Value * 256
        Case 65364 'HDMA4 dl
        hdmaD = (hdmaD And 65280) + Value
        Case 65365 'HDMA5 lms
        hdmaD = (hdmaD And 8176) + 32768
        hdmaS = hdmaS And 65520
        'If nwr Then
        If Hdma = True Then If (Value And 128) = 0 Then Hdma = False: RAM(65365, 0) = 128 + 70: Exit Sub Else Exit Sub
        If Value And 128 Then Hdma = True: Hdmal = Value And 127: tHdmal = Value And 127: RAM(65365, 0) = Hdmal: Exit Sub
        'End If
        j = hdmaD
        For i = hdmaS To hdmaS + (Value And 127) * 16 + 15
        RAM(j, vRamB) = readM(i)
        j = j + 1
        Next i
        RAM(65365, 0) = 255
        Case 65366  'Rp
        'Stop
        Case 65384  'Bg pal indx
        bgpi = Value And 63
        bgai = Value And 128
        If bgpi Mod 2 Then RAM(65385, 0) = bgp(bgpi \ 8, (bgpi \ 2) Mod 4) \ 256 Else RAM(65385, 0) = bgp(bgpi \ 8, (bgpi \ 2) Mod 4) And 255
        
        Case 65385 'BG Pal Val
        i = bgpi Mod 2
        If i = 0 Then ' 1st byte
        bgp(bgpi \ 8, (bgpi \ 2) Mod 4) = (bgp(bgpi \ 8, (bgpi \ 2) Mod 4) And 65280) + Value
        Else '2nd byte
        bgp(bgpi \ 8, (bgpi \ 2) Mod 4) = ((bgp(bgpi \ 8, (bgpi \ 2) Mod 4) And 255) + Value * 256) And 32767
        End If
        
        
        If bgai Then bgpi = bgpi + 1
        WriteM 65384, (RAM(65384, 0) And 128) Or (bgpi And 63)
        
        Case 65386  'OBJ pal indx
        objpi = Value And 63
        objai = Value And 128
        If objpi Mod 2 Then RAM(65387, 0) = objp(objpi \ 8, (objpi \ 2) Mod 4) \ 256 Else RAM(65387, 0) = objp(objpi \ 8, (objpi \ 2) Mod 4) And 255
        
        Case 65387 'OBJ Pal Val
        
        i = objpi Mod 2
        If i = 0 Then ' 1st byte
        objp(objpi \ 8, (objpi \ 2) Mod 4) = (objp(objpi \ 8, (objpi \ 2) Mod 4) And 65280) + Value
        Else '2nd byte
        objp(objpi \ 8, (objpi \ 2) Mod 4) = ((objp(objpi \ 8, (objpi \ 2) Mod 4) And 255) + Value * 256) And 32767
        End If
        
        
        If objai Then objpi = objpi + 1
        WriteM 65386, (RAM(65386, 0) And 128) Or (objpi And 63)
        
        Case 65392  'SVBK
        wRamB = Value And 7
        If wRamB < 1 Then wRamB = 1
        RAM(65392, 0) = wRamB
        End Select
        End If
        
    End If
    Else    'rom
            Select Case rominfo.Ctype
            
            Case 1, 2, 3 ' mbc1
             
            If memptr >= &H2000 And memptr < &H4000 Then    ' Bank Switch
                If Value > 0 Then
                    CurROMBank = Value And (Rosn(rominfo.romsize) - 1)
                Else
                    CurROMBank = 1
                End If
            ElseIf memptr >= &H4000 And memptr < &H6000 Then    ' Bank Switch
            Value = Value And 3
            If mbc1mode = 0 Then CurROMBank = 127 And (CurROMBank + Value * 32) Else CurRAMBank = Value
            ElseIf memptr >= &H6000 And memptr < &H8000 Then
            mbc1mode = Value And 1
            End If
            
            Case &H19, &H1A, &H1B, &HC, &H1D, &H1E 'mbc5
            If memptr >= &H2000 And memptr < &H3000 Then
            CurROMBank = Value Mod Rosn(rominfo.romsize)
            ElseIf memptr < &H4000 Then
            CurROMBank = ((Value And 1) * 256 + (255 And CurROMBank)) Mod Rosn(rominfo.romsize)
            ElseIf memptr < &H6000 Then
            If rominfo.ramsize Then CurRAMBank = Value Mod Rasn(rominfo.ramsize)
            End If
            
            Case &HF, &H10, &H11, &H12, &H13 'mbc3
            If memptr >= &H2000 And memptr < &H4000 Then 'rom
            If Value > 0 Then CurROMBank = Value Mod Rosn(rominfo.romsize) Else CurROMBank = 1
            ElseIf memptr >= &H4000 And memptr < &H6000 Then 'ram
            Value = Value And 3
            CurRAMBank = Value
            End If
            
        End Select
    End If
End Sub

Public Sub initCI()
Dim i As Long, i2 As Long
'ReDim RAM(32768 To 65535)
For i2 = 0 To 7
For i = 32768 To 65535
RAM(i, i2) = 0
Next i
Next i2
For i = 0 To 7
bgp(i, 0) = 32767: objp(i, 0) = 32767
bgp(i, 1) = 32767: objp(i, 1) = 32767
bgp(i, 2) = 32767: objp(i, 2) = 32767
bgp(i, 3) = 32767: objp(i, 3) = 32767
Next i
CurROMBank = 1
'On Error Resume Next
Ct(0) = "Rom Only": Ct(&H12) = "Rom+MBC3+Ram"
Ct(1) = "Rom+MBC1": Ct(&H13) = "Rom+MBC3+Ram+Batt"
Ct(2) = "Rom+MBC1+Ram": Ct(&H19) = "Rom+MBC5"
Ct(3) = "Rom+MBC1+Ram+Batt": Ct(&H1A) = "Rom+MBC5+Ram"
Ct(5) = "Rom+MBC2": Ct(&H1B) = "Rom+MBC5+Ram+Batt"
Ct(6) = "Rom+MBC2+Batt": Ct(&H1C) = "Rom+MBC5+Rumble"
Ct(8) = "Rom+Ram": Ct(&H1D) = "Rom+MBC3+Rumble+Sram"
Ct(9) = "Rom+Ram+Batt": Ct(&H1E) = "Rom+MBC3+Rumble+Sram+Batt"
Ct(&HB) = "Rom+MMO1": Ct(&H1F) = "Pocet Camera"
Ct(&HC) = "Rom+MMO1+Sram": Ct(&HFD) = "Bandai TAMA5"
Ct(&HD) = "Rom+MMO1+Sram+Batt": Ct(&HFE) = "Hudson HuC-3"
Ct(&HF) = "Rom+MBC3+Timer+Batt": Ct(&HFF) = "Hudson HuC-1"
Ct(&H10) = "Rom+MBC3+Timer+Ram+Batt"
Ct(&H11) = "Rom+MBC3"
rominfo.Ctype = ROM(&H147, 0)
For i = &H134 To &H142
    rominfo.titleB(i - &H134) = ROM(i, 0)
Next i
rominfo.title = StrConv(rominfo.titleB, vbUnicode)
rominfo.romsize = ROM(&H148, 0)
rominfo.ramsize = ROM(&H149, 0)
Ros(0) = "32 Kbyte": Rosn(0) = 2
Ros(1) = "64 Kbyte": Rosn(1) = 4
Ros(2) = "128 Kbyte": Rosn(2) = 8
Ros(3) = "256 Kbyte": Rosn(3) = 16
Ros(4) = "512 Kbyte": Rosn(4) = 32
Ros(5) = "1 Mbyte": Rosn(5) = 64
Ros(6) = "2 Mbyte": Rosn(6) = 128
Ros(52) = "1.1 Mbyte": Rosn(52) = 72
Ros(53) = "1.2 Mbyte": Rosn(53) = 80
Ros(54) = "1.5 Mbyte": Rosn(54) = 96

Ras(0) = "None": Rasn(0) = 0
Ras(1) = "2 Kbyte": Rasn(1) = 1
Ras(2) = "8 Kbyte": Rasn(2) = 1
Ras(3) = "32 Kbyte": Rasn(3) = 4
Ras(4) = "128 Kbyte": Rasn(4) = 16

 m_CurrentClockSpeed = 1024
 utu = True
'ReDim RAM(32768 To 65535, Rasn(rominfo.ramsize))
End Sub

Sub wrRam()
Dim tRam() As Byte
On Error GoTo enf
If Len(ro) > 0 Then
ReDim tRam(Rasn(rominfo.ramsize) * 8192 - 1)
CopyMemory tRam(0), bRam(0, 0), UBound(tRam) + 1
Open ro For Binary As #1
Put #1, , tRam
Close #1
enf:
Close #1
ro = ""
End If
End Sub
Sub rdRam()
Dim tRam() As Byte
On Error GoTo enf
If Len(ro) > 0 Then wrRam
'ro = Form1.cd.Filename & ".sav"

ro = "gameboy.sav"

ReDim tRam(Rasn(rominfo.ramsize) * 8192 - 1)
Open ro For Binary As #1
Get #1, , tRam
Close #1
CopyMemory bRam(0, 0), tRam(0), UBound(tRam) + 1
enf:
Close #1
End Sub
Public Function pb() As Long
If GBM = 0 Then
    Select Case pc
        Case Is < 16384
            pb = ROM(pc, 0)      ' Read from ROM
        Case Is < 32768
            pb = ROM(pc - 16384, CurROMBank)      ' Read from ROM
        Case Else
            If pc > 40959 And pc < 49152 Then
            pb = bRam(pc - 40960, CurRAMBank)    ' Read from sRAM
            Else
            pb = RAM(pc, 0)      ' Read from RAM
            End If
    End Select
Else
    Select Case pc
        Case Is < 16384
            pb = ROM(pc, 0)      ' Read from ROM
        Case Is < 32768
            pb = ROM(pc - 16384, CurROMBank)      ' Read from ROM
        Case Is < 40960 'read Vram
        pb = RAM(pc, vRamB)
        Case Is < 49152 'read sRam
        pb = bRam(pc - 40960, CurRAMBank)
        Case Is < 53248 'read wRam(0)
        pb = RAM(pc, 0)
        Case Is < 57344 'read wRam(1-7)
        pb = RAM(pc, wRamB)
        Case Else 'read ram
        pb = RAM(pc, 0)      ' Read from RAM
    End Select
End If
pc = pc + 1
End Function
Public Function pw() As Long
If GBM = 0 Then
    Select Case pc
        Case Is < 16384
            pw = ROM(pc, 0)      ' Read from ROM
        Case Is < 32768
            pw = ROM(pc - 16384, CurROMBank)      ' Read from ROM
        Case Else
            If pc > 40959 And pc < 49152 Then
            pw = bRam(pc - 40960, CurRAMBank)    ' Read from sRAM
            Else
            pw = RAM(pc, 0)      ' Read from RAM
            End If
    End Select
    pc = pc + 1
        Select Case pc
        Case Is < 16384
            pw = pw + ROM(pc, 0) * 256  ' Read from ROM
        Case Is < 32768
            pw = pw + ROM(pc - 16384, CurROMBank) * 256  ' Read from ROM
        Case Else
            If pc > 40959 And pc < 49152 Then
            pw = pw + bRam(pc - 40960, CurRAMBank) * 256 ' Read from sRAM
            Else
            pw = pw + RAM(pc, 0) * 256  ' Read from RAM
            End If
    End Select
    pc = pc + 1
Else
    Select Case pc
        Case Is < 16384
            pw = ROM(pc, 0)      ' Read from ROM
        Case Is < 32768
            pw = ROM(pc - 16384, CurROMBank)      ' Read from ROM
        Case Is < 40960 'read Vram
        pw = RAM(pc, vRamB)
        Case Is < 49152 'read sRam
        pw = bRam(pc - 40960, CurRAMBank)
        Case Is < 53248 'read wRam(0)
        pw = RAM(pc, 0)
        Case Is < 57344 'read wRam(1-7)
        pw = RAM(pc, wRamB)
        Case Else 'read ram
        pw = RAM(pc, 0)      ' Read from RAM
    End Select
    pc = pc + 1
    Select Case pc
        Case Is < 16384
        pw = pw + ROM(pc, 0) * 256  ' Read from ROM
        Case Is < 32768
        pw = pw + ROM(pc - 16384, CurROMBank) * 256  ' Read from ROM
        Case Is < 40960 'read Vram
        pw = pw + RAM(pc, vRamB) * 256
        Case Is < 49152 'read sRam
        pw = pw + bRam(pc - 40960, CurRAMBank) * 256
        Case Is < 53248 'read wRam(0)
        pw = pw + RAM(pc, 0) * 256
        Case Is < 57344 'read wRam(1-7)
        pw = pw + RAM(pc, wRamB) * 256
        Case Else 'read ram
        pw = pw + RAM(pc, 0) * 256  ' Read from RAM
    End Select
    pc = pc + 1
End If


End Function
