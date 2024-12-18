VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDirectDraw7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'class variables
Dim objDX As DirectX7
Dim objDD As DirectDraw7
Public cddsScreen As clsDDSurface7
Dim objDDClip As DirectDrawClipper
Dim rScreen As DxVBLib.RECT
Dim hSourceWindow As Long

Private Sub Class_Initialize()
    'create the DirectX7 object
    Set objDX = New DirectX7
    'create the DirectDraw object
    Set objDD = objDX.DirectDrawCreate("")
End Sub

Private Sub Class_Terminate()
    Set objDD = Nothing
    Set objDX = Nothing
End Sub

Public Function Startup(ByVal hWnd As Long, ByVal lWidth As Long, _
    ByVal lHeight As Long, ByVal lBPP As Long, _
    ByVal bFullscreen As Boolean) As Boolean
    
    hSourceWindow = hWnd
    If bFullscreen Then
        'set the color depth
        objDD.SetCooperativeLevel hWnd, DDSCL_FULLSCREEN Or _
            DDSCL_EXCLUSIVE
        If lBPP < 1 Or lBPP > 32 Then lBPP = 32
        
        'set the display mode
        objDD.SetDisplayMode lWidth, lHeight, lBPP, 0, DDSDM_DEFAULT
    Else
        'windowed mode
        objDD.SetCooperativeLevel hWnd, DDSCL_NORMAL
    End If
   
    'set up the primary display surface
    Set cddsScreen = New clsDDSurface7
    
    'create the primary display surface
    Startup = cddsScreen.Create(objDD, lWidth, lHeight, True)
    If Startup Then
        'create the clipper object
        Set objDDClip = objDD.CreateClipper(0)
        
        'assoiciate the window handle with the clipper
        objDDClip.SetHWnd hSourceWindow
        
        'clip blitting routines to the window
        cddsScreen.Surface.SetClipper objDDClip
    End If
End Function

Public Sub Shutdown()
    objDD.RestoreDisplayMode
    objDD.SetCooperativeLevel hSourceWindow, DDSCL_NORMAL
End Sub

Public Sub Blt(ByRef ddsSource As DirectDrawSurface7, _
    ByRef srcRect As DxVBLib.RECT)
    cddsScreen.Blt ScreenRect, ddsSource, srcRect
End Sub
Public Sub ddcls(ByVal collor As Long)
    cddsScreen.BltColorFill collor
End Sub

Public Sub Draw(ByRef cddsSource As clsDDSurface7, _
    ByRef srcRect As DxVBLib.RECT)
    cddsScreen.Blt ScreenRect, cddsSource.Surface, srcRect
End Sub

Public Property Get DXObj() As DirectX7
    Set DXObj = objDX
End Property

Public Property Get ddObj() As DirectDraw7
    Set ddObj = objDD
End Property

Public Property Get ScreenRect() As DxVBLib.RECT
    objDX.GetWindowRect hSourceWindow, rScreen
    ScreenRect = rScreen
End Property


