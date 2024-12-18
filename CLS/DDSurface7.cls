VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDDSurface7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'class variables
Public Surface As DirectDrawSurface7
Private ddsd As DDSURFACEDESC2
Private rSurface As DxVBLib.RECT
Private ddcColorKey As DDCOLORKEY

Public Function Create(ByRef objDD As DirectDraw7, _
    ByVal lWidth As Long, ByVal lHeight As Long, _
    Optional bPrimary As Boolean = False) As Boolean
    Create = False
    On Error GoTo create_error

    If bPrimary Then
        'create the primary display surface
        ddsd.lFlags = DDSD_CAPS
        ddsd.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Else
        'create a normal surface
        ddsd.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    End If

    'set surface resolution
    ddsd.lWidth = lWidth
    ddsd.lHeight = lHeight
    
    'create the requested surface
    Set Surface = objDD.CreateSurface(ddsd)
    
    'set transparent color to black
    ddcColorKey.low = 0
    ddcColorKey.high = 0
    Surface.SetColorKey DDCKEY_SRCBLT, ddcColorKey
    
    'set surface rectangle
    rSurface.Bottom = ddsd.lHeight
    rSurface.Right = ddsd.lWidth
    Create = True
create_error:
End Function

Public Function Load(ByRef DDraw As clsDirectDraw7, _
    ByVal sFilename As String) As Boolean

    Load = False
    On Error GoTo load_error
    
    'create surface and load bitmap
    Set Surface = DDraw.DDObj.CreateSurfaceFromFile(sFilename, ddsd)
    If Surface Is Nothing Then Exit Function

    'set transparent color to black
    ddcColorKey.low = 0
    ddcColorKey.high = 0
    Surface.SetColorKey DDCKEY_SRCBLT, ddcColorKey

    'set surface rectangle
    rSurface.Bottom = ddsd.lHeight
    rSurface.Right = ddsd.lWidth
    Load = True
load_error:
End Function

Public Sub Blt(ByRef dstRect As DxVBLib.RECT, _
    ByRef ddsSource As DirectDrawSurface7, _
    ByRef srcRect As DxVBLib.RECT)
    Surface.Blt dstRect, ddsSource, srcRect, DDBLT_WAIT
End Sub

Public Sub BltTrans(ByRef dstRect As DxVBLib.RECT, _
    ByRef ddsSource As DirectDrawSurface7, _
    ByRef srcRect As DxVBLib.RECT)
    Surface.Blt dstRect, ddsSource, srcRect, DDBLT_WAIT Or _
        DDBLT_KEYSRC
End Sub

Public Sub BltFast(ByVal X As Long, ByVal Y As Long, _
    ByRef ddsSource As DirectDrawSurface7, _
    ByRef srcRect As DxVBLib.RECT)
    
    Surface.BltFast X, Y, ddsSource, srcRect, DDBLTFAST_WAIT
End Sub

Public Sub BltColorFill(ByVal lColor As Long)
    Surface.BltColorFill rSurface, lColor
End Sub

Public Sub DrawText(ByVal X As Long, ByVal Y As Long, _
    ByVal sText As String)
    Surface.DrawText X, Y, sText, False
End Sub

Public Property Get SurfaceRect() As DxVBLib.RECT
    SurfaceRect = rSurface
End Property

Public Property Get Width() As Long
    Width = ddsd.lWidth
End Property

Public Property Get Height() As Long
    Height = ddsd.lHeight
End Property
