VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDirectInput8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim objDX As New DirectX8
Dim objDI As DirectInput8

Private Sub Class_Terminate()
    Set objDI = Nothing
    Set objDX = Nothing
End Sub

Public Property Get DIObj() As DirectInput8
    Set DIObj = objDI
End Property

Public Property Get DXObj() As DirectX8
    Set DXObj = objDX
End Property

Public Sub Startup(ByRef lWindowHandle As Long)
    'create the DirectX8 object
    Set objDX = New DirectX8
    
    'initialize DirectInput
    Set objDI = objDX.DirectInputCreate()
End Sub



