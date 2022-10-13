VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ComFuncAddr Lib "test.dll" _
(ByVal obj As Long, ByVal sMethodName As Long) As Long

 
Private Sub Form_Load()
    Dim addr As Long
    Dim method As String
    Dim c As New CTest
    Dim fso As New FileSystemObject
    Dim o As Object
    
    'MsgBox c.x
    method = "test"
    
    'Set o = CreateObject("vbSample.CTest") 'works...
    'If o Is Nothing Then End
    'addr = ComFuncAddr(ObjPtr(CreateObject("vbSample.CTest")), StrPtr("test"))
    
    'Dim i As IDispatch 'restricted...
    'Set i = c
   '
    'InputBox "IUnk, IDisp", "", Hex(ObjPtr(c)) & " " & Hex(ObjPtr(i))
    
    CallByName c, "test", VbMethod, CLng(21)
    End
    
    addr = ComFuncAddr(ObjPtr(c), StrPtr(method))
    
    'addr = ComFuncAddr(ObjPtr(fso), StrPtr("FileExists")) 'works...
    MsgBox Hex(addr)
    End
    
End Sub

Public Function test(x As Long)
    MsgBox x
End Function
