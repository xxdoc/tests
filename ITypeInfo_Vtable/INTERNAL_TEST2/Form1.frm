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

Private Declare Function ComFuncAddr Lib "test.dll" ( _
    ByVal obj As Long, _
    ByVal sMethodName As Long, _
    addr As Long, _
    offset As Long _
) As Long

Private Sub Form_Load()
    Dim addr As Long, offset As Long, rv As Long
    Dim c As New CTest
    Dim fso As New Scripting.FileSystemObject
    
    'InputBox "Objptr(c) = ", "", Hex(ObjPtr(c))
    
    rv = ComFuncAddr(ObjPtr(fso), StrPtr("FileExists"), addr, offset)
    MsgBox "fso.FileExists: vtable[" & Hex(offset) & "] = " & Hex(addr)
    
    rv = ComFuncAddr(ObjPtr(c), StrPtr("test"), addr, offset)
    MsgBox "CTest.test: vtable[" & Hex(offset) & "] = " & Hex(addr)
    
    rv = ComFuncAddr(ObjPtr(Me), StrPtr("FormTest"), addr, offset)
    MsgBox "Me.FormTest: vtable[" & Hex(offset) & "] = " & Hex(addr)
    End
    
End Sub

Public Function FormTest()
    MsgBox "in form test"
End Function

