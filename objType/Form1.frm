VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Function dbg(txt As String, x As Long)
    'MsgBox caption & " = " & Hex(x)
    List1.AddItem txt & " = " & Hex(x)
End Function

Private Sub Form_Load()
    
    Dim vtable As Long
    Dim objInfo As Long
    Dim mainObjStruct As Long
    Dim objType As Long
    
     
    If TypeOf Me Is Form Then
       List1.AddItem "TypeOf me is form"
    End If

    dbg "objptr(me)", ObjPtr(Me)
    
    CopyMemory vtable, ByVal ObjPtr(Me), 4
    dbg "vtable", vtable
    
    CopyMemory objInfo, ByVal vtable - 4, 4
    dbg "objinfo", objInfo
    
    CopyMemory mainObjStruct, ByVal objInfo + &H18, 4
    dbg "mainObjStruct", mainObjStruct
    
    CopyMemory objType, ByVal mainObjStruct + &H28, 4
    dbg "objType", objType
    
    List1.AddItem "isForm = " & isForm(objType)
    List1.AddItem "isUserCtl = " & isUserControl(objType)
    List1.AddItem "hasOptInfo = " & hasOptInfo(objType)
    List1.AddItem "isClass = " & isClass(objType)
    

End Sub

'https://github.com/VBGAMER45/Semi-VB-Decompiler/blob/035aebf7b72ac1193704a43655acaed3e9e09428/Semi%20VB%20Decompiler/modGlobals.bas
'
'mostly from moog
'https://www.rapidtables.com/convert/number/hex-to-binary.html?x=50430083
'tObject.ObjectTyper Properties...24 bits used
'0x80 =                                 1000 0000
'hexcodes: 18083, 1118803,180A3,180C3
'#########################################################
'form               0000 0001 1000 0000 1000 0011 --> 18083
'                   0000 0001 1000 0000 1010 0011 --> 180A3
'                   0000 0001 1000 0000 1100 0011 --> 180C3
'module             0000 0001 1000 0000 0000 0001 --> 18001
'                   0000 0001 1000 0000 0100 0001 --> 18041
'                   0000 0001 1000 0000 0010 0001 --> 18021
'class              0001 0001 1000 0000 0000 0011 --> 118003
'                   0001 0001 1000 0000 0100 0011 --> 118043
'                   0001 0011 1000 0000 0000 0011 --> 138003
'                   0000 0001 1000 0000 0010 0011 --> 18023
'                   0000 0001 1000 1000 0000 0011 --> 18803
'                   0001 0001 1000 1000 0000 0011 --> 118803
'usercontrol        0001 1101 1010 0000 0000 0011 --> 1DA003
'                   0001 1101 1010 0000 0010 0011 --> 1DA023
'                   0001 1101 1010 1000 0000 0011 --> 1DA803
'                   0001 1101 1010 0000 0100 0011 --> 1DA043
'propertypage       0001 0101 1000 0000 0000 0011 --> 158003
'                   0001 0101 1000 0000 0100 0011 --> 158043
'datareport    0001 0000 0001 1000 0000 1000 0011 -->1018083
'              0001 0000 0001 1000 0000 1100 0011 -->10180C3
'DataEnv  0101 0000 0100 0011 0000 0000 1000 0011 -->50430083 https://flylib.com/books/en/3.405.1.73/1/
'                 |    | ||     |  |    |||    |
'data report   ---+    | ||     |  |    |||    |
'                      | ||     |  |    |||    |
'HasPublicInterface ---+ ||     |  |    |||    |
'HasPublicEvents --------+|     |  |    |||    |
'IsCreatable/Visible? ----+     |  |    |||    |
'Same as "HasPublicEvents" -----+  |    |||    |
'                               |  |    |||    |
'usercontrol     ---------------+  |    |||    |
'ocx/dll     ----------------------+    |||    |
'form      -----------------------------+||    |
'opt cmp text ---------------------------+|    |
'vb5     ---------------------------------+    |
'HasOptInfo     -------------------------------+
'                                              |
'module    ------------------------------------+

Function isForm(ByVal v As Long) As Boolean
    If (v And &H80) = &H80 Then isForm = True
End Function

Function hasOptInfo(ByVal v As Long) As Boolean
    If (v And 2) = 2 Then hasOptInfo = True
End Function

Function hasPubEvents(ByVal v As Long) As Boolean
    If (v And &H80000) = &H80000 Then hasPubEvents = True
End Function

Function hasPubiFace(ByVal v As Long) As Boolean
    If (v And &H100000) = &H100000 Then hasPubiFace = True
End Function

Function isUserControl(ByVal v As Long) As Boolean
    If (v And &H2000) = &H2000 Then isUserControl = True
End Function

'some known values we can seen, what is the unique bit mask?
Function isClass(v As Long) As Boolean
    If v = 1146883 Then isClass = True
    If v = 1277955 Then isClass = True
    If v = 98339 Then isClass = True
    If v = 100355 Then isClass = True
    If v = 1148931 Then isClass = True
    If v = &H118043 Then isClass = True
End Function
