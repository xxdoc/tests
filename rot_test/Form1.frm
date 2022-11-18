VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Register myClass for direct use"
      Height          =   375
      Left            =   2070
      TabIndex        =   3
      Top             =   3645
      Width           =   2490
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make myClass public"
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   3645
      Width           =   1725
   End
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   225
      TabIndex        =   1
      Top             =   1935
      Width           =   4065
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   4200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

'rot code from wqweto
'https://www.vbforums.com/showthread.php?879529-project-one-workng-with-project2

Private Declare Function GetRunningObjectTable Lib "ole32" (ByVal dwReserved As Long, pResult As IUnknown) As Long
Private Declare Function CreateFileMoniker Lib "ole32" (ByVal lpszPathName As Long, pResult As IUnknown) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal newVal As Long)
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Const RWE = &H40

Private m_lCookie As Long
Private m_lCookie2 As Long

Public myClass As New Class1

Function dbg(x)
    Debug.Print x
    List2.AddItem x
    'MsgBox x
End Function

'mostly from elroy
Function makePublic(obj As Object, Optional ByRef orgVal As Long, Optional andPatch As Boolean = True) As Boolean

    Dim pVTbl As Long, pObjInfo As Long, pObj As Long, newVal As Long, oldMemProt As Long, r As Long
    Dim pObjTypeField As Long
    
    Const flag = &H800 '1000 0000 0000
    
    dbg "Typename(obj) = " & TypeName(obj)
    
    GetMem4 ByVal ObjPtr(obj), pVTbl                 ' Pointer to vTable.
    'dbg "pVTbl=" & Hex(pVTbl)
    
    GetMem4 ByVal pVTbl - 4&, pObjInfo              ' Pointer to tObjectInfo structure.
    'dbg "pObjInfo =" & Hex(pObjInfo)
     
    GetMem4 ByVal pObjInfo + &H18&, pObj            ' Pointer to tObject     structure.
    'dbg "pObj=" & Hex(pObj)
     
    pObjTypeField = pObj + &H28&
    GetMem4 ByVal pObjTypeField, orgVal             ' objType value
    
    dbg "Current ObjType: " & Hex(orgVal)
    
    If andPatch Then
        If (orgVal And flag) = 0 Then  'in IDE = 0x18883 flag is set?
            
            newVal = orgVal Xor flag
            
            If VirtualProtect(ByVal pObjTypeField, 4, RWE, oldMemProt) <> 0 Then
                dbg "Patching to: " & Hex(newVal)
                PutMem4 ByVal pObjTypeField, newVal

                GetMem4 ByVal pObjTypeField, newVal
                dbg "Sanity Check: " & Hex(newVal)
                
                makePublic = True
                VirtualProtect ByVal pObjTypeField, 4, oldMemProt, r
            Else
                dbg "virt prot failed"
            End If
            
        Else
            dbg "Cant patch flag already set?"
        End If
    End If

End Function

Private Sub Command1_Click()

    'we needed a live class to find the objtype but it was already created with the original flags
    If makePublic(myClass) Then
        dbg "patched ok, creating new class for changes to take effect"
        Set myClass = New Class1
        'now we can use it from the form reference, but not yet registered for direct use..
    End If

End Sub

Private Sub Command2_Click()
    
    If m_lCookie2 = 0 Then
        m_lCookie2 = PutObject(myClass, "MySpecialProject.myClass")
        dbg "MySpecialProject.myClass cookie: " & Hex(m_lCookie2)  'fails
    End If
        
End Sub

Private Sub Form_Load()

    If App.PrevInstance Then
        MsgBox "prev instance"
        End
    End If
    
    List1.AddItem "test"
    List1.AddItem Now

    m_lCookie = PutObject(Me, "MySpecialProject.Form1")
    'm_lCookie2 = PutObject(myClass, "MySpecialProject.myClass") 'this will fail class not ready yet..
    
    List1.AddItem "MySpecialProject.Form1 cookie: " & Hex(m_lCookie)
    'List1.AddItem "MySpecialProject.myClass cookie: " & Hex(m_lCookie2)  '0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RevokeObject m_lCookie
    RevokeObject m_lCookie2
End Sub

Public Function PutObject(oObj As Object, sPathName As String, Optional ByVal Flags As Long) As Long
    Const ROTFLAGS_REGISTRATIONKEEPSALIVE As Long = 1
    Const IDX_REGISTER  As Long = 3
    Dim hResult         As Long
    Dim pROT            As IUnknown
    Dim pMoniker        As IUnknown
    
    hResult = GetRunningObjectTable(0, pROT)
    If hResult < 0 Then
        Err.Raise hResult, "GetRunningObjectTable"
    End If
    hResult = CreateFileMoniker(StrPtr(sPathName), pMoniker)
    If hResult < 0 Then
        Err.Raise hResult, "CreateFileMoniker"
    End If
    DispCallByVtbl pROT, IDX_REGISTER, ROTFLAGS_REGISTRATIONKEEPSALIVE Or Flags, ObjPtr(oObj), ObjPtr(pMoniker), VarPtr(PutObject)
End Function

Public Sub RevokeObject(ByVal lCookie As Long)
    Const IDX_REVOKE    As Long = 4
    Dim hResult         As Long
    Dim pROT            As IUnknown
    
    hResult = GetRunningObjectTable(0, pROT)
    If hResult < 0 Then
        Err.Raise hResult, "GetRunningObjectTable"
    End If
    DispCallByVtbl pROT, IDX_REVOKE, lCookie
End Sub

Private Function DispCallByVtbl(pUnk As IUnknown, ByVal lIndex As Long, ParamArray a() As Variant) As Variant
    Const CC_STDCALL    As Long = 4
    Dim lIdx            As Long
    Dim vParam()        As Variant
    Dim vType(0 To 63)  As Integer
    Dim vPtr(0 To 63)   As Long
    Dim hResult         As Long
    
    vParam = a
    For lIdx = 0 To UBound(vParam)
        vType(lIdx) = VarType(vParam(lIdx))
        vPtr(lIdx) = VarPtr(vParam(lIdx))
    Next
    hResult = DispCallFunc(ObjPtr(pUnk), lIndex * 4, CC_STDCALL, vbLong, lIdx, vType(0), vPtr(0), DispCallByVtbl)
    If hResult < 0 Then
        Err.Raise hResult, "DispCallFunc"
    End If
End Function
