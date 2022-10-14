Attribute VB_Name = "Module1"
Option Explicit
'AngelV
'  https://www.vbforums.com/showthread.php?898004-Opposite-of-dereferencing-Obj-pointers

'todo: dynamically look up function index from name rather than hard coded offset..

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Const PAGE_EXECUTE_READWRITE = &H40&

Sub Main()
 
    Dim oClass As New Class1
    Dim initAddr As Long
    
    MsgBox "ObjPtr(oClass) = " & Hex(ObjPtr(oClass))
    
    'Hook the first class Method (SayHello)
    initAddr = HookClassMethod(oClass, 7, AddressOf OverrideFunction)
    
    oClass.SayHello 'test our hook
    
    'UnHook the Method.
    Call HookClassMethod(oClass, 7, initAddr)

    oClass.SayHello 'call the original to test unhook
 
End Sub

Function HookClassMethod(oClassInstance As Object, ByVal IDX_CLASSMETHOD As Long, newVal As Long) As Long

    Dim pVTable As Long, orgVal As Long, targetAddress As Long, oldProtect As Long
    
    CopyMemory pVTable, ByVal ObjPtr(oClassInstance), 4
    targetAddress = pVTable + (IDX_CLASSMETHOD * 4)
    
    CopyMemory orgVal, ByVal targetAddress, 4
    
    VirtualProtect targetAddress, 4, PAGE_EXECUTE_READWRITE, oldProtect
    CopyMemory ByVal targetAddress, newVal, 4
    VirtualProtect targetAddress, 4, oldProtect, 0&
     
    HookClassMethod = orgVal
    
End Function

'Sub OverrideFunction(ByVal pObjPtr As Long)
Function OverrideFunction(ByVal pObjPtr As Long) As Long
    MsgBox "Hooked this = " & Hex(pObjPtr)
    OverrideFunction = 0 'S_OK
End Function



