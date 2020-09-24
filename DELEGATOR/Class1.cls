VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'KIND OF PublicVarNumber
Public DUMMY1 As Long
Public DUMMY2 As String
Public DUMMY3 As Integer
Public DUMMY4 As Currency
Public DUMMY5 As Long
'5-PublicVarNumber!!!! 2-nd Parameter For GetObjectFunctionsPointers FUNCTION

'KIND OF PublicObjVariantNumber
Public DUMMY6 As Object
Public DUMMY7 As Variant
'2 PublicObjVariantNumber!!!! 3-rd Parameter For GetObjectFunctionsPointers FUNCTION

Private TIMERPROC As Long  'DELEGATION FUNCTION ADDRESS FOR TIMER
Private HMEM2 As Long
Private HTIMER As Long
Private TMRhwnd As Long


Private OLDPROC As Long
Private CALLMETHODPROC As Long 'DELEGATION FUNCTION ADDRESS
Private HMEM As Long
Private SUBhwnd As Long

'Required !!!
Private TempVariable As Long
Private TempVariable2 As Long


Private ObjectProcedures() As FunctionSPointerS


'METHOD 1! DELEGATED FOR SUBCLASSING..
'SUGGESTION:
'Always use FIRST METHOD Because there is no need to calculate which Number is That Method! It is always method 1!!!!
Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wparam As Long, ByVal lparam As Long) As Long

Debug.Print "HWND:" & hwnd & " -MSG:" & uMsg & " -WPARAM:" & wparam & " -LPARAM:" & lparam
WndProc = CallWindowProc(OLDPROC, hwnd, uMsg, wparam, lparam)
End Function

'METHOD 2! DELEGATED FOR TIMER..
Public Sub TimerProcedura(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)

Debug.Print "TIMER HWND:" & hwnd & ",-" & dwTime
End Sub

Public Sub PerformSubclassCalling()
If HMEM = 0 Then
HMEM = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, CalculateSpaceForDelegation(4))
CALLMETHODPROC = GlobalLock(HMEM)
'Allocate Memory for Subclassing Delegated Procedure!

'*****************************************************
DelegateFunction CALLMETHODPROC, Me, ObjectProcedures(0).FunctionAddress, 4, TempVariable
'1st Parameter is AddressOf Allocated Memory
'2nd Parameter is Object which Delegate Function!
'3rd Parameter is AddressOf Delegated Function! (Previously calculated by GetObjectFunctionsPointers Function)
'4th Parameter is Temporary Variable Required for Delegation Procedure!
'*****************************************************
End If

End Sub
Public Sub PerformTimerCalling()
'Timer Delegator
If HMEM2 = 0 Then
HMEM2 = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, CalculateSpaceForDelegation(4))
TIMERPROC = GlobalLock(HMEM2)
DelegateFunction TIMERPROC, Me, ObjectProcedures(1).FunctionAddress, 4, TempVariable2
End If
End Sub

Public Sub StartTimer(ByVal hwnd As Long, ByVal uElapse As Long)
If HTIMER = 0 And TIMERPROC <> 0 Then
TMRhwnd = hwnd
HTIMER = SetTimer(hwnd, 1, uElapse, TIMERPROC)
End If
End Sub

Public Sub StopTimer()
If HTIMER <> 0 Then
KillTimer TMRhwnd, HTIMER
HTIMER = 0
End If
End Sub


Public Sub SubClass(ByVal hwnd As Long)
If OLDPROC = 0 And HMEM <> 0 Then
SUBhwnd = hwnd
OLDPROC = SetWindowLong(SUBhwnd, -4, CALLMETHODPROC)
End If
End Sub

Public Sub UnSubClass()
If OLDPROC <> 0 Then
Call SetWindowLong(SUBhwnd, -4, OLDPROC)
End If
OLDPROC = 0
SUBhwnd = 0
End Sub

Private Sub Class_Initialize()
'**************************************************
ObjectProcedures = GetObjectFunctionsPointers(Me, 2, 5, 2)
'1st Parameter is Object
'2nd Parameter is Number Of Methods!
'3rd Parameter is Number Of Public VarNumber
'4th Parameter is Number Of PublicObjVariantNumber
'THAT MUST BE CALCULATED CORRECTLY!!!!!!!!!!!!!! IF YOU USE ANY PUBLIC STUFF!
'***************************************************
'This object have 5 PublicVarNumber (3rd Parameter)
'And 2 PublicObjVariantNumber (4th Parameter)
'If you use First Method there is no need calculate all of methods!
'In that case 2nd parameter for GetObjectFunctionsPointers Function =1!!!!!!!!!

End Sub

Private Sub Class_Terminate()
If OLDPROC <> 0 Then UnSubClass
If HTIMER <> 0 Then StopTimer

If HMEM <> 0 Then
Call GlobalUnlock(HMEM)
Call GlobalFree(HMEM)
End If

If HMEM2 <> 0 Then
Call GlobalUnlock(HMEM2)
Call GlobalFree(HMEM2)
End If
End Sub
