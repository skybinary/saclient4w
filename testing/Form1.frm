VERSION 5.00
Object = "{0375EA14-9C5D-4504-91A2-526ACCD762AF}#10.0#0"; "SAClient.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin prjSAClient.SAClient SAClient1 
      Height          =   585
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1032
      Application     =   "http://192.168.0.3/homework/pactivex.php"
      IP              =   "192.168.0.3"
      Node            =   "1"
      Register        =   "4000"
      Count           =   "8"
      Interval        =   5000
      Enabled         =   -1
      Value           =   "-1"
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   1800
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Check1_Click()
    SAClient1.Enabled = CBool(Check1.Value)
End Sub

Private Sub Form_Load()
    Check1.Value = CInt(Not SAClient1.Enabled) + 1
End Sub

'Private Sub SAClient1_OnChange(ByVal Value As String)
'    Dim anArray As Variant
'    anArray = Split(Value, ",")
'
'    Debug.Print " Count:" & CStr(anArray(0))
'    Debug.Print "Second:" & CStr(anArray(1))
'    Debug.Print "Minute:" & CStr(anArray(2))
'    Debug.Print "  Hour:" & CStr(anArray(3))
'    Debug.Print "   Day:" & CStr(anArray(4))
'    Debug.Print " Month:" & CStr(anArray(5))
'    Debug.Print "  Year:" & CStr(anArray(6))
'    Debug.Print "   ???:" & CStr(anArray(7))
'
'End Sub
Private Sub SAClient1_OnChange(ByVal Value As String)
    Dim anArray As Variant
    anArray = Split(Value, ",")

    Debug.Print " Count:" & CStr(anArray(0))
    Debug.Print "Second:" & CStr(anArray(1))
    Debug.Print "Minute:" & CStr(anArray(2))
    Debug.Print "  Hour:" & CStr(anArray(3))
    Debug.Print "   Day:" & CStr(anArray(4))
    Debug.Print " Month:" & CStr(anArray(5))
    Debug.Print "  Year:" & CStr(anArray(6))
    Debug.Print "   ???:" & CStr(anArray(7))


End Sub
