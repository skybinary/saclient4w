VERSION 5.00
Object = "{0375EA14-9C5D-4504-91A2-526ACCD762AF}#14.0#0"; "SAClient.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAClient"
   ClientHeight    =   4710
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   2490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   2490
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4700
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   4665
      ScaleWidth      =   2445
      TabIndex        =   0
      Top             =   0
      Width           =   2480
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   2055
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000099&
         Caption         =   "Enabled"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   660
         MaskColor       =   &H000000C0&
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
      Begin prjSAClient.SAClient SAClient1 
         Height          =   180
         Left            =   360
         TabIndex        =   1
         Top             =   157
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   318
         Application     =   "http://192.168.0.3/homework/pactivex.php"
         IP              =   "192.168.0.3"
         Node            =   "1"
         Register        =   "4000"
         Count           =   "8"
         Interval        =   1000
         Enabled         =   -1
         Value           =   "-1"
         ConnStatus      =   -1  'True
      End
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
    If Value = "" Then
        Debug.Print "Empty Response! is the server accessible?"
    Else
        Dim anArray As Variant
        Dim outText As String
        anArray = Split(Value, ",")
    
        outText = outText & vbCrLf
        outText = outText & "Count:" & CStr(anArray(0)) & vbCrLf
        outText = outText & "Second:" & CStr(anArray(1)) & vbCrLf
        outText = outText & "Minute:" & CStr(anArray(2)) & vbCrLf
        outText = outText & "Hour:" & CStr(anArray(3)) & vbCrLf
        outText = outText & "Day:" & CStr(anArray(4)) & vbCrLf
        outText = outText & "Month:" & CStr(anArray(5)) & vbCrLf
        outText = outText & "Year:" & CStr(anArray(6)) & vbCrLf
        outText = outText & "???:" & CStr(anArray(7))
        Text1.Text = outText
    End If
End Sub
