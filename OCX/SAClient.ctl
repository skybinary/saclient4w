VERSION 5.00
Begin VB.UserControl SAClient 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2895
   Enabled         =   0   'False
   Picture         =   "SAClient.ctx":0000
   PropertyPages   =   "SAClient.ctx":208E
   ScaleHeight     =   1844.102
   ScaleMode       =   0  'User
   ScaleWidth      =   2895
   ToolboxBitmap   =   "SAClient.ctx":209F
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1560
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1080
      Top             =   0
   End
   Begin VB.Shape Highlight 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   61
      Left            =   60
      Shape           =   3  'Circle
      Top             =   30
      Width           =   60
   End
   Begin VB.Shape Indicator 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   128
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   192
   End
End
Attribute VB_Name = "SAClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private p_Application As String
Private p_IP As String
Private p_Node As String
Private p_Register As String
Private p_Count As String
Private p_Interval As Long
Private p_Enabled As Boolean
Private p_Value As String

Private statuses(7) As ColorConstants

Public XHRequest As MSXML2.XMLHTTP
Private firstPass As Boolean

Public Event OnChange(ByVal Value As String)

' properties
Public Property Let Application(ByVal NewVal As String)
    p_Application = NewVal
    PropertyChanged "Application"
End Property

Public Property Get Application() As String
Attribute Application.VB_ProcData.VB_Invoke_Property = "Properties"
    Application = p_Application
End Property

Public Property Let IP(ByVal NewVal As String)
    p_IP = NewVal
    PropertyChanged "IP"
End Property

Public Property Get IP() As String
Attribute IP.VB_ProcData.VB_Invoke_Property = "Properties"
    IP = p_IP
End Property

Public Property Let Node(ByVal NewVal As String)
    p_Node = NewVal
    PropertyChanged "Node"
End Property

Public Property Get Node() As String
Attribute Node.VB_ProcData.VB_Invoke_Property = "Properties"
    Node = p_Node
End Property

Public Property Let Register(ByVal NewVal As String)
    p_Register = NewVal
    PropertyChanged "Register"
End Property

Public Property Get Register() As String
Attribute Register.VB_ProcData.VB_Invoke_Property = "Properties"
    Register = p_Register
End Property

Public Property Let Count(ByVal NewVal As String)
    p_Count = NewVal
    PropertyChanged "Count"
End Property

Public Property Get Count() As String
Attribute Count.VB_ProcData.VB_Invoke_Property = "Properties"
    Count = p_Count
End Property

Public Property Let Interval(ByVal NewVal As Long)
    p_Interval = NewVal
    Timer1.Interval = p_Interval
    PropertyChanged "Interval"
End Property

Public Property Get Interval() As Long
Attribute Interval.VB_ProcData.VB_Invoke_Property = "Properties"
    Interval = p_Interval
End Property

Public Property Let Enabled(ByVal NewVal As Boolean)
    p_Enabled = NewVal
    If Not Ambient.UserMode Then
        setIndicator CInt(p_Enabled) + 1
    Else
        setIndicator CInt(p_Enabled) + 1
    End If
    Timer1.Enabled = p_Enabled And Ambient.UserMode
    PropertyChanged "Enabled"
End Property

Public Property Get Enabled() As Boolean
    Enabled = p_Enabled
End Property

Public Property Let Value(ByVal NewVal As String)
    p_Value = NewVal
    PropertyChanged "Value"
    If Ambient.UserMode Then Call Change(p_Value)
End Property

Public Property Get Value() As String
    Value = p_Value
End Property

'subs
Private Sub Change(ByVal Value As String)
    RaiseEvent OnChange(p_Value)
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Dim myXHR As XHR
    Dim varString As String
    
    Set myXHR = New XHR
    myXHR.setParent Me
    
    varString = varString & "ip=" & p_IP
    varString = varString & "&node=" & p_Node
    varString = varString & "&register=" & p_Register
    varString = varString & "&count=" & p_Count
    
    Set XHRequest = New MSXML2.XMLHTTP60
    XHRequest.onreadystatechange = myXHR
    XHRequest.open "POST", p_Application, True
    XHRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    XHRequest.setRequestHeader "Cache-Control", "no-cache"
    XHRequest.send (varString)
    
End Sub

Public Sub hasResponse(ByVal response As String)
    Me.Value = response
    Timer1.Enabled = p_Enabled And Ambient.UserMode
    If Ambient.UserMode Then
        Timer2.Interval = CInt(p_Interval * 0.25)
        Timer2.Enabled = True
    End If
End Sub

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    If p_Enabled Then setIndicator 6
End Sub

Private Sub UserControl_Initialize()
    p_Application = "http://192.168.0.3/homework/pactivex.php"
    p_IP = "0.0.0.0"
    p_Node = "1"
    p_Register = "4000"
    p_Count = "2"
    p_Interval = 5000
    p_Enabled = True
    p_Value = "-1"
    
    statuses(0) = &H80FF&     ' control enabled
    statuses(1) = vbBlack ' control disabled
    statuses(2) = vbRed ' http response code 1
    statuses(3) = vbWhite ' http response code 2
    statuses(4) = vbCyan ' http response code 3
    statuses(5) = vbGreen ' http response code 4
    statuses(6) = &H80FF&    ' content in response
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        p_Application = CStr(.ReadProperty("Application", "0.0.0.0"))
        p_IP = CStr(.ReadProperty("IP", "0.0.0.0"))
        p_Node = CStr(.ReadProperty("Node", "1"))
        p_Register = CStr(.ReadProperty("Register", "4000"))
        p_Count = CStr(.ReadProperty("Count", "2"))
        p_Interval = CLng(.ReadProperty("Interval", 5000))
        p_Enabled = CBool(.ReadProperty("Enabled", True))
        p_Value = CStr(.ReadProperty("Value", "-1"))
    End With
    If firstPass = False Then
        firstPass = True
        Timer1.Interval = p_Interval
        Timer1.Enabled = p_Enabled And Ambient.UserMode
        setIndicator CInt(p_Enabled) + 1
    End If
End Sub

Private Sub UserControl_Resize()
    Dim cWidth As Integer
    Dim cHeight As Integer
    cWidth = 1050
    cHeight = 580
    ' max width
    If UserControl.Width > cWidth Then UserControl.Width = cWidth
    'max height
    If UserControl.Height > cHeight Then UserControl.Height = cHeight
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Application", CStr(p_Application)
        .WriteProperty "IP", CStr(p_IP)
        .WriteProperty "Node", CStr(p_Node)
        .WriteProperty "Register", CStr(p_Register)
        .WriteProperty "Count", CStr(p_Count)
        .WriteProperty "Interval", CLng(p_Interval)
        .WriteProperty "Enabled", CInt(p_Enabled)
        .WriteProperty "Value", CStr(p_Value)
    End With
End Sub

Public Sub setIndicator(ByRef num As Integer)
    Indicator.FillColor = statuses(num)
End Sub
