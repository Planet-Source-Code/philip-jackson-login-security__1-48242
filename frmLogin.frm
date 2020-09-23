VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4440
   ClientLeft      =   2790
   ClientTop       =   3045
   ClientWidth     =   6060
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmLogin.frx":1CFA
   ScaleHeight     =   2623.3
   ScaleMode       =   0  'User
   ScaleWidth      =   5690.013
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox txtUserName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   2160
      TabIndex        =   0
      ToolTipText     =   "Users Name"
      Top             =   1200
      Width           =   2295
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   4185
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   23481
            MinWidth        =   23481
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1920
      TabIndex        =   2
      ToolTipText     =   "Click Ok To Login"
      Top             =   3720
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3360
      TabIndex        =   3
      ToolTipText     =   "Click Cancel To Cancel Login"
      Top             =   3720
      Width           =   900
   End
   Begin VB.TextBox txtPassword 
      DataField       =   "Password"
      DataSource      =   "data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Enter Password"
      Top             =   2640
      Width           =   2325
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Tango BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   2040
      MouseIcon       =   "frmLogin.frx":2004
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3360
      Width           =   2115
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Please Login"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   345
      Left            =   2400
      TabIndex        =   7
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   1020
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   2160
      TabIndex        =   5
      Top             =   2400
      Width           =   900
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   2775
      Left            =   960
      Shape           =   2  'Oval
      Top             =   480
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00B9913C&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   360
      Shape           =   2  'Oval
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Dim strSQL As String
Dim ctr As Integer
Dim xText

Private Sub Form_Activate()
Dim ws As Workspace
Dim db As Database
Dim rs As Recordset
Dim max As Long
Dim i As Integer
 
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path + "\Login.mdb", False, False, ";pwd=AdmiN")
Set rs = db.OpenRecordset("Password_Security", dbOpenTable)
max = rs.RecordCount
rs.MoveFirst

For i = 1 To max
txtUserName.AddItem rs!User_Name
Singlelize txtUserName
rs.MoveNext
Next i
db.Close
End Sub

Private Sub Form_Load()
Dim hr&, dl&
    Dim usew&, useh&
    usew& = Me.Width / Screen.TwipsPerPixelX
    useh& = Me.Height / Screen.TwipsPerPixelY
    hr& = CreateEllipticRgn(0, 0, usew, useh)
    dl& = SetWindowRgn(Me.hwnd, hr, True)
End Sub

Private Sub cmdCancel_Click()
 End
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(1).Text = "Click Cancel"
End Sub

Private Sub cmdOK_Click()

Dim db As Database
Dim rs As Recordset
Dim Doc_ENGINE As Doc_ENGINE
Set Doc_ENGINE = New Doc_ENGINE

LogOnResult = Doc_ENGINE.LogOnValidate(Trim(txtUserName.Text), Trim(txtPassword.Text))
If Trim(LogOnResult) <> " " Then
      LogOnResult = LogOnResult
      AssignUserInfoToGlobalVar (LogOnResult)
      
    Set db = OpenDatabase(App.Path & "\Login.mdb", False, False, ";pwd=AdmiN")
    Set rs = db.OpenRecordset("Password_Security")


    Do While Not rs.EOF
             
      If rs.Fields("user_name") = (txtUserName.Text) And _
        rs.Fields("User_password") = (txtPassword.Text) Then
        fMain.Show
        UserName = txtUserName.Text
        Unload Me
       Exit Sub
    Else
        rs.MoveNext
        End If
    Loop
     
         ctr = ctr + 1
        If ctr = 4 Then
           End
        Else
    txtPassword.SetFocus
    txtPassword.Text = ""
            xText = "You have" + str(4 - ctr) + " tries left"
            If ctr = 3 Then
                xText = "This is your last chance!!"
            End If
            MsgBox "Access Denied!!" & vbCrLf & _
                   xText, vbOKOnly + vbCritical, progname
             
            SendKeys "{Home}+{End}"
        End If
        End If
End Sub

Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(1).Text = "Click OK"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(1).Text = ""
End Sub

Private Sub Label2_Click()
fChangePassword.Show
End Sub

Private Sub txtPassword_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(1).Text = "Password"
End Sub

Private Sub txtUserName_Click()
txtPassword.SetFocus
End Sub

Private Sub txtUserName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
StatusBar1.Panels(1).Text = "UserName"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub
