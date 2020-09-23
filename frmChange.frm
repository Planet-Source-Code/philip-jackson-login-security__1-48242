VERSION 5.00
Begin VB.Form fChangePassword 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3585
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5490
   ControlBox      =   0   'False
   Icon            =   "frmChange.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtConfirm 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Confirm Password"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtNewPassword 
      DataField       =   "Password"
      DataSource      =   "Data1"
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "New Password"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtOldPassword 
      DataField       =   "User_Password"
      DataSource      =   "Data1"
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "Old Password"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Tango BT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   1650
      TabIndex        =   8
      Top             =   600
      Width           =   2355
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   960
      TabIndex        =   7
      Top             =   2040
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "&New Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1200
      TabIndex        =   6
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "&Old Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1320
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   2895
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Width           =   4575
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00B9913C&
      BackStyle       =   1  'Opaque
      Height          =   3135
      Left            =   120
      Shape           =   2  'Oval
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "fChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
  If frmLogin.Visible Then
  frmLogin.Show
  ElseIf FAdministrator.Visible Then
  FAdministrator.Show
  Else
  fMain.Show
  End If
  End Sub

Private Sub cmdChange_Click()
 Dim db As Database
 Dim rs As Recordset
    
    
    Set db = OpenDatabase(App.Path & "\Login.mdb", False, False, ";pwd=AdmiN")
    Set rs = db.OpenRecordset("Password_Security")
    
    If (Len(txtOldPassword) < 1) Then
    MsgBox "Old password Cann't be blank" & vbCrLf & _
    "Please enter the old password", vbInformation, progname
    txtOldPassword.SetFocus
    Else
    If (Len(txtNewPassword) < 1) Then
    MsgBox "New password Cann't be blank" & vbCrLf & _
    "Please enter a new password", vbInformation, progname
    txtNewPassword.SetFocus
    ElseIf (Len(txtConfirm) < 1) Then
    MsgBox "Confirm password cann't be blank" & vbCrLf & _
    "Please enter the correct password", vbInformation, progname
    txtConfirm.SetFocus
    Else
    
  Do While (Not rs.EOF)
  
    If (txtOldPassword.Text) = rs.Fields("User_Password") Then
       
       If (txtNewPassword.Text) = (txtConfirm.Text) Then
 
      rs.Edit
     rs.Fields("User_Password") = (txtNewPassword.Text)
     rs.Update
     
     MsgBox "Password changed.", vbInformation, progname
    
     Else
     
     MsgBox "Password does not match!!!"
     txtConfirm.SetFocus
     txtConfirm.Text = ""
     
    End If
     Unload Me
     FAdministrator.Show
      Exit Sub
  Else
      rs.MoveNext
     End If
  Loop
  End If
  End If
 DoEvents
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub
