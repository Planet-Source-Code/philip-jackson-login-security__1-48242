VERSION 5.00
Begin VB.Form FAdministrator 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Administrator Level"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "FAdministrator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Change Password"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "View Password"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New User"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Administrator Level"
      BeginProperty Font 
         Name            =   "Tango BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1170
      TabIndex        =   4
      Top             =   600
      Width           =   2355
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   2535
      Left            =   480
      Shape           =   2  'Oval
      Top             =   240
      Width           =   4095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00B9913C&
      BackStyle       =   1  'Opaque
      Height          =   2775
      Left            =   0
      Shape           =   2  'Oval
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "FAdministrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmAddNew.Show
Unload Me
End Sub

Private Sub Command2_Click()
Dim Doc_ENGINE As Doc_ENGINE
Set Doc_ENGINE = New Doc_ENGINE

If Doc_ENGINE.CheckPermission(Val(gVarAccessLevel), 7) = False Then Exit Sub
Unload Me

frmWelcome.Show
End Sub

Private Sub Command3_Click()
fChangePassword.Show
Unload Me
End Sub

Private Sub Command4_Click()
Unload Me
frmLogin.Show
End Sub

Private Sub Form_Load()
Me.Caption = "Login Database - " & App.CompanyName
End Sub
