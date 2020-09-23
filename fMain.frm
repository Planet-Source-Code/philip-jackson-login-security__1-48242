VERSION 5.00
Begin VB.Form fMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   2280
   End
   Begin VB.Label loginbar 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2190
      TabIndex        =   0
      Top             =   1320
      Width           =   75
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2415
      Left            =   600
      Shape           =   2  'Oval
      Top             =   360
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00B9913C&
      FillColor       =   &H00B9913C&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   240
      Shape           =   2  'Oval
      Top             =   120
      Width           =   4095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnulogoff 
         Caption         =   "LogOff"
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "View"
      Begin VB.Menu mnuadmin 
         Caption         =   "Administrator Level"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Me.Caption = "Login Database - " & App.CompanyName

End Sub

Private Sub mnuabout_Click()
MsgBox "Thanks for using this login program " & vbCrLf & _
                    "Please Vote for this small but good program. " & vbCrLf & _
                    "I got some of the code off  " & vbCrLf & _
                    "WWW.Plante-Soucre-Code.com  " & vbCrLf & _
                    "When i was looking for a Security level program  " & vbCrLf & _
                    "I was looking to set different level of access  ", vbOKOnly + vbCritical, progname
             
End Sub

Private Sub mnuadmin_Click()
Dim Doc_ENGINE As Doc_ENGINE
Set Doc_ENGINE = New Doc_ENGINE

If Doc_ENGINE.CheckPermission(Val(gVarAccessLevel), 6) = False Then Exit Sub
Unload Me

FAdministrator.Show
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnulogoff_Click()
Unload Me
frmLogin.Show
End Sub

Private Sub Timer1_Timer()
loginbar.Caption = "User = " & UserName
End Sub
