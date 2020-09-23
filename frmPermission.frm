VERSION 5.00
Begin VB.Form frmPermission 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Assign User(s)  Permission(s)"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8940
   Icon            =   "frmPermission.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   8940
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H000040C0&
      Caption         =   "&Exit"
      Height          =   375
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdSaveSettings 
      BackColor       =   &H00C0C000&
      Caption         =   "&Save Settings"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Check to enable "
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   41
         Left            =   7920
         TabIndex        =   50
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   40
         Left            =   6960
         TabIndex        =   49
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   39
         Left            =   6000
         TabIndex        =   48
         Top             =   3240
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   38
         Left            =   5040
         TabIndex        =   47
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   37
         Left            =   4080
         TabIndex        =   46
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   36
         Left            =   3000
         TabIndex        =   45
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   35
         Left            =   2040
         TabIndex        =   44
         Top             =   3240
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   34
         Left            =   7920
         TabIndex        =   43
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   33
         Left            =   6960
         TabIndex        =   42
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   32
         Left            =   6000
         TabIndex        =   41
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   31
         Left            =   5040
         TabIndex        =   40
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   30
         Left            =   4080
         TabIndex        =   39
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   29
         Left            =   3000
         TabIndex        =   38
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   28
         Left            =   2040
         TabIndex        =   37
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   27
         Left            =   7920
         TabIndex        =   36
         Top             =   2280
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   26
         Left            =   6960
         TabIndex        =   35
         Top             =   2280
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   25
         Left            =   6000
         TabIndex        =   34
         Top             =   2280
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   24
         Left            =   5040
         TabIndex        =   33
         Top             =   2280
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   23
         Left            =   4080
         TabIndex        =   32
         Top             =   2280
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   22
         Left            =   3000
         TabIndex        =   31
         Top             =   2280
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   21
         Left            =   2040
         TabIndex        =   30
         Top             =   2280
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   20
         Left            =   7920
         TabIndex        =   29
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   19
         Left            =   6960
         TabIndex        =   28
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   18
         Left            =   6000
         TabIndex        =   27
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   17
         Left            =   5040
         TabIndex        =   26
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   16
         Left            =   4080
         TabIndex        =   25
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   15
         Left            =   3000
         TabIndex        =   24
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   14
         Left            =   2040
         TabIndex        =   23
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   13
         Left            =   7920
         TabIndex        =   22
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   12
         Left            =   6960
         TabIndex        =   21
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   11
         Left            =   6000
         TabIndex        =   20
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   10
         Left            =   5040
         TabIndex        =   19
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   9
         Left            =   4080
         TabIndex        =   18
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   17
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   16
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   7920
         TabIndex        =   15
         Top             =   840
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   6960
         TabIndex        =   14
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   4
         Left            =   6000
         TabIndex        =   13
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   3
         Left            =   5040
         TabIndex        =   12
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   11
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   10
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00B9913C&
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   720
         TabIndex        =   56
         Top             =   3270
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   55
         Top             =   2800
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   54
         Top             =   2310
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   53
         Top             =   1830
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   52
         Top             =   1350
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   51
         Top             =   860
         Width           =   120
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Access Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View Users"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7440
         TabIndex        =   7
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Users"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6720
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5880
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T,Bar Buttons"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4560
         TabIndex        =   4
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Menu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   3
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   405
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         FillColor       =   &H00B9913C&
         FillStyle       =   0  'Solid
         Height          =   3375
         Left            =   480
         Shape           =   2  'Oval
         Top             =   240
         Width           =   7095
      End
   End
End
Attribute VB_Name = "frmPermission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub initDB()
Dim loop1 As Integer
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase(App.Path & "\Login.mdb", False, False, ";pwd=AdmiN")
             
Set rs = db.OpenRecordset("PermissionTable", dbOpenTable)

For loop1 = 1 To 6
    rs.AddNew
    rs.Fields("AccessLevel") = loop1
    rs.Fields("Permission") = "0000000"
    rs.Update
Next loop1

Set db = Nothing
Set rs = Nothing
End Sub

Sub UpdateDB()

Dim loop1, loop3, counter As Integer
Dim strn As String
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase(App.Path & "\Login.mdb", False, False, ";pwd=AdmiN")
             
Set rs = db.OpenRecordset("PermissionTable", dbOpenTable)
    counter = 0
    rs.MoveFirst
    strn = ""
For loop1 = 1 To 6
      rs.Edit
      For loop3 = 1 To 7
        strn = strn & Trim(str(Check1(counter + loop3 - 1).Value))
      Next loop3
      rs.Fields("Permission") = strn
      strn = ""
      rs.Update
      counter = counter + 7
    If rs.EOF = False Then rs.MoveNext
Next loop1

Set db = Nothing
Set rs = Nothing
End Sub

Sub loadvaluesToChkbox()
Dim loop1, loop2, counter As Integer
Dim str As String
Dim db As Database
Dim rs As Recordset
Set db = OpenDatabase(App.Path & "\Login.mdb", False, False, ";pwd=AdmiN")
            
Set rs = db.OpenRecordset("PermissionTable", dbOpenTable)
counter = 0
    rs.MoveFirst
For loop1 = 1 To 6
    str = Trim(rs.Fields("Permission"))
    For loop2 = 1 To 7
        Check1(counter).Value = Int(Val(Mid(str, loop2, 1)))
        counter = counter + 1
    Next loop2
    If rs.EOF = False Then rs.MoveNext
Next loop1

Set db = Nothing
Set rs = Nothing
End Sub

Private Sub Check1_Click(Index As Integer)
Check1(40).Value = 1
Check1(41).Value = 1
End Sub

Private Sub cmdSaveSettings_Click()
Call UpdateDB
MsgBox "Settings have been successfully changed.  ", vbInformation, "Settings changed"
Check1(0).SetFocus
End Sub

Private Sub Command1_Click()
Unload Me
FAdministrator.Show
End Sub

Private Sub Form_Activate()
Dim Doc_ENGINE As Doc_ENGINE
Set Doc_ENGINE = New Doc_ENGINE
Dim loop1 As Integer
' Check if file exist
If Doc_ENGINE.ReportFileStatus(App.Path & "\Login.mdb") = False Then
       Call initDB
    Call loadvaluesToChkbox
End If
'Save to DB
Call loadvaluesToChkbox
Check1(0).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub
