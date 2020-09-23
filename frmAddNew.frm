VERSION 5.00
Begin VB.Form frmAddNew 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AddNew User"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6435
   Icon            =   "frmAddNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Save"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "Edit"
      Height          =   375
      Left            =   4320
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cboUsers 
      Height          =   315
      Left            =   4440
      TabIndex        =   5
      ToolTipText     =   "User"
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdSetPermission 
      BackColor       =   &H008080FF&
      Caption         =   "Permission"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAccessLevel 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "Access Level"
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0080FF80&
      Caption         =   "Add User"
      Height          =   375
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtDescripition 
      Height          =   735
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "User Description"
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FF80FF&
      Caption         =   "Exit"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtConfirmPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   2
      Tag             =   "1"
      ToolTipText     =   "Confirm Password"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Password"
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "User Name"
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select A User"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4560
      TabIndex        =   14
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AccessLevel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   12
      Top             =   1800
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Descripition"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   10
      Top             =   1320
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   9
      Top             =   840
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   8
      Top             =   360
      Width           =   945
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00B9913C&
      BackStyle       =   1  'Opaque
      Height          =   3375
      Left            =   360
      Shape           =   2  'Oval
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "frmAddNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim frmGVAR_cboUsersText As String

Private Sub cboUsers_Change()
If Len(cboUsers.Text) > 0 Then
cmdSetPermission.Visible = True
Command1.Visible = True
Else
cmdSetPermission.Visible = False
Command1.Visible = False
End If
End Sub

Private Sub cmdSave_Click()
  Dim samepassword As Boolean
  Dim db As Database
  Dim rs As Recordset
    
  Set db = OpenDatabase(App.Path & "\Login.mdb", False, False, ";pwd=AdmiN")
  Set rs = db.OpenRecordset("Password_Security")
  
  If Trim(txtUserName.Text) = "" Or Trim(txtPassword.Text) = "" Then
  MsgBox "Please enter Username & Password"
  txtUserName.SetFocus
  Exit Sub
  End If
    If txtPassword.Text = txtConfirmPassword.Text Then
    
        samepassword = True
    Else
        samepassword = False
        MsgBox "The passwords do not match", , progname
       
        txtConfirmPassword.SetFocus
        Exit Sub
    End If
   
 Do While (Not rs.EOF)
    rs.MoveLast
    rs.AddNew
    rs.Fields!User_Name = txtUserName.Text
    rs.Fields!User_Password = txtPassword.Text
    rs.Fields!AccessLevel = txtAccessLevel.Text
    rs.Fields!User_Descripition = txtDescripition.Text
    rs.Update
    
    MsgBox "The User Has Been added"
  
        Call LockFields(True)
        Call ClearFields
     
      Exit Sub
      rs.MoveNext
  Loop
        cboUsers.Enabled = False
        DoEvents
End Sub

Private Sub cmdexit_Click()
    Unload Me
FAdministrator.Show
End Sub

Private Sub Command1_Click()
Call LockFields(False)
Command2.Visible = True
End Sub

Private Sub Command2_Click()
  Dim samepassword As Boolean
  Dim db As Database
  Dim rs As Recordset
    
  Set db = OpenDatabase(App.Path & "\Login.mdb", False, False, ";pwd=AdmiN")
  Set rs = db.OpenRecordset("Password_Security")
  
  If Trim(txtConfirmPassword.Text) = "" Or Trim(txtPassword.Text) = "" Then
  MsgBox "Please enter confirmpassword & Password"
  txtConfirmPassword.SetFocus
  Exit Sub
  End If
    If txtPassword.Text = txtConfirmPassword.Text Then
        samepassword = True
    Else
        samepassword = False
        MsgBox "The passwords do not match", , progname
       
        txtConfirmPassword.SetFocus
        txtConfirmPassword.Text = ""
        Exit Sub
    End If
 Do While (Not rs.EOF)
    rs.MoveLast
    rs.Edit
    rs.Fields!User_Name = txtUserName.Text
    rs.Fields!User_Password = txtPassword.Text
    rs.Fields!AccessLevel = txtAccessLevel.Text
    rs.Fields!User_Descripition = txtDescripition.Text
    rs.Update
    MsgBox "The User Has Been updated"
   txtConfirmPassword.Text = ""
     Command2.Visible = False
      
      rs.MoveNext
       Call LockFields(True)
       Call ClearFields
      
       Exit Sub
  Loop
  
        DoEvents
End Sub

Private Sub Form_Load()
Me.Caption = "Login Database - " & App.CompanyName
  
  cboUsers.Clear '' Clears combobox
  Dim Doc_ENGINE As Doc_ENGINE
  Set Doc_ENGINE = New Doc_ENGINE
  
  Call Doc_ENGINE.LoadUsers(cboUsers)
 
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub ClearFields()
Dim indx As Integer
Dim TempMask As String
 
With Me.Controls

For indx = 0 To .Count - 1
If Me.Controls(indx).Tag = "1" Then
If (TypeOf Me.Controls(indx) Is TextBox) Then
    Me.Controls(indx).Text = ""
        End If
    End If
Next
End With
    DoEvents
End Sub

Public Sub LockFields(bDoLock As Boolean)
Dim indx As Integer
For indx = 0 To Me.Controls.Count - 1
If Me.Controls(indx).Tag = "1" Then
If (TypeOf Me.Controls(indx) Is TextBox) Then
If (bDoLock = True) Then
    Me.Controls(indx).Locked = True
    Me.Controls(indx).BackColor = vbWhite
Else
    Me.Controls(indx).Locked = False
    Me.Controls(indx).BackColor = vbWhite
End If
End If
End If
Next
DoEvents
End Sub


Private Sub cboUsers_Click()
Dim Doc_ENGINE As Doc_ENGINE
Set Doc_ENGINE = New Doc_ENGINE
    
Call Doc_ENGINE.getUserInfo(txtUserName, txtPassword, _
                txtAccessLevel, txtDescripition, cboUsers)

Call LockFields(True)
frmGVAR_cboUsersText = cboUsers.Text

If Len(cboUsers.Text) > 0 Then
cmdSetPermission.Visible = True
Command1.Visible = True
Else
cmdSetPermission.Visible = False
Command1.Visible = False
End If

End Sub

Private Sub txtUserName_Change()
txtUserName.Text = UCase(txtUserName.Text)
txtUserName.SelStart = (Len(UCase(txtUserName.Text)))
End Sub

Public Sub LoadUsers(cbo As ComboBox)  ''load UserNames in a listbox or combo box
    On Error GoTo ErrorHandler:
    Dim db As Database
    Dim rs As Recordset
    Dim TDM As Variant
    Dim loop1 As Integer
    cboUsers.Clear
    
    Set db = OpenDatabase(App.Path & "\Login.mdb", False, False, ";pwd=AdmiN")
             
    Set rs = db.OpenRecordset("Password_security", dbOpenTable)
    
    Singlelize cboUsers
    
    rs.MoveFirst
    If rs.RecordCount > 0 Then
        For loop1 = 1 To rs.RecordCount
            TDM = DoEvents()
            cboUsers.AddItem rs.Fields("User_Name")
            rs.MoveNext
        Next loop1
     End If
     
  db.Close
  Exit Sub
  
ErrorHandler:
    db.Close

End Sub

Private Sub cmdSetPermission_Click()
If gVarAccessLevel < 6 Then
    
    If cboUsers.Enabled = True Then
        cboUsers.SetFocus
    Else
        txtUserName.SetFocus
    End If
    Exit Sub
End If
frmPermission.Show
If cboUsers.Enabled = True Then
cboUsers.SetFocus
Else
txtUserName.SetFocus
End If
Unload Me
End Sub

Private Sub txtAccessLevel_LostFocus()

If Val(txtAccessLevel.Text) > 6 Then
    MsgBox "'Access Level' should not be greater than 6. ", vbInformation, "Invalid Entry"
    txtAccessLevel.Text = ""
    txtAccessLevel.SetFocus
    Exit Sub
End If

If Val(txtAccessLevel.Text) > gVarAccessLevel Then
   MsgBox "'Access Level' should not be greater than your own access level. ", vbInformation, "Invalid Entry"
    txtAccessLevel.Text = ""
    txtAccessLevel.SetFocus
    Exit Sub
End If
End Sub
