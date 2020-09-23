VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmWelcome 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Viewer Login"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2895
   Icon            =   "frmWelcome.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   2895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Delete"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000080FF&
      Caption         =   "&Exit"
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "View Users"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "frmWelcome.frx":1CFA
      Height          =   1575
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2778
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      Appearance      =   0
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "Hide Users"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome Admin"
      BeginProperty Font 
         Name            =   "Tango BT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   90
      TabIndex        =   4
      Top             =   240
      Width           =   2685
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00B9913C&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   120
      Shape           =   2  'Oval
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Call MSFlexGrid1_KeyUp(vbKeyDelete, 0)
End Sub

Private Sub Command2_Click()
    frmWelcome.Width = 6915
    Command3.Visible = True
    Command2.Visible = False
If Command3.Visible = True Then
Command1.Visible = True
Else
Command1.Visible = False
End If
    
End Sub
Private Sub Command3_Click()
    frmWelcome.Width = 3030
    Command2.Visible = True
    Command3.Visible = False
    Command1.Visible = Fals
End Sub

Private Sub Command4_Click()
Unload Me
FAdministrator.Show
End Sub

Private Sub Form_Load()
Call strRefresh
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
SendKeys "{TAB}"
KeyAscii = 0
End If
End Sub

Private Sub strRefresh()
Dim db As Database
Dim SQL As String
Dim rs As Recordset

    MSFlexGrid1.Visible = False
    
    Set db = OpenDatabase(App.Path & "\Login.mdb", False, False, ";pwd=AdmiN")
    SQL = ("SELECT * From Password_Security")
    Set rs = db.OpenRecordset(SQL, dbOpenDynaset)
        
    With MSFlexGrid1
        .Cols = 5
        .Rows = 1
        .Visible = True
        
        .ColWidth(0) = 600
        .Col = 0
        .Row = 0
        .Text = "UserID"
        
        .ColWidth(1) = 1000
        .Col = 1
        .Row = 0
        .Text = "User_Name"
        
        .ColWidth(2) = 1200
        .Col = 2
        .Row = 0
        .Text = "User_Password"
        
        .ColWidth(3) = 1200
        .Col = 3
        .Row = 0
        .Text = "Access Level"
        
        .ColWidth(4) = 1550
        .Col = 4
        .Row = 0
        .Text = "User_Description"
        
   Do While Not rs.EOF
            .AddItem rs![UserID] & _
                          vbTab & rs![User_Name] & _
                          vbTab & rs![User_Password] & _
                          vbTab & rs![AccessLevel] & _
                          vbTab & rs![User_Descripition]
                          
            rs.MoveNext
        Loop
'
        .Visible = True
    End With
End Sub


Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim db As Database
Dim SQL As String
Dim rs As Recordset
Dim i As Long
Dim iStep As Integer
Dim strWhere As String
Dim RowCount As Integer
Dim Ans As String

Set db = OpenDatabase(App.Path & "\Login.mdb", False, False, ";pwd=AdmiN")
       
    If KeyCode = vbKeyDelete Then
       
        With MSFlexGrid1
            
            If .Row > .RowSel Then
                RowCount = (.Row - .RowSel) + 1
                iStep = -1
            Else
                iStep = 1
                RowCount = (.RowSel - .Row) + 1
            End If
                        
            Ans = MsgBox("Are you sure you want to deleted the selected " & RowCount & " record(s)?", vbYesNo + vbCritical + vbDefaultButton2)
            
            Select Case Ans
                Case vbYes
                    MSFlexGrid1.Visible = False
                        For i = .Row To .RowSel Step iStep
                            If Len(strWhere) > 0 Then
                                strWhere = strWhere & " or "
                            End If
                            
                            strWhere = strWhere & "UserID =  " & .TextMatrix(i, 0)
                            SQL = "DELETE From PAssword_Security Where " & strWhere
                            db.Execute SQL
                        Next
                        db.Close
                        Me.MousePointer = vbNormal
                Case vbNo
                    db.Close
                    Me.MousePointer = vbNormal
                    Exit Sub
            End Select
        End With
       Call strRefresh
        MsgBox "Delete completed."
    End If
End Sub
