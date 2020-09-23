Attribute VB_Name = "Module1"
Option Explicit
   

Public db As Database

Public Const progname = "Login Database "

Public rsLoginTable As Recordset

Public UserName As String


Public Function OpenTheDatabase() As Boolean

Dim dbPath As String

On Error GoTo dbErrors

dbPath = (App.Path) & "\Login.mdb"
Set db = OpenDatabase(App.Path & "\Login.mdb", False, False, ";pwd=AdmiN")

Set db = DBEngine.Workspaces(0).OpenDatabase(dbPath, False)

Set rsLoginTable = db.OpenRecordset("Password_Security", dbOpenTable)

OpenTheDatabase = True

Exit Function

dbErrors:

    OpenTheDatabase = False
    MsgBox (Err.Description)
    
End Function

'stop duplicates in combo box or list box
Public Sub Singlelize(ListObject As Object)
    Dim i As Integer
    Dim X As Integer
    
    For i = 0 To ListObject.ListCount - 1
        For X = 0 To ListObject.ListCount - 1
            If i <> X Then
                If LCase(ListObject.List(X)) = LCase(ListObject.List(i)) Then
                    ListObject.RemoveItem X
                    X = i
                End If
            End If
        Next X
    Next i
End Sub

Public Sub Main()
Load frmLogin
frmLogin.Show
End Sub
