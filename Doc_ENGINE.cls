VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Doc_ENGINE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function CheckPermission(Level As Integer, Fcode As Integer) As Boolean
Dim loop1, loop2, counter As Integer
Dim str As String
Dim db As Database
Dim rec As Recordset
Set db = OpenDatabase(App.Path & "\Login.mdb", False, False, ";pwd=AdmiN")
             
Set rec = db.OpenRecordset("PermissionTable", dbOpenTable)
counter = 0
    rec.MoveFirst
For loop1 = 1 To Level
    str = Trim(rec.Fields("Permission"))
    For loop2 = 1 To 7
        If loop2 = Fcode And loop1 = Level Then
           If Int(Val(Mid(str, loop2, 1))) = 1 Then
              CheckPermission = True
              Exit Function
           Else
              MsgBox "You don't have the permission to use this feature.  ", vbInformation, progname & "-" & " !!!!  Access Denied  !!!!"
              CheckPermission = False
              Exit Function
           End If
        End If
        counter = counter + 1
    Next loop2
    If rec.EOF = False Then rec.MoveNext
Next loop1

Set db = Nothing
Set rec = Nothing
End Function

Function LogOnValidate(UserName As String, Password As String) As String

    Dim db As Database
    Dim rs As Recordset
    Dim TDM As Variant
    Dim loop1 As Integer

    LogOnValidate = 0
    
    Set db = OpenDatabase(App.Path & "\Login.mdb", False, False, ";pwd=AdmiN")
    Set rs = db.OpenRecordset("Password_security", dbOpenTable)
    
    rs.MoveFirst
    If rs.RecordCount > 0 Then
        For loop1 = 1 To rs.RecordCount
            TDM = DoEvents()
            
            If rs.Fields("User_Name") = UserName And _
                rs.Fields("User_Password") = Password Then
               
            LogOnValidate = "txtUserName : " & rs.Fields("User_Name") & vbCrLf _
                                & "txtPassword : " & rs.Fields("User_Password") & vbCrLf _
                                 & "txtAccessLevel : " & rs.Fields("AccessLevel") & vbCrLf _
                                 & "txtDescripition : " & rs.Fields("User_Descripition")
                Exit For
            Else
                LogOnValidate = ""
            End If
            rs.MoveNext  ' Move to the next rsord
        Next loop1
    End If
    
    db.Close
End Function

Function ReportFileStatus(filespec) As Boolean '' Check if file is present
   Dim fso, Msg
   Set fso = CreateObject("Scripting.FileSystemObject")
   If (fso.FileExists(filespec)) Then
      Msg = -1
   Else
      Msg = 0
   End If
   ReportFileStatus = Msg
End Function

Sub LoadUsers(cbo As ComboBox)  ''load UserNames in a combobox
    On Error GoTo ErrorHandler:
    Dim db As Database
    Dim rs As Recordset
    Dim TDM As Variant
    Dim loop1 As Integer
    cbo.Clear
    
    Set db = OpenDatabase(App.Path & "\Login.mdb", False, False, ";pwd=AdmiN")
             
    Set rs = db.OpenRecordset("Password_security", dbOpenTable)
    
    rs.MoveFirst
    If rs.RecordCount > 0 Then
        For loop1 = 1 To rs.RecordCount
            TDM = DoEvents()
            cbo.AddItem rs.Fields("User_Name")
            rs.MoveNext
        Next loop1
     End If
     
  db.Close
  Exit Sub
  
ErrorHandler:
    db.Close
End Sub

Sub getUserInfo(txtUserName As TextBox, txtPassword As TextBox, _
                txtAccessLevel As TextBox, txtDescripition As TextBox, cbo As ComboBox)
                
On Error GoTo Err:
    Dim db As Database
    Dim rs As Recordset
    Dim TDM As Variant
    Dim loop1 As Integer
    
    Set db = OpenDatabase(App.Path & "\Login.mdb", False, False, ";pwd=AdmiN")

    Set rs = db.OpenRecordset("Password_security", dbOpenTable)
    
If rs.BOF = True And rs.EOF = True Then Exit Sub 'chk if rs exists

rs.MoveFirst
    If rs.RecordCount > 0 Then
        For loop1 = 1 To rs.RecordCount
            TDM = DoEvents()
            If cbo.Text = rs.Fields("User_Name") Then
                    txtUserName.Text = rs.Fields("User_Name")
                    txtPassword.Text = rs.Fields("User_Password")
                    txtAccessLevel.Text = rs.Fields("AccessLevel")
                    txtDescripition.Text = rs.Fields("User_Descripition")
                    db.Close
                    Exit Sub
            End If
            rs.MoveNext
        Next loop1
     End If
     
  db.Close
Err:
End Sub
