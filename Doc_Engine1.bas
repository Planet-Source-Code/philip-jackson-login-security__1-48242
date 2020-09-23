Attribute VB_Name = "Doc_Engine1"
Option Explicit

Global LogOnResult As String ''User Full Info

Global gVarUserName, gVarPassword, gVarAccessLevel, _
                gVarDescripition As String

Sub AssignUserInfoToGlobalVar(UserInfo As String)
    Dim TDM
    Dim loop1, counter As Integer
    Dim tmpString, Char As String
    counter = 1
    For loop1 = 1 To Len(UserInfo)
        TDM = DoEvents()
        Char = Mid(UserInfo, loop1, 1)
        If Char = Chr(10) Then
        
        tmpString = Trim(Mid(tmpString, InStr(1, tmpString, ":", 1) + 2, Len(tmpString) - InStr(1, tmpString, ":", 1)))
              If counter = 1 Then gVarUserName = tmpString
              If counter = 2 Then gVarPassword = tmpString
              If counter = 3 Then gVarAccessLevel = tmpString
              If counter = 4 Then gVarDescripition = tmpString
              
           tmpString = ""
           counter = counter + 1
        End If
        If Char <> Chr(13) And Char <> Chr(10) Then tmpString = tmpString & Char
    Next loop1
End Sub
