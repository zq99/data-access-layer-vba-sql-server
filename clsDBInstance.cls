Option Explicit

Private objSharedClass As clsDatabase

Public Function GetNewDatabase(Optional ByVal DBName As String) As clsDatabase
    Set GetNewDatabase = CreateDatabase(DBName)
End Function

Public Function GetSharedDatabase(Optional ByVal DBName As String) As clsDatabase
    If objSharedClass Is Nothing Then
        Set objSharedClass = CreateDatabase(DBName)
    End If
    Set GetSharedDatabase = objSharedClass
End Function

Public Function CloseSharedDatabase() As Boolean
    CloseSharedDatabase = objSharedClass.CloseDatabase
    Set objSharedClass = Nothing
End Function

Private Function CreateDatabase(Optional ByVal Name As String) As clsDatabase

    Dim objDatabase      As clsDatabase
    Dim objDBCredentials As clsDBCredentials
    
    Set objDBCredentials = New clsDBCredentials
    Set objDatabase = New clsDatabase
    objDBCredentials.SetServer = "WORK-PC"
    objDBCredentials.SetDBName = "DataAccessDemo"
    objDBCredentials.SetTrustedSecurity = True
    objDatabase.OpenDatabase objDBCredentials
    Set CreateDatabase = objDatabase

End Function