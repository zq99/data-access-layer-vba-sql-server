Option Explicit


Public Sub GetListOfCountries()

    Dim rstData     As New ADODB.Recordset
    Dim oDBInstance As New clsDBInstance
    Dim oDatabase   As clsDatabase
    
    Set oDatabase = oDBInstance.GetSharedDatabase()
    Set rstData = oDatabase.GetRecordsetFromStoredProc("spGetItems")
    rstData.MoveFirst
    Do While Not rstData.EOF
        Debug.Print rstData.Fields("country_name").Value & " " & rstData.Fields("country_region").Value
        rstData.MoveNext
    Loop
    
    oDBInstance.CloseSharedDatabase
    Set rstData = Nothing
    Set oDBInstance = Nothing
    Set oDatabase = Nothing

End Sub




Public Sub GetListOfCountriesInAsia()

    Dim rstData     As New ADODB.Recordset
    Dim oDBInstance As New clsDBInstance
    Dim oDatabase   As clsDatabase
    Dim strRegion   As String
    
    Set oDatabase = oDBInstance.GetSharedDatabase()
    strRegion = "ASIA"
    oDatabase.AddToParamList "region", adChar, Len(strRegion), strRegion, adParamInput
    Set rstData = oDatabase.GetRecordsetFromStoredProc("spGetItems")
    rstData.MoveFirst
    Do While Not rstData.EOF
        Debug.Print rstData.Fields("country_name").Value & " " & rstData.Fields("country_region").Value
        rstData.MoveNext
    Loop
    
    oDBInstance.CloseSharedDatabase
    Set rstData = Nothing
    Set oDBInstance = Nothing
    Set oDatabase = Nothing

End Sub



Public Sub GetListOfCountriesInAsia()

    Dim rstData     As New ADODB.Recordset
    Dim oDBInstance As New clsDBInstance
    Dim oDatabase   As clsDatabase
    Dim strRegion   As String
    
    Set oDatabase = oDBInstance.GetSharedDatabase()
    strRegion = "ASIA"
    oDatabase.AddToParamList "region", adChar, Len(strRegion), strRegion, adParamInput
    Set rstData = oDatabase.GetRecordsetFromStoredProc("spGetItems")
    rstData.MoveFirst
    Do While Not rstData.EOF
        Debug.Print rstData.Fields("country_name").Value & " " & rstData.Fields("country_region").Value
        rstData.MoveNext
    Loop
    
    oDBInstance.CloseSharedDatabase
    Set rstData = Nothing
    Set oDBInstance = Nothing
    Set oDatabase = Nothing

End Sub



