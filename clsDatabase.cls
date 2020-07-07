Option Explicit

Private Const CLNG_DEFAULT_TIMEOUT          As Long = 0
Private Const CINT_PARAMETER_FIELDS_COUNT   As Integer = 5

Private moConn                  As New ADODB.Connection
Private mstrConnectionString    As String
Private mstrADOErrors           As String
Private mlngCmdTimeOut          As Long
Private mvarParams()            As Variant

Private Enum EnuParameters
    Name = 0
    DataType = 1
    Length = 2
    Value = 3
    Direction = 4
End Enum

Public Property Get ConnectionErrors() As String
    ConnectionErrors = mstrADOErrors
End Property

Private Sub Class_Initialize()
    mstrADOErrors = Empty
    mlngCmdTimeOut = CLNG_DEFAULT_TIMEOUT
    ReDim mvarParams(CINT_PARAMETER_FIELDS_COUNT, 0)
End Sub

Public Function OpenDatabase(ByVal objCred As clsDBCredentials) As Boolean

On Error GoTo ERR_HANDLER:
    OpenDatabase = False
    If (moConn Is Nothing) Or (moConn.State = adStateClosed) Then
        mstrConnectionString = objCred.GetConnectionString
        moConn.Open mstrConnectionString
    End If
    OpenDatabase = True
    
EXIT_HERE:
    Exit Function
    
ERR_HANDLER:
    OpenDatabase = False
    mstrADOErrors = GetADOErrorInformation()
    mstrADOErrors = mstrADOErrors & vbCrLf & "Err Description:" & Err.Description
    Set moConn = Nothing
    GoTo EXIT_HERE
End Function

Public Function CloseDatabase() As Boolean

On Error GoTo ERR_HANDLER:
    CloseDatabase = False
    If (Not moConn Is Nothing) Or (moConn.State = adStateOpen) Then
        moConn.Close
        Set moConn = Nothing
    End If
    CloseDatabase = True
    
EXIT_HERE:
    Exit Function
    
ERR_HANDLER:
    CloseDatabase = False
    GoTo EXIT_HERE
    
End Function

Public Sub ClearParamList()
    ReDim mvarParams(CINT_PARAMETER_FIELDS_COUNT, 0)
End Sub

Public Sub AddToParamList(ByVal strParamName As String, _
                          ByVal dblDataType As Double, _
                          ByVal dblDataLength As Double, _
                          ByVal varParamValue As Variant, _
                          ByVal intParamDirection As Integer)

    Dim intArrayElement         As Integer
    Dim intArrayBound           As Integer

    intArrayBound = UBound(mvarParams(), 2)
    intArrayElement = intArrayBound
    intArrayBound = intArrayBound + 1
    
    ReDim Preserve mvarParams(CINT_PARAMETER_FIELDS_COUNT, intArrayBound)
    mvarParams(EnuParameters.Name, intArrayElement) = strParamName
    mvarParams(EnuParameters.DataType, intArrayElement) = dblDataType
    mvarParams(EnuParameters.Length, intArrayElement) = dblDataLength
    mvarParams(EnuParameters.Value, intArrayElement) = varParamValue
    mvarParams(EnuParameters.Direction, intArrayElement) = intParamDirection
  
End Sub

Public Function GetParamValue(ByVal strParamName As String) As Variant
    Dim intIndex As Integer
    For intIndex = LBound(mvarParams, 2) To UBound(mvarParams, 2) - 1
        If mvarParams(EnuParameters.Name, intIndex) = strParamName Then
            GetParamValue = mvarParams(EnuParameters.Value, intIndex)
            Exit Function
        End If
    Next intIndex
End Function

Private Function GetADOErrorInformation() As String
    Dim lngErrorCount   As Long
    Dim lngErrorIndex   As Long
    Dim oError          As ADODB.Error
    Dim oErrorColl      As ADODB.Errors
    Dim strErr          As String
    
    If moConn Is Nothing Then
        GoTo EXIT_HERE
    Else
        Set oErrorColl = moConn.Errors
    End If

    lngErrorCount = oErrorColl.Count
    If (lngErrorCount > 0) Then
        strErr = "Errors reported by ADO" & vbCrLf
    End If
    For lngErrorIndex = 0 To (lngErrorCount - 1)
        Set oError = oErrorColl.Item(lngErrorIndex)
        With oError
            strErr = strErr & "(" & lngErrorIndex + 1 & ") "
            strErr = strErr & "Error#: " & .Number & vbCrLf
            strErr = strErr & vbTab & "Desc : " & .Description & vbCrLf
            strErr = strErr & vbTab & "Source: " & .Source & vbCrLf
            strErr = strErr & vbTab & "Native Error: " & .NativeError & vbCrLf
            strErr = strErr & vbTab & "SQL State: " & .SqlState & vbCrLf
            strErr = strErr & vbTab & "Help Context: " & .HelpContext & vbCrLf
            strErr = strErr & vbTab & "Help File: " & .HelpFile & vbCrLf
        End With
    Next lngErrorIndex
    GetADOErrorInformation = strErr
EXIT_HERE:
    Set oError = Nothing
    Set oErrorColl = Nothing
End Function

Public Function IsConnected() As Boolean
    IsConnected = False
    If Not moConn Is Nothing Then
        IsConnected = IIf(moConn.State = adStateOpen, True, False)
    End If
End Function

Public Function ExecuteStoredProc(ByVal strQueryName As String) As Boolean

On Error GoTo ERR_HANDLER

    Dim intParam        As Integer
    Dim prmParameter    As ADODB.Parameter
    Dim objCommand      As ADODB.Command
    
    Set objCommand = New ADODB.Command
    ExecuteStoredProc = False
    mstrADOErrors = Empty
    
    If Not IsConnected Then
        GoTo EXIT_HERE
    End If
    
    With objCommand
        .ActiveConnection = moConn
        .CommandTimeout = mlngCmdTimeOut
        .CommandType = adCmdStoredProc
        .CommandText = "[" & strQueryName & "]"
         For intParam = LBound(mvarParams, 2) To UBound(mvarParams, 2) - 1
            Set prmParameter = objCommand.CreateParameter(mvarParams(EnuParameters.Name, intParam), mvarParams(EnuParameters.DataType, intParam), mvarParams(EnuParameters.Direction, intParam), mvarParams(EnuParameters.Length, intParam), mvarParams(EnuParameters.Value, intParam))
            objCommand.Parameters.Append prmParameter
         Next intParam
        .Execute
    End With
    
    For intParam = LBound(mvarParams, 2) To UBound(mvarParams, 2) - 1
        mvarParams(EnuParameters.Value, intParam) = objCommand.Parameters(intParam).Value
    Next intParam
    ExecuteStoredProc = True

EXIT_HERE:
    Set objCommand = Nothing
    Set prmParameter = Nothing
    Exit Function
ERR_HANDLER:
    mstrADOErrors = GetADOErrorInformation()
    mstrADOErrors = mstrADOErrors & vbCrLf & "Err Description:" & Err.Description
    ExecuteStoredProc = False
    GoTo EXIT_HERE
End Function

Public Function GetRecordsetFromStoredProc(ByVal strQueryName As String, _
                                           Optional ByVal intCursorType = adOpenKeyset, _
                                           Optional ByVal intLockType = adLockOptimistic, _
                                           Optional ByVal intCursorLocation = adUseClient) As ADODB.Recordset

On Error GoTo ERR_HANDLER

    Dim intParam        As Integer
    Dim prmParameter    As ADODB.Parameter
    Dim objCommand      As ADODB.Command
    Dim rstOutput       As ADODB.Recordset
    Set objCommand = New ADODB.Command
    Set rstOutput = New ADODB.Recordset
    
    Set GetRecordsetFromStoredProc = Nothing
    mstrADOErrors = Empty
    
    If Not IsConnected Then
        GoTo EXIT_HERE
    End If

    With objCommand
        .ActiveConnection = moConn
        .CommandTimeout = mlngCmdTimeOut
        .CommandType = adCmdStoredProc
        .CommandText = strQueryName
         For intParam = LBound(mvarParams, 2) To UBound(mvarParams, 2) - 1
            Set prmParameter = objCommand.CreateParameter(mvarParams(EnuParameters.Name, intParam), mvarParams(EnuParameters.DataType, intParam), mvarParams(EnuParameters.Direction, intParam), mvarParams(EnuParameters.Length, intParam), mvarParams(EnuParameters.Value, intParam))
            objCommand.Parameters.Append prmParameter
         Next intParam
        objCommand.ActiveConnection.CursorLocation = intCursorLocation
        rstOutput.CursorType = intCursorType
        rstOutput.LockType = intLockType
        Set rstOutput = .Execute
    End With
    
    For intParam = LBound(mvarParams, 2) To UBound(mvarParams, 2) - 1
        mvarParams(EnuParameters.Value, intParam) = objCommand.Parameters(intParam).Value
    Next intParam
    Set GetRecordsetFromStoredProc = rstOutput
    
EXIT_HERE:
    Set objCommand = Nothing
    Set prmParameter = Nothing
    Set rstOutput = Nothing
    Exit Function
ERR_HANDLER:
    mstrADOErrors = GetADOErrorInformation()
    mstrADOErrors = mstrADOErrors & vbCrLf & "Err Description:" & Err.Description
    Set GetRecordsetFromStoredProc = Nothing
    GoTo EXIT_HERE
End Function


Public Function GetDataFromSQLStatement(ByVal strSQL As String) As ADODB.Recordset

On Error GoTo Error_Handler

	Dim oRst As New ADODB.Recordset
	Set GetDataFromSQLStatement = Nothing
	If Not IsConnected Then
		GoTo Exit_Here
	End If
	oRst.Open strSQL, moConn, adOpenKeyset, adLockOptimistic
	Set GetDataFromSQLStatement = oRst

EXIT_HERE:
	Set oRst = Nothing
	Exit Function
ERROR_HANDLER:
	mstrADOErrors = GetADOErrorInformation()
	mstrADOErrors = mstrADOErrors & vbCrLf & “Err Description:” & Err.Description
	GoTo EXIT_HERE
End Function

