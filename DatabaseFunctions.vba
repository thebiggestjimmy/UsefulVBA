Function ProcessDBRange(vData As Variant) As Variant

    Dim Output() As Variant
    Dim rowArr() As Variant
    
    Dim iOutRows As Integer
    
    Dim irc As Integer
    Dim icc As Integer
    
    For irc = 0 To UBound(vData, 2)
              
        ReDim rowArr(UBound(vData, 1))
        For icc = 0 To UBound(vData, 1)
            If (IsNull(vData(icc, irc))) Then
                rowArr(icc) = ""
            Else
                rowArr(icc) = vData(icc, irc)
            End If
        Next
        
        ReDim Preserve Output(iOutRows)
        Output(iOutRows) = rowArr
        iOutRows = iOutRows + 1
    Next
    
    ProcessDBRange = Output
    
End Function


Function CallProcAndReturnValues(valueDate As Date, qrUnderlying As String) As Variant
    
    Dim sDate As String
    sDate = Format(valueDate, "dd MMM yyyy")
    
    Dim cnn As ADODB.Connection, cmd As ADODB.Command, rs As ADODB.Recordset
    Set cnn = New ADODB.Connection
    
    cnn.connectionString = [rngConnectionString]
    cnn.Open
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cnn
    
    cmd.CommandText = "EXEC ..."
    Set rs = cmd.Execute()
    
    Dim vData As Variant
    vData = rs.GetRows
       
    GetImpliedForwardTenors = ProcessDBRange(vData)
    
On Error GoTo errh
        
    
    
    Exit Function
    
errh:
    MsgBox Err.Description
    
    
End Function