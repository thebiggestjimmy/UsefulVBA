Function ReturnUniqueValues(inputRange As Variant, Optional bIgnoreCase As Boolean = False) As Variant

    Dim row As Variant
    Dim cell As Variant
    
    Dim iOutRows As Integer
    Dim iOutCol As Integer
        
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    
    If (bIgnoreCase) Then
        dict.CompareMode = TextCompare
    End If
    
    Dim Output() As Variant
    Dim rowArr() As Variant
    
    Dim rowString As String

    Dim irc As Integer
    Dim icc As Integer
    
    For irc = 1 To UBound(inputRange, 1)
       
       rowString = ""
    
       For icc = 1 To UBound(inputRange, 2)
           If (Not (IsError(inputRange(irc, icc)))) Then
               rowString = rowString + CStr(inputRange(irc, icc)) + ","
           End If
       Next
       
       If (Not (dict.Exists(rowString))) Then
           dict.Add rowString, ""
                     
           ReDim rowArr(UBound(inputRange, 2) - 1)
           For icc = 1 To UBound(inputRange, 2)
                rowArr(icc - 1) = inputRange(irc, icc)
           Next
           
           
           ReDim Preserve Output(iOutRows)
           Output(iOutRows) = rowArr
           iOutRows = iOutRows + 1
           
       End If
    Next
        
    ReturnUniqueValues = Output

End Function

Public Function SortRange(inputRange As Variant, iSortCol As Integer, bSortAscending As Boolean) As Variant

    If (TypeOf inputRange Is Range) Then
        inputRange = inputRange.Value
    End If

    Dim iRows, iCols, iCurrentRow, iCurrentCol, iSwap As Integer
    Dim bSwapped, bSwap As Boolean
    
    iRows = UBound(inputRange, 1)
    iCols = UBound(inputRange, 2)
    
    ReDim aSortOrder(iRows - 1, 0) As Integer
    For iCurrentRow = 1 To iRows
        aSortOrder(iCurrentRow - 1, 0) = iCurrentRow
    Next
    
    Dim vCurrent, vNext As Variant
    bSwapped = True
    
    Do Until bSwapped = False
    
        bSwapped = False
        For iCurrentRow = 1 To iRows - 1
        
            vCurrent = inputRange(aSortOrder(iCurrentRow - 1, 0), iSortCol)
            vNext = inputRange(aSortOrder(iCurrentRow, 0), iSortCol)
            
            If (IsError(vCurrent) <> IsError(vNext)) Then
                bSwap = IsError(vCurrent)
            Else
                bSwap = (vCurrent > vNext And bSortAscending = True) Or (vCurrent < vNext And bSortAscending = False)
            End If
            
            If (bSwap = True) Then
                iSwap = aSortOrder(iCurrentRow - 1, 0)
                aSortOrder(iCurrentRow - 1, 0) = aSortOrder(iCurrentRow, 0)
                aSortOrder(iCurrentRow, 0) = iSwap
            End If
            
            bSwapped = bSwapped Or bSwap
        Next
    Loop
    
    ReDim OutputRange(iRows - 1, iCols - 1) As Variant
    For iCurrentRow = 0 To iRows - 1
        For iCurrentCol = 0 To iCols - 1
            OutputRange(iCurrentRow, iCurrentCol) = inputRange(aSortOrder(iCurrentRow, 0), iCurrentCol + 1)
        Next
    Next
    
    SortRange = OutputRange

End Function


Function FilterInput(inputRange As Variant, iColumn As Integer, bMatch As Boolean, vValue As Variant, bSkipHeader As Boolean, Optional bIsWildcardMatch As Boolean = False) As Variant

    Dim row As Variant
    Dim cell As Variant
    
    Dim iOutRows As Integer
    Dim iOutCol As Integer
            
    Dim Output() As Variant
    Dim rowArr() As Variant
       
    Dim irc As Integer
    Dim icc As Integer
    
    Dim ircFirst As Integer
    Dim bIsMatch As Boolean
    
    If (TypeOf inputRange Is Range) Then
        inputRange = inputRange.Value
    End If
    
    If (bSkipHeader) Then ircFirst = 2 Else ircFirst = 1
    
    For irc = ircFirst To UBound(inputRange, 1)
       If (Not IsError(inputRange(irc, iColumn))) Then
            cell = inputRange(irc, iColumn)
                        
            If (bIsWildcardMatch) Then
                bIsMatch = (cell Like vValue)
            Else
                bIsMatch = (cell = vValue)
            End If
                        
            If (bIsMatch = bMatch) Then
                 ReDim rowArr(UBound(inputRange, 2) - 1)
                 For icc = 1 To UBound(inputRange, 2)
                      rowArr(icc - 1) = inputRange(irc, icc)
                 Next
                 
                 ReDim Preserve Output(iOutRows)
                 Output(iOutRows) = rowArr
                 iOutRows = iOutRows + 1
            End If
        End If
    Next
    
    If (iOutRows = 1) Then
        ReDim rowArr(UBound(inputRange, 2) - 1)
        For icc = 1 To UBound(inputRange, 2)
            rowArr(icc - 1) = CVErr(xlErrNA)
        Next
                 
        ReDim Preserve Output(iOutRows)
        Output(iOutRows) = rowArr
    End If
        
    FilterInput = Output

End Function
