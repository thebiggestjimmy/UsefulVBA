VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private dictStartTimes As Dictionary
Private dictAverageTimes As Dictionary

Private Sub Class_Initialize()
    Set dictStartTimes = New Dictionary
    Set dictAverageTimes = New Dictionary
End Sub

Public Function LogStart(sName As String)
    dictStartTimes(sName) = Timer
End Function

Public Function LogEnd(sName As String)
    Dim vCompletionTime
    
    If (dictStartTimes.Exists(sName)) Then
        
        vCompletionTime = (Timer - dictStartTimes(sName))
        
        If (Not dictAverageTimes.Exists(sName)) Then
            dictAverageTimes(sName) = vCompletionTime
        Else
            dictAverageTimes(sName) = (dictAverageTimes(sName) + vCompletionTime) / 2
        End If
    Else
        MsgBox "Timer " & sName & " not started"
    End If
End Function

Public Function GetTime(sName As String) As Date
    GetTime = dictAverageTimes(sName)
End Function

Public Function GetDetails() As String
    Dim sSeconds As Double
    Dim sDetails As String
    
    For i = 0 To dictAverageTimes.Count - 1
        sSeconds = dictAverageTimes.Items(i)
        sDetails = sDetails & dictAverageTimes.Keys(i) & "=" & Format(sSeconds, "0.00") & vbNewLine
    Next i
    
    GetDetails = sDetails
    
End Function
