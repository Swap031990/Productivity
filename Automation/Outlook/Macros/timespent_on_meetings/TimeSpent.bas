Attribute VB_Name = "TimeSpent"
'=================================================================
'Description: Outlook macro which allows you to calculate the
'	      total amount of time spent on the selected Calendar,
'	      Task and Journal items.
'
'author : Robert Sparnaaij
'version: 1.0
'website: https://www.howto-outlook.com/howto/timespent.htm
'=================================================================

'Limitation; This code does not work for recurring meeting items
'in a List view (like All Appointments) since recurring items
'are only listed once.
'To work with recurring items and for full reporting features,
'you can use a reporting add-in;
'https://www.howto-outlook.com/tag/reporting

==================================================================

Public Sub TimeSpentReport()
    Dim objOL As Outlook.Application
    Dim objSelection As Outlook.Selection
    Dim objItem As Object
    Dim Duration As Long
    Dim TotalWork As Long
    Dim Mileage As Long
    Dim Result As Integer
    Dim ShowMileage As Boolean
    
    'Change to True if you also want to calculate
    'the total mileage in the report.
    ShowMileage = False
    
    Duration = 0
    TotalWork = 0
    Mileage = 0
    
    On Error Resume Next

    Set objOL = Outlook.Application
    Set objSelection = objOL.ActiveExplorer.Selection

    For Each objItem In objSelection
        If objItem.Class = olAppointment Then
            Duration = Duration + objItem.Duration
            Mileage = Mileage + objItem.Mileage
        ElseIf objItem.Class = olTask Then
            Duration = Duration + objItem.ActualWork
            TotalWork = TotalWork + objItem.TotalWork
            Mileage = Mileage + objItem.Mileage
        ElseIf objItem.Class = Outlook.olJournal Then
            Duration = Duration + objItem.Duration
            Mileage = Mileage + objItem.Mileage
        Else
            Result = MsgBox("No Calendar, Task or Journal item selected.", vbCritical, "Time Spent")
            Exit Sub
        End If
    Next
    
    'Building the message box text
    Dim MsgBoxText As String
    MsgBoxText = "Total time spent on the selected items; " & vbNewLine & Duration & " minutes"
    
    If Duration > 60 Then
        MsgBoxText = MsgBoxText & HoursMinsMsg(Duration)
    End If
    
    If TotalWork > 0 Then
        MsgBoxText = MsgBoxText & vbNewLine & vbNewLine & "Total work recorded for the selected Tasks; " & vbNewLine & TotalWork & " minutes"
        
        If TotalWork > 60 Then
            MsgBoxText = MsgBoxText & HoursMinsMsg(TotalWork)
        End If
    End If
    
    If ShowMileage = True Then
        MsgBoxText = MsgBoxText & vbNewLine & vbNewLine & "Total mileage; " & Mileage
    End If

    Result = MsgBox(MsgBoxText, vbInformation, "Time spent")

ExitSub:
    Set objItem = Nothing
    Set objSelection = Nothing
    Set objOL = Nothing
End Sub

Public Function HoursMinsMsg(TotalMinutes As Long) As String
        Dim Hours As Long
        Dim Minutes As Long
        Hours = TotalMinutes \ 60
        Minutes = TotalMinutes Mod 60
        HoursMinsMsg = " (" & Hours & " hours and " & Minutes & " minutes)"
End Function