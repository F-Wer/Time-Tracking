Attribute VB_Name = "Modul1"
Dim TimerActive As Boolean
Dim Total As Date
Dim Latestleave As Date
Dim Clockin As Date

Sub StartTimer()

 TimerActive = True
    Start_Timer
    Worktime_Check
End Sub
Sub stopTImer()
Call Stop_Timer
End Sub
Private Sub Stop_Timer()
    TimerActive = False
End Sub

Private Sub Start_Timer()
    Range("F6:G7").ClearContents
    Range("H3") = InputBox("Wann bist du gekommen?")
    TimerActive = True
    Application.OnTime Now() + TimeValue("00:00:01"), "Timerino"
End Sub
Private Sub Timerino()
    If TimerActive Then
        ActiveSheet.Cells(1, 7).Value = Time
        Application.OnTime Now() + TimeValue("00:00:01"), "Timerino"
    End If
  End Sub
Private Sub Worktime_Check()

    If Range("N2").Value >= 7.8 Then
    MsgBox "Idiot! Nimm n Dispo!"

End If
Latestleave = CDate(Range("H15").Value)
    Clockin = CDate(Range("G1").Value)
    Total = TimeValue(Latestleave) - TimeValue(Clockin)
 If Total <= TimeValue("00:10:00") Then
MsgBox "GEH!"
 End If

End Sub





