Attribute VB_Name = "Modul3"
Sub Auto_Close()


    Application.DisplayAlerts = False

        ActiveWorkbook.Close

        Application.DisplayAlerts = True
    End Sub
Sub sbWriteCellWhenClosing()

    ActiveSheet.Range("N4") = ActiveSheet.Range("N2")
    ActiveWorkbook.Save

End Sub
