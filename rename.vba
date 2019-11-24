Sub fileRename()

    Dim J As Long
    Dim xFileDlg As FileDialog
    Dim xFileDlgItem As Variant
    On Error Resume Next
    Application.ScreenUpdating = False
    Set xFileDlg = Application.FileDialog(msoFileDialogFolderPicker)
    If xFileDlg.Show = -1 Then
        xFileDlgItem = xFileDlg.SelectedItems.Item(1)
        ActiveSheet.Cells(1, "A") = xFileDlgItem
    End If
    Let r = ActiveSheet.UsedRange.Rows.Count - 1
    For J = 2 To r + 1
        ActiveSheet.Cells(J, "C") = ActiveSheet.Cells(1, "A").Value & "\" & ActiveSheet.Cells(J, "A").Value
    Next
    Application.ScreenUpdating = True
    
    Dim xNumLeft, xNumRight As Long
    Dim xOldName, xNewName As String
    On Error Resume Next
    xAddress = ActiveWindow.RangeSelection.Address
    Application.ScreenUpdating = False
    For J = 2 To r + 1
        xOldName = ActiveSheet.Cells(J, "C").Value
        xNumLeft = InStrRev(xOldName, "\")
        xNumRight = InStrRev(xOldName, ".")
        xNewName = ActiveSheet.Cells(J, "B").Value
        If xNewName <> "" Then
            xNewName = Left(xOldName, xNumLeft) & xNewName & Mid(xOldName, xNumRight)
            Name xOldName As xNewName
        End If
    Next
    Application.ScreenUpdating = True

End Sub
