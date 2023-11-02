Sub Combine()
    Dim wb1 As Workbook
    Dim wb2 As Workbook
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim outputPath As String
    Dim file1Path As String
    Dim file2Path As String
    Dim combinedWb As Workbook

    ' Prompt for the paths to the two .xls files
    file1Path = Application.GetOpenFilename("Excel Files (*.xls), *.xls")
    file2Path = Application.GetOpenFilename("Excel Files (*.xls), *.xls")

    ' Check if both files are selected
    If file1Path = "False" Or file2Path = "False" Then
        MsgBox "You need to select both files."
        Exit Sub
    End If

    ' Prompt for the output path and file name
    outputPath = Application.GetSaveAsFilename("Output File", "Excel Files (*.xlsx), *.xlsx")

    ' Check if an output path is selected
    If outputPath = "False" Then
        MsgBox "You need to select an output path."
        Exit Sub
    End If

    ' Open the first workbook
    Set wb1 = Workbooks.Open(file1Path)
    ' Open the second workbook
    Set wb2 = Workbooks.Open(file2Path)

    ' Set references to the first and second worksheets
    Set ws1 = wb1.Sheets(1)
    Set ws2 = wb2.Sheets(1)

    ' Rename the sheets
    ws1.Name = "MeasureFA"
    ws2.Name = "LQCTuning"

    ' Combine the two sheets into a new workbook
    ws1.Copy
    Set ws1 = ActiveSheet
    wb2.Sheets(1).Copy After:=ws1

    ' Save the combined workbook with the given output path as .xlsx
    Set combinedWb = ActiveWorkbook
    combinedWb.SaveAs outputPath, FileFormat:=51 ' File format 51 represents .xlsx

    ' Close the workbooks without saving changes
    wb1.Close SaveChanges:=False
    wb2.Close SaveChanges:=False

    ' Close the combined workbook
    combinedWb.Close SaveChanges:=False
End Sub
