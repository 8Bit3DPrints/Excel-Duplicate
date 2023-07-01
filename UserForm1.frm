VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Update QTY"
   ClientHeight    =   1875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4335
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private wireLabel As String
Private quantity As String

Private Sub TextBox1_Change()
    wireLabel = TextBox1.Text
End Sub

Private Sub TextBox2_Change()
    quantity = TextBox2.Text
End Sub

Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Check if the 'Enter' key was pressed
    If KeyCode = 13 Then
        ' Ensure both wireLabel and quantity are entered
        If wireLabel <> "" And quantity <> "" Then
            WireLabelQuantityChange
        End If
    End If
End Sub

Private Sub CommandButton1_Click()
    On Error GoTo ErrorHandler
    
    ' Confirmation dialog before clearing the data
    If MsgBox("Are you sure you want to clear the data?", vbQuestion + vbYesNo, "Confirmation") = vbNo Then
        Exit Sub
    End If
    
    ClearData
    
    Exit Sub
    
ErrorHandler:
    ThisWorkbook.LogError "An error occurred while clearing the data: " & Err.Description
    MsgBox "An error occurred while clearing the data. The error has been logged.", vbExclamation, "Error"
    Exit Sub
End Sub

Private Sub WireLabelQuantityChange()
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    Dim i As Long
    Dim skipRow As Boolean
    
    ' Get the last row with data in column A
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' If the last row is 1 and there's no data, we don't skip a row
    If lastRow = 1 And Cells(lastRow, "A").Value = "" Then
        skipRow = False
    ' If there's no data in the last row (and last row isn't 1), we skip a row
    Else
        skipRow = True
    End If

    For i = 1 To CInt(quantity)
        If skipRow Then
            Range("A" & lastRow + i + 1).Value = wireLabel
        Else
            Range("A" & lastRow + i).Value = wireLabel
        End If
    Next i

    Rows(lastRow + CInt(quantity) + IIf(skipRow, 2, 1)).Insert Shift:=xlDown

    TextBox1.Text = ""
    TextBox2.Text = ""
    TextBox1.SetFocus
    
    Exit Sub
    
ErrorHandler:
    ThisWorkbook.LogError "An error occurred while updating the wire label quantity: " & Err.Description
    MsgBox "An error occurred while updating the wire label quantity. The error has been logged.", vbExclamation, "Error"
    Exit Sub
End Sub

Private Sub CommandButton2_Click()
    On Error GoTo ErrorHandler
    
    ExportSheetData
    
    Exit Sub
    
ErrorHandler:
    ThisWorkbook.LogError "An error occurred while exporting the sheet data: " & Err.Description
    MsgBox "An error occurred while exporting the sheet data. The error has been logged.", vbExclamation, "Error"
    Exit Sub
End Sub

Private Sub ExportSheetData()
    On Error GoTo ErrorHandler
    
    Dim NewWorkbook As Workbook
    Dim SaveFileName As Variant
    Dim FileFilter As String
    
    ' Set the file filter
    FileFilter = "Excel Files (*.xlsx), *.xlsx"
    
    ' Prompt the user for the save file name and location, default to "Do Not Overwrite"
    SaveFileName = Application.GetSaveAsFilename(InitialFileName:="Choose a New Name!", FileFilter:=FileFilter, Title:="Save As")
    
    ' Check if the user canceled the Save As dialog box
    If SaveFileName = False Then Exit Sub
    
    ' Create a new workbook
    Set NewWorkbook = Workbooks.Add
    
    ' Copy the active sheet's data to the new workbook
    ThisWorkbook.ActiveSheet.UsedRange.Copy Destination:=NewWorkbook.Sheets(1).Cells(1, 1)
    
    ' Save the new workbook
    NewWorkbook.SaveAs SaveFileName
    
    ' Close the new workbook
    NewWorkbook.Close SaveChanges:=False
    
    Exit Sub
    
ErrorHandler:
    ThisWorkbook.LogError "An error occurred while exporting the sheet data: " & Err.Description
    MsgBox "An error occurred while exporting the sheet data. The error has been logged.", vbExclamation, "Error"
    Exit Sub
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    
    ' Set error handling for the entire UserForm
    On Error Resume Next
    
    ' Add any additional initialization code here
    
    Exit Sub
    
ErrorHandler:
    ThisWorkbook.LogError "An error occurred while initializing the UserForm: " & Err.Description
    MsgBox "An error occurred while initializing the UserForm. The error has been logged.", vbExclamation, "Error"
    Me.Hide
End Sub

Private Sub UserForm_Terminate()
    ' Add any cleanup code or finalization steps here
End Sub

Private Sub Worksheet_Activate()
    On Error GoTo ErrorHandler
    
    UserForm1.Show
    
    Exit Sub
    
ErrorHandler:
    ThisWorkbook.LogError "An error occurred while activating the worksheet: " & Err.Description
    MsgBox "An error occurred while activating the worksheet. The error has been logged.", vbExclamation, "Error"
    Exit Sub
End Sub

Private Sub ClearData()
    On Error GoTo ErrorHandler
    
    ActiveSheet.Cells.ClearContents ' Clear only the cell contents, not the formatting
    TextBox1.Text = ""
    TextBox2.Text = ""
    TextBox1.SetFocus
    
    Exit Sub
    
ErrorHandler:
    ThisWorkbook.LogError "An error occurred while clearing the data: " & Err.Description
    MsgBox "An error occurred while clearing the data. The error has been logged.", vbExclamation, "Error"
    Exit Sub
End Sub
