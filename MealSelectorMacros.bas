Attribute VB_Name = "MealSelectorMacros"
Sub UpdateInventory_and_SendEmail_Generic()

Dim og As Worksheet
Dim ns As Worksheet
Dim count_col As Integer
Dim count_row As Integer


Set og = ThisWorkbook.Sheets("Meal Selector")
Set ns = ThisWorkbook.Sheets("Inventory")

ns.Cells.ClearContents
og.Activate

count_col = WorksheetFunction.CountA(Range("D2", Range("D2").End(xlToRight)))
count_row = 6

For i = 1 To count_col + 2
    For j = 1 To count_row + 1
    ns.Cells(i, j) = og.Cells(j, i).Text
    
    Next j
Next i

'For i = 1 To count_col
    'ns.Cells(i, 1) = ns.Cells(i, count_col)
'Next i

ns.Activate

Range("G5", Range("G5").End(xlDown)).Select
Selection.Cut
Range("A5").Select
ActiveSheet.Paste
Range("E1:Q100").Select
Selection.ClearContents

' Add option to send email if answer yes to prompt
Answer = MsgBox("Do you also want to send an email?", vbYesNo)

If Answer = vbYes Then
    Dim OutApp As Object
    Dim OutMail As Object
    Dim SourceWorksheet As Worksheet
    Dim TempWorkbook As Workbook
    Dim TempFilePath As String

    ' Set the active worksheet
    Set SourceWorksheet = ActiveSheet
        
    ' Copy the active sheet to a new workbook
    SourceWorksheet.Copy
    Set TempWorkbook = ActiveWorkbook
        
    ' Save the new workbook to a temporary file
    TempFilePath = Environ$("temp") & "\" & SourceWorksheet.Name & Format(Date, "MMddyyyy") & ".xlsx"
    TempWorkbook.SaveAs Filename:=TempFilePath, FileFormat:=xlOpenXMLWorkbook


    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
    .To = ""
    .CC = ""
    .BCC = ""
    .Subject = "Inventory for " & Format(Date, "MM/dd/yyyy")
    .Body = "Here is the inventory for " & Format(Date, "MM/dd/yyyy")
    .Attachments.Add TempFilePath
    .Display
    End With
    
    'Clean up
    TempWorkbook.Close SaveChanges:=False
    Kill TempFilePath ' deletes the temporary file
    Set OutApp = Nothing
    Set OutMail = Nothing
    Set TempWorkbook = Nothing
Else
Exit Sub
End If

End Sub

Sub SendEmail_Generic()

' Add option to send email if answer yes to prompt
Answer = MsgBox("Do you want to send an email of the inventory?", vbYesNo)

If Answer = vbYes Then
    Dim OutApp As Object
    Dim OutMail As Object
    Dim SourceWorksheet As Worksheet
    Dim TempWorkbook As Workbook
    Dim TempFilePath As String

    ' Set the active worksheet
    Set SourceWorksheet = ActiveSheet
        
    ' Copy the active sheet to a new workbook
    SourceWorksheet.Copy
    Set TempWorkbook = ActiveWorkbook
        
    ' Save the new workbook to a temporary file
    TempFilePath = Environ$("temp") & "\" & SourceWorksheet.Name & Format(Date, "MMddyyyy") & ".xlsx"
    TempWorkbook.SaveAs Filename:=TempFilePath, FileFormat:=xlOpenXMLWorkbook


    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
    .To = ""
    .CC = ""
    .BCC = ""
    .Subject = "Inventory for " & Format(Date, "MM/dd/yyyy")
    .Body = "Here is the inventory for " & Format(Date, "MM/dd/yyyy")
    .Attachments.Add TempFilePath
    .Display
    End With
    
    'Clean up
    TempWorkbook.Close SaveChanges:=False
    Kill TempFilePath ' deletes the temporary file
    Set OutApp = Nothing
    Set OutMail = Nothing
    Set TempWorkbook = Nothing
Else
Exit Sub
End If

End Sub

Sub ClearSelection()

Dim count_col As Integer
Dim count_row As Integer

count_col = WorksheetFunction.CountA(Range("E7", Range("E7").End(xlToRight))) - 2
count_row = WorksheetFunction.CountA(Range("C8", Range("C8").End(xlDown)))

Answer = MsgBox("Are you sure you want to reset the meal selector?", vbYesNo)

If Answer = vbYes Then
Range(Cells(8, 5), Cells(count_row + 8, count_col + 4)).ClearContents
Else
Exit Sub
End If
End Sub


Sub DeleteMeals()

Dim count_col As Integer

count_col = WorksheetFunction.CountA(Range("E7", Range("E7").End(xlToRight))) - 1

Answer = MsgBox("This action deletes any meals with 0 quantity. Are you sure you want to do this?", vbYesNo)

If Answer = vbYes Then

    For i = 1 To count_col
     If Cells(4, 4 + i) = 0 Then
        Columns(4 + i).EntireColumn.Delete
        i = i - 1
        count_col = WorksheetFunction.CountA(Range("E7", Range("E7").End(xlToRight))) - 1
    Else: If Cells(4, 4 + i) = "Stop" Then Exit For
    End If
    Next i
    
Else
Exit Sub
End If

End Sub

Sub UpdateQuantities()

Worksheets("Meal Selector").Range("E4", Range("E4").End(xlToRight)).Select
Selection.Copy
Range("E2").Select
Worksheets("Meal Selector").Range("E2", Range("E2").End(xlToRight)).PasteSpecial xlPasteValues
Application.CutCopyMode = False

End Sub

Sub AddMeal()

mealName = Application.InputBox(prompt:="Meal", Type:=2)
num = Application.InputBox(prompt:="Number", Type:=1)

If (mealName = "" Or num <= 0) Then
    MsgBox ("Please enter a meal name and quantity > 0.")
    Exit Sub
Else
    ' InsertColumn
    Range("F:F").EntireColumn.Insert Shift:=xlToRight
    '
    ' Copy over formulas
    Range("E3:E6").Select
    Selection.Copy
    Range("F3").Select
    Selection.PasteSpecial xlFormulas
    Application.CutCopyMode = False
    
    Cells(2, 6) = num
    Cells(7, 6) = mealName
End If

End Sub
