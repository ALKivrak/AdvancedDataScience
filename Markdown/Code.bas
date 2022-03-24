Attribute VB_Name = "Module1"
Sub Code()
'blue to keep columns'
'orange to not include text in the cell'
'yellow to turn hyperlink into normal text'


Dim lCol As Long
Dim rng As Range
Dim icntr As Long
Dim shape As Excel.shape
Dim FileName, sLine, Deliminator, url As String
Dim LastCol, LastRow, FileNumber As Integer
Dim lnk As Hyperlink
FileName = "MarkdownTable.md"
Deliminator = "|"

lCol = 17

For icntr = lCol To 1 Step -1
If Not Cells(1, icntr).Interior.ColorIndex = 37 Then

Columns(icntr).Delete
End If

Next

For Each rng In ActiveSheet.UsedRange
If rng.Interior.ColorIndex = 45 Then
rng.ClearContents
rng.Interior.Color = xlColorIndexNone
End If

If rng.Interior.ColorIndex = 6 Then
rng.Hyperlinks.Delete
End If

Next rng

For Each shape In ActiveSheet.Shapes
shape.Delete
Next shape

LastCol = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column
LastRow = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row - 5
FileNumber = FreeFile

Open FileName For Output As FileNumber

For i = 1 To LastRow
For j = 1 To LastCol


 
If j = LastCol Then

    If (Cells(i, j).Hyperlinks.Count <> 1) Then
        url = ""
    Else
        url = Cells(i, j).Hyperlinks(1).Address
    End If
    
    If url = "" Then
        sLine = sLine & Cells(i, j).Value
        
    Else
        sLine = sLine & "[" & Cells(i, j).Value & "]" & "(" & url & ")"
        url = ""
    End If
Else
    If (Cells(i, j).Hyperlinks.Count <> 1) Then
        url = ""
    Else
        url = Cells(i, j).Hyperlinks(1).Address
    End If
    
    If url = "" Then
        sLine = sLine & Cells(i, j).Value & Deliminator
    Else
        sLine = sLine & "[" & Cells(i, j).Value & "]" & "(" & url & ")" & Deliminator
        url = ""
    End If
End If

If i = 2 And j = 2 Then
sLine = "|:---- |:----:| ----:|"
End If

Next j

Print #FileNumber, sLine
sLine = ""

Next i

Close #FileNumber
MsgBox "Markdown table has been generated"

End Sub
