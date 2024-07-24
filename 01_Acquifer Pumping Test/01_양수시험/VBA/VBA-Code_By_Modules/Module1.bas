Attribute VB_Name = "Module1"
Sub MergeNextColumn()
Attribute MergeNextColumn.VB_ProcData.VB_Invoke_Func = " \n14"


    Range(ActiveCell, ActiveCell.Offset(0, 1)).Select
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
End Sub

Sub MergeNextColumn2()
Attribute MergeNextColumn2.VB_ProcData.VB_Invoke_Func = "d\n14"

' 바로 가기 키: Ctrl+d
'

    Range(ActiveCell, ActiveCell.Offset(0, 2)).Select
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
End Sub

