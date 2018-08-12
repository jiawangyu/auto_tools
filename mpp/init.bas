Sub AddNewColum()
    TableEditEx Name:="项(&E)", TaskTable:=True, NewFieldName:="Text1", Title:="风险", ColumnPosition:=8
    TableEditEx Name:="项(&E)", TaskTable:=True, NewFieldName:="Text2", Title:="交付件", ColumnPosition:=9
    TableApply Name:="项(&E)"
End Sub