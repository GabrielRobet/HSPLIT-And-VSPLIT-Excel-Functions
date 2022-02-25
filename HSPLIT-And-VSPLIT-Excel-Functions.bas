Function HSPLIT(MyText As String, MyDelimiter As String)
    
    HSPLIT = Split(MyText, MyDelimiter)
    
End Function


Function VSPLIT(MyText As String, MyDelimiter As String)
    
    VSPLIT = Application.WorksheetFunction.Transpose(Split(MyText, MyDelimiter))
    
End Function
