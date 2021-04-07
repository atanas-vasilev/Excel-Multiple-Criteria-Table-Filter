Sub filerTableByList()
    Dim criteriaList As Range
    Set criteriaList = Application.InputBox("Select a Criteria Range", "Obtain Range Object", Type:=8)

    Dim tableColRNG As Range
    Set tableColRNG = Application.InputBox("Select a Cell to Which to Apply the Criteria", "Obtain Range Object", Type:=8)
    
        'Get column number and table name.
        'TODO: Add Error Handling in case there is no table.
        Dim tableName As String
        tableName = tableColRNG.ListObject
        Dim colNum As Integer
        colNum = tableColRNG.Column


    Dim sArray() As String
    Dim I As Long
    Dim var1 As Variant
    var1 = criteriaList.Value
    
   ReDim sArray(1 To UBound(var1))
    
    For I = 1 To (UBound(var1))
      sArray(I) = var1(I, 1)
    Next
    
    ActiveSheet.ListObjects(tableName).Range.AutoFilter Field:=colNum, Criteria1:=sArray, Operator:=xlFilterValues
End Sub
