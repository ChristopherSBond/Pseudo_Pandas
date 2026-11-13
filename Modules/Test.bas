Attribute VB_Name = "Test"
'@Folder("VBAProject")
Option Explicit

Private m_bigDf As clsDataFrame

Public Sub testSmallDataset()
    Dim df As clsDataFrame: Set df = New clsDataFrame
'    df.fromRange Sheet1.Range("D8:F27"), headers:=False
    df.FromRange Sheet1.Range("D7:F27"), True
    df.ToRange Sheet1.Range("H7")
    Dim dividingText As String: dividingText = String(200, "*")
'    df.head
'    df.tail
    Debug.Print dividingText
    Debug.Print "df.columns"
    df.Columns
    
    Debug.Print dividingText
    Debug.Print df.Shape(0) & " " & df.Shape(1)
    
    Debug.Print dividingText
    Debug.Print "df.selectCols(""A"")"
    df.SelectCols("A").Head
    
    Debug.Print dividingText
    Debug.Print "df.selectRows(""1"", ""2,3"", ""4:5"")"
    df.SelectRows("1", "2,3", "4:5").Head
    
    Debug.Print dividingText
    Debug.Print "df.dropDuplicates(""A"")"
    df.DropDuplicates("A").Head
    
    Debug.Print dividingText
    Debug.Print "df.dropColumns(""A"")"
    df.DropColumns("A").Head
    
    Debug.Print dividingText
    Debug.Print "df.dropNull"
    df.DropNull.Head
    
    Debug.Print dividingText
    Debug.Print "df.filter("" 'A' == 'B' "")"
    df.Filter(" 'A' == 'B' ").Head
    
    Debug.Print dividingText
    Debug.Print "df.mutate("" 'D' = 4*5 "")"
    df.Mutate("'D' = 4*5 ").Head
        
    Debug.Print dividingText
    Debug.Print "df.renameColumns(Array(""A"", ""B""), Array(""E"", ""F""))"
    df.RenameColumns(Array("A", "B"), Array("E", "F")).Head
    
    Debug.Print dividingText
    Debug.Print " df.sort(""A"") "
    df.Sort("A").Head
    
    Debug.Print dividingText
    Debug.Print " df.unique ""A"" "
    df.Unique "A"
    
    Debug.Print dividingText
    Debug.Print " df.filter ("" 'A' == abs('C') "")"
    df.Filter(" 'A' == abs('C') ").Head

    df.Filter(" 'A' == abs(round('C'*3, 2)) ").Head

    df.ListOperationsForTextExpressions

End Sub

Public Sub testErrors()
    Dim df As clsDataFrame: Set df = New clsDataFrame
    df.FromRange Sheet1.Range("D7:F27"), True

'//// PANDAS_ERROR_COL_HAS_FORBIDDEN_CHARS
'    df.RenameColumns("A", "A""").Head
    
'//// PANDAS_ERROR_EVALUATE_SYNTAX
     ' mismatched "("
'    df.Filter(" 'A' == 3*(2").Head
    
    ' arithmetic between string and numeric
'    df.Filter(" 'A' == 3*""Hello"" ").Head
    
    ' unknown function
'    df.Filter(" 'A' == func('B')").Head
    
'//// PANDAS_ERROR_MISMATCHED_NUMBER_COLS
    ' These two should NOT fail
'    df.RenameColumns(Array("A", "C"), Array("E", "F")).Head
'    df.RenameColumns("C", "E").Head
    
    ' These two SHOULD fail
'    df.RenameColumns("C", Array("E", "F")).Head
'    df.RenameColumns(Array("A", "C"), "E").Head

'//// PANDAS_ERROR_COL_EXISTS_TWICE
    ' For single column renaming
'    df.RenameColumns("B", "A").Head
    
    ' For multi-column renaming
'    df.RenameColumns(Array("A", "B"), Array("D", "D")).Head

'//// PANDAS_ERROR_SELECTED_ROW_OUT_OF_BOUNDS
'    df.SelectRows("1000000").Head
'    df.SelectRows("0:10000000").Head

'//// PANDAS_ERROR_ROW_NOT_NUM
'    df.SelectRows("Hi").Head
'    df.SelectRows("0:A").Head
'    df.SelectRows("0,1,A").Head

    ' Should NOT error
'    df.SelectRows("").Head


'//// PANDAS_ERROR_SELECTED_COL_OUT_OF_BOUNDS
'    df.SelectCols("1000").Head
'    df.SelectCols("0,1,2,10000").Head
'    df.SelectCols("0:1000").Head

'//// PANDAS_ERROR_COL_NAME_NOT_EXISTS
'    df.SelectCols("EF").Head

'//// PANDAS_ERROR_AGG_FUNC_NOT_EXISTS
'    df.GroupByAgg("A", "B", "maes").Head
    
    ' To test, change the name of 'A' to 'A_max' and 'B' to 'A'
'    df.GroupByAgg("A_max", "A", "max").Head

End Sub

Public Sub test135124()
    Dim df As clsDataFrame: Set df = New clsDataFrame
    df.FromRange Sheet1.Range("D7:F27"), True
    
    df.DropNull.Mutate(" 'B' = 2*'C' ").Head
'
'    df.ListOperationsForTextExpressions
'    df.filter (" 'A' <> Null ")
'    df.groupByAgg("A", "B", Array("mean", "sum", "count", "max", "min")).head
'    df. _
'        GroupByAgg("A", "B", Array("mean", "sum", "count")) _
'        .RenameColumns(Array("B_mean", "B_sum", "B_count"), Array("Mean", "Sum", "Count")) _
'        .Mutate(" 'Mean' = round('Mean',2) ").Head
'
    Dim df2 As clsDataFrame: Set df2 = New clsDataFrame
'    df.dropNull(Array("A", "B")).head


'    df.groupByAgg("A", "B", "sum").head
    
End Sub


Public Sub testBigDataset()
    Dim df As clsDataFrame: Set df = New clsDataFrame
    If m_bigDf Is Nothing Then
        Dim bigDataFileName As String: bigDataFileName = "County_Health_Rankings.csv"
        Dim fso As FileSystemObject: Set fso = New FileSystemObject
        Dim bigDataFileLocation As String: bigDataFileLocation = fso.BuildPath(fso.GetFolder(ThisWorkbook.Path), bigDataFileName)
        df.FromDelimited bigDataFileLocation
        Set m_bigDf = df
    Else
        Set df = m_bigDf
    End If

    Dim startTime As Double
    startTime = Timer
'    df.filter(" 'State' == ""MO"" ").head
'    Dim df2 As clsDataFrame
'    df.head
'    df.filter(" 'col1' == abs('col2')").head
    'df.head
    Dim df2 As clsDataFrame: Set df2 = New clsDataFrame
    Set df2 = df.GroupByAgg(Array("State", "County"), "State", "count").Head

    Debug.Print Timer - startTime


End Sub

Public Sub test252352()
    Dim df As clsDataFrame: Set df = New clsDataFrame
'    df.FromArray Sheet1.Range("H7:J27").Value2
    df.FromArray Sheet1.Range("H7:J27")
    df.Head
End Sub

Public Sub test42425()
    Dim res As Variant
    res = Widen(Sheet1.Range("S7:W46").Value2, Sheet1.Range("S6:w6"), Array(1, 3, 5), 2, 4)
 
    Sheet1.Range("Y6").Resize(UBound(res, 1), UBound(res, 2)).Value2 = res
 
End Sub

' data: 2D 1-based array (e.g. Range.Value, without the header row)
' headers: 1-row 2D array from Range.Value (headers(1, i))
' idCols: Array(1, 3) -> Vehicle and Lane positions; hdrCol: 2; valCol: 4
Public Function Widen(data As Variant, headers As Variant, idCols As Variant, _
                      hdrCol As Long, valCol As Long) As Variant
    Dim nIds As Long: nIds = UBound(idCols) - LBound(idCols) + 1
    Dim nR As Long: nR = UBound(data, 1)

    Dim rowKeys As Object, colKeys As Object
    Set rowKeys = CreateObject("Scripting.Dictionary")
    Set colKeys = CreateObject("Scripting.Dictionary")

    Dim rowOf() As Long, colOf() As Long
    ReDim rowOf(1 To nR): ReDim colOf(1 To nR)
    Dim idParts() As Variant
    ReDim idParts(1 To nR, 1 To nIds) ' upper bound; only first rowKeys.Count rows used

    Dim r As Long, j As Long, key As String, k As String
    For r = 1 To nR
        ' composite row key
        key = ""
        For j = 0 To nIds - 1
            key = key & CStr(data(r, idCols(LBound(idCols) + j))) & Chr(0)
        Next j
        If Not rowKeys.Exists(key) Then
            rowKeys.Add key, rowKeys.Count + 1
            For j = 0 To nIds - 1
                idParts(rowKeys.Count, j + 1) = data(r, idCols(LBound(idCols) + j))
            Next j
        End If
        rowOf(r) = rowKeys(key)

        ' header key
        k = CStr(data(r, hdrCol))
        If Not colKeys.Exists(k) Then colKeys.Add k, colKeys.Count + 1
        colOf(r) = colKeys(k)
    Next r

    Dim out() As Variant, filled() As Boolean
    ReDim out(1 To rowKeys.Count + 1, 1 To nIds + colKeys.Count)
    ReDim filled(1 To rowKeys.Count, 1 To colKeys.Count)

    ' header row
    For j = 0 To nIds - 1
        out(1, j + 1) = headers(1, idCols(LBound(idCols) + j))
    Next j
    Dim kv As Variant
    For Each kv In colKeys.keys
        out(1, nIds + colKeys(kv)) = kv
    Next kv

    ' id columns
    Dim i As Long
    For i = 1 To rowKeys.Count
        For j = 1 To nIds
            out(i + 1, j) = idParts(i, j)
        Next j
    Next i

    ' place values
    For r = 1 To nR
        If filled(rowOf(r), colOf(r)) Then
            Err.Raise vbObjectError + 3, , "Duplicate entry at data row " & r
        End If
        out(rowOf(r) + 1, nIds + colOf(r)) = data(r, valCol)
        filled(rowOf(r), colOf(r)) = True
    Next r

    Widen = out
End Function
