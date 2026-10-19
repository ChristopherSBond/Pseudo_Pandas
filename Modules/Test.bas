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


