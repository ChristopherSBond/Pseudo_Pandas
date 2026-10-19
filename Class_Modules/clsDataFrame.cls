VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule RedundantByRefModifier
'@Folder("VBAProject")
Option Explicit
''*********************************************************************************
'' ***   This class module contains all functionality for Pseudo_Pandas.        ***
'' ***   GitHub Link: https://github.com/ChristopherSBond/Pseudo_Pandas         ***
''*********************************************************************************

' Set to false if you want to see what line code errors on
' Otherwise, all errors will go through error handling routines
Private Const debugMode As Boolean = True

' Pandas specific values
Private Const PANDAS_VALUES_ERROR As Variant = Null
Private Const PANDAS_VALUES_FORBIDDEN_COL_CHARS As String = "'\""" ' characters that cannot exist in column names

Private dict_df As Scripting.Dictionary
Private dict_col_types As Scripting.Dictionary

Private m_fnNames As Scripting.Dictionary
Private m_opNames As Scripting.Dictionary
Private m_aggFnNames As Scripting.Dictionary

Private Enum errors
    PANDAS_ERROR_SELECTED_ROW_OUT_OF_BOUNDS = vbObjectError + 1
    PANDAS_ERROR_SELECTED_COL_OUT_OF_BOUNDS
    PANDAS_ERROR_COL_NAME_NOT_EXISTS
    PANDAS_ERROR_EVALUATE_SYNTAX
    PANDAS_ERROR_COL_HAS_FORBIDDEN_CHARS
    PANDAS_ERROR_COL_EXISTS_TWICE
    PANDAS_ERROR_COL_TYPES_DONT_MATCH
    PANDAS_ERROR_INVALID_ARITH_TYPES
    PANDAS_ERROR_SELECTED_ROW_INVALID
    PANDAS_ERROR_MISMATCHED_NUMBER_COLS
    PANDAS_ERROR_ROW_NOT_NUM
    PANDAS_ERROR_AGG_FUNC_NOT_EXIST
    PANDAS_ERROR_FILE_NOT_FOUND
End Enum

Private Enum rpnOp
    OP_NONE = 0
    OP_ADD
    OP_SUB
    OP_MUL
    OP_DIV
    OP_LT
    OP_LE
    OP_EQ
    OP_GE
    OP_GT
    OP_NE
    OP_AND
    OP_OR
    OP_POW
    OP_MOD
End Enum

Private Enum rpnFn
    FN_NONE = 0
    FN_ABS
    FN_SQRT
    FN_LEN
    FN_ROUND
End Enum

Private Enum tokenTypes
    TK_NULL = 0
    TK_NULLSTR
    TK_STR
    TK_NUM
    TK_OP
    TK_PAR
    TK_FN
End Enum

Private Enum tokenAssoc
    ASS_L = 0
    ASS_R
End Enum

Private Enum aggOp
    AGG_UNK = 0
    agg_sum
    AGG_MEAN
    AGG_MIN
    AGG_MAX
    AGG_COUNT
End Enum

'************************************************************************************************************************
' fromDict
'   Purpose:
'       Gets data from a dictionary and stores in the current df object
'   Parameters:
'                 dict: dictionary to create dataframe from
'       nullNormalized: flag if data has already been normalized (excel errors and blanks = NULL). True if yes.
'                       Normalizing is an expensive operation, so this cuts down the time when creating dataframes that are already normalized
'             colTypes: dictionary containing the data types for each column
'                       Determining types is an expensive operation, so this cuts down the time when created dataframes where these are already known
'   Note:
'       dict_df is a module-wide variable that stores all of the actual data used in the dataframe
'       the creation of a dataframe the process of storing the values of dictionary in dict_df
'---------------------------------------------------------------------------------
Public Sub FromDict(ByVal dict As Scripting.Dictionary, Optional ByVal nullNormalized As Boolean = False, Optional ByVal colTypes As Scripting.Dictionary)
    Set dict_col_types = New Scripting.Dictionary

    Dim headerName As Variant
    For Each headerName In dict.keys
        If Not isValidColumnName(headerName) Then pandasRaiseError "fromDict", PANDAS_ERROR_COL_HAS_FORBIDDEN_CHARS, headerName
        
        If colTypes Is Nothing Then dict_col_types.Add headerName, getArrayDataType(dict(headerName))
        
        If Not nullNormalized Then dict(headerName) = getArrWithNormalizedMissing(dict(headerName))
    
    Next headerName
    
    If Not colTypes Is Nothing Then Set dict_col_types = colTypes
    
    Set dict_df = dict
End Sub

'************************************************************************************************************************
' fromRange
'   Purpose:
'       Gets data from a range and stores in the current df object
'   Parameters:
'           rng: range storing all the data
'       headers: flag; True if the range includes headers
'   Note:
'       dict_df is a module-wide variable that stores all of the actual data used in the dataframe
'       The creation of a dataframe is the process of storing the values of the range in dict_df
'       Dataframes do NOT store the reference of the ranges, but the actual values of the range
'---------------------------------------------------------------------------------
Public Sub FromRange(ByVal rng As Range, Optional ByVal headers As Boolean = True)
    Set dict_df = New Scripting.Dictionary
    Set dict_col_types = New Scripting.Dictionary
    
    Dim colNum As Long
    For colNum = 1 To rng.Columns.count
        Dim headerName As String
        Dim arrColData As Variant
        
        If headers = True Then
            headerName = rng.Cells(1, colNum).Value2
            
            If Not isValidColumnName(headerName) Then pandasRaiseError "fromRange", PANDAS_ERROR_COL_HAS_FORBIDDEN_CHARS, headerName
            
            arrColData = getArr1dFromRng(rng.Cells(1, colNum).Offset(1).Resize(rng.rows.count - 1))
        Else
            headerName = "Col_" & colNum
            arrColData = getArr1dFromRng(rng.Cells(1, colNum).Resize(rng.rows.count))
        End If
            
        If Not debugMode And dict_df.Exists(headerName) Then pandasRaiseError "fromRange", PANDAS_ERROR_COL_EXISTS_TWICE, headerName
        
        dict_df.Add headerName, getArrWithNormalizedMissing(arrColData)
        dict_col_types.Add headerName, getArrayDataType(arrColData)
        
    Next colNum

End Sub

Public Sub FromDelimited(ByVal filePath As String, Optional ByVal headers As Boolean = True, Optional ByVal delimiter As String = ",")
    Dim arr2dCSV As Variant: arr2dCSV = getArr2dFromDelimited(filePath)
    FromDict getDictFromArr2d(arr2dCSV, headers)
End Sub

Public Sub ToDelimited(ByVal filePath As String, Optional ByVal delimiter As String = ",")
    writeArrToDelimited filePath, getArrExcel2DFromDict(dict_df), delimiter
End Sub

Public Sub ToRange(ByVal rngStartCell As Range)
    Dim arr2d As Variant: arr2d = getArrExcel2DFromDict(dict_df)
    rngStartCell.Resize(Me.Shape(0) + 1, Me.Shape(1)).Value2 = arr2d
End Sub

Public Function ToArray(Optional ByVal dropHeaders As Boolean = False) As Variant
    Dim arr2d As Variant: arr2d = getArrExcel2DFromDict(dict_df, dropHeaders)
        
    ToArray = arr2d
End Function

Public Function ToDict() As Scripting.Dictionary
    Set ToDict = dict_df
End Function

Public Function ListOperationsForTextExpressions() As Variant
    Dim availableFunctions As Variant: availableFunctions = functionRegistry().keys
    
    Debug.Print "Available functions: "; Join(availableFunctions, ", ")
    
    Dim availableOperations As Variant: availableOperations = operatorRegistry().keys
    Debug.Print "Available operations: "; Join(availableOperations, ", ")
    
    Dim availableAggFunctions As Variant: availableAggFunctions = aggFunctionRegistry().keys
    Debug.Print "Available aggregate functions: "; Join(availableAggFunctions, ", ")
    
End Function

'************************************************************************************************************************
' shape
'   Purpose:
'       Gets the shape of the data frame as an array
'   Returns:
'       Array(numberOfRows, numberOfColumns) of the dataframe
'---------------------------------------------------------------------------------
Public Function Shape() As Variant
    Dim numRows As Long
    If isEmptyArray(dict_df(dict_df.keys(0))) Then
        numRows = 0
    Else
        numRows = UBound(dict_df(dict_df.keys(0))) - LBound(dict_df(dict_df.keys(0))) + 1
    End If
    
    Dim numCols As Long
    If isEmptyArray(dict_df.keys) Then
        numCols = 0
    Else
        numCols = UBound(dict_df.keys) - LBound(dict_df.keys) + 1
    End If
    Shape = Array(numRows, numCols)
End Function

'************************************************************************************************************************
' columns
'   Purpose:
'       Gets the column names of the dataframe
'   Returns:
'       Array(colName1, colName2, ... colNamen) of the dataframe
'---------------------------------------------------------------------------------
Public Function Columns() As Variant
    Columns = dict_df.keys
    Debug.Print Join(dict_df.keys, ", ")
End Function

'************************************************************************************************************************
' dropColumns
'   Purpose:
'       Drops the inputted column names from the dataframe
'   Parameter:
'       columnsToDrop: array (or string if dropping one) containing the column name(s) to drop
'   Returns:
'       Dataframe with the dropped columns
'---------------------------------------------------------------------------------
Public Function DropColumns(ByVal columnsToDrop As Variant) As clsDataFrame
    Dim newDict As Scripting.Dictionary: Set newDict = getDictCopy(dict_df)
    
    If isEmptyArray(columnsToDrop) Then
        If dict_df.Exists(columnsToDrop) Then newDict.Remove columnsToDrop
    
    Else
        Dim idx As Long
        For idx = LBound(columnsToDrop) To UBound(columnsToDrop)
            If dict_df.Exists(columnsToDrop(idx)) Then newDict.Remove columnsToDrop(idx)
        Next idx
    End If
    
    Dim newDf As clsDataFrame: Set newDf = New clsDataFrame
    newDf.FromDict newDict
    Set DropColumns = newDf
End Function

'************************************************************************************************************************
' renameColumns
'   Purpose:
'       Renames the inputted columns to the new column names
'   Parameters:
'       originalColNames: array (or string) of the original column name(s)
'            newColNames: array (or string) of the new column name(s) ordered the same as the original name(s)
'   Example:
'       df.renameColumns(Array("A", "B", "C"), Array("D", "E", "F")) will rename column "A" to "D", "B" to "E", and "C" to "F"
'       df.renameColumns("A","B") will rename column "A" to "B"
'   Returns:
'       A dataframe with the changed column names
'---------------------------------------------------------------------------------
Public Function RenameColumns(ByVal originalColNames As Variant, ByVal newColNames As Variant) As clsDataFrame
    Dim newDict As Scripting.Dictionary: Set newDict = getDictCopy(dict_df)
    
    If Not ((isEmptyArray(originalColNames) And isEmptyArray(newColNames)) Or Not (isEmptyArray(originalColNames) Or isEmptyArray(newColNames))) Then
        pandasRaiseError "RenameColumns", PANDAS_ERROR_MISMATCHED_NUMBER_COLS
    End If
    
    If Not isEmptyArray(originalColNames) Then
        Dim numCols As Long: numCols = UBound(originalColNames) - LBound(originalColNames) + 1
        
        Dim lbNew As Long: lbNew = LBound(newColNames)
        Dim lbOrig As Long: lbOrig = LBound(originalColNames)
    
        If numCols <> (UBound(newColNames) - LBound(newColNames) + 1) Then pandasRaiseError "RenameColumns", PANDAS_ERROR_MISMATCHED_NUMBER_COLS
        
        Dim idx As Long
        For idx = 0 To numCols - 1
            If Not dict_df.Exists(originalColNames(idx + lbOrig)) Then pandasRaiseError "renameColumns", PANDAS_ERROR_COL_NAME_NOT_EXISTS, originalColNames(idx + lbOrig)
            
            If dict_df.Exists(newColNames(idx + lbNew)) Or newDict.Exists(newColNames(idx + lbNew)) Then
                pandasRaiseError "renameColumns", PANDAS_ERROR_COL_EXISTS_TWICE, "'" & newColNames(idx + lbNew) & "' already exists in dataframe"
            End If
            
            newDict.key(originalColNames(idx + lbOrig)) = newColNames(idx + lbNew)
        Next idx
    
    Else
        If Not dict_df.Exists(originalColNames) Then pandasRaiseError "renameColumns", PANDAS_ERROR_COL_NAME_NOT_EXISTS, originalColNames
        
        If dict_df.Exists(newColNames) Then pandasRaiseError "renameColumns", PANDAS_ERROR_COL_EXISTS_TWICE, "'" & newColNames & "' already exists in dataframe"

        newDict.key(originalColNames) = newColNames
    End If
    
    Dim newDf As clsDataFrame: Set newDf = New clsDataFrame
    newDf.FromDict newDict
    
    Set RenameColumns = newDf
End Function

'************************************************************************************************************************
' selectCols
'   Purpose:
'       Gets a data frame with only the selected columns
'   Parameter:
'       colIdxs: parameter array with the column indexes or column names
'                colIdxs can be input as a single column per index of paramArray, a comma-separated string of colidxs,
'                   or as a range of column indexes or column names in the form "firstCol:lastCol" inclusive
'   Returns:
'       clsDataFrame with only the selected columns
'---------------------------------------------------------------------------------
Public Function SelectCols(ParamArray colIdxs() As Variant) As clsDataFrame
    ' Stores all of the selected column indexes. Each column index = value in the array
    Dim arrSelectedColIdxs As Variant
    
    ' Parse each of the inputs with the following options in addition to selecting the column numbers entered directly as paras
    '   ":" means to select a continuous range of columns
    '   "," means to select each single column
    Dim colIdx As Variant
    For Each colIdx In colIdxs
        If InStr(1, colIdx, ":") Then
            Dim firstCol As Variant: firstCol = split(colIdx, ":")(0)
            Dim lastCol As Variant: lastCol = split(colIdx, ":")(1)
            
            If Not IsNumeric(firstCol) Then
                If Not debugMode And dict_df.Exists(firstCol) Then pandasRaiseError "selectCols", PANDAS_ERROR_COL_NAME_NOT_EXISTS, firstCol
                
                If Not debugMode And dict_df.Exists(lastCol) Then pandasRaiseError "selectCols", PANDAS_ERROR_COL_NAME_NOT_EXISTS, lastCol
                
                ' Get the index of the header instead of the column name
                firstCol = getIndexOfDictKey(dict_df, firstCol)
                lastCol = getIndexOfDictKey(dict_df, lastCol)
            End If
                              
            ' Get an array with each number in a range (e.g., getArrRngBtNums(0,5) returns Array(0,1,2,3,4,5))
            Dim arrRngBtNums As Variant: arrRngBtNums = getArrRngBtNums(CLng(firstCol), CLng(lastCol))
            
            arrSelectedColIdxs = getMergedArr(arrSelectedColIdxs, arrRngBtNums)
        
        ElseIf InStr(1, colIdx, ",") Then
            arrSelectedColIdxs = getMergedArr(arrSelectedColIdxs, split(colIdx, ","))
        
        ElseIf colIdx <> vbNullString Then
            arrSelectedColIdxs = getMergedArr(arrSelectedColIdxs, Array(colIdx))
        
        End If
    Next colIdx
    
    ' Dict later used to create the df
    Dim dict As Scripting.Dictionary: Set dict = New Scripting.Dictionary
    
    Dim col As Variant
    For Each col In arrSelectedColIdxs
        If IsNumeric(col) Then
            ' Check if the row idx is valid for the dataframe
            If Not inValRange(col, 0, Me.Shape(1) - 1) Then pandasRaiseError "selectCols", PANDAS_ERROR_SELECTED_COL_OUT_OF_BOUNDS, col
            
            ' Initialize column name
            dict(dict_df.keys(col)) = dict_df(dict_df.keys(col))
             
        Else
            If Not dict_df.Exists(col) Then pandasRaiseError "selectCols", PANDAS_ERROR_COL_NAME_NOT_EXISTS, col
        
            dict(col) = dict_df(col)
        
        End If
    Next col
    
    ' Create the new df and return it
    Dim dfNew As clsDataFrame: Set dfNew = New clsDataFrame
    dfNew.FromDict dict
    
    Set SelectCols = dfNew

End Function

'************************************************************************************************************************
' selectRows
'   Purpose:
'       Gets a data frame with only the selected rows
'   Parameter:
'       rowIdxs: parameter array with the row indexes
'                rowIdxs can be input as a single row per index of paramArray, a comma-separated string of rowIdx,
'                   or as a range of row indexes in the form "firstRow:lastRow" inclusive
'   Returns:
'       clsDataFrame with only the selected rows
'---------------------------------------------------------------------------------
Public Function SelectRows(ParamArray rowIdxs() As Variant) As clsDataFrame
    Dim newDict As Scripting.Dictionary: Set newDict = New Scripting.Dictionary

    Dim countIdxs As Long
    Dim rowIdx As Variant
    For Each rowIdx In rowIdxs
        rowIdx = Replace(rowIdx, " ", vbNullString)
        If InStr(1, rowIdx, ",") Then
            If hasNonNumericChar(rowIdx, allowComma:=True) Then pandasRaiseError "SelectRows", PANDAS_ERROR_ROW_NOT_NUM, rowIdx
            
            countIdxs = countIdxs + UBound(split(rowIdx, ",")) + 1
        ElseIf InStr(1, rowIdx, ":") Then
            If hasNonNumericChar(rowIdx, allowColon:=True) Then pandasRaiseError "SelectRows", PANDAS_ERROR_ROW_NOT_NUM, rowIdx
            
            Dim firstNum As Long: firstNum = split(rowIdx, ":")(0)
            Dim lastNum As Long: lastNum = split(rowIdx, ":")(1)
            
            If firstNum > lastNum Then pandasRaiseError "selectRows", PANDAS_ERROR_SELECTED_ROW_INVALID, firstNum & ":" & lastNum
            countIdxs = countIdxs + CLng(lastNum) - CLng(firstNum) + 1
        Else
            If hasNonNumericChar(rowIdx) Then pandasRaiseError "SelectRows", PANDAS_ERROR_ROW_NOT_NUM, rowIdx
            If rowIdx <> vbNullString Then countIdxs = countIdxs + 1
        End If
    Next rowIdx

    If countIdxs = 0 Then
        Dim k As Variant
        For Each k In dict_df.keys: newDict(k) = vbNullString: Next k
        GoTo endFunct
    End If

    ' arr to store all rowIdxs
    Dim arrSelectedRowIdxs() As Long: ReDim arrSelectedRowIdxs(0 To countIdxs - 1)
    
    Dim minIdx As Long: minIdx = -1
    Dim maxIdx As Long: maxIdx = -1
    Dim tempIdx As Long
    Dim tempArr As Variant
    Dim tempVal As Long
    Dim selectedIdx As Long
    For Each rowIdx In rowIdxs
        If InStr(1, rowIdx, ",") Then
            tempArr = split(rowIdx, ",")
            For tempIdx = LBound(tempArr) To UBound(tempArr)
                tempVal = CLng(tempArr(tempIdx))
                arrSelectedRowIdxs(selectedIdx) = tempVal
                selectedIdx = selectedIdx + 1
                
                If minIdx = -1 Or minIdx > tempVal Then minIdx = tempVal
                If maxIdx = -1 Or maxIdx < tempVal Then maxIdx = tempVal
            Next tempIdx

        ElseIf InStr(1, rowIdx, ":") Then
            For tempIdx = split(rowIdx, ":")(0) To split(rowIdx, ":")(1)
                arrSelectedRowIdxs(selectedIdx) = tempIdx
                selectedIdx = selectedIdx + 1
                
                If minIdx = -1 Or minIdx > tempIdx Then minIdx = tempIdx
                If maxIdx = -1 Or maxIdx < tempIdx Then maxIdx = tempIdx
            Next tempIdx
        Else
            arrSelectedRowIdxs(selectedIdx) = CLng(rowIdx)
            selectedIdx = selectedIdx + 1
            If minIdx = -1 Or minIdx > rowIdx Then minIdx = rowIdx
            If maxIdx = -1 Or maxIdx < rowIdx Then maxIdx = rowIdx
        End If

    Next rowIdx

    Dim arrDictKeys As Variant: arrDictKeys = dict_df.keys
    
    If minIdx < 0 Then pandasRaiseError "selectRows", PANDAS_ERROR_SELECTED_ROW_OUT_OF_BOUNDS, minIdx
    If maxIdx > Me.Shape(0) - 1 Then pandasRaiseError "selectRows", PANDAS_ERROR_SELECTED_ROW_OUT_OF_BOUNDS, maxIdx
    
    Dim colIdx As Long
    For colIdx = LBound(arrDictKeys) To UBound(arrDictKeys)
        Dim arrCol As Variant: arrCol = dict_df(arrDictKeys(colIdx))
        Dim newArrCol As Variant: ReDim newArrCol(0 To countIdxs - 1)
        Dim idx As Long
        
        For idx = LBound(arrSelectedRowIdxs) To UBound(arrSelectedRowIdxs)
            newArrCol(idx) = arrCol(arrSelectedRowIdxs(idx))
        Next idx
        
        newDict(arrDictKeys(colIdx)) = newArrCol
    Next colIdx

endFunct:
    'Create the new df and return it
    Dim dfNew As clsDataFrame: Set dfNew = New clsDataFrame
    dfNew.FromDict newDict, nullNormalized:=True, colTypes:=dict_col_types
    Set SelectRows = dfNew

End Function

Private Function hasNonNumericChar(ByVal s As String, Optional ByVal allowComma As Boolean = False, Optional ByVal allowColon As Boolean = False) As Boolean
    Dim i As Long, ch As Integer
    
    For i = 1 To Len(s)
        ch = asc(mid$(s, i, 1))
        
        ' Allow: digits (48 57) and comma (44)
        If Not ((ch >= 48 And ch <= 57) Or IIf(allowComma, ch = 44, False) Or IIf(allowColon, ch = 58, False)) Then
            hasNonNumericChar = True
            Exit Function
        End If
    Next i
End Function
'************************************************************************************************************************
' selectVals
'   Purpose:
'       Gets a dataframe of only the provided indexes. Similar to Python .loc
'   Parameter:
'       rows: Array (or single string) of row indexes.
'             Passed directly to selectRows, see that procedure for more information on the expected syntax
'       cols: Array (or single string) of col indexes.
'             Passed directly to selectCols, see that procedure for more information on the expected syntax
'             Differs from Python loc as both indexes and names can be used
'   Returns:
'       clsDataFrame with only the selected rows and columns
'---------------------------------------------------------------------------------
Public Function SelectVals(Optional ByVal rows As Variant, Optional ByVal cols As Variant) As clsDataFrame
Attribute SelectVals.VB_UserMemId = 0
    Dim rowIdxs As String
    
    If IsMissing(rows) Then
        rowIdxs = "0:" & Me.Shape(0) - 1
    Else
        If IsArray(rows) Then
            rowIdxs = Join(rows, ",")
        Else
            If rows = ":" Or rows = vbNullString Then
                rowIdxs = "0:" & Me.Shape(0) - 1
            Else
                rowIdxs = rows
            End If
        End If
    End If
    
    Dim colIdxs As String
    If IsMissing(cols) Then
        colIdxs = "0:" & Me.Shape(1) - 1
    Else
        If IsArray(cols) Then
            colIdxs = Join(cols, ",")
        Else
            If cols = ":" Then
                colIdxs = "0:" & Me.Shape(1) - 1
            Else
                colIdxs = cols
            End If
        End If
    End If
    
    Set SelectVals = Me.SelectCols(colIdxs).SelectRows(rowIdxs)

End Function

'************************************************************************************************************************
' head
'   Purpose:
'       Gets and debug.prints a dataframe only including the first numRows rows
'   Parameter:
'       numRows: number of rows to include in the returned/printed dataframe
'   Returns:
'       clsDataFrame with only the first numRows rows
'---------------------------------------------------------------------------------
Public Function Head(Optional ByVal numRows As Long = 5, Optional ByVal colSize As Long = 20) As clsDataFrame
    Dim dict As Scripting.Dictionary: Set dict = New Scripting.Dictionary
    
    Dim actualNumRows As Long: actualNumRows = WorksheetFunction.Min(numRows, Me.Shape(0))
    
    ' Create dictionary for new df, limiting the number of rows to the first n for head
    Dim col As Variant
    For Each col In dict_df
        Dim tempArr As Variant: tempArr = dict_df(col)
                
        ' We still want the columns to be printed, even if aren't any rows.
        ' We will skip "truncating" the array if there isn't anything to trunctate, but still treat it as an array
        If actualNumRows > 0 Then ReDim Preserve tempArr(0 To actualNumRows - 1)
        
        Set dict(col) = tempArr
    Next col
        
    print_dict dict, colSize:=colSize, dictColTypes:=dict_col_types
    Debug.Print "(Rows: " & Me.Shape(0); ", Columns:" & Me.Shape(1) & ")"
    
    Dim dfNew As clsDataFrame: Set dfNew = New clsDataFrame
    dfNew.FromDict dict
    Set Head = dfNew
End Function

'************************************************************************************************************************
' tail
'   Purpose:
'       Gets and debug.prints a dataframe only including the last numRows rows
'   Parameter:
'       numRows: number of rows to include in the returned/printed dataframe
'   Returns:
'       clsDataFrame with only the last numRows rows
'---------------------------------------------------------------------------------
Public Function Tail(Optional ByVal numRows As Long = 5, Optional ByVal colSize As Long = 20) As clsDataFrame
    Dim dict As Scripting.Dictionary: Set dict = New Scripting.Dictionary
    
    Dim actualNumRows As Long: actualNumRows = WorksheetFunction.Min(numRows, Me.Shape(0))
    
    ' Create dictionary for new df, limiting the number of rows to the last n for tail
    Dim col As Variant
    For Each col In dict_df
        If actualNumRows = 0 Then
            dict(col) = vbNullString
            GoTo nextCol
        End If
        
        Dim tempArr As Variant: ReDim tempArr(0 To actualNumRows - 1)
        
        Dim ub As Long: ub = UBound(dict_df(col))
        
        Dim rowIdx As Long: Dim idxTemp As Long: idxTemp = 0
        For rowIdx = ub - actualNumRows + 1 To ub
            tempArr(idxTemp) = dict_df(col)(rowIdx)
            idxTemp = idxTemp + 1
        Next rowIdx
        
        dict(col) = tempArr

nextCol:
    Next col
    
    print_dict dict, colSize:=colSize, dictColTypes:=dict_col_types
    Debug.Print "(Rows: " & Me.Shape(0); ", Columns:" & Me.Shape(1) & ")"

    Dim dfNew As clsDataFrame: Set dfNew = New clsDataFrame
    dfNew.FromDict dict, nullNormalized:=True
    
    Set Tail = dfNew
End Function

Public Function DropNull(Optional ByVal arrSubset As Variant) As clsDataFrame
    Dim arrStrFilterNull As Variant
    Dim idx As Long
      
    ' If subset parameter was not entered, then drop null from all columns
    If isEmptyArray(arrSubset) Then
        ReDim arrStrFilterNull(0 To dict_df.count - 1)
        Dim col As Variant
        For Each col In dict_df.keys
            arrStrFilterNull(idx) = "'" & col & "'" & "<> Null"
            idx = idx + 1
        Next col
    
    Else
        Dim lbSub As Long: lbSub = LBound(arrSubset)
        Dim ubSub As Long: ubSub = UBound(arrSubset)
        Dim numCols As Long: numCols = ubSub - lbSub + 1
        ReDim arrStrFilterNull(0 To numCols - 1)
                
        Dim idxSub As Long
        For idxSub = lbSub To ubSub
            Dim colName As String: colName = arrSubset(idxSub)
            arrStrFilterNull(idx) = "'" & colName & "'" & "<> Null"
            idx = idx + 1
        Next idxSub
    End If
        
    Dim strFilterNull As String: strFilterNull = Join(arrStrFilterNull, ") & (")

    strFilterNull = "(" & strFilterNull & ")"
    
    Set DropNull = Me.Filter(strFilterNull)
End Function

Public Function DropDuplicates(ByVal colName As String) As clsDataFrame
    If Not dict_df.Exists(colName) Then pandasRaiseError "dropDuplicates", PANDAS_ERROR_COL_NAME_NOT_EXISTS, colName
    
    Dim arrCol As Variant: arrCol = dict_df(colName)
    
    Dim dictUniqueRows As Scripting.Dictionary: Set dictUniqueRows = New Scripting.Dictionary
    Dim arrUniqueIdxs As Variant: ReDim arrUniqueIdxs(LBound(arrCol) To UBound(arrCol))
        
    Dim rowIdx As Long, countUnique As Long
    For rowIdx = LBound(arrCol) To UBound(arrCol)
        If dictUniqueRows.Exists(arrCol(rowIdx)) Then GoTo nextRow
            
        dictUniqueRows(arrCol(rowIdx)) = vbNullString
        arrUniqueIdxs(countUnique) = rowIdx
        countUnique = countUnique + 1
nextRow:
    Next rowIdx
    
    ReDim Preserve arrUniqueIdxs(0 To countUnique - 1)
    
    Set DropDuplicates = SelectRows(Join(arrUniqueIdxs, ","))
End Function

Public Function Sort(ByVal colName As String, Optional ByVal ascending As Boolean = True) As clsDataFrame
    Dim dict As Scripting.Dictionary: Set dict = New Scripting.Dictionary
    Dim col As Variant
    For Each col In dict_df
        dict(col) = dict_df(col)
    Next col
    
    sortDictByColumn dict, colName, ascending:=ascending
    
    Dim newDf As clsDataFrame: Set newDf = New clsDataFrame
    
    newDf.FromDict dict, nullNormalized:=True
    
    Set Sort = newDf
    
End Function


Public Function Unique(ByVal colName As String) As Variant
    If Not dict_df.Exists(colName) Then pandasRaiseError "unique", PANDAS_ERROR_COL_NAME_NOT_EXISTS, colName
       
    Dim dict As Scripting.Dictionary: Set dict = New Scripting.Dictionary
    
    Dim arrRowData As Variant: arrRowData = dict_df(colName)
    
    Dim uniqueVals As Variant: ReDim uniqueVals(0 To UBound(arrRowData))

    Dim row As Long, uniqueCount As Long
    For row = 0 To UBound(arrRowData)
        If IsNull(arrRowData(row)) Then GoTo nextRow
    
        If Not dict.Exists(arrRowData(row)) Then
            dict(arrRowData(row)) = vbNullString
            uniqueVals(uniqueCount) = arrRowData(row)
            uniqueCount = uniqueCount + 1
        End If
nextRow:
    Next row
    ReDim Preserve uniqueVals(0 To uniqueCount - 1)
    
    Debug.Print Join(uniqueVals, ", ")
    Unique = uniqueVals
End Function

'************************************************************************************************************************
' filter
'   Purpose:
'       Filters a data frame based on the input filterStr
'   Parameter:
'       filterStr: a string containing the filter conditions
'                  The string can contain expressions (e.g., " 'col1' == 3 " or " 'col1' == 'col2' " or " 'col1'^3-5 =< 'col2'^3-5/2)
'                  The string will be evaluated for each row (i.e., " 'col1' == 3 " will filter the data frame to only rows where df.selectVals(rowNum, "col3") equals 3)
'                  "And" conditions can be inputted using "&" and "Or" conditions inputted using "|"
'                  Very similar to R filter function and syntax
'   Examples:
'       df.filter(" 'col1' == 'col3' ") will return all rows where df.selectVals(rowNum, col1) equals df.selectVals(rowNum, col3)
'       df.filter(" 'col1' == 3 | 'col2' <5 ) will return all rows where OR(df.selectVals(rowNum, Col1) equals 3, df.selectVals(rowNum, Col2) <5)
'       df.filter(" ('col1' <= 'col2') & Col3 <> ""No"" ") will return all rows where AND(df.selectVals(rowNum, Col1) equals df.selectVals(rowNum, Col2), df.selectVals(rowNum, Col3) not equal to the literal "No")
'   Notes:
'       The following symbols can be used as operators:
'           + - * / < <= > >=
'       PEMDAS will be followed (parentheses are also permitted)
'       String values can be check for by encasing them in (") ("" needed if typing a literal in a string)
'           E.g. df.filter(" 'col1' == ""Hello World"" ") will return all rows where (df.selectVals(rowNum, Col1) has string "Hello World")
'       Column names must be entered enclosed them in (')
'           E.g. df.columns = Array("Col A","Col B"), df.filter(" 'Col A' == 'Col B' ")
'   Returns:
'       clsDataFrame filtered as entered
'---------------------------------------------------------------------------------
Public Function Filter(ByVal filterStr As String) As clsDataFrame
    Dim n As Long: n = Me.Shape(0)
    If n = 0 Then
        Set Filter = Me.SelectRows(vbNullString)
        Exit Function
    End If

    Dim arrRpnExpression As Variant: arrRpnExpression = getArrRpnNotation(getArrTokenizedStr(filterStr))
    
    Dim keep() As Boolean: keep = evalRpnColumns(arrRpnExpression, n)

    Dim rowIdx() As Variant: ReDim rowIdx(0 To n - 1)
    Dim cnt As Long, r As Long
    For r = 0 To n - 1
        If keep(r) Then
            rowIdx(cnt) = r
            cnt = cnt + 1
        End If
    Next r

    If cnt = 0 Then
        Set Filter = Me.SelectRows(vbNullString)
    Else
        ReDim Preserve rowIdx(0 To cnt - 1)
        Set Filter = Me.SelectRows(Join(rowIdx, ","))
    End If
End Function

'************************************************************************************************************************
' mutate
'   Purpose:
'       Mutates a data frame (either by adding a new column, or changing an existing column of data) based on the input expression
'   Parameter:
'       expression: a string containing the mutate conditions
'   Examples:
'       df.columns = Array("col1", "col2", "col3")
'       df.mutate(" 'col1' = 'col2' * 'col3' ") will change all values in col1 to col2*col3 for each row
'       df.mutate(" 'col4' = 'col2' * 'col3' ") will create a new column, col4, with each row value = col2*col3
'       df.mutate(" 'col4' = 'col1' <> 'col3' ") will create a new column, col4, with each row value = the boolean result of (col <> col3)
'   Notes:
'       Filter expects an expression "columnName = columnValues"
'       "columnValues" can be a constant or an expression defined using the same math syntax as used for the filter function
'   Returns:
'       clsDataFrame filtered as entered
'---------------------------------------------------------------------------------
Public Function Mutate(ByVal expression As String) As clsDataFrame
    Dim parts As Variant: parts = splitAtAssignment(expression)
    Dim newColName As String: newColName = Replace(Trim$(parts(0)), "'", vbNullString)
    Dim newColExpr As String: newColExpr = Trim$(parts(1))

    Dim n As Long: n = Me.Shape(0)

    ' Evaluate the whole expression once, column-wise
    Dim arrRpnExpression As Variant: arrRpnExpression = getArrRpnNotation(getArrTokenizedStr(newColExpr))
    Dim newCol As Variant: newCol = toColumn(evalRpnColumns(arrRpnExpression, n), n)

    ' Build the result on a copy so the original frame isn't touched
    Dim newDict As Scripting.Dictionary: Set newDict = New Scripting.Dictionary
    Dim k As Variant
    For Each k In dict_df.keys
        newDict(k) = dict_df(k)
    Next k
    
    newDict(newColName) = newCol

    Dim newDf As clsDataFrame: Set newDf = New clsDataFrame
    newDf.FromDict newDict, nullNormalized:=True
    Set Mutate = newDf
End Function

'************************************************************************************************************************
' groupByAgg
'   Purpose:
'       Allows user to group by columns then aggregate by another column
'       Equivalent to Python Panda's df.groupby().agg(), but all 1 function due to limitations of vba
'   Parameters:
'        groupByCols: array of each column to be grouped by in order. If there's only one column name, a string is also accepted.
'       colToMeasure: column name to measure / apply aggregate function to
'           aggFuncs: array of aggregate functions to use (see ListOperationsForTextExpressions for available agg functions).
'                     If there's only one agg function, a string is also accepted.
'   Examples:
'       Say we have a dataframe that stores data of death information around the country described with the following columns:
'           df.columns = array("State", "County", "Year", "Number of Fatalities")
'       Where:
'           "State" is the state of the record
'           "County" is the county of the state
'           "Year" is the year of the record
'           "Number of Fatalities" is the number of fatalities for that record
'       If we want to find the total number of recorded deaths for each county, we can do this very easily with groupby
'       We can simply do df.groupByAgg(Array("State","County"), "Number of Fatalities", "sum")
'       This will group the data by the "State" and "County" columns, then sum the number of fatalities in each of the rows
'       The returned data frame will be all states and counties with a value for the total number of fatalities associated with them
'   Notes:
'       This is roughly the same syntax as seen in Pandas Python, just all in one function.
'       More information about groupby can be found online by searching for Python examples.
'   Returns:
'       clsDataFrame grouped and aggregated as entered
'---------------------------------------------------------------------------------
Public Function GroupByAgg(ByVal groupByCols As Variant, ByVal colToMeasure As String, ByVal aggFuncs As Variant) As clsDataFrame
    Dim colName As Variant
    Dim colIdx As Long
    Dim rowIdx As Long
    
    Dim numCols As Long

    ' Key column names as a 0-based array
    Dim keyNames() As String
    If IsArray(groupByCols) Then
    
        numCols = UBound(groupByCols) - LBound(groupByCols) + 1
        
        If numCols < 1 Then pandasRaiseError "groupByAgg", PANDAS_ERROR_COL_NAME_NOT_EXISTS, Join(groupByCols, "', '")
        
        ReDim keyNames(0 To numCols - 1)
        For Each colName In groupByCols
            keyNames(colIdx) = CStr(colName)
            colIdx = colIdx + 1
        Next colName
        
    Else
        numCols = 1
        
        ReDim keyNames(0 To 0)
        
        keyNames(0) = CStr(groupByCols)
    End If
    
    ' Validate and fetch columns
    Dim arrColData() As Variant: ReDim arrColData(0 To numCols - 1)
    
    Dim dfNullFree As clsDataFrame: Set dfNullFree = Me.DropNull(getMergedArr(keyNames, Array(colToMeasure)))
    
    Dim dict_dfNullFree As Scripting.Dictionary: Set dict_dfNullFree = dfNullFree.ToDict
    
    For colIdx = 0 To numCols - 1
        If Not dict_dfNullFree.Exists(keyNames(colIdx)) Then pandasRaiseError "groupByAgg", PANDAS_ERROR_COL_NAME_NOT_EXISTS, keyNames(colIdx)
        
        arrColData(colIdx) = dict_dfNullFree(keyNames(colIdx))
    
    Next colIdx
    
    If Not dict_dfNullFree.Exists(colToMeasure) Then pandasRaiseError "groupByAgg", PANDAS_ERROR_COL_NAME_NOT_EXISTS, colToMeasure
    Dim colDataToMeasure As Variant: colDataToMeasure = dict_dfNullFree(colToMeasure)

    Dim numRows As Long: numRows = dfNullFree.Shape(0)
    If numRows = 0 Then
        Set GroupByAgg = Me.SelectRows(vbNullString) ' placeholder: return an empty frame
        Exit Function
    End If

    ' One string key per row
    Dim rowKey() As String: ReDim rowKey(0 To numRows - 1)

    For colIdx = 0 To numCols - 1
        Dim arrCurrCol As Variant: arrCurrCol = arrColData(colIdx)
        Dim lbCol As Long: lbCol = LBound(arrCurrCol)
        Dim piece As String

        For rowIdx = 0 To numRows - 1
            Dim v As Variant: v = arrCurrCol(lbCol + rowIdx)
            
            If VarType(v) <= vbNull Then piece = Chr$(30) Else piece = CStr(v)
            
            If colIdx = 0 Then
                rowKey(rowIdx) = piece
            Else
                rowKey(rowIdx) = rowKey(rowIdx) & Chr$(31) & piece
            End If
            
        Next rowIdx
        
    Next colIdx

    ' Assign group ids; remember each group's first row
    Dim rowGroupId() As Long: ReDim rowGroupId(0 To numRows - 1)
    Dim firstRow() As Long: ReDim firstRow(0 To numRows - 1)
    
    Dim dictGroups As Scripting.Dictionary: Set dictGroups = New Scripting.Dictionary
    
    ' Look for groups case-insensitive
    dictGroups.CompareMode = TextCompare
    
    Dim countGroups As Long
    For rowIdx = 0 To numRows - 1
        If dictGroups.Exists(rowKey(rowIdx)) Then
            rowGroupId(rowIdx) = dictGroups(rowKey(rowIdx))
        Else
            dictGroups.Add rowKey(rowIdx), countGroups
            firstRow(countGroups) = rowIdx
            rowGroupId(rowIdx) = countGroups
            countGroups = countGroups + 1
        End If
    Next rowIdx
    
    ReDim Preserve firstRow(0 To countGroups - 1)

    ' Aggregate
    Dim aggResults As Variant
    Dim arrAggFuncs As Variant

    Dim numAggFuncs As Long
    If isEmptyArray(aggFuncs) Then
        numAggFuncs = 1
        ReDim arrAggFuncs(0 To 0)
        arrAggFuncs(0) = aggFuncs
    Else
        numAggFuncs = UBound(aggFuncs) - LBound(aggFuncs) + 1
        ReDim arrAggFuncs(0 To numAggFuncs - 1)
        Dim aggFunc As Variant, aggFuncIdx As Long
        For Each aggFunc In aggFuncs
            arrAggFuncs(aggFuncIdx) = aggFunc
            aggFuncIdx = aggFuncIdx + 1
        Next aggFunc
    End If
    
    ReDim aggResults(0 To numAggFuncs - 1)

    Dim aggIdx As Long
    For aggIdx = 0 To numAggFuncs - 1
    
        If getAggFunctionCode(arrAggFuncs(aggIdx)) = AGG_UNK Then pandasRaiseError "groupByAgg", PANDAS_ERROR_AGG_FUNC_NOT_EXIST, arrAggFuncs(aggIdx)
        aggResults(aggIdx) = applyAgg(rowGroupId, colDataToMeasure, countGroups, getAggFunctionCode(arrAggFuncs(aggIdx)))
    Next aggIdx
    
    ' Assemble the result
    Dim groupIdx As Long
    
    Dim result As Scripting.Dictionary: Set result = New Scripting.Dictionary
    For colIdx = 0 To numCols - 1
        Dim src As Variant: src = arrColData(colIdx)
        Dim keyOut() As Variant: ReDim keyOut(0 To countGroups - 1)

        lbCol = LBound(src)
        
        For groupIdx = 0 To countGroups - 1
            ' Original typed value
            keyOut(groupIdx) = src(lbCol + firstRow(groupIdx))
        Next groupIdx
        
        result(keyNames(colIdx)) = keyOut
    Next colIdx

    Dim aggOut() As Variant: ReDim aggOut(0 To countGroups - 1)
    
    For aggIdx = 0 To numAggFuncs - 1
        Dim outName As String: outName = colToMeasure & "_" & arrAggFuncs(aggIdx)
        Dim currentAggResults As Variant:  currentAggResults = aggResults(aggIdx)
        
        If result.Exists(outName) Then pandasRaiseError "groupByAgg", PANDAS_ERROR_COL_EXISTS_TWICE, outName
        
        Dim lbAgg As Long: lbAgg = LBound(currentAggResults)
        
        For groupIdx = 0 To countGroups - 1
            aggOut(groupIdx) = currentAggResults(lbAgg + groupIdx)
        Next groupIdx
        
        result(outName) = aggOut
    Next aggIdx

    Dim newDf As clsDataFrame: Set newDf = New clsDataFrame
    newDf.FromDict result, nullNormalized:=True
    Set GroupByAgg = newDf
End Function

Private Function applyAgg(ByRef rowGroupId() As Long, ByRef colDataToMeasure As Variant, ByVal countGroups As Long, ByVal aggFn As Long) As Variant
    Select Case aggFn
        Case agg_sum
            applyAgg = aggSum(rowGroupId, colDataToMeasure, countGroups)
            
        Case AGG_COUNT
            applyAgg = aggCount(rowGroupId, colDataToMeasure, countGroups)
        
        Case AGG_MAX
            applyAgg = aggExtreme(rowGroupId, colDataToMeasure, countGroups, True)
            
        Case AGG_MIN
            applyAgg = aggExtreme(rowGroupId, colDataToMeasure, countGroups, False)
            
        Case AGG_MEAN
            applyAgg = aggMean(rowGroupId, colDataToMeasure, countGroups)

    End Select
End Function

Private Function aggSum(ByRef rowGroupId() As Long, ByRef colDataToMeasure As Variant, ByVal countGroups As Long) As Variant
    Dim sumAcc() As Variant: ReDim sumAcc(0 To countGroups - 1)
    
    Dim numRows As Long: numRows = UBound(rowGroupId) - LBound(rowGroupId) + 1
    Dim rowIdx As Long
    For rowIdx = 0 To numRows - 1
        Dim rowValue As Variant: rowValue = colDataToMeasure(rowIdx)
        
        ' scalarArith handles Null values, so prefer instead of manually adding the two
        sumAcc(rowGroupId(rowIdx)) = scalarArith(OP_ADD, sumAcc(rowGroupId(rowIdx)), rowValue)
    
    Next rowIdx
    
    aggSum = sumAcc
End Function

Private Function aggCount(ByRef rowGroupId() As Long, ByRef colDataToMeasure As Variant, ByVal countGroups As Long) As Long()
    Dim cntAcc() As Long: ReDim cntAcc(0 To countGroups - 1)
    
    Dim numRows As Long: numRows = UBound(rowGroupId) - LBound(rowGroupId) + 1
    Dim rowIdx As Long
    For rowIdx = 0 To numRows - 1
        cntAcc(rowGroupId(rowIdx)) = cntAcc(rowGroupId(rowIdx)) + 1
    Next rowIdx
    
    aggCount = cntAcc
End Function

Private Function aggMean(ByRef rowGroupId() As Long, ByRef colDataToMeasure As Variant, ByVal countGroups As Long) As Variant
    Dim sumAcc() As Variant: ReDim sumAcc(0 To countGroups - 1)
    Dim cntAcc() As Long: ReDim cntAcc(0 To countGroups - 1)
    
    Dim numRows As Long: numRows = UBound(rowGroupId) - LBound(rowGroupId) + 1
    Dim rowIdx As Long
    For rowIdx = 0 To numRows - 1
        Dim rowValue As Variant: rowValue = colDataToMeasure(rowIdx)
        
        ' scalarArith handles Null values, so prefer instead of manually adding the two
        sumAcc(rowGroupId(rowIdx)) = scalarArith(OP_ADD, sumAcc(rowGroupId(rowIdx)), rowValue)
        
        cntAcc(rowGroupId(rowIdx)) = cntAcc(rowGroupId(rowIdx)) + 1
    Next rowIdx
    
    Dim meanAcc() As Double: ReDim meanAcc(0 To countGroups - 1)
    Dim groupIdx As Long
    For groupIdx = 0 To countGroups - 1
        ' scalarArith handles Null values, so prefer instead of dividing the two
        meanAcc(groupIdx) = scalarArith(OP_DIV, sumAcc(groupIdx), cntAcc(groupIdx))
    Next groupIdx
    
    aggMean = meanAcc
End Function

Private Function aggExtreme(ByRef rowGroupId() As Long, ByRef colDataToMeasure As Variant, ByVal countGroups As Long, getMax As Boolean) As Variant
    ' Tracks if a number has been stored for each group. There is no safe way to compare max without storing the first value in the group
    Dim seen() As Boolean: ReDim seen(0 To countGroups - 1)
    Dim extreme() As Variant: ReDim extreme(0 To countGroups - 1)
    Dim target As Long: target = IIf(getMax, 4, 1) ' 4 = greater, 1 = less
    
    Dim numRows As Long: numRows = UBound(rowGroupId) - LBound(rowGroupId) + 1
    Dim rowIdx As Long
    For rowIdx = 0 To numRows - 1
        Dim rowValue As Variant: rowValue = colDataToMeasure(rowIdx)
        
        If VarType(rowValue) > vbNull Then
            If Not IsNumeric(rowValue) Then pandasRaiseError "groupByAgg", PANDAS_ERROR_ROW_NOT_NUM, rowValue
        
            If Not seen(rowGroupId(rowIdx)) Then
                extreme(rowGroupId(rowIdx)) = rowValue
                seen(rowGroupId(rowIdx)) = True
            ElseIf compareBit(rowValue, extreme(rowGroupId(rowIdx))) = target Then
                extreme(rowGroupId(rowIdx)) = rowValue
            End If
        End If
    Next rowIdx
        
    Dim groupIdx As Long
    For groupIdx = 0 To countGroups - 1
        If Not seen(groupIdx) Then extreme(groupIdx) = Null
    Next groupIdx
    
    aggExtreme = extreme
End Function

Private Function aggFunctionRegistry() As Scripting.Dictionary
    If m_aggFnNames Is Nothing Then
        Dim aggFnNames As Scripting.Dictionary: Set aggFnNames = New Scripting.Dictionary
        
        aggFnNames.Add "sum", agg_sum
        aggFnNames.Add "count", AGG_COUNT
        aggFnNames.Add "max", AGG_MAX
        aggFnNames.Add "min", AGG_MIN
        aggFnNames.Add "mean", AGG_MEAN
        
        Set m_aggFnNames = aggFnNames
    End If
    Set aggFunctionRegistry = m_aggFnNames
End Function

Private Function getAggFunctionCode(ByVal aggFunctionName As String) As Long
    If aggFunctionRegistry().Exists(aggFunctionName) Then
        getAggFunctionCode = aggFunctionRegistry(aggFunctionName)
    Else
        getAggFunctionCode = AGG_UNK
    End If
End Function

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ ERROR HANDLING //////////////////////////////////////////////////////////////////////////////////////
Private Sub pandasRaiseError(ByVal procedureName As String, ByVal errID As Long, Optional ByVal additionalComment As String)
    Dim errInfo As Variant: errInfo = pandasGetErrorInfo(procedureName, errID, additionalComment)
    Err.Raise errID, errInfo(0), errInfo(1)
End Sub

Private Function pandasGetErrorInfo(ByVal procedureName As String, ByVal errID As Long, Optional ByVal additionalComment As String) As Variant
    Dim errDescription As String
    Dim errResolution As String
    Dim errSource As String: errSource = "clsDataFrame." & procedureName
    
    Select Case errID
        Case PANDAS_ERROR_SELECTED_COL_OUT_OF_BOUNDS
            errDescription = "The selected column index " & "'" & additionalComment & "'" & " is out of bounds of the dataframe! (indexes available: [0, " & Me.Shape(1) - 1 & "]"
            
            errResolution = "Check the selected column indexes and ensure each is valid."
                
        Case PANDAS_ERROR_SELECTED_ROW_OUT_OF_BOUNDS
            errDescription = "The selected row index " & "'" & additionalComment & "'" & " is out of bounds of the dataframe! (indexes available: [0, " & Me.Shape(0) - 1 & "]"
            errResolution = "Check the selected row indexes and ensure each is valid."
                
        Case PANDAS_ERROR_COL_NAME_NOT_EXISTS
            errDescription = "The entered column " & "'" & additionalComment & "'" & " does not exist!"
            errResolution = "Check spelling of each column name entered and ensure they exist in the dataframe."
        
        Case PANDAS_ERROR_EVALUATE_SYNTAX
            errDescription = "Invalid syntax entered in the math expression! Error: " & additionalComment
            errResolution = "Ensure expression follows PEMDAS, has any literals for comparison enclosed in (""""), and column names enclosed in ('), and functions written as plain text."
        
        Case PANDAS_ERROR_INVALID_ARITH_TYPES
            errDescription = "Invalid operation between parameters! " & additionalComment
            errResolution = vbNullString
        
        Case PANDAS_ERROR_COL_HAS_FORBIDDEN_CHARS
            errDescription = "Column name '" & additionalComment & "' is invalid!"
            errResolution = "Ensure none of the following characters are used in column names:" _
                            & vbNewLine & _
                            Join(split(PANDAS_VALUES_FORBIDDEN_COL_CHARS, "\"), " ")
        
        Case PANDAS_ERROR_COL_EXISTS_TWICE
            errDescription = "Duplicate column name " & "'" & additionalComment & "'" & " in data!"
            errResolution = "Ensure all column names are have unique names."
            
        Case PANDAS_ERROR_COL_TYPES_DONT_MATCH
            errDescription = "Column types are not compatible with the inputted operation! " & additionalComment
            errResolution = "Check all expressions passed as a string only use columns with compatible data types for the operation."
        
        Case PANDAS_ERROR_MISMATCHED_NUMBER_COLS
            errDescription = "Number of new column names do not match with the number of old column names in rename."
            
        Case PANDAS_ERROR_ROW_NOT_NUM
            errDescription = "Inputted row index " & "'" & additionalComment & "' has non-numeric values!"
        
        Case PANDAS_ERROR_AGG_FUNC_NOT_EXIST
            errDescription = "Aggregate function " & "'" & additionalComment & "' is not recognized!"
        
        Case PANDAS_ERROR_FILE_NOT_FOUND
            errDescription = "File not found at path " & "'" & additionalComment & "'"
        
    End Select
    
    pandasGetErrorInfo = Array(errSource, errDescription & vbNewLine & vbNewLine & errResolution)

End Function

'********************************************************************************************************************************************************************
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ EVALUATE EXPR /////////////////////////////////////////////////////////////////////////////////////////
'********************************************************************************************************************************************************************

'************************************************************************************************************************
' getArrTokenizedStr
'   Purpose:
'       Get an array where each index is a token of an expression passed as a string
'   Parameter:
'       expression: a string containing the expression to tokenize
'   Examples:
'       getArrTokenizedStr("1+2*(3+2)") -> array("1", "+", "2", "*", "(", "3", "+", "2", ")")
'       getArrTokenizedStr(" 'col1' == ""Hello"" ") -> array("'col1'", "==", ""Hello"")
'   Returns:
'       Array of the tokenized input
'---------------------------------------------------------------------------------
Private Function getArrTokenizedStr(ByVal expression As String) As Variant
    Dim arrStackTokenizedStr As Variant: ReDim arrStackTokenizedStr(0 To 127)
    Dim stackPtr As Long: stackPtr = 0
    
    Dim cIdx As Long: cIdx = 1
    Do While cIdx <= Len(expression)
        
        Dim currChar As String: currChar = mid$(expression, cIdx, 1)
        
        If currChar = " " Then GoTo nextChar
        
        Dim currCharType As Long: currCharType = getTokenType(currChar)
                
        Dim token As Variant: token = vbNullString
       
        ' Characters enclosed in ticks (') will be lumped into its own token (including the ticks). Syntax for column names.
        If currChar = "'" Then
            Dim inSingleQuotes As Boolean: inSingleQuotes = True
            Do While inSingleQuotes
                ' Triggers if there's a mismatched "'"
                If Not debugMode And cIdx = Len(expression) Then pandasRaiseError "getArrTokenizedStr", PANDAS_ERROR_EVALUATE_SYNTAX, "Mismatched (')!"
                
                token = token & currChar
                
                cIdx = cIdx + 1
                
                currChar = mid$(expression, cIdx, 1)
                
                inSingleQuotes = (currChar <> "'")
            Loop
            token = token & "'"
            pushArrStack arrStackTokenizedStr, stackPtr, token

            GoTo nextChar
            
        ' Characters enclosed in quotes (") will be lumped into its own token (including the quotes). Syntax for literals.
        ElseIf currChar = """" Then
            Dim inDoubleQuotes As Boolean: inDoubleQuotes = True
            Do While inDoubleQuotes
                ' Triggers if there's a mismatched """"
                If Not debugMode And cIdx = Len(expression) Then pandasRaiseError "getArrTokenizedStr", PANDAS_ERROR_EVALUATE_SYNTAX, "Mistmatched ("""")!"
                
                token = token & currChar
                
                cIdx = cIdx + 1
                
                currChar = mid$(expression, cIdx, 1)
                
                inDoubleQuotes = (currChar <> """")
            Loop
            ' Add the closing quote mark
            token = token & """"
            
            pushArrStack arrStackTokenizedStr, stackPtr, token

            GoTo nextChar
        End If
       
        Dim nextChar As String: nextChar = vbNullString

        If cIdx < Len(expression) Then nextChar = mid$(expression, cIdx + 1, 1)
        
        Dim nextCharType As Long: nextCharType = getTokenType(nextChar)
       
        token = currChar
        
        ' Adjacent characters of the same type will be lumped into its own token (e.g., ajacent numbers will be read as digits of a bigger number, not each number be its own token)
        Do
            currChar = mid$(expression, cIdx, 1)
            currCharType = getTokenType(currChar)
            
            If cIdx < Len(expression) Then nextChar = mid$(expression, cIdx + 1, 1)
            nextCharType = getTokenType(nextChar)
            
            If currCharType <> nextCharType Or currCharType = TK_PAR Then Exit Do
            
            token = token & nextChar

            cIdx = cIdx + 1
        Loop
 
        If Trim$(token) = "Null" Then token = Null

        pushArrStack arrStackTokenizedStr, stackPtr, token
        
nextChar:
        cIdx = cIdx + 1
    Loop
    
    ReDim Preserve arrStackTokenizedStr(0 To stackPtr - 1)
    
    ' Ensure tokens are not surrounded in whitespace
    Dim idx As Variant
    For idx = LBound(arrStackTokenizedStr) To UBound(arrStackTokenizedStr)
        If Not IsNull(arrStackTokenizedStr(idx)) Then arrStackTokenizedStr(idx) = Trim(arrStackTokenizedStr(idx))
    Next idx
    
    getArrTokenizedStr = arrStackTokenizedStr
End Function

'************************************************************************************************************************
' getArrRpnNotation
'   Purpose:
'       Get an array where each index is a token, ordered in rpnNotation for evaluation
'   Parameter:
'       tokens: array of tokens in order of a normal expression (e.g., array tokenized version of "1+2*(3-1)" )
'   Returns:
'       Array of the input, in rpn order (read left to right for rpn evaluation)
'---------------------------------------------------------------------------------
Private Function getArrRpnNotation(ByRef tokens As Variant) As Variant
    Dim size As Long: size = UBound(tokens) - LBound(tokens) + 1
    Dim outArr As Variant: ReDim outArr(0 To size)
    Dim outPtr As Long
    
    Dim opArr As Variant:  ReDim opArr(0 To size)
    Dim opPtr As Long
    
    Dim idx As Long
    For idx = LBound(tokens) To UBound(tokens)
        Dim token As Variant: token = tokens(idx)

        Select Case True
            Case IsNull(token)
                pushArrStack outArr, outPtr, token

            Case isFunctionToken(token), token = "("
                pushArrStack opArr, opPtr, token

            Case token = ","
                ' Argument separator: flush operators back to the opening "("
                Do While opPtr > 0
                    If opArr(opPtr - 1) = "(" Then Exit Do
                    pushArrStack outArr, outPtr, popArrStack(opArr, opPtr)
                Loop
                If opPtr = 0 Then pandasRaiseError "getArrRpnNotation", PANDAS_ERROR_EVALUATE_SYNTAX, "Mismatched ""(""!"

            Case token = ")"
                Do While opPtr > 0
                    If opArr(opPtr - 1) = "(" Then Exit Do
                    pushArrStack outArr, outPtr, popArrStack(opArr, opPtr)
                Loop
                If opPtr = 0 Then pandasRaiseError "getArrRpnNotation", PANDAS_ERROR_EVALUATE_SYNTAX, "Mismatched "")""! "
                
                ' Discard the "("
                popArrStack opArr, opPtr
                If opPtr > 0 Then
                    ' Check if there's a function associated with the parentheses
                    If isFunctionToken(opArr(opPtr - 1)) Then
                        pushArrStack outArr, outPtr, popArrStack(opArr, opPtr)
                    End If
                End If

            ' Number or string (will also catch functions that don't exist or strings not enclosed in (") )
            Case getTokenPrecedence(token) = -1
                ' Any string token not enclosed in "" or '' is assumed to be a function. If it was not found above, then the function does not exist
                If InStr(1, token, """") = 0 And InStr(1, token, "'") = 0 Then
                    If Not IsNumeric(token) Then pandasRaiseError "getArrRpnNotation", PANDAS_ERROR_EVALUATE_SYNTAX, "Unknown function " & "'" & token & "'"
                End If
                
                pushArrStack outArr, outPtr, Replace(token, """", vbNullString)
                
            ' Operator
            Case Else
                Do While opPtr > 0
                    If opArr(opPtr - 1) = "(" Then Exit Do
                    If Not shouldPopBefore(token, opArr(opPtr - 1)) Then Exit Do
                    pushArrStack outArr, outPtr, popArrStack(opArr, opPtr)
                Loop
                pushArrStack opArr, opPtr, token
        End Select
    Next idx

    Do While opPtr > 0
        If opArr(opPtr - 1) = "(" Then pandasRaiseError "getArrRpnNotation", PANDAS_ERROR_EVALUATE_SYNTAX, "Mismatched ""(""! "
        pushArrStack outArr, outPtr, popArrStack(opArr, opPtr)
    Loop

    If outPtr = 0 Then pandasRaiseError "getArrRpnNotation", PANDAS_ERROR_EVALUATE_SYNTAX, "General syntax issues."
    ReDim Preserve outArr(0 To outPtr - 1)
    getArrRpnNotation = outArr
End Function

'************************************************************************************************************************
' shouldPopBefore
'   Purpose:
'       Check if operator on top of stack must be output before pushing token
'   Parameters:
'       token: token pending placement in stack
'    topToken: token at the top of the stack
'   Returns:
'       True if the operator on top of the stack must be output before pushing `token`
'---------------------------------------------------------------------------------
Private Function shouldPopBefore(ByRef token As Variant, ByRef topToken As Variant) As Boolean
    Dim prec As Long: prec = getTokenPrecedence(token)
    Dim topPrec As Long: topPrec = getTokenPrecedence(topToken)

    shouldPopBefore = (topPrec > prec) Or _
                      (topPrec = prec And getTokenAssociativity(token) = ASS_L)
End Function

'************************************************************************************************************************
' popArrStack
'   Purpose:
'       Pop the arrStack (decrement stack pointer and return last item)
'   Parameters:
'       arrStack: array acting as a stack
'       stackPtr: stack pointer (assumes the top of the stack is the open spot above the last placed item)
'   Returns:
'       The item previously on top of the stack
'---------------------------------------------------------------------------------
Private Function popArrStack(ByRef arrStack As Variant, ByRef stackPtr As Long) As Variant
    If stackPtr = 0 Then Exit Function
    arrStack(stackPtr) = Empty
    stackPtr = stackPtr - 1
    popArrStack = arrStack(stackPtr)
End Function

'************************************************************************************************************************
' pushArrStack
'   Purpose:
'       Push to the arrStack (place item stackVal on top of stack, increment stack pointer)
'   Parameters:
'       arrStack: array acting as a stack
'       stackPtr: stack pointer (assumes the top of the stack is the open spot above the last placed item)
'       stackVal: value to push to stack
'---------------------------------------------------------------------------------
Private Sub pushArrStack(ByRef arrStack As Variant, ByRef stackPtr As Long, ByVal stackVal As Variant)
    If stackPtr > UBound(arrStack) Then ReDim Preserve arrStack(0 To UBound(arrStack) * 2 + 1)
    
    arrStack(stackPtr) = stackVal
    stackPtr = stackPtr + 1
End Sub

'************************************************************************************************************************
' getTokenType
'   Purpose:
'       Get the type (e.g., operator, parentheses, etc.) of a token
'   Parameter:
'       token: token to get the type of
'   Returns:
'       The token type code in enum tokenTypes
'---------------------------------------------------------------------------------
Private Function getTokenType(ByVal token As Variant) As Long
    If IsNumeric(token) Then
        getTokenType = TK_NUM
        Exit Function
    End If

    Select Case token
        Case "+", "-", "<", ">", "^", "=", "*", "/", "==", "<=", ">=", "&", "|", "<>", "^", "%"
            getTokenType = TK_OP
        Case "(", ")"
            getTokenType = TK_PAR
        Case vbNullString
            getTokenType = TK_NULLSTR
        Case PANDAS_VALUES_ERROR
            getTokenType = TK_NULL
        Case Else
            If isFunctionToken(token) Then
                getTokenType = TK_FN
            Else
                getTokenType = TK_STR
            End If
    End Select
End Function

'************************************************************************************************************************
' getTokenPrecedence
'   Purpose:
'       Get the precedence of a token (needed for rpn evaluation)
'   Parameter:
'       token: token to get the precedence of
'   Returns:
'       The token's precedence
'---------------------------------------------------------------------------------
Private Function getTokenPrecedence(ByVal token As Variant) As Long
    Select Case token
        Case "^"
            getTokenPrecedence = 4
        Case "*", "/", "%"
            getTokenPrecedence = 3
        Case "+", "-"
            getTokenPrecedence = 2
        Case ">", ">=", "<", "<=", "==", "<>"
            getTokenPrecedence = 1
        Case "&", "|"
            getTokenPrecedence = 0
        Case Else   ' numbers or characters
            getTokenPrecedence = -1
    End Select

End Function

'************************************************************************************************************************
' getTokenAssociativity
'   Purpose:
'       Get the associativity of a token (needed for rpn evaluation)
'   Parameter:
'       token: token to get the associativity of
'   Returns:
'       The token's associativity
'---------------------------------------------------------------------------------
Private Function getTokenAssociativity(ByVal token As String) As Long
    Select Case token
        Case "^"
            getTokenAssociativity = ASS_R
        Case "*", "/", "+", "-", ">", ">=", "<", "<=", "==", "<>", "&", "|", "%"
            getTokenAssociativity = ASS_L
        Case Else   ' numbers or characters
            getTokenAssociativity = -1
    End Select

End Function

'************************************************************************************************************************
' functionRegistry
'   Purpose:
'       Acts as a register of all functions implemented in program
'   Returns:
'       A dictionary with all function codes (defined in enum rpnFn) as keys
'---------------------------------------------------------------------------------
Private Function functionRegistry() As Scripting.Dictionary
    If m_fnNames Is Nothing Then
        Set m_fnNames = New Scripting.Dictionary
        
        ' Case insensitive
        m_fnNames.CompareMode = TextCompare
        m_fnNames.Add "abs", FN_ABS
        m_fnNames.Add "len", FN_LEN
        m_fnNames.Add "sqrt", FN_SQRT
        m_fnNames.Add "round", FN_ROUND
    End If
    
    Set functionRegistry = m_fnNames
End Function

'************************************************************************************************************************
' isFunctionToken
'   Purpose:
'       Check if token is a function
'   Parameter:
'       token: token to check if it's a function
'   Returns:
'       True if token is a function, False if not.
'---------------------------------------------------------------------------------
Private Function isFunctionToken(ByVal token As Variant) As Boolean
    If VarType(token) <> vbString Then Exit Function
    isFunctionToken = functionRegistry.Exists(token)
End Function

'************************************************************************************************************************
' getFunctionArity
'   Purpose:
'       Get the arity (number of input parameters) of a function based on its code
'   Parameter:
'       fnCode: code of the function to get the arity of
'   Returns:
'       Arity (number of input paras) expected by function
'---------------------------------------------------------------------------------
Private Function getFunctionArity(ByVal fnCode As Long) As Long
    Select Case fnCode
        Case FN_ROUND:  getFunctionArity = 2
        Case Else:      getFunctionArity = 1
    End Select
End Function

'************************************************************************************************************************
' getFunctionCode
'   Purpose:
'       Get the code a function based on its name
'   Parameter:
'       token: token that is a function
'   Returns:
'       Token code of the inputted token name. Defaults to FN_NONE if not an implemented function
'---------------------------------------------------------------------------------
Private Function getFunctionCode(ByVal token As Variant) As Long
    If functionRegistry().Exists(token) Then
        getFunctionCode = functionRegistry(token)
    Else
        getFunctionCode = FN_NONE
    End If
End Function

'********************************************************************************************************************************************************************
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ EVALUATE EXPR /////////////////////////////////////////////////////////////////////////////////////////
'********************************************************************************************************************************************************************

'************************************************************************************************************************
' evalRpnColumns
'   Purpose:
'       Analyzes entire array of values per the expression given
'   Parameter:
'       arrRpnExpression: array with an rpn expression. Each expression token is stored in its own index.
'   Returns:
'       Array of operation outcomes per row. These can be numbers or booleans both as arrays or single values, depending on what was passed thru.
'---------------------------------------------------------------------------------
Private Function evalRpnColumns(ByRef arrRpnExpression As Variant, ByVal n As Long) As Variant
    Dim lo As Long: lo = LBound(arrRpnExpression)
    
    Dim stack() As Variant: ReDim stack(0 To UBound(arrRpnExpression) - lo + 1)
    Dim sp As Long
        
    Dim colName As String
    
    Dim i As Long
    For i = lo To UBound(arrRpnExpression)
        Dim token As Variant: token = arrRpnExpression(i)
        
        If IsNull(token) Then
            stack(sp) = token
            sp = sp + 1

        ElseIf IsColumnToken(token) Then
            ' Column names are enclosed with (')
            colName = mid$(token, 2, Len(token) - 2)
            
            If Not dict_df.Exists(colName) Then pandasRaiseError "evalRpnColumns", PANDAS_ERROR_COL_NAME_NOT_EXISTS, colName
            
            stack(sp) = dict_df(colName)
            
            sp = sp + 1

        Else
            Select Case getTokenType(token)
                Case TK_NUM
                    If VarType(token) = vbString Then stack(sp) = val(token) Else stack(sp) = token
                    sp = sp + 1

                Case TK_STR, TK_NULLSTR
                    stack(sp) = token
                    sp = sp + 1

                Case TK_OP
                    Dim op As Long: op = getOperatorCode(token)
                    If op = OP_NONE Or sp < 2 Then pandasRaiseError "evalRpnColumns", PANDAS_ERROR_EVALUATE_SYNTAX, "Invalid syntax for operator '" & operatorRegistry(op) & "'!"
                    
                    stack(sp - 2) = applyOperator(op, stack(sp - 2), stack(sp - 1), n)
                    stack(sp - 1) = Empty            ' release the popped operand
                    sp = sp - 1
                
                Case TK_FN
                    Dim fn As Long: fn = getFunctionCode(token)
                    
                    ' If the function name is unrecognized
                    If fn = FN_NONE Then pandasRaiseError "evalRpnColumns", PANDAS_ERROR_EVALUATE_SYNTAX, "Unknown function '" & functionRegistry(fn) & "'"
                    
                    Dim arity As Long: arity = getFunctionArity(fn)
                    
                    ' If the function requires more paras than given
                    If sp < arity Then pandasRaiseError "evalRpnColumns", PANDAS_ERROR_EVALUATE_SYNTAX, "Only " & sp & " parameters entered for " & functionRegistry(fn) & " but " & arity & " parameters expected"
                    
                    If arity = 1 Then
                        stack(sp - 1) = applyFunc1(fn, stack(sp - 1), n)
                    ElseIf arity = 2 Then
                        stack(sp - 2) = applyFunc2(fn, stack(sp - 2), stack(sp - 1), n)
                        stack(sp - 1) = Empty
                        sp = sp - 1
                    End If
                
                Case Else
                    pandasRaiseError "evalRpnColumns", PANDAS_ERROR_EVALUATE_SYNTAX, "Unexpected token '" & token & "'"
                    
            End Select
        End If
    Next i

    If sp <> 1 Then pandasRaiseError "evalRpnColumns", PANDAS_ERROR_EVALUATE_SYNTAX, "General syntax issues."
    
    evalRpnColumns = stack(0)
End Function

'************************************************************************************************************************
' IsColumnToken
'   Purpose:
'       Checks if the token is a column name
'   Parameter:
'       token: token in a tokenized array
'   Returns:
'       True if a token represents a column, False if not
'---------------------------------------------------------------------------------
Private Function IsColumnToken(ByRef token As Variant) As Boolean
    If VarType(token) <> vbString Then Exit Function
    If Len(token) < 2 Then Exit Function
    IsColumnToken = (left$(token, 1) = "'" And right$(token, 1) = "'")
End Function

'************************************************************************************************************************
' getOperatorCode
'   Purpose:
'       Gets the operator code of the operator represented by the token
'   Parameter:
'       token: token in a tokenized array
'   Returns:
'       Operator code based on enum RpnOp
'---------------------------------------------------------------------------------
Private Function getOperatorCode(ByRef token As Variant) As Long
    If operatorRegistry().Exists(token) Then
        getOperatorCode = operatorRegistry(token)
    Else
        getOperatorCode = OP_NONE
    End If
End Function

'************************************************************************************************************************
' operatorRegistry:
'   Purpose:
'       Get all available operations
'   Returns:
'       Dictionary where each key is the available operation
'---------------------------------------------------------------------------------
Private Function operatorRegistry() As Scripting.Dictionary
    If m_opNames Is Nothing Then
        Set m_opNames = New Scripting.Dictionary
        m_opNames.Add "+", OP_ADD
        m_opNames.Add "-", OP_SUB
        m_opNames.Add "*", OP_MUL
        m_opNames.Add "/", OP_DIV
        m_opNames.Add "<", OP_LT
        m_opNames.Add "<=", OP_LE
        m_opNames.Add "==", OP_EQ
        m_opNames.Add ">=", OP_GE
        m_opNames.Add ">", OP_GT
        m_opNames.Add "<>", OP_NE
        m_opNames.Add "&", OP_AND
        m_opNames.Add "|", OP_OR
        m_opNames.Add "^", OP_POW
        m_opNames.Add "%", OP_MOD
    End If
    
    Set operatorRegistry = m_opNames
End Function

'************************************************************************************************************************
' applyOperator
'   Purpose:
'       Gets the results of the operation of the passed parameters by passing a mask to other function
'   Parameters:
'       op: operator code representing the operator being used
'        a: left parameter (array or individual value accepted)
'        b: right parameter (array or individual value accepted)
'        n: number of indexes in a and b (1 if individual value)
'   Returns:
'       Results of the operation, a op b
'       Array if a and b are arrays, scalar if a and b are scalars
'---------------------------------------------------------------------------------
Private Function applyOperator(ByVal op As Long, ByRef a As Variant, ByRef b As Variant, ByVal n As Long) As Variant
    Select Case op
        ' Comparison masks: 1 = less, 2 = equal, 4 = greater, 8 = unordered (Null/Empty)
        Case OP_LT: applyOperator = compareOp(1, a, b, n)
        Case OP_EQ: applyOperator = compareOp(2, a, b, n)
        Case OP_GT: applyOperator = compareOp(4, a, b, n)
        Case OP_LE: applyOperator = compareOp(3, a, b, n)
        Case OP_GE: applyOperator = compareOp(6, a, b, n)
        Case OP_NE: applyOperator = compareOp(13, a, b, n)
        Case OP_AND: applyOperator = logicalOp(True, a, b, n)
        Case OP_OR:  applyOperator = logicalOp(False, a, b, n)
        Case Else:   applyOperator = arithOp(op, a, b, n)
    End Select
End Function

'************************************************************************************************************************
' compareOp
'   Purpose:
'       Applys the logic operator to the passed parameters
'   Parameter:
'       op: operator code representing the operator being used
'        a: left parameter (array or individual value accepted)
'        b: right parameter (array or individual value accepted)
'        n: number of indexes in a and b (1 if individual value)
'   Notes:
'       Uses masks for quicker evaluation
'   Returns:
'       Results of the operation, a op b
'       Array if a and b are arrays, scalar if a and b are scalars
'---------------------------------------------------------------------------------
Private Function compareOp(ByVal mask As Long, ByRef a As Variant, ByRef b As Variant, ByVal n As Long) As Variant
    Dim aArr As Boolean: aArr = IsArray(a)
    Dim bArr As Boolean: bArr = IsArray(b)

    ' Scalar vs Scalar
    If Not aArr And Not bArr Then
        compareOp = ((mask And compareBit(a, b)) <> 0)
        Exit Function
    End If

    Dim res() As Boolean: ReDim res(0 To n - 1)
    
    Dim i As Long
    Dim la As Long
    Dim lb As Long
    
    ' Column vs Scalar
    If aArr And Not bArr Then
        la = LBound(a)
        For i = 0 To n - 1
            res(i) = ((mask And compareBit(a(la + i), b)) <> 0)
        Next i
    
    ' Scalar vs Column
    ElseIf Not aArr Then
        lb = LBound(b)
        For i = 0 To n - 1
            res(i) = ((mask And compareBit(a, b(lb + i))) <> 0)
        Next i
        
    ' Column vs Column
    Else
        la = LBound(a)
        lb = LBound(b)
        
        For i = 0 To n - 1
            res(i) = ((mask And compareBit(a(la + i), b(lb + i))) <> 0)
        Next i
        
    End If
    
    compareOp = res
End Function

'************************************************************************************************************************
' compareBit
'   Purpose:
'       Compares x and y, returns value asssuming mask applied to the results
'   Parameter:
'       x: left parameter
'       y: right parameter
'   Returns:
'       1 (x<y)
'       2 (x=y)
'       4 (x>y)
'       8 (unordered: Null/Empty involved)
'---------------------------------------------------------------------------------
Private Function compareBit(ByRef x As Variant, ByRef y As Variant) As Long
    Dim tx As Long: tx = VarType(x)
    Dim ty As Long: ty = VarType(y)

    If tx <= vbNull And ty <= vbNull Then
        compareBit = 2
    
    ElseIf tx <= vbNull Or ty <= vbNull Then
        compareBit = 8
        
    ' Any text involved -> use text compare
    ElseIf tx = vbString Or ty = vbString Then
        Select Case StrComp(CStr(x), CStr(y), vbTextCompare)
            Case Is < 0: compareBit = 1
            Case 0:      compareBit = 2
            Case Else:   compareBit = 4
        End Select
    Else
        If x < y Then
            compareBit = 1
        ElseIf x > y Then
            compareBit = 4
        Else
            compareBit = 2
        End If
    End If
End Function

'************************************************************************************************************************
' logicalOp
'   Purpose:
'       Perform and/or operation for a column
'   Parameters:
'       isAnd: True if doing an AND comparsion, false if OR
'           a: left parameter (array)
'           b: right parameter (array)
'           n: number of indexes in array
'   Returns:
'       Boolean array with the truth values of the operation at each index
'---------------------------------------------------------------------------------
Private Function logicalOp(ByVal isAnd As Boolean, ByRef a As Variant, ByRef b As Variant, ByVal n As Long) As Variant
    Dim av() As Boolean: av = toBoolArray(a, n)
    Dim bv() As Boolean: bv = toBoolArray(b, n)
    Dim i As Long
    If isAnd Then
        For i = 0 To n - 1
            av(i) = av(i) And bv(i)
        Next i
    
    Else
        For i = 0 To n - 1
            av(i) = av(i) Or bv(i)
        Next i
    End If
    logicalOp = av
End Function

'************************************************************************************************************************
' toBoolArray
'   Purpose:
'       Convert the value (or array of values) to a boolean array
'       If a single value is passed, this will return an array of that constant boolean value based on n array indexes
'   Parameters:
'       v: value or array
'       n: number of indexes in array
'   Returns:
'       Boolean array with the truth values of the operation at each index
'---------------------------------------------------------------------------------
Private Function toBoolArray(ByRef v As Variant, ByVal n As Long) As Boolean()
    Dim res() As Boolean: ReDim res(0 To n - 1)
    Dim i As Long
    Dim x As Variant
    If Not IsArray(v) Then
        If VarType(v) > vbNull Then
            If CBool(v) Then
                For i = 0 To n - 1
                    res(i) = True
                Next i
            End If
        End If
        
    ElseIf VarType(v) = (vbBoolean Or vbArray) And LBound(v) = 0 Then
        res = v
        
    Else
        Dim lb As Long
        lb = LBound(v)
        For i = 0 To n - 1
            x = v(lb + i)
            If VarType(x) > vbNull Then res(i) = CBool(x)
        Next i
    End If
    
    toBoolArray = res
End Function

'************************************************************************************************************************
' arithOp
'   Purpose:
'       Performs inputted arithmetic operation for either a scalar or array of values
'   Parameters:
'       op: arithmetic operation token
'       a: left parameter (array or scalar)
'       b: right parameter (array or scalar)
'       n: number of indexes in array
'   Returns:
'       Array or scalar with the value of the arithmetic operation
'---------------------------------------------------------------------------------
Private Function arithOp(ByVal op As Long, ByRef a As Variant, ByRef b As Variant, ByVal n As Long) As Variant
    Dim aArr As Boolean: aArr = IsArray(a)
    Dim bArr As Boolean: bArr = IsArray(b)

    If Not aArr And Not bArr Then
        arithOp = scalarArith(op, a, b)
        Exit Function
    End If

    Dim res() As Variant: ReDim res(0 To n - 1)
    
    Dim i As Long
    Dim x As Variant
    Dim y As Variant
    Dim la As Long: If aArr Then la = LBound(a)
    Dim lb As Long: If bArr Then lb = LBound(b)

    For i = 0 To n - 1
        If aArr Then x = a(la + i) Else x = a
        If bArr Then y = b(lb + i) Else y = b
        res(i) = scalarArith(op, x, y)
    Next i
    arithOp = res
End Function

'************************************************************************************************************************
' scalarArith
'   Purpose:
'       Performs inputted arithmetic operation of the inputted scalars
'   Parameters:
'       op: arithmetic operation token
'       a: left parameter (scalar)
'       b: right parameter (scalar)
'       n: number of indexes in array
'   Returns:
'       Scalar with the value of the arithmetic operation
'---------------------------------------------------------------------------------
Private Function scalarArith(ByVal op As Long, ByRef x As Variant, ByRef y As Variant) As Variant
    If IsNull(x) Or IsNull(y) Then
        scalarArith = Null
        Exit Function
    End If
    
    If Not (IsNumeric(x) And IsNumeric(y)) Then pandasRaiseError "scalarArith", PANDAS_ERROR_INVALID_ARITH_TYPES, "Invalid operation between " & TypeName(x) & " and " & TypeName(y)
    
    Select Case op
        Case OP_ADD: scalarArith = x + y
        Case OP_SUB: scalarArith = x - y
        Case OP_MUL: scalarArith = x * y
        Case OP_DIV
            'Prevents dividing by zero
            If y = 0 Then
                scalarArith = Null
            Else
                scalarArith = x / y
            End If
            
        Case OP_POW
            If x = 0 And y < 0 Then
                scalarArith = Null
            ElseIf x < 0 And y <> Int(y) Then
                scalarArith = Null
            Else
                On Error Resume Next
                scalarArith = x ^ y
                If Err.Number <> 0 Then scalarArith = Null
                On Error GoTo 0
            End If
            
        Case OP_MOD
            If y = 0 Then
                scalarArith = Null
            Else
                scalarArith = x - y * Int(x / y)    ' modulo manually computed bc VBA's is whack
            End If
    End Select
    
End Function

'************************************************************************************************************************
' applyFunc1
'   Purpose:
'       Applies the defined functions with a single arity (only one input parameter)
'   Parameters:
'       fn: function code per enum RpnOp
'        a: array or scalar as an input to the function
'        n: number of indexes in the array (=1 if scalar)
'   Returns:
'       Array or scalar with the value of the function
'---------------------------------------------------------------------------------
Private Function applyFunc1(ByVal fn As Long, ByRef a As Variant, ByVal n As Long) As Variant
    If Not IsArray(a) Then
        applyFunc1 = scalarFunc1(fn, a)
        Exit Function
    End If

    Dim res() As Variant: ReDim res(0 To n - 1)
    Dim la As Long: la = LBound(a)
    Dim i As Long
    For i = 0 To n - 1
        res(i) = scalarFunc1(fn, a(la + i))
    Next i
    applyFunc1 = res
End Function

'************************************************************************************************************************
' scalarFunc1
'   Purpose:
'       Performs the calculation of the defined functions with single arity (one input parameter)
'   Parameters:
'       fn: function code per enum RpnOp
'        x: parameter used for the function
'   Returns:
'       Scalar value of the function output
'---------------------------------------------------------------------------------
Private Function scalarFunc1(ByVal fn As Long, ByRef x As Variant) As Variant
    If VarType(x) <= vbNull Then          ' Empty or Null in -> Null out
        scalarFunc1 = Null
        Exit Function
    End If

    Select Case fn
        Case FN_ABS:  scalarFunc1 = Abs(x)
        Case FN_LEN:  scalarFunc1 = Len(CStr(x))
        Case FN_SQRT
            If x < 0 Then scalarFunc1 = Null Else scalarFunc1 = Sqr(x)
    End Select
End Function

'************************************************************************************************************************
' applyFunc2
'   Purpose:
'       Applies the defined functions with double arity (two input parameters)
'   Parameters:
'       fn: function code per enum RpnOp
'        a: array or scalar as the first input to the function
'        b: array or scalar as the second input to the function
'        n: number of indexes in the arrays (=1 if scalar)
'   Returns:
'       Array or scalar with the output value of the function
'---------------------------------------------------------------------------------
Private Function applyFunc2(ByVal fn As Long, ByRef a As Variant, ByRef b As Variant, ByVal n As Long) As Variant
    Dim aArr As Boolean: aArr = IsArray(a)
    Dim bArr As Boolean: bArr = IsArray(b)

    If Not aArr And Not bArr Then
        applyFunc2 = scalarFunc2(fn, a, b)
        Exit Function
    End If

    Dim res() As Variant: ReDim res(0 To n - 1)
    Dim la As Long: If aArr Then la = LBound(a)
    Dim lb As Long: If bArr Then lb = LBound(b)
    Dim i As Long, x As Variant, y As Variant

    For i = 0 To n - 1
        If aArr Then x = a(la + i) Else x = a
        If bArr Then y = b(lb + i) Else y = b
        res(i) = scalarFunc2(fn, x, y)
    Next i
    applyFunc2 = res
End Function

'************************************************************************************************************************
' scalarFunc2
'   Purpose:
'       Performs the calculation of the defined functions with single arity (one input parameter)
'   Parameters:
'       fn: function code per enum RpnOp
'        x: parameter 1 used for the function
'        y: parameter 2 used for the function
'   Returns:
'       Scalar value of the function output
'---------------------------------------------------------------------------------
Private Function scalarFunc2(ByVal fn As Long, ByRef x As Variant, ByRef y As Variant) As Variant
    If VarType(x) <= vbNull Or VarType(y) <= vbNull Then
        scalarFunc2 = Null
        Exit Function
    End If

    Select Case fn
        Case FN_ROUND: scalarFunc2 = Round(x, CLng(y))
    End Select
End Function

'************************************************************************************************************************
' toColumn
'   Purpose:
'       Gets a 0-based array of the input
'   Parameters:
'       v: scalar or array of values to turn into a column
'       n: number of indexes in output array
'   Returns:
'       Array of values
'       If v was a scalar, the output array will have a constant value for all indexes
'       If v was an array, the output array will be a copy of the inpurt array, just ensuring it's 0-based
'---------------------------------------------------------------------------------
Private Function toColumn(ByRef v As Variant, ByVal n As Long) As Variant
    Dim res() As Variant: ReDim res(0 To n - 1)
    Dim i As Long
    
    ' If v is a scalar, return an array with each index having the constant, v
    If Not IsArray(v) Then
        For i = 0 To n - 1: res(i) = v: Next i
    
    ' If v is an array, return a 0-based index copy of the array, v
    Else
        Dim lb As Long: lb = LBound(v)
        For i = 0 To n - 1: res(i) = v(lb + i): Next i
    End If
    toColumn = res
End Function

'********************************************************************************************************************************************************************
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ SORTING FUNCTIONALITY /////////////////////////////////////////////////////////////////////////////////
'********************************************************************************************************************************************************************

' Sorts a Scripting.Dictionary of {colName: 1D array} in place, ordering ALL
' columns by the values in column `keyCol`. Uses a stable merge sort on an
' index permutation, then applies that permutation to every column.
Private Sub sortDictByColumn(ByVal dict As Object, ByVal keyCol As String, Optional ByVal ascending As Boolean = True)
    If Not dict.Exists(keyCol) Then pandasRaiseError "sortDictByColumn", PANDAS_ERROR_COL_NAME_NOT_EXISTS, keyCol
    
    Dim keyArr As Variant: keyArr = dict(keyCol)
    Dim lb As Long: lb = LBound(keyArr)
    Dim n As Long: n = UBound(keyArr) - lb + 1
    If n < 2 Then Exit Sub

    ' idx(i) = original (0-based) position of the item that belongs at slot i
    Dim idx() As Long: ReDim idx(0 To n - 1)
    Dim tmp() As Long: ReDim tmp(0 To n - 1)
    
    Dim i As Long
    For i = 0 To n - 1: idx(i) = i: Next i
    
    mergeSortIdx keyArr, lb, idx, tmp, 0, n - 1, ascending

    ' Apply the permutation to every column
    Dim colName As Variant
    For Each colName In dict.keys
        Dim src As Variant: src = dict(colName)
        
        Dim clb As Long: clb = LBound(src)
        
        Dim dst() As Variant: ReDim dst(clb To clb + n - 1)
        
        For i = 0 To n - 1
            dst(clb + i) = src(clb + idx(i))
        Next i
        
        dict(colName) = dst
    Next colName
End Sub

Private Sub mergeSortIdx(ByRef keys As Variant, ByVal lb As Long, ByRef idx() As Long, ByRef tmp() As Long, ByVal lo As Long, ByVal hi As Long, ByVal ascending As Boolean)
    If lo >= hi Then Exit Sub

    Dim middle As Long: middle = (lo + hi) \ 2
    mergeSortIdx keys, lb, idx, tmp, lo, middle, ascending
    mergeSortIdx keys, lb, idx, tmp, middle + 1, hi, ascending

    Dim i As Long: i = lo
    Dim j As Long: j = middle + 1
    Dim k As Long: k = lo

    Do While i <= middle And j <= hi
        Dim c As Long: c = compareValues(keys(lb + idx(i)), keys(lb + idx(j)))
        
        If Not ascending Then c = -c
        If c <= 0 Then              ' <= keeps the sort stable
            tmp(k) = idx(i): i = i + 1
        Else
            tmp(k) = idx(j): j = j + 1
        End If
        
        k = k + 1
    Loop
    
    Do While i <= middle: tmp(k) = idx(i): i = i + 1: k = k + 1: Loop
    Do While j <= hi:  tmp(k) = idx(j): j = j + 1: k = k + 1: Loop

    For k = lo To hi: idx(k) = tmp(k): Next k
End Sub

' Returns -1, 0, 1. Order: Empty/Null < numbers/dates < text (case-insensitive).
Private Function compareValues(ByVal a As Variant, ByVal b As Variant) As Long
    Dim ra As Long: ra = TypeRank(a)
    Dim rb As Long: rb = TypeRank(b)
    
    If ra <> rb Then
        compareValues = IIf(ra < rb, -1, 1)
    ElseIf ra = 0 Then
        compareValues = 0
    ElseIf ra = 1 Then
        If a < b Then
            compareValues = -1
        ElseIf a > b Then
            compareValues = 1
        End If
    Else
        compareValues = StrComp(CStr(a), CStr(b), vbTextCompare)
    End If
End Function

Private Function TypeRank(ByVal v As Variant) As Long
    Select Case VarType(v)
        Case vbEmpty, vbNull: TypeRank = 0
        Case vbString:        TypeRank = 2
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbDecimal, vbByte, vbBoolean
            TypeRank = 1
        Case Else:            TypeRank = 2   ' errors, etc. compared as text
    End Select
End Function

'********************************************************************************************************************************************************************
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ MISC HELPER FUNCTIONS /////////////////////////////////////////////////////////////////////////////////
'********************************************************************************************************************************************************************
Private Function getArrayDataType(ByRef arr As Variant) As String
    If Not IsArray(arr) Then
        getArrayDataType = getVarTypeName(vbEmpty)
        Exit Function
    End If
        
    Dim i As Long
    Dim t As Long
    Dim result As Long: result = vbEmpty   ' vbEmpty = "nothing seen yet"

    For i = LBound(arr) To UBound(arr)
        t = VarType(arr(i))

        ' Skip blanks so a few empty cells don't make the column "mixed"
        If t <> vbEmpty And t <> vbNull Then
            If result = vbEmpty Then
                result = t
            ElseIf t <> result Then
                If IsNumericType(t) And IsNumericType(result) Then
                    result = vbDouble          ' e.g. Long + Double -> Double
                Else
                    getArrayDataType = getVarTypeName(vbString)   ' genuinely mixed, call it a string
                    Exit Function
                End If
            End If
        End If
    Next i

    getArrayDataType = getVarTypeName(result)   ' vbEmpty if the array was all blank
End Function

Private Function IsNumericType(ByVal t As Long) As Boolean
    Select Case t
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            IsNumericType = True
    End Select
End Function

Private Function getVarTypeName(ByVal t As Long) As String
    Dim isArr As Boolean
    If t And vbArray Then
        isArr = True
        t = t And Not vbArray
    End If

    Dim s As String
    Select Case t
        Case vbEmpty:      s = "Empty"
        Case vbNull:       s = "Null"
        Case vbInteger:    s = "Integer"
        Case vbLong:       s = "Long"
        Case vbSingle:     s = "Single"
        Case vbDouble:     s = "Double"
        Case vbCurrency:   s = "Currency"
        Case vbDate:       s = "Date"
        Case vbString:     s = "String"
        Case vbObject:     s = "Object"
        Case vbError:      s = "Error"
        Case vbBoolean:    s = "Boolean"
        Case vbVariant:    s = "Variant"
        Case vbDataObject: s = "DataObject"
        Case vbDecimal:    s = "Decimal"
        Case vbByte:       s = "Byte"
        Case vbLongLong:   s = "LongLong"
        Case Else:         s = "Unknown(" & t & ")"
    End Select

    If isArr Then s = s & "()"
    getVarTypeName = s
End Function

Private Function isMissingVal(ByRef v As Variant) As Boolean
    Select Case VarType(v)
        Case vbEmpty, vbNull, vbError
            isMissingVal = True
        Case vbString
            isMissingVal = (LenB(Trim$(v)) = 0)
    End Select
End Function

Private Function getArrWithNormalizedMissing(ByRef arr As Variant) As Variant
    If isEmptyArray(arr) Then GoTo endFunct
    
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If isMissingVal(arr(i)) Or IsError(arr(i)) Then arr(i) = PANDAS_VALUES_ERROR
    Next i
    
endFunct:
    getArrWithNormalizedMissing = arr
End Function

'************************************************************************************************************************
' isValidColumnName
'   Purpose:
'       Check if the headerName is valid
'       A column name is invalid if it:
'           a) has any of the forbidden characters in the string PANDAS_VALUES_FORBIDDEN_COL_CHARS
'   Parameter:
'       str: String to check for the forbidden characters
'---------------------------------------------------------------------------------
Private Function isValidColumnName(ByVal headerName As String) As Boolean
    Dim arrForbiddenChars As Variant:  arrForbiddenChars = split(PANDAS_VALUES_FORBIDDEN_COL_CHARS, "\")

    If isEmptyArray(arrForbiddenChars) Then
        isValidColumnName = InStr(1, headerName, arrForbiddenChars) > 0
    Else
        Dim idx As Long
        For idx = LBound(arrForbiddenChars) To UBound(arrForbiddenChars)
            If InStr(1, headerName, arrForbiddenChars(idx)) > 0 Then isValidColumnName = False: Exit Function
        Next idx
        isValidColumnName = True
    End If
End Function

'************************************************************************************************************************
' getIndexOfDictKey
'   Purpose:
'       Gets the index of a key in a dictionary
'   Parameter:
'           dict: dictionary
'       dictKey : dictionary key to get index of
'   Returns:
'       Index of the dictionary key
'---------------------------------------------------------------------------------
Private Function getIndexOfDictKey(ByVal dict As Scripting.Dictionary, ByVal dictKey As String) As Long
    Dim idx As Long
    For idx = LBound(dict.keys) To UBound(dict.keys)
        If dict.keys(idx) = dictKey Then
            getIndexOfDictKey = idx
            Exit Function
        End If
    Next idx
    getIndexOfDictKey = -1
End Function

'************************************************************************************************************************
' print_dict
'   Purpose:
'       Print the inputted dictionary to the debug screen
'   Parameter:
'          dict: dictionary to print
'       colSize: number of characters to display in each column when excluding border
'---------------------------------------------------------------------------------
Private Sub print_dict(ByVal dict As Scripting.Dictionary, Optional ByVal colSize As Long = 20, Optional ByVal dictColTypes As Scripting.Dictionary)
    Dim allText As String
    Dim bufferedColSize As Long: bufferedColSize = colSize + 1
    Dim separatingLine As String: separatingLine = "|"

    Dim colTruncated As Boolean

    Dim topBorderLine As String: topBorderLine = "|"
    Dim headerRow As String: headerRow = "|"
    Dim colTypesRow As String: colTypesRow = "|"

    ' Print the column names
    Dim col As Variant
    For Each col In dict
        If Len(col) > colSize Then colTruncated = True

        topBorderLine = topBorderLine & String(bufferedColSize, "-")
        headerRow = headerRow & getCenteredText(bufferedColSize, left$(col, colSize))
        If Not dictColTypes Is Nothing Then colTypesRow = colTypesRow & getCenteredText(bufferedColSize, "[" & left$(dictColTypes(col), colSize - 2) & "]")
        
        separatingLine = separatingLine & String(bufferedColSize - 1, "-") & "+"

    Next col

    ' Change the last "+" border mark to a "|"
    separatingLine = mid$(separatingLine, 1, Len(separatingLine) - 1) & "|"
    topBorderLine = mid$(topBorderLine, 1, Len(topBorderLine) - 1) & "|"
    If dictColTypes Is Nothing Then
        allText = allText & topBorderLine & vbNewLine & headerRow & vbNewLine & separatingLine
    Else
        allText = allText & topBorderLine & vbNewLine & headerRow & vbNewLine & colTypesRow & vbNewLine & separatingLine
    End If
    
    ' Exit without looping thru the rows if there isn't any row data
    If isEmptyArray(dict(dict.keys(0))) Then GoTo endSub

    ' Print the row values
    Dim rowIdx As Long
    For rowIdx = 0 To UBound(dict(dict.keys(0)))
        Dim rowText As String: rowText = "|"

        For Each col In dict
            Dim cellValue As String
            If IsNull(dict(col)(rowIdx)) Then
                cellValue = "Null"
            Else
                cellValue = dict(col)(rowIdx)
            End If
            If Len(cellValue) > colSize Then colTruncated = True
            rowText = rowText & getCenteredText(bufferedColSize, left$(cellValue, colSize))
        Next col

        allText = allText & vbNewLine & rowText

    Next rowIdx
    allText = allText & vbNewLine & separatingLine

endSub:
    If colTruncated Then allText = allText & vbNewLine & "Cell values with greater than colSize=" & colSize & " characters were truncated"
    allText = allText & vbNewLine
    Debug.Print allText
End Sub

Private Function getCenteredText(ByVal totalSpace As Long, ByVal text As String) As String
    Dim numSpacesLeft As Long: numSpacesLeft = (totalSpace - Len(text)) \ 2

    Dim numSpacesRight As Long: numSpacesRight = totalSpace - Len(text) - numSpacesLeft - 1
    
    getCenteredText = Space(numSpacesLeft) & text & Space(numSpacesRight) & "|"
    
End Function

' Gets a 1D array from a range (assumes range is only 1 column wide)
Private Function getArr1dFromRng(ByVal rng As Range) As Variant
    Dim arr As Variant
    Dim result() As Variant
    Dim i As Long

    'Load the 2D array from the range
    arr = rng.Value

    'Resize the 1D array
    ReDim result(0 To UBound(arr, 1) - 1)

    'Copy values from column 1 into the 1D array
    For i = 1 To UBound(arr, 1)
        If IsError(arr(i, 1)) Then arr(i, 1) = PANDAS_VALUES_ERROR
        result(i - 1) = arr(i, 1)
    Next i

    getArr1dFromRng = result
End Function

Private Function getArrRngBtNums(ByVal num1 As Long, ByVal num2 As Long) As Variant
    Dim lbNum As Long: lbNum = WorksheetFunction.Min(num1, num2)
    Dim ubNum As Long: ubNum = WorksheetFunction.max(num1, num2)

    Dim arr As Variant: ReDim arr(0 To ubNum - lbNum)
    Dim arrIdx As Long: arrIdx = 0
    Dim num As Long
    For num = lbNum To ubNum
        arr(arrIdx) = num
        arrIdx = arrIdx + 1
    Next num
    
    getArrRngBtNums = arr
End Function

Private Function getMergedArr(ParamArray arrs() As Variant) As Variant
    Dim numIdxs As Long

    Dim arr As Variant
    For Each arr In arrs
        If Not isEmptyArray(arr) Then
            numIdxs = numIdxs + UBound(arr) - LBound(arr) + 1
        End If
    Next arr
    
    Dim arrMerged As Variant: ReDim arrMerged(0 To numIdxs - 1)
    
    Dim idxMerged As Long: idxMerged = 0
    For Each arr In arrs
        If isEmptyArray(arr) Then GoTo nextArr

        Dim lb As Long: lb = LBound(arr)
        Dim ub As Long: ub = UBound(arr)
                
        Dim idxCurr As Long
        For idxCurr = lb To ub
            arrMerged(idxMerged) = arr(idxCurr)
            idxMerged = idxMerged + 1
        Next idxCurr
nextArr:
    Next arr
        
    getMergedArr = arrMerged
End Function

Private Function isEmptyArray(ByVal arr As Variant) As Boolean
    ' Not an array? Consider it empty.
    If Not IsArray(arr) Then
        isEmptyArray = True
        Exit Function
    End If

    ' Uninitialized dynamic array ? UBound throws error
    On Error Resume Next
    Dim lb As Long: lb = LBound(arr)
    Dim ub As Long: ub = UBound(arr)
    
    If Err.Number <> 0 Then
        ' Error = uninitialized or invalid array
        isEmptyArray = True
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    ' Zero-length array (e.g., ub = -1)
    isEmptyArray = (ub < lb)

End Function

Private Function getDictCopy(ByVal dict As Scripting.Dictionary) As Scripting.Dictionary
    Dim dictCopy As Scripting.Dictionary: Set dictCopy = New Scripting.Dictionary
    
    Dim dictKey As Variant
    For Each dictKey In dict
        dictCopy(dictKey) = dict(dictKey)
    Next dictKey
    
    Set getDictCopy = dictCopy
End Function

' Check if testVal in range(lBoundVal, uBoundVal)
Private Function inValRange(ByVal testVal As Double, ByVal lBoundVal As Double, ByVal uBoundVal As Double) As Boolean
    inValRange = True
    If testVal > uBoundVal Or testVal < lBoundVal Then inValRange = False
End Function

'************************************************************************************************************************
' getArrExcel2DFromDict
'   Purpose:
'       Get a 2D array from a dictionary where the keys are column names and the values are rows
'   Parameter:
'       dict: dict with {colName: arrRowData}
'   Returns:
'       Array(row, col) w/ all data from dict
'---------------------------------------------------------------------------------
Private Function getArrExcel2DFromDict(ByVal dict As Scripting.Dictionary, Optional ByVal dropHeaders As Boolean = False) As Variant
    Dim numCols As Long: numCols = dict.count
    If numCols = 0 Then
        getArrExcel2DFromDict = Empty
        Exit Function
    End If

    Dim arrColNames As Variant: arrColNames = dict.keys

    ' Cache each column once
    Dim arrCols() As Variant: ReDim arrCols(0 To numCols - 1)
    Dim colIdx As Long
    For colIdx = 0 To numCols - 1
        arrCols(colIdx) = dict(arrColNames(colIdx))
    Next colIdx

    If Not dropHeaders Then
    Dim numRows As Long: numRows = UBound(arrCols(0)) - LBound(arrCols(0)) + 1
    
        ' Row 0 = headers, rows 1..numRows = data
        Dim arrOut() As Variant: ReDim arrOut(0 To numRows, 0 To numCols - 1)
    
        For colIdx = 0 To numCols - 1
            arrOut(0, colIdx) = arrColNames(colIdx)
        Next colIdx
    End If

    ' Headers only if there's no data. If no headers, then no data.
    If numRows = 0 Then
        getArrExcel2DFromDict = arrOut
        Exit Function
    End If

    Dim r As Long, lb As Long, col As Variant
    For colIdx = 0 To numCols - 1
        col = arrCols(colIdx)
        lb = LBound(col)
        For r = 0 To numRows - 1
            arrOut(r + 1, colIdx) = col(lb + r)
        Next r
    Next colIdx

    getArrExcel2DFromDict = arrOut
End Function

'************************************************************************************************************************
' writeArrToDelimited
'   Purpose:
'       Writes all values stored in arr to a delimited file
'   Parameters:
'       filePath: path to the csv file to store ranges
'       arr     : array of values to write to csv
'---------------------------------------------------------------------------------
Private Sub writeArrToDelimited(ByVal filePath As String, ByRef arr As Variant, Optional ByVal delimiter As String = ",")
    Dim rL As Long: rL = LBound(arr, 1)
    Dim cL As Long: cL = LBound(arr, 2)
    Dim rU As Long: rU = UBound(arr, 1)
    Dim cU As Long: cU = UBound(arr, 2)
    
    Dim numRows As Long: numRows = rU - rL + 1
    Dim numCols As Long: numCols = cU - cL + 1
    
    Dim lines() As String: ReDim lines(1 To numRows)
    Dim buff() As String: ReDim buff(1 To numCols)

    Dim r As Long, c As Long
    For r = rL To rU
        For c = cL To cU
            buff(1 + c - cL) = arr(r, c)
        Next c

        lines(1 + r - rL) = Join(buff, delimiter)

    Next r
    
    Dim ff As Long: ff = FreeFile
    Open filePath For Output As #ff
    Print #ff, Join(lines, vbCrLf)
    Close #ff

End Sub

'************************************************************************************************************************
' writeCSVToArray2d
'   Purpose:
'       Write the contents of a csv file to a 2D array
'   Parameter:
'       filePath: path to csv to create arr of
'---------------------------------------------------------------------------------
Private Function getArr2dFromDelimited(ByVal filePath As String, Optional ByVal delimiter As String = ",") As Variant
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    If Not fso.FileExists(filePath) Then pandasRaiseError getArr2dFromDelimited, PANDAS_ERROR_FILE_NOT_FOUND, filePath

    ' Store file in memory
    Open filePath For Binary As #1
        Dim textData As String: textData = Space$(LOF(1))
        Get #1, , textData
    Close #1
    
    ' Remove UTF-8 BOM if present
    If left$(textData, 3) = Chr$(239) & Chr$(187) & Chr$(191) Then
        textData = mid$(textData, 4)
    End If
    
    Dim lines() As String: lines = split(textData, vbCrLf)
    
    Dim numRows As Long: numRows = UBound(lines) + 1
    
    Dim parts() As String: parts = split(lines(0), delimiter)
    
    Dim numCols As Long: numCols = UBound(parts) + 1
    
    Dim arr2d() As Variant: ReDim arr2d(0 To numRows - 1, 0 To numCols - 1)
    
    Dim r As Long
    For r = 0 To numRows - 1
        parts = split(lines(r), delimiter)
        Dim c As Long
        For c = 0 To numCols - 1
            If c <= UBound(parts) Then
                arr2d(r, c) = parts(c)
            Else
                arr2d(r, c) = vbNullString
            End If
        Next
    Next

    getArr2dFromDelimited = arr2d
End Function

Private Function getDictFromArr2d(ByVal arr As Variant, Optional ByVal headers As Boolean = True) As Scripting.Dictionary
    Dim dict As Scripting.Dictionary: Set dict = New Scripting.Dictionary

    ' Determine bounds
    Dim colCount As Long: colCount = UBound(arr, 2)
    Dim rowCount As Long: rowCount = UBound(arr, 1)
    
    Dim colArray() As Variant   ' array to store column data
    Dim startRowData As Long    ' index of first row of data excluding headers
    If headers Then
        ReDim colArray(0 To rowCount - 1)
        startRowData = 1
    Else
        ReDim colArray(0 To rowCount)
        startRowData = 0
    End If
    
    ' Create an array for each column
    Dim col As Long
    For col = 0 To colCount
        Dim colName As String
        If headers Then
            colName = CStr(arr(0, col))
        Else
            colName = "Column_" & col
        End If
        
        Dim row As Long
        For row = startRowData To rowCount
            colArray(row - startRowData) = arr(row, col)
        Next row
        
        dict(colName) = colArray
    Next col

    Set getDictFromArr2d = dict
End Function

Private Function splitAtAssignment(ByVal expression As String) As Variant
    Dim pos As Long: pos = InStr(expression, "=")
    ' Skip double-equals
    If mid$(expression, pos + 1, 1) = "=" Then
        ' Find next "=" that isn't doubled
        Do
            pos = InStr(pos + 1, expression, "=")
            If pos = 0 Then Exit Do
        
        Loop While mid$(expression, pos + 1, 1) = "="
    
    End If
    
    Dim leftStr As String: leftStr = Trim$(left$(expression, pos - 1))
    Dim rightStr As String: rightStr = Trim$(mid$(expression, pos + 1))

    splitAtAssignment = Array(leftStr, rightStr)
    
End Function


