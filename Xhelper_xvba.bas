Attribute VB_Name = "Xhelper"

'namespace=xvba_modules\Xhelper

'/*
'[Xhelper]
'@verion 1.0.0
'This Package has some helper functions for vba
'
'*/



'/*
'
'This Function calculate the array length for any array declaration
'
'@return {Array} arrayInfo 
' arrayInfo(1) returns dimension
' arrayInfo(2) returns array CountA
' arrayInfo(3) returns TypeName
'*/
Public Function ArrayLength(arr) 
    
  Dim xArryDim As Integer
  Dim yArryDim As Integer
  Dim arrayInfo(3) As Long

  xArryDim = 1
  yArryDim = 1


  Const ERROR_INVALID_DATA  As Long = vbObjectError + 513

  If (IsArray(arr)) Then
    xArryDim = UBound(arr,1) - LBound(arr,1) + 1
    On Error Resume Next 'For one dimension Arrays resume next
    yArryDim = UBound(arr,2) - LBound(arr,2) + 1
    
    arrayInfo(1) = xArryDim * yArryDim  
    arrayInfo(2) = Application.CountA(arr)
    arrayInfo(3) = TypeName(arr)
    
    length = arrayInfo

    Exit Function
  End If

  Err.Raise ERROR_INVALID_DATA, "lenth","The param must be an Array Type and not: " & TypeName(arr)

End Function


'/*
'
'Clear All Sheets Formulas
'
'@param {Array:String} ignoreSheets - List of sheets name to ignore clear formulas
'
'*/
Public Function clearFormulas(Optional ignoreSheets)

  Dim ws As Worksheet
  Set ws = ActiveSheet
  Dim ignore As Boolean
  For Each ws In Worksheets
   
    If IsMissing(ignoreSheets) Then
      Call clearFormulasHandler(ws)
    Else
      'Check ignored sheets
      ignore = checkIgnoreSheet(ignoreSheets,ws.Name)
      if(Not ignore) then
        Call clearFormulasHandler(ws)
      End If
    End If
  

  Next
End Function

'/*
'Handler for clear excel Sheet UsedRange formulas
'@ {Worksheet} ws
'*/
Private  Function clearFormulasHandler(ws As Worksheet)
  ws.Activate

  With ws.UsedRange
    .Copy
    .PasteSpecial Paste:=xlPasteValues, _
    Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
  End With
End Function


'/*
'
'Delete Sheets on worksheet
'
'@param {Array:String} ignoreSheets - List of sheets name to ignore clear formulas
'*/
Public Sub deleteSheets(Optional ignoreSheetsArray As Variant)
    
  Dim ws As Worksheet
  Set ws = ActiveSheet
  Dim ignoreSheet As Boolean
  ignoreSheet = FALSE
  Dim isMissingIgnoredSheets AS Boolean
  
  isMissingIgnoredSheets =  IsMissing(ignoreSheetsArray)

  For Each ws In Worksheets
    ws.Activate
    
    if (NOT isMissingIgnoredSheets) Then
      ignoreSheet = checkIgnoreSheet(ignoreSheetsArray ,ws.Name) 
    End If
   
            
    If (NOT ignoreSheet) Then
    
      Application.DisplayAlerts = False
      ws.Delete
      Application.DisplayAlerts = True
     
    End If

  Next
  Application.Sheets(1).Select
End Sub


'/*
'Check if the sheet (by name) is ignored
'@param {Variant} ignoreSheets Array of string
'@return{Boolean} ignore - Ture is the sheet is in ignored Array
'                        - False if Sheet is not in Ignore Array
'*/  
Private  Function checkIgnoreSheet(ignoreSheetsArray As Variant ,sheetName As String) As Boolean

  Dim name As Variant
  Dim  ignore As Boolean

  For Each name In ignoreSheetsArray
    ignore = InStr(1, name,sheetName, vbTextCompare) > 0
    If (ignore) Then Exit For
    Next

    checkIgnoreSheet =  ignore

End Function


'/*
'
'Handller For set Application States off
' DisplayAlerts
' Calculation
' ScreenUpdating
' EnableEvents
'
'@Param {Boolean} state = false
'
'*/
Public Function setAppUpdateState(Optional state As Boolean = False)

  Application.DisplayAlerts = state
  Application.ScreenUpdating = state
  Application.EnableEvents = state

  If (state) Then
    Application.Calculation = xlAutomatic
  else
    Application.Calculation = xlCalculationManual
  End If


End Function

'/*
' This FunctionDelete coloumns between NamedRanges
'
'
'*/
Public Function delColBetweenNameRanges(startNamedRange,endNamedRange)


  Dim startColumn As Variant
  Dim endColumn As Variant
    
  startColumn = Columns(Split(Replace(Split(Range(startNamedRange).Address, ":")(1), "$", "", 1, 1), "$")(0)).Column + 1
  endColumn = Columns(Split(Replace(Split(Range(endNamedRange).Address, ":")(0), "$", "", 1, 1), "$")(0)).Column - 1
   
  Columns(Split(Cells(1, startColumn).Address, "$")(1) & ":" & Split(Cells(1, endColumn).Address, "$")(1)).Delete Shift:=xlToLeft
   
  
End Function





'/*
' This Function Delete rows between top and Bottom NamedRanges
'
'
'*/
Public Function delRowsBetweenNameRanges(topNamedRange As String,bottomNamedRange As String)

  Dim startRow As Variant
  Dim endRow As Variant
    
  startRow = Split(Replace(Split(Range(topNamedRange).Address, ":")(1), "$", "", 1, 1), "$")(1) + 1
  On Error Resume Next
  endRow = Split(Replace(Split(Range(bottomNamedRange).Address, ":")(0), "$", "", 1, 1), "$")(1) - 1
  endRow = Replace(Split(Range(bottomNamedRange).Address, ":")(0), "$", "", 1, 1) - 1
 
  Rows(startRow & ":" & endRow).Delete Shift:=xlToLeft
      
End Function



  
'/*
' This Function Delete rows abouve NamedRanges
'
'
'*/
Public Function delRowsAboveNameRange(bottomNamedRange As String,Optional startRow As Integer = 1)

  Dim endRow As Variant

  endRow = Split(Replace(Split(Range(bottomNamedRange).Address, ":")(0), "$", "", 1, 1), "$")(1) - 1
 
  Rows(startRow & ":" & endRow).Delete Shift:=xlToLeft
      
End Function


'/*
' This Function Delete rows below NamedRanges
'
'
'*/
Public Function delRowsBelowNameRange(nameRange As String)

  Dim startRow As Variant
  Dim endRow As Variant
  
  endRow = Cells.SpecialCells(xlCellTypeLastCell).Row
  On Error Resume Next
  startRow = Split(Replace(Split(Range(nameRange).Address, ":")(1), "$", "", 1, 1), "$")(1) + 1
  startRow = Replace(Split(Range(nameRange).Address, ":")(1), "$", "", 1, 1) + 1
 
  Rows(startRow & ":" & endRow).Delete Shift:=xlToLeft
      
End Function



'/*
' This Function Delete coloumns beAfter  NamedRange
'
'@param {String} startNamedRange : Name Range
'
'*/
Public Function delColAfterNameRanges(startNamedRange)


  Dim startColumn As Variant
  Dim endColumn As Variant

  Dim sht As Worksheet
  Set sht = ThisWorkbook.ActiveSheet
  
     
  startColumn = Columns(Split(Replace(Split(Range(startNamedRange).Address, ":")(1), "$", "", 1, 1), "$")(0)).Column + 1
  'endColumn = Cells(1, Columns.Count).End(xlToLeft).Column
  endColumn = sht.UsedRange.Columns(sht.UsedRange.Columns.Count).Column
    
    
  Columns(Split(Cells(1, startColumn).Address, "$")(1) & ":" & Split(Cells(1, endColumn).Address, "$")(1)).Delete Shift:=xlToLeft
    
End Function


'/*
' This Function Delete coloumns before  NamedRange
'
'@param {String} namedRange : Name Range
'@param {Long} startColumn : First Column number
'
'*/
Public Function delColBeforeNameRange(namedRange,Optional startColumn As Long = 1)

  Dim endColumn As Variant
     
  endColumn = Columns(Split(Replace(Split(Range(namedRange).Address, ":")(0), "$", "", 1, 1), "$")(0)).Column - 1

  Columns(Split(Cells(1, startColumn).Address, "$")(1) & ":" & Split(Cells(1, endColumn).Address, "$")(1)).Delete Shift:=xlToLeft
    
End Function




'/*
'
'Delete Rows by check cell value condition
'
'@param {String} startRow : Start row for check the cell
'@param {Long} columnCheckEmpty [Optional = 1] the cell column for check isEmpty
'@param {Variant} condition [Optional = ""] the cell condition for check
'/*
Public Function delRowsByCheckCellValue(Optional ByVal startRow As Long = 1,Optional ByVal columnCheckEmpty As Long = 1,Optional ByVal condition As Variant = "")


  Dim rowNumber As Variant
  Dim cellValue As Variant
  Dim firstRow As Long
  firstRow = 0
  Dim lastRow As Long
  lastRow = 0

  Dim sheetLastRow As Long
  sheetLastRow =  Cells(Rows.Count,columnCheckEmpty).End(xlUp).Row


  For rowNumber = startRow To sheetLastRow


    cellValue =  Cells(rowNumber,columnCheckEmpty).Value

    If(cellValue = condition And firstRow = 0 ) then
      firstRow = rowNumber
    End IF

    if(firstRow <> 0 And cellValue<> condition And lastRow = 0) then
      lastRow = rowNumber
    End if

    if(firstRow<> 0 And  lastRow <> 0)Then

      Range(firstRow & ":" & lastRow).Delete Shift:=xlUp

      Call delRowsByCheckCellValue(firstRow, columnCheckEmpty,condition)
      
      firstRow = 0
      lastRow = 0
    End If

  Next rowNumber


End Function
    

Public Function delColumnsUntil(nameRange As String, Optional ByVal condition As Variant = 0)
 
  Dim sht As Worksheet
  Dim sheetLastCol As Long
  Set sht = ThisWorkbook.ActiveSheet
    
  sheetLastCol = sht.UsedRange.Columns(sht.UsedRange.Columns.Count).Column
  Dim lastCol As Long
  lastCol = sheetLastCol
  Dim deleted As Boolean
  deleted = False
  
  Dim startRow As Variant


  On Error Resume Next
  startRow = Split(Replace(Split(Range(nameRange).Address, ":")(1), "$", "", 1, 1), "$")(1) 
  startRow = Replace(Split(Range(nameRange).Address, ":")(1), "$", "", 1, 1) 

  While (sheetLastCol > 0 And deleted = False)
    If (Cells(startRow, sheetLastCol) > condition) Then
      Range(Cells(1, sheetLastCol+1).entirecolumn, Cells(1, lastCol).entirecolumn).Delete Shift:=xlToLeft
      deleted = True
    End If
    
    sheetLastCol = sheetLastCol - 1
  Wend
    
End Function


'/*
'
'Save the actual workbook state and create a TMP 
'workbook for process the changes
'
'*/
Public  Function saveWorkbookFileTemp(Optional tmpSaveNamed As String = "workbook_temp",Optional extension As String = "xlsb")

  Dim App As Application
  Set App = Application
  Dim fileTemp As String


  'Save ThisWorkbook
  App.ThisWorkbook.Save
  'Save this workbook on TEMP folder for make changes
  Dim tempFolder As String
  tempFolder = Environ("Temp")
  'Temp file
  fileTemp = tempFolder & "\" & tmpSaveNamed & "." &  extension
  'Clear Temp CPU Temp file if Exist
  On Error Resume Next
  Kill fileTemp
  'Save CPU Temp Files
  App.ThisWorkbook.SaveAs (fileTemp)

End Function


'/*
'
' This Function Save new workbook with without formulas in xlsx
'
'*/
Public  Function saveFileWithoutFormulas(filename As String,filePath As String)
  Dim App As Application
  Set App = Application
  'Save CPU file
  Application.DisplayAlerts = False
  Dim CpuFileName As String
  CpuFileName = filePath & "\" & filename & ".xlsx"
  App.ThisWorkbook.SaveAs Filename:=CpuFileName, FileFormat:=xlOpenXMLWorkbook
  App.ThisWorkbook.Save
   
End Function


'/*
' Test if the sheet exist on AsctiveWorkbook
'
'@param {String} name: Sheet Name
'
'@return {Boolean} True if exist
'                  False if not  
'*/
Public Function sheetExist(name As String)
  
  Dim check As Boolean
  check = False

  On Error Resume Next
  check =  (ActiveWorkbook.Sheets(name).Index > 0)
  
  sheetExist = check
  
End Function

'/*
'
' Get param from a existing sheet or return default param
'
'@param {String} sheetName : A sheet name where params are store
'@param {Long} row : row number 
'@param {Long} col : col number 
'@param {Variant} defaultValue : user defined default value number 
'
'*/
Function getSheetParam(sheetName As String,row, col, defaultValue)
  
  Dim checkSheet As Boolean
  Dim sheetValue As Variant
  sheetValue = ""

  checkSheet = Xhelper.sheetExist(sheetName)
  
  'If sheet exist get the value
  If (checkSheet) Then
    sheetValue = Application.Sheets(sheetName).Cells(row, col).Value
  End If
  'Return the value
  getSheetParam = IIf(sheetValue <> "", sheetValue, defaultValue)

End Function



'/*
' Delete Columns Range by check a Condition on cell 
'
'@param{Long}  startColumn:  Start column number for check the range
'@param{Long}  lastColumn:  Last column number for check range of columns
'@param{Long}  condition:  Condition to check for delete
'@param{Long}  rowConditonCheck:  Therow where to check the condition
'
'*/
Public  Function DeleteColumnsByCondition(startColumn,lastColumn,condition,rowConditonCheck)

  Dim colNumber As Long
  Dim firstColRange As Long
  Dim lasColRange As Long
  firstColRange = 0
  lasColRange = 0

  Dim col As String
  For colNumber = startColumn To lastColumn
    if(Cells(rowConditonCheck,colNumber).Value = condition And firstColRange = 0) Then
      firstColRange = colNumber
    End If

    if(Cells(rowConditonCheck,colNumber).Value <> condition And lasColRange = 0 And firstColRange <> 0) Then
      lasColRange = colNumber
    End If

    If( lasColRange <> 0 And firstColRange <> 0) Then

      
      Columns(getColumnLetter(firstColRange) & ":" & getColumnLetter(lasColRange - 1)).Delete Shift:=xlToLeft
    
      Call DeleteColumnsByCondition(colNumber,lastColumn - colNumber,condition,rowConditonCheck)
      lasColRange = 0 
      firstColRange = 0
    End IF
  Next colNumber

End Function

'/*
' Get column letter by giving column number
'
'@param {Long} colNumber: Column number
'@return {String} column letter
'*/
Public Function getColumnLetter(colNumber As Long) As String
  getColumnLetter = Split(Cells(1, colNumber).Address, "$")(1)
End Function

  
'/*
'This Function Return Range Columns and Rows Values
'
'@param {String} rangeName
'
'@return {Dictionary} rangeData : rangeData("START_COL_NB")  Start Column number
'                                 rangeData("START_COL")  Start Column Letter
'                                 rangeData("END_COL_NB")
'                                 rangeData("END_COL") 
'                                 rangeData("START_ROW")
'                                 rangeData("END_ROW")
'                                 rangeData("Address") 
'*/
Public Function getRangeAddress(rangeName As String) As Object

  
  Dim rangeData As Object
  Set rangeData = CreateObject("Scripting.Dictionary") 

  rangeData("Address") = Range(rangeName).Address
  
  Dim SplitedRangeAddress As Variant
  SplitedRangeAddress = Split(rangeData("Address"), ":")

  rangeData("START_COL_NB") = Columns(Split(Replace(SplitedRangeAddress(0), "$", "", 1, 1), "$")(0)).Column
  rangeData("START_COL") =  Split(Cells(1, rangeData("START_COL_NB")).Address, "$")(1)
  rangeData("END_COL_NB") = Columns(Split(Replace(SplitedRangeAddress(1), "$", "", 1, 1), "$")(0)).Column
  rangeData("END_COL") =  Split(Cells(1, rangeData("END_COL_NB")).Address, "$")(1)

  rangeData("START_ROW") = Split(Replace(SplitedRangeAddress(0), "$", "", 1, 1), "$")(1)
  rangeData("END_ROW") = Split(Replace(SplitedRangeAddress(1), "$", "", 1, 1), "$")(1)
  
  Set getRangeAddress = rangeData

End Function