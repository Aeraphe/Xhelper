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
   
 Columns(Split(Cells(1, startColumn).Address, "$")(1) & ":" & Split(Cells(1, endColumn).Address, "$")(1)).Select
 Selection.Delete Shift:=xlToLeft
   
  
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
 
  Rows(startRow & ":" & endRow).Select
  Selection.Delete Shift:=xlToLeft
      
 End Function



  
'/*
' This Function Delete rows abouve NamedRanges
'
'
'*/
Public Function delRowsAboveNameRange(bottomNamedRange As String,Optional startRow As Integer = 1)

  Dim endRow As Variant

  endRow = Split(Replace(Split(Range(bottomNamedRange).Address, ":")(0), "$", "", 1, 1), "$")(1) - 1
 
  Rows(startRow & ":" & endRow).Select
  Selection.Delete Shift:=xlToLeft
      
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
 
  Rows(startRow & ":" & endRow).Select
  Selection.Delete Shift:=xlToLeft
      
 End Function


