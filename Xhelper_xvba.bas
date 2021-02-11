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
    If (ignore) Then
      ignoreSheet =  ignore
      Exit Function
    End If
   
  Next

  ignoreSheet =  ignore

End Function