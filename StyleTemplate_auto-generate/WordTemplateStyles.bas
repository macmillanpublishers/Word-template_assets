Attribute VB_Name = "WordTemplateStyles"
Option Explicit
Dim rngList As Range

' Subs and Functions in this module require:
' - JsonConverter module be loaded in the same project
' - Dictionary Class module be loaded in the same project

Private Function get24bitColorXLS(p_strRGB) As String
    ' An Excel version of a macro we are using for color conversion for outputting styles from json
    Dim arrRGB() As String
    Dim lngBitColor As Long
    
    'split rgb string into array
    arrRGB = Split(p_strRGB, ",")
    
    'use RGB function to get 24Bit color value
    lngBitColor = RGB(CInt(arrRGB(0)), CInt(arrRGB(1)), CInt(arrRGB(2)))
    get24bitColorXLS = lngBitColor
    
End Function

Private Function CalcTargetRange(p_strFindString As String, p_lngHeaderRow As Long) As Range
    ' This takes a search term and a number value for header row, it searches the header row
    ' and returns a range, for any columns where the searchstring matched the header, from the
    ' cell below the header row to the last cell in use + 100 rows (for room to grow at the bottom)
    Dim lngLastColumn As Long
    Dim rngMyRange As Range
    Dim rngCell As Range
    Dim rngFoundRange As Range
    Dim rngNewRange As Range
    Dim arrFoundArray() As String
    Dim lngN As Long
    Dim rngTotalRange As Range
    Dim lngRowsUsed As Long
    Dim shtMainSheet As Worksheet
    
    Workbooks("WordTemplateStyles.xlsm").Activate
    Worksheets("Styles").Activate
    
    lngN = 1
    Set shtMainSheet = Excel.ActiveWorkbook.Sheets("Styles")
    Excel.ActiveWorkbook.Sheets("Styles").UsedRange     'refresh used range
    lngLastColumn = shtMainSheet.UsedRange.Columns.Count
    lngRowsUsed = shtMainSheet.UsedRange.Rows.Count
    
    ' set a range based on input param. for header row and columns in use
    Set rngMyRange = Range(Cells(p_lngHeaderRow, 1), Cells(p_lngHeaderRow, lngLastColumn))
    
    For Each rngCell In rngMyRange
        '  Search the header row for search string
        Set rngFoundRange = rngCell.Find(p_strFindString)
        ' Debug.Print rngFoundRange.Address
        ' Make sure we found something
        If Not rngFoundRange Is Nothing Then
            Debug.Print "Found search string at " & rngFoundRange.Column
            ' set found range, from header cell where string was found to cell from last row used in that column + 100 (room to grow)
            If lngN = 1 Then
                Set rngTotalRange = Range(Cells(4, rngFoundRange.Column), Cells(lngRowsUsed + 100, rngFoundRange.Column))
            ' merge any new found range with previous found range(s)
            ElseIf lngN > 1 Then
                Set rngNewRange = Range(Cells(4, rngFoundRange.Column), Cells(lngRowsUsed + 100, rngFoundRange.Column))
                Set rngTotalRange = Union(rngTotalRange, rngNewRange)
            End If
            lngN = lngN + 1
        End If
        ' reset found range to nuttin'
        Set rngFoundRange = Nothing
    Next
     
    Set CalcTargetRange = rngTotalRange

End Function

Function getColumnByHeadingValue(p_strFindString As String, p_lngHeaderRow As Long) As Long
    ' Returns index # for a column with Contents in a given row,
    ' exactly matching given string
        
    getColumnByHeadingValue = Sheet1.Cells(p_lngHeaderRow, 1).EntireRow.Find(What:=p_strFindString, LookIn:=xlValues, LookAt:=xlWhole).Column
  
End Function

Public Sub UpdateRGBsamples()
Attribute UpdateRGBsamples.VB_ProcData.VB_Invoke_Func = "u\n14"
    ' This is set for use with a shortcut-key (ctrl-u for PC) to update rgb previews for the style worksheet
    ' "u" for "update colors!"
    
    Call ChangeColorRange(CalcTargetRange("color", 3), False)
    Call ChangeColorRange(CalcTargetRange("TextColor", 3), True)

End Sub

Public Sub applyDataValidations()
    ' Setup validation for a number of columns, based on search strings
    ' in rows 2 & 3, with a variety of validation types
    
    Dim shtMainSheet As Worksheet
    Dim lngRowsUsed As Long
    Dim rngTF As Range
    Dim rngTF_B As Range
    Dim rngTF_C As Range
    Dim rngNextPara As Range
    Dim rngType As Range
    Dim rngPoints As Range
    Dim rngLineStyle As Range
    Dim rngOutline As Range
    Dim rngParaAlign As Range
    Dim rngParaSpaceRule As Range
    Dim rngLineWidth As Range
    Dim rngBaseStyle As Range
    Dim rngLLNum As Range
    Dim rngPriority As Range
    Dim rngWdKey As Range
    Dim rngFontNames As Range
    Dim rngStyleNameAndCode As Range
    
    Workbooks("WordTemplateStyles.xlsm").Activate
    Worksheets("Styles").Activate
    
    Set shtMainSheet = Excel.ActiveWorkbook.Sheets("Styles")
    Excel.ActiveWorkbook.Sheets("Styles").UsedRange
    lngRowsUsed = shtMainSheet.UsedRange.Rows.Count
    
    ' Set validation ranges
    Set rngTF = CalcTargetRange("TRUE / FALSE", 2)
    Set rngTF_B = CalcTargetRange("TRUE/FALSE", 2)
    Set rngNextPara = CalcTargetRange("NextParagraphStyle", 3)
    Set rngType = CalcTargetRange("1 is para, 2 is span", 2)
    Set rngPoints = CalcTargetRange("unit is points", 2)
    Set rngLineStyle = CalcTargetRange("LineStyle", 3)
    Set rngOutline = CalcTargetRange("OutlineLevel", 3)
    Set rngParaAlign = CalcTargetRange("ParagraphFormat.Alignment", 3)
    Set rngParaSpaceRule = CalcTargetRange("ParagraphFormat.LineSpacingRule", 3)
    Set rngLineWidth = CalcTargetRange("LineWidth", 3)
    Set rngBaseStyle = CalcTargetRange("BaseStyle", 3)
    Set rngLLNum = CalcTargetRange("ListLevelNumber", 3)
    Set rngPriority = CalcTargetRange("Priority", 3)
    Set rngTF_C = CalcTargetRange("_tf", 3)
    Set rngWdKey = CalcTargetRange("shortcut_keys__letter", 3)
    Set rngFontNames = CalcTargetRange("Font.Name", 3)
    Set rngStyleNameAndCode = CalcTargetRange("Style_", 3)
    
    ActiveSheet.Unprotect
    Application.ScreenUpdating = False
    
    'Reset existing validations!
    Cells.Validation.Delete
    
    ' Validations were easy to line up by recording validations to test
    ' Apply True/false validation
    With rngTF.Validation
    .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=validation_menus!$A$2:$A$3"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    With rngTF_B.Validation
    .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=validation_menus!$A$2:$A$3"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    With rngTF_C.Validation
    .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=validation_menus!$A$2:$A$3"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    'Apply NextParagraph validation - any existing style name
    With rngNextPara.Validation
    .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Styles!$A$4:$A$" & lngRowsUsed
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    'Apply type validation : accept "1" or "2"
    With rngType.Validation
    .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="1", Formula2:="2"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    'Apply points validation - any number >= 0
    With rngPoints.Validation
    .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator _
        :=xlGreaterEqual, Formula1:="0"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    'Apply outlinelevel validation - any number between 1 & 10
    With rngOutline.Validation
    .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="1", Formula2:="10"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    'Apply linestyle enumeration validation - any int between 0 & 24
    With rngLineStyle.Validation
    .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="0", Formula2:="24"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    'Apply para Alignment enumeration validation - any int between 0 & 9
    With rngParaAlign.Validation
    .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="0", Formula2:="9"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    'Apply para spacing enumeration validation - any int between 0 & 5
    With rngParaSpaceRule.Validation
    .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="0", Formula2:="5"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    'Apply line width validation - refer to range on validation_menus sheet
    With rngLineWidth.Validation
    .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=validation_menus!$B$2:$B$10"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    'Apply basestyle validation - refer to range on validation_menus sheet
    With rngBaseStyle.Validation
    .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=validation_menus!$C$2:$C$3"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    'Apply priority validation - any integer greater than 0
    With rngPriority.Validation
    .Delete
'        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
'        Operator:=xlBetween, Formula1:="1", Formula2:="2"
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator _
        :=xlGreaterEqual, Formula1:="1"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    'Apply llnumber enumeration validation -must be 0 I think?
    With rngLLNum.Validation
    .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="0", Formula2:="0"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    'Apply wdkey enumeration validation -refer to range on validation_menus sheet
    With rngWdKey.Validation
    .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=validation_menus!$D$2:$D$95"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    'Apply wdkey enumeration validation - refer to range on validation_menus sheet
    With rngWdKey.Validation
    .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=validation_menus!$D$2:$D$95"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    'Apply fontName validation - refer to range on validation_menus sheet
    With rngFontNames.Validation
    .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=validation_menus!$E$2:$E$8"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    'Apply validation for unique values for Style Names & Codes.  These are hard-coded to specific columns b/c alternatives seemed onerous
    With rngStyleNameAndCode.Validation
    .Delete
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=COUNTIF($C:$D, C4)<=1"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    Application.ScreenUpdating = True
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows _
        :=True, AllowDeletingColumns:=True, AllowDeletingRows:=True, AllowSorting _
        :=True, AllowFiltering:=True

End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
    ' This sub should automatically update cells in Excel for Windows..  not Mac :(
    
    If Target.Address = CalcTargetRange("color", 3) Then
        Call UpdateRGBsamples
    End If

End Sub
Private Function ChangeColorRange(p_rngTarget As Range, p_boolFontOnly As Boolean)
    Dim rngCell As Range
    Dim strRGB As String
    Dim strRGBsplit() As String
    Dim lngA As Long
    Dim lngB As Long
    Dim lngColor As Long
    Dim lngOppColor As Long
    
    For Each rngCell In p_rngTarget
        If Not IsEmpty(rngCell.Value) Then
            strRGB = rngCell.Value
            lngB = 0
            ' this whole next block is to validate our rgb string
            ' make sure we have a string that has two commas
            If Len(strRGB) - Len(Replace(strRGB, ",", "")) = 2 Then
                strRGBsplit = Split(strRGB, ",")
                For lngA = 0 To 2
                    ' make each string item  (split on commas) is numeric
                    If IsNumeric(strRGBsplit(lngA)) Then
                        ' make sure its in a valid range for an rgb value
                        If 0 <= strRGBsplit(lngA) < 256 Then
                            ' this conditional is to prep the split string array for use as "opposite color"
                            If strRGBsplit(lngA) > 125 Then
                                strRGBsplit(lngA) = 0
                            Else
                                strRGBsplit(lngA) = 255
                            End If
                            lngB = lngB + 1
                        End If
                    End If
                Next lngA
            End If
    '        Debug.Print rngCell.Value
    '        Debug.Print lngB
            If lngB = 3 Then
                ' get color values from rgb string & 'opposite' value from split string
                lngColor = get24bitColorXLS(strRGB)
                lngOppColor = get24bitColorXLS(Join(strRGBsplit, ","))
                ' set font colors for font-only column(s)
                If p_boolFontOnly Then
                    rngCell.Font.color = lngColor
                    If lngColor = 16777215 Then            'white font, make dark background
                        rngCell.Interior.color = 6579300
                    Else
                        rngCell.Interior.ColorIndex = 0
                    End If
                ' set interior color and "opposite color" for fonts, so they pop
                Else
                    rngCell.Interior.color = lngColor
                    rngCell.Font.color = lngOppColor
                End If
            End If
        End If
    Next

End Function
Private Function StylesColumnLoop(RowNum As Long, StartColumn As Long) As Dictionary
  ' This sub borrowed and pared down from Erica's for creating jsons frm excel
  ' Creates dictionary of row contents (key = column heading)
  Dim colCount As Long
  Dim strKey As String
  Dim strValue As String
  Dim dict_Return As Dictionary
  Set dict_Return = New Dictionary

  For colCount = StartColumn To rngList.Columns.Count

  ' key is always column header
    strKey = rngList.Cells(3, colCount).Value
    'make sure our key has a value
    If Not (strKey Like "no_export.*") And Not (strKey Like "html.*") Then
        Debug.Print strKey
        strValue = rngList.Cells(RowNum, colCount).Value
        Debug.Print strValue
        dict_Return.Item(strKey) = strValue
    End If
  Next colCount

  Set StylesColumnLoop = dict_Return
End Function


Private Function HTMLcolumnLoopB(ColNum As Long, StartRow As Long) As Collection
  ' This sub borrowed and pared down from Erica's for creating jsons frm excel
  ' Creates collection of Classes in 'Class' column based on "TRUE" markers in
  ' in a given column's cells
  ' HTMLcolumnLoopA does the nested dict version of this for top level heads
  Dim rowCount As Long
  Dim coll As New Collection
  Dim lngClassCol As Long
  
  ' Get the index # for column with contents in Row 2 exactly matching: "Class"
  lngClassCol = getColumnByHeadingValue("Class", 2)

  For rowCount = StartRow To rngList.Rows.Count
    If rngList.Cells(rowCount, ColNum).Value = True Then
        coll.Add "." & rngList.Cells(rowCount, lngClassCol).Value
    End If
  Next rowCount

  Set HTMLcolumnLoopB = coll
End Function



Private Function HTMLcolumnLoopA(ColNum As Long, StartRow As Long) As Dictionary
  ' This sub borrowed and pared down from Erica's for creating jsons frm excel
  ' Creates dictionary of column contents (key = ClassName)
  Dim rowCount As Long
  Dim strKey As String
  Dim strValue As String
  Dim dict_Return As Dictionary
  Set dict_Return = New Dictionary
    
  ' Get the index # for column with contents in Row 2 exactly matching: "Class"
  Dim lngClassCol As Long
  lngClassCol = getColumnByHeadingValue("Class", 2)

  For rowCount = StartRow To rngList.Rows.Count
    If rngList.Cells(rowCount, ColNum).Value <> vbNullString Then
        strKey = "." & rngList.Cells(rowCount, lngClassCol).Value
        strValue = rngList.Cells(rowCount, ColNum).Value
        dict_Return.Item(strKey) = strValue
    End If
  Next rowCount

  Set HTMLcolumnLoopA = dict_Return
End Function




Public Sub StylesToJSON(Optional p_boolUserInteract As Boolean = True)
  ' This sub borrowed and pared down from Erica's for creating jsons frm excel
    ' Creates an array of objects, each object uses header for key and
    ' one object for each row. Column headers are keys.
    ' "p_boolUserInteract" is so we can disable msgbox if autorunning via powershell
    
    Workbooks("WordTemplateStyles.xlsm").Activate
    Worksheets("Styles").Activate
    
    ActiveSheet.Unprotect
    Application.ScreenUpdating = False
    
    ' Get active range
    Range("A1").Activate
    Set rngList = ActiveCell.CurrentRegion
    
    ' Determine which sheet we're working with, set variables
    Dim strSheet As String
    strSheet = ActiveSheet.Name
    Debug.Print "strSheet is : " & strSheet
    
    Dim lngColStart As Long
    lngColStart = 2
    ' we're returning an object, not an array, so create dictionary
    Dim dict_Defaults As Dictionary
    Set dict_Defaults = New Dictionary

    ' Create dictionary to hold each record/row
    Dim dict_Record As Dictionary

    ' Loop through rows
    Dim rowCount As Long
    Dim lngIndex As Long
    Dim strKey1 As String

    ' Start at 4, header row, property info rows: don't count
    For rowCount = 4 To rngList.Rows.Count
        ' set the key to value in col A
        strKey1 = rngList.Cells(rowCount, 1).Value
        ' make sure our row has a value in A1
        If strKey1 <> vbNullString Then
            ' Loop through each column in row and write to Dictionary
            Set dict_Record = StylesColumnLoop(RowNum:=rowCount, StartColumn:=lngColStart)
            ' Add dictionary to array or dictionary
            Debug.Print strKey1
            Set dict_Defaults.Item(strKey1) = dict_Record
        End If

    Next rowCount

    ' Convert to json
    Dim strJson As String
    Dim fnum As Long
    Dim strPath As String

    strJson = JsonConverter.ConvertToJson(dict_Defaults, Whitespace:=2)

    ' Create output file path
    'strPath = ThisWorkbook.Path & Application.PathSeparator & strSheet & ".json"
    strPath = ThisWorkbook.Path & Application.PathSeparator & "macmillan.json"

    ' write string to file
    fnum = FreeFile
    ' creates the file if it doesn't exist, overwrites if it does
    Open strPath For Output Access Write As #fnum
    Print #fnum, strJson
    Close #fnum


    If p_boolUserInteract = True Then
        MsgBox "Done exporting to json!"
    End If
    
    Application.ScreenUpdating = True
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows _
        :=True, AllowDeletingColumns:=True, AllowDeletingRows:=True, AllowSorting _
        :=True, AllowFiltering:=True
    
End Sub

Public Sub HTMLmappingsToJSON(Optional p_boolUserInteract As Boolean = True)
  ' This sub borrowed and pared down from Erica's for creating jsons frm excel
  ' Scans for two types of columns: "html.toplevelheads"
  '   (writes a nested hash with toplevelheads as key)
  ' or others with values matching "html.*" (creates a hash with colname as key)
  ' Writes dict to json.
  ' "p_boolUserInteract" parameter is so we can disable msgbox if autorunning via powershell
    
    Workbooks("WordTemplateStyles.xlsm").Activate
    Worksheets("Styles").Activate
    
    ActiveSheet.Unprotect
    Application.ScreenUpdating = False
    
    ' Get active range
    Range("A1").Activate
    Set rngList = ActiveCell.CurrentRegion
    
    ' Determine which sheet we're working with, set variables
    Dim strSheet As String
    strSheet = ActiveSheet.Name
    Debug.Print "strSheet is : " & strSheet
    
    'set first row of non-header contents in sheet, & header row
    Dim lngHeader_KeyRow As Long
    Dim lngRowStart As Long
    lngRowStart = 4
    lngHeader_KeyRow = 3
    
    ' we're returning an object, not an array, so create dictionary
    Dim dict_Defaults As Dictionary
    Set dict_Defaults = New Dictionary

    ' Dim dict and collection for each record/row's hash or array- for dict "value"
    Dim dict_Record As Dictionary
    Dim coll_Record As Collection
    
    ' Loop through columns
    Dim colCount As Long
    Dim lngIndex As Long
    Dim strColHeader As String
    Dim strKey1 As String

    ' Populate our htmlmappings json with data from columns!
    For colCount = 1 To rngList.Columns.Count
        ' get column header name
        strColHeader = rngList.Cells(lngHeader_KeyRow, colCount).Value
        ' rename for use as key, without the "html." signifier
        strKey1 = Replace(strColHeader, "html.", "")
        
        ' find Col with header "toplevelheads" first - needs to be a nested hash (aka dict)
        If strColHeader = "html.toplevelheads" Then
            ' Loop through each row in column and write any values to Dictionary
            Set dict_Record = HTMLcolumnLoopA(ColNum:=colCount, StartRow:=lngRowStart)
            ' Add dictionary as value for toplevelheads Key
            Set dict_Defaults.Item(strKey1) = dict_Record
            
        'find all other Cols with "html.*- these are arrays so we use a collection"
        ElseIf strColHeader Like "html.*" Then
            ' Loop through each row in column and write to Collection
            Set coll_Record = HTMLcolumnLoopB(ColNum:=colCount, StartRow:=lngRowStart)
            ' Add collection as value for key
            Set dict_Defaults.Item(strKey1) = coll_Record
            
        End If
    Next colCount
    
    ' Add hardcoded key/value to json for footnotes
    Dim coll_footnoteSelector As New Collection
    coll_footnoteSelector.Add ("div.footnote")
    Set dict_Defaults.Item("footnotetextselector") = coll_footnoteSelector

    ' Convert to json
    Dim strJson As String
    Dim fnum As Long
    Dim strPath As String

    strJson = JsonConverter.ConvertToJson(dict_Defaults, Whitespace:=2)

    ' Create output file path
    strPath = ThisWorkbook.Path & Application.PathSeparator & "style_config.json"

    ' write string to file
    fnum = FreeFile
    ' creates the file if it doesn't exist, overwrites if it does
    Open strPath For Output Access Write As #fnum
    Print #fnum, strJson
    Close #fnum

    If p_boolUserInteract = True Then
        MsgBox "Done exporting to json!"
    End If

    Application.ScreenUpdating = True
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows _
        :=True, AllowDeletingColumns:=True, AllowDeletingRows:=True, AllowSorting _
        :=True, AllowFiltering:=True

End Sub


Private Sub autorun_HTMLmappingsToJSON()
    'So if we call this script from outside of excel, the toJson macro doesn't hang on the msgbox!
    Call HTMLmappingsToJSON(False)
End Sub

Public Sub WriteHTMLmappingsToJSON()
    ' This is so we can still run the HTMLmappings Macro directly from the "View Macros" menu-
    ' Even though its public it wasn't appearing b/c of its parameter
    Call HTMLmappingsToJSON
End Sub


Private Sub autorun_StylesToJSON()
    'So if we call this script from outside of excel, the toJson macro doesn't hang on the msgbox!
    Call StylesToJSON(False)
End Sub

Public Sub WriteStylesToJson()
    ' This is so we can still run the WriteStyles Macro directly from the "View Macros" menu-
    ' Even though its public it wasn't appearing b/c of its parameter
    Call StylesToJSON
End Sub

