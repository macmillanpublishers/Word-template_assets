Attribute VB_Name = "StyleTemplateCreator"
Option Explicit
Option Base 1

' Subs and Functions in this module require:
' - JsonConverter module be loaded in the same project
' - Dictionary Class module be loaded in the same project
' - Microsoft Office & Word 15.0 Object Library References be enabled for Project
'   (under Tools>References)

Private Function CreateKeybindings(styleDict As Dictionary, sname As String, docTemplate As Document)
    ' apply keybindings based on entries from the styles json
    Dim boolShift As Boolean
    Dim boolCtrl As Boolean
    Dim boolAlt As Boolean
    Dim lngLetter As Long
    Dim objWdKeyDict As New Dictionary
    
    ' load wdKey Dict
    Set objWdKeyDict = SetupWdKeyDict
    ' set context for keybindings
    CustomizationContext = docTemplate
    
    boolShift = CBool(styleDict(sname).Item("shortcut_keys__shift_tf"))
    boolCtrl = CBool(styleDict(sname).Item("shortcut_keys__control_tf"))
    boolAlt = CBool(styleDict(sname).Item("shortcut_keys__alt_tf"))
    lngLetter = objWdKeyDict(styleDict(sname).Item("shortcut_keys__letter"))
    
    ' Apply keybindings!
    If Abs(CInt(boolShift) + CInt(boolCtrl) + CInt(boolAlt)) = 3 Then
        KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyControl, wdKeyShift, lngLetter), KeyCategory:=5, Command:=sname
    ElseIf Abs(CInt(boolShift) + CInt(boolCtrl) + CInt(boolAlt)) = 2 Then
        If boolShift = False Then
            KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, lngLetter), KeyCategory:=5, Command:=sname
        ElseIf boolCtrl = False Then
            KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyShift, wdKeyAlt, lngLetter), KeyCategory:=5, Command:=sname
        ElseIf boolAlt = False Then
            KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyShift, wdKeyControl, lngLetter), KeyCategory:=5, Command:=sname
        End If
    ElseIf Abs(CInt(boolShift) + CInt(boolCtrl) + CInt(boolAlt)) = 1 Then
        If boolShift = True Then
            KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyShift, lngLetter), KeyCategory:=5, Command:=sname
        ElseIf boolCtrl = True Then
            KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, lngLetter), KeyCategory:=5, Command:=sname
        ElseIf boolAlt = True Then
            KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, lngLetter), KeyCategory:=5, Command:=sname
        End If
    End If

End Function

Private Function SetupWdKeyDict()
    ' There is no good way to convert stored wdKey string to wdKey type in Word VBA
    ' Creating a dict to compare stored values to and get enumeration.
    Dim objWdKeyDict As New Dictionary
    
    objWdKeyDict.Add "wdKey0", "48"
    objWdKeyDict.Add "wdKey1", "49"
    objWdKeyDict.Add "wdKey2", "50"
    objWdKeyDict.Add "wdKey3", "51"
    objWdKeyDict.Add "wdKey4", "52"
    objWdKeyDict.Add "wdKey5", "53"
    objWdKeyDict.Add "wdKey6", "54"
    objWdKeyDict.Add "wdKey7", "55"
    objWdKeyDict.Add "wdKey8", "56"
    objWdKeyDict.Add "wdKey9", "57"
    objWdKeyDict.Add "wdKeyA", "65"
    objWdKeyDict.Add "wdKeyAlt", "1024"
    objWdKeyDict.Add "wdKeyB", "66"
    objWdKeyDict.Add "wdKeyBackSingleQuote", "192"
    objWdKeyDict.Add "wdKeyBackSlash", "220"
    objWdKeyDict.Add "wdKeyBackspace", "8"
    objWdKeyDict.Add "wdKeyC", "67"
    objWdKeyDict.Add "wdKeyCloseSquareBrace", "221"
    objWdKeyDict.Add "wdKeyComma", "188"
    objWdKeyDict.Add "wdKeyCommand", "512"
    objWdKeyDict.Add "wdKeyControl", "512"
    objWdKeyDict.Add "wdKeyD", "68"
    objWdKeyDict.Add "wdKeyDelete", "46"
    objWdKeyDict.Add "wdKeyE", "69"
    objWdKeyDict.Add "wdKeyEnd", "35"
    objWdKeyDict.Add "wdKeyEquals", "187"
    objWdKeyDict.Add "wdKeyEsc", "27"
    objWdKeyDict.Add "wdKeyF", "70"
    objWdKeyDict.Add "wdKeyF1", "112"
    objWdKeyDict.Add "wdKeyF10", "121"
    objWdKeyDict.Add "wdKeyF11", "122"
    objWdKeyDict.Add "wdKeyF12", "123"
    objWdKeyDict.Add "wdKeyF13", "124"
    objWdKeyDict.Add "wdKeyF14", "125"
    objWdKeyDict.Add "wdKeyF15", "126"
    objWdKeyDict.Add "wdKeyF16", "127"
    objWdKeyDict.Add "wdKeyF2", "113"
    objWdKeyDict.Add "wdKeyF3", "114"
    objWdKeyDict.Add "wdKeyF4", "115"
    objWdKeyDict.Add "wdKeyF5", "116"
    objWdKeyDict.Add "wdKeyF6", "117"
    objWdKeyDict.Add "wdKeyF7", "118"
    objWdKeyDict.Add "wdKeyF8", "119"
    objWdKeyDict.Add "wdKeyF9", "120"
    objWdKeyDict.Add "wdKeyG", "71"
    objWdKeyDict.Add "wdKeyH", "72"
    objWdKeyDict.Add "wdKeyHome", "36"
    objWdKeyDict.Add "wdKeyHyphen", "189"
    objWdKeyDict.Add "wdKeyI", "73"
    objWdKeyDict.Add "wdKeyInsert", "45"
    objWdKeyDict.Add "wdKeyJ", "74"
    objWdKeyDict.Add "wdKeyK", "75"
    objWdKeyDict.Add "wdKeyL", "76"
    objWdKeyDict.Add "wdKeyM", "77"
    objWdKeyDict.Add "wdKeyN", "78"
    objWdKeyDict.Add "wdKeyNumeric0", "96"
    objWdKeyDict.Add "wdKeyNumeric1", "97"
    objWdKeyDict.Add "wdKeyNumeric2", "98"
    objWdKeyDict.Add "wdKeyNumeric3", "99"
    objWdKeyDict.Add "wdKeyNumeric4", "100"
    objWdKeyDict.Add "wdKeyNumeric5", "101"
    objWdKeyDict.Add "wdKeyNumeric5Special", "12"
    objWdKeyDict.Add "wdKeyNumeric6", "102"
    objWdKeyDict.Add "wdKeyNumeric7", "103"
    objWdKeyDict.Add "wdKeyNumeric8", "104"
    objWdKeyDict.Add "wdKeyNumeric9", "105"
    objWdKeyDict.Add "wdKeyNumericAdd", "107"
    objWdKeyDict.Add "wdKeyNumericDecimal", "110"
    objWdKeyDict.Add "wdKeyNumericDivide", "111"
    objWdKeyDict.Add "wdKeyNumericMultiply", "106"
    objWdKeyDict.Add "wdKeyNumericSubtract", "109"
    objWdKeyDict.Add "wdKeyO", "79"
    objWdKeyDict.Add "wdKeyOpenSquareBrace", "219"
    objWdKeyDict.Add "wdKeyOption", "1024"
    objWdKeyDict.Add "wdKeyP", "80"
    objWdKeyDict.Add "wdKeyPageDown", "34"
    objWdKeyDict.Add "wdKeyPageUp", "33"
    objWdKeyDict.Add "wdKeyPause", "19"
    objWdKeyDict.Add "wdKeyPeriod", "190"
    objWdKeyDict.Add "wdKeyQ", "81"
    objWdKeyDict.Add "wdKeyR", "82"
    objWdKeyDict.Add "wdKeyReturn", "13"
    objWdKeyDict.Add "wdKeyS", "83"
    objWdKeyDict.Add "wdKeyScrollLock", "145"
    objWdKeyDict.Add "wdKeySemiColon", "186"
    objWdKeyDict.Add "wdKeyShift", "256"
    objWdKeyDict.Add "wdKeySingleQuote", "222"
    objWdKeyDict.Add "wdKeySlash", "191"
    objWdKeyDict.Add "wdKeySpacebar", "32"
    objWdKeyDict.Add "wdKeyT", "84"
    objWdKeyDict.Add "wdKeyTab", "9"
    objWdKeyDict.Add "wdKeyU", "85"
    objWdKeyDict.Add "wdKeyV", "86"
    objWdKeyDict.Add "wdKeyW", "87"
    objWdKeyDict.Add "wdKeyX", "88"
    objWdKeyDict.Add "wdKeyY", "89"
    objWdKeyDict.Add "wdKeyZ", "90"
    objWdKeyDict.Add "wdNoKey", "255"
    
    Set SetupWdKeyDict = objWdKeyDict
End Function

Private Function localReadTextFile(Path As String, Optional FirstLineOnly As Boolean = True) As String
  ' Local version of Erica's ReadTextFile fnction from GeneralHelpers module

    Dim fnum As Long
    Dim strTextWeWant As String
    
    fnum = FreeFile()
    Open Path For Input As fnum
    
    If FirstLineOnly = False Then
        strTextWeWant = Input$(LOF(fnum), #fnum)
    Else
        Line Input #fnum, strTextWeWant
    End If
    
    Close fnum
    
    localReadTextFile = strTextWeWant
  Exit Function

End Function

Private Function localReadJson(JsonPath As String) As Dictionary
  ' Local version of Erica's readJSON fnction from Class Helpers module
  Dim dictJson As Dictionary
  Dim strJson As String
  
  strJson = localReadTextFile(JsonPath, False)
  If strJson <> vbNullString Then
      Set dictJson = JsonConverter.ParseJson(strJson)
  Else
      ' If file exists but has no content, return empty dictionary
      Set dictJson = New Dictionary
  End If

  If dictJson Is Nothing Then
        Debug.Print "ReadJson fail"
  End If
  
  Set localReadJson = dictJson
  Exit Function
  
End Function

Public Sub ReadStylestoJson()
    ' NOTE: requirements for this sub are listed in comments under this moedule's "Declarations"
    ' This macro is to read all of the files from a Word Document
    ' It expcets some propoerties that are not available on Word for Mac
    ' So should be run on Word for PC
    ' It reads styles from the ActiveDocument & writes to a 'macmillan.json' file ..
    ' .. Located in the same dir as the active Document
    Dim dictStyle_dict As Dictionary
    Dim dictEmpty As Dictionary
    Dim styleThis_style As Style
    Dim strJson As String
    Dim strJsonPath As String
    Dim dictInner_dict As Dictionary
    Dim i As WdBorderType
    Dim a As Long
    Dim b As Long
    Dim k As Long
    Dim lngKeycode As Long
    Dim lngIncrement As Long
    
    'how often we save and undo.clear on doc
    lngIncrement = 25
    ' JSON file path
    strJsonPath = Word.ActiveDocument.Path & Application.PathSeparator & "macmillan.json"
    ' initialize dicts & counter
    Set dictStyle_dict = New Dictionary
    Set dictEmpty = New Dictionary
    b = 1
    ' set Context for Keybindings
    CustomizationContext = Word.ActiveDocument
    
    ' Create new document for template
    
    For a = 1 To Word.ActiveDocument.Styles.Count
        Set styleThis_style = Word.ActiveDocument.Styles(a)
        If styleThis_style.BuiltIn = False Then
            Debug.Print styleThis_style.NameLocal
            Set dictInner_dict = New Dictionary
            '''All Styles:
            'Top Level properties
            'dictInner_dict.Item("Description") = styleThis_style.Description
            dictInner_dict.Item("BaseStyle") = styleThis_style.BaseStyle
            dictInner_dict.Item("ListLevelNumber") = styleThis_style.ListLevelNumber
            dictInner_dict.Item("Priority") = styleThis_style.Priority
            dictInner_dict.Item("NoProofing") = CBool(styleThis_style.NoProofing)     'Endnote Text (ntx) NoProofing = -1
            dictInner_dict.Item("NextParagraphStyle") = styleThis_style.NextParagraphStyle
            dictInner_dict.Item("QuickStyle") = styleThis_style.QuickStyle
            dictInner_dict.Item("Type") = styleThis_style.Type           'Type 1 is paragraph, Type 2 is char
            'Font properties
            dictInner_dict.Item("Font.name") = styleThis_style.Font.Name
            dictInner_dict.Item("Font.Size") = styleThis_style.Font.Size
            dictInner_dict.Item("Font.Italic") = CBool(styleThis_style.Font.Italic)
            dictInner_dict.Item("Font.Bold") = CBool(styleThis_style.Font.Bold)
            dictInner_dict.Item("Font.TextColor") = getRGB(styleThis_style.Font.TextColor)
            dictInner_dict.Item("Font.Underline") = CBool(styleThis_style.Font.Underline)
            dictInner_dict.Item("Font.SmallCaps") = CBool(styleThis_style.Font.SmallCaps)
            dictInner_dict.Item("Font.Superscript") = CBool(styleThis_style.Font.Superscript)
            dictInner_dict.Item("Font.Subscript") = CBool(styleThis_style.Font.Subscript)
            dictInner_dict.Item("Font.StrikeThrough") = CBool(styleThis_style.Font.StrikeThrough)
            'set defaults for Keybindings & record actual Keybindings
            dictInner_dict.Item("shortcut_keys_enabled_tf") = False
            'dictInner_dict.Item("shortcut_key__kcode") = ""
            dictInner_dict.Item("shortcut_keys__shift_tf") = False
            dictInner_dict.Item("shortcut_keys__control_tf") = False
            dictInner_dict.Item("shortcut_keys__alt_tf") = False
            dictInner_dict.Item("shortcut_keys__letter") = ""
            For k = 1 To KeyBindings.Count
                If KeyBindings.Item(k).Command = styleThis_style.NameLocal Then
                    dictInner_dict.Item("shortcut_keys_enabled_tf") = True
                    lngKeycode = KeyBindings.Item(k).KeyCode
                    ' keycode value of 1920 or greater is invalid for reassignment
                    ' (this can happen if more than one keycode is assigned to a style)
                    If lngKeycode < 1920 Then
                        ' commenting this value write-out, as we aren't using it anymore
    '                    dictInner_dict.Item("shortcut_key__kcode") = lngKeycode
                        If lngKeycode > 1024 Then
                            dictInner_dict.Item("shortcut_keys__alt_tf") = True
                            lngKeycode = lngKeycode - 1024
                        End If
                        If lngKeycode > 512 Then
                            dictInner_dict.Item("shortcut_keys__control_tf") = True
                            lngKeycode = lngKeycode - 512
                        End If
                        If lngKeycode > 256 Then
                            dictInner_dict.Item("shortcut_keys__shift_tf") = True
                            lngKeycode = lngKeycode - 256
                        End If
                        dictInner_dict.Item("shortcut_keys__letter") = Chr(lngKeycode)
                    End If
                    Exit For
                End If
            Next
            '''Paragraph Styles only:
            If styleThis_style.Type = 1 Then
                ''Top Level Para-only Properties
                dictInner_dict.Item("NoSpaceBetweenParagraphsOfSameStyle") = styleThis_style.NoSpaceBetweenParagraphsOfSameStyle
                If InStr(styleThis_style.NameLocal, "Bullet") > 0 Then
                    dictInner_dict.Item("Bullet") = True
                Else
                    dictInner_dict.Item("Bullet") = False
                End If
                If InStr(styleThis_style.NameLocal, "Checklist") > 0 Then
                    dictInner_dict.Item("Checklist") = True
                Else
                    dictInner_dict.Item("Checklist") = False
                End If
                ''ParagraphFormat properties
                dictInner_dict.Item("ParagraphFormat.LeftIndent") = styleThis_style.ParagraphFormat.LeftIndent
                dictInner_dict.Item("ParagraphFormat.RightIndent") = styleThis_style.ParagraphFormat.RightIndent
                dictInner_dict.Item("ParagraphFormat.SpaceBefore") = styleThis_style.ParagraphFormat.SpaceBefore
                dictInner_dict.Item("ParagraphFormat.SpaceAfter") = styleThis_style.ParagraphFormat.SpaceAfter
                dictInner_dict.Item("ParagraphFormat.LineSpacingRule") = styleThis_style.ParagraphFormat.LineSpacingRule
                dictInner_dict.Item("ParagraphFormat.LineSpacing") = styleThis_style.ParagraphFormat.LineSpacing
                dictInner_dict.Item("ParagraphFormat.Alignment") = styleThis_style.ParagraphFormat.Alignment
                dictInner_dict.Item("ParagraphFormat.FirstLineIndent") = styleThis_style.ParagraphFormat.FirstLineIndent
                dictInner_dict.Item("ParagraphFormat.OutlineLevel") = styleThis_style.ParagraphFormat.OutlineLevel
                dictInner_dict.Item("ParagraphFormat.PageBreakBefore") = CBool(styleThis_style.ParagraphFormat.PageBreakBefore)
                dictInner_dict.Item("ParagraphFormat.WidowControl") = CBool(styleThis_style.ParagraphFormat.WidowControl)
                ''Border & Shading properties
                'cycle through top, left, bottom, right border properties respectively:
                For i = -1 To -4 Step -1
                    dictInner_dict.Item("Borders(" & i & ")") = styleThis_style.Borders(i)
    '                dictInner_dict.Item("Borders(" & i & ").ColorIndex") = styleThis_style.Borders(i).ColorIndex
                    dictInner_dict.Item("Borders(" & i & ").Color") = getRGB(styleThis_style.Borders(i).Color)
                    dictInner_dict.Item("Borders(" & i & ").LineStyle") = styleThis_style.Borders(i).LineStyle
                    dictInner_dict.Item("Borders(" & i & ").LineWidth") = styleThis_style.Borders(i).lineWidth
                Next i
                dictInner_dict.Item("Borders.DistanceFromLeft") = styleThis_style.Borders.DistanceFromLeft
                dictInner_dict.Item("Borders.DistanceFromRight") = styleThis_style.Borders.DistanceFromRight
                dictInner_dict.Item("Borders.DistanceFromTop") = styleThis_style.Borders.DistanceFromTop
                dictInner_dict.Item("Borders.DistanceFromBottom") = styleThis_style.Borders.DistanceFromBottom
                dictInner_dict.Item("ParagraphFormat.Shading.BackgroundPatternColor") = getRGB(styleThis_style.ParagraphFormat.Shading.BackgroundPatternColor)
    '            dictInner_dict.Item("ParagraphFormat.Shading.ForegroundPatternColor") = styleThis_style.ParagraphFormat.Shading.ForegroundPatternColor
    '            dictInner_dict.Item("ParagraphFormat.Shading.Texture") = styleThis_style.ParagraphFormat.Shading.Texture
            '''Character Styles only:
            ElseIf styleThis_style.Type = 2 Then
                ''Border & Shading properties
                dictInner_dict.Item("Font.Borders(1).Color") = getRGB(styleThis_style.Font.Borders(1).Color)
                dictInner_dict.Item("Font.Borders(1).LineStyle") = styleThis_style.Font.Borders(1).LineStyle
                dictInner_dict.Item("Font.Borders(1).LineWidth") = styleThis_style.Font.Borders(1).lineWidth
    '            dictInner_dict.Item("Font.Borders(1).ColorIndex") = styleThis_style.Font.Borders(1).ColorIndex
                dictInner_dict.Item("Font.Shading.BackgroundPatternColor") = getRGB(styleThis_style.Font.Shading.BackgroundPatternColor)
    '            dictInner_dict.Item("Font.Shading.ForegroundPatternColor") = styleThis_style.Font.Shading.ForegroundPatternColor
    '            dictInner_dict.Item("Font.Shading.Texture") = styleThis_style.Font.Shading.Texture
            End If
            
            '''Set all this to the style name Key in outer dict
            Set dictStyle_dict(styleThis_style.NameLocal) = dictInner_dict
        
    ''''''TESTING top level styles to see if they are uniformly applied in our situation (& can be ignored)
    '    ''''-------------Checking styles!
    '    If Not styleThis_style.LanguageID = 1033 Then
    '               Debug.Print styleThis_style.NameLocal & " LanguageID = " & styleThis_style.LanguageID
    '    End If
    '    If Not styleThis_style.Locked = False Then
    '               Debug.Print styleThis_style.NameLocal & " Locked = " & styleThis_style.Locked
    '    End If
    '    If Not styleThis_style.InUse = True Then
    '               Debug.Print styleThis_style.NameLocal & " InUse = " & styleThis_style.InUse
    '    End If
    '    ''''-------Checking paragraph only styles!
    '    If styleThis_style.Type = 1 Then
    '        If Not styleThis_style.Linked = False Then
    '               Debug.Print styleThis_style.NameLocal & " Linked = " & styleThis_style.Linked
    '        End If
    '        If Not styleThis_style.LinkStyle = "Normal" Then
    '            Debug.Print styleThis_style.NameLocal & " LinkStyle = " & styleThis_style.LinkStyle
    '        End If
    '        '''paragraphformatting styles:
    ''        If Not styleThis_style.ParagraphFormat.SpaceAfterAuto = False Then
    ''            Debug.Print styleThis_style.NameLocal & " SpaceAfterAuto = " & styleThis_style.ParagraphFormat.SpaceAfterAuto
    ''        End If
    ''        If Not styleThis_style.ParagraphFormat.CharacterUnitLeftIndent = 0 Then
    ''            Debug.Print styleThis_style.NameLocal & " CharacterUnitLeftIndent = " & styleThis_style.ParagraphFormat.CharacterUnitLeftIndent
    ''        End If
    ''        If Not styleThis_style.ParagraphFormat.CharacterUnitRightIndent = 0 Then
    ''            Debug.Print styleThis_style.NameLocal & " CharacterUnitRightIndent = " & styleThis_style.ParagraphFormat.CharacterUnitRightIndent
    ''        End If
    ''        If Not styleThis_style.ParagraphFormat.CharacterUnitFirstLineIndent = 0 Then
    ''            Debug.Print styleThis_style.NameLocal & " CharacterUnitFirstLineIndent = " & styleThis_style.ParagraphFormat.CharacterUnitFirstLineIndent
    ''        End If
    ''        If Not styleThis_style.ParagraphFormat.LineUnitBefore = False Then
    ''            Debug.Print styleThis_style.NameLocal & " LineUnitBefore = " & styleThis_style.ParagraphFormat.LineUnitBefore
    ''        End If
    ''        If Not styleThis_style.ParagraphFormat.LineUnitAfter = False Then
    ''            Debug.Print styleThis_style.NameLocal & " LineUnitAfter = " & styleThis_style.ParagraphFormat.LineUnitAfter
    ''        End If
    '
    '    End If
        '''' ------Checking char only styles?
    '    If styleThis_style.Type = 2 Then
    '       If styleThis_style.Linked = False Then
    '               Debug.Print styleThis_style.NameLocal & " Linked = " & styleThis_style.Linked
    '       End If
    '    End if
    
    
    '    ElseIf InStr(styleThis_style.NameLocal, "(") > 0 Then
    '        Debug.Print "ALERT: Found a builtin style wiht a paren! :" & styleThis_style.NameLocal
        End If
        
        If a Mod lngIncrement = 0 Or a = Word.ActiveDocument.Styles.Count Then
            ActiveDocument.UndoClear
        End If
    Next
    
    strJson = JsonConverter.ConvertToJson(dictStyle_dict, Whitespace:=2)
    OverwriteTextFile strJsonPath, strJson
    Debug.Print "Scanned " & a & " styles, wrote info for " & dictStyle_dict.Count & " Macmillan Styles to json"

End Sub

Public Sub WriteNoColorTemplatefromJson()
    ' So we can invoke WriteTemplatefromJson Sub for no color template, without rewriting hte whole thing over
    Call WriteTemplatefromJsonCore(True)

End Sub
Public Sub WriteTemplatefromJson()
    ' So we can invoke WriteTemplatefromJson Sub for regular template;
    ' cuz optional parameters don't work in Word vba :(
    Call WriteTemplatefromJsonCore(False)

End Sub

Private Function WriteTemplatefromJsonCore(Optional p_boolNoColor As Boolean = False)
    ' NOTE: requirements for this sub are listed in comments under this moedule's "Declarations"
    ' When invoked this sub looks for a "macmillan.json" file in the same dir as
    ' "StyleTemplateCreator.docm".  It loads all of the styles form that json file
    ' and creates corresponding styles in a new macmillan.dotx template file
    ' (also in the same dir).  Then it writes cycles through the new styles in the .dotx
    ' and writes a paragraph for each one to the dotx.
    ' It can be run from Mac for Word, but cannot write a handful of properties via Mac for Word:
    ' The one that's really a problem is the Quickstyles property
    Dim strJsonPath As String
    Dim strNewFilePath As String
    Dim strNewFileName As String
    Dim dictStyle_dict As Dictionary
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim n As Long
    Dim k As Long
    Dim lngIncrement As Long
    Dim strStylename As String
    Dim strStylename_B As String
    Dim strStylename_C As String
    Dim docTemplate As Document
    Dim objBulletLT As ListTemplate
    Dim objChecklistLT As ListTemplate
    
    ' file paths
    strJsonPath = ThisDocument.Path & Application.PathSeparator & "macmillan.json"
    If p_boolNoColor = False Then
        strNewFilePath = ThisDocument.Path & Application.PathSeparator & "macmillan.dotx"
    ElseIf p_boolNoColor = True Then
        strNewFilePath = ThisDocument.Path & Application.PathSeparator & "macmillan_NoColor.dotx"
    End If

    Application.ScreenUpdating = False

    ' create and save new template doc - existing file is overwritten
    Set docTemplate = Documents.Add
    docTemplate.SaveAs2 FileName:=strNewFilePath, FileFormat:= _
        wdFormatXMLTemplate, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=15

    ' Add the version of the template
    AddVersionNumber docTemplate
    
    ' read in the json data
    Set dictStyle_dict = localReadJson(strJsonPath)
    ' create list templates for adding bullets to styles
    Set objBulletLT = CreateListTemplate("objBulletLT", "61623", "Symbol")
    Set objChecklistLT = CreateListTemplate("objChecklistLT", "61692", "Wingdings")
    ' how often we save and undo.clear on doc to prevent memory errs
    If IsMac Then
        lngIncrement = 10
    Else
        lngIncrement = 50
    End If
    

    ' Cycle through all styles in style_dict and create corresponding styles in new .dotx
    For a = 0 To dictStyle_dict.Count - 1
        strStylename = dictStyle_dict.Keys(a)
        Debug.Print strStylename

        ''CREATE THE STYLE!
        docTemplate.Styles.Add Name:=strStylename, Type:=dictStyle_dict(strStylename).Item("Type")
            With docTemplate.Styles(strStylename).Font
                .Name = dictStyle_dict(strStylename).Item("Font.name")
                .Size = Val(dictStyle_dict(strStylename).Item("Font.Size"))
                .Bold = CBool(dictStyle_dict(strStylename).Item("Font.Bold"))
                .Italic = CBool(dictStyle_dict(strStylename).Item("Font.Italic"))
                .Underline = CBool(dictStyle_dict(strStylename).Item("Font.Underline"))
                .UnderlineColor = wdColorAutomatic
                .StrikeThrough = CBool(dictStyle_dict(strStylename).Item("Font.StrikeThrough"))
                .DoubleStrikeThrough = False
                .Outline = False
                .Emboss = False
                .Shadow = False
                .Hidden = False
                .SmallCaps = CBool(dictStyle_dict(strStylename).Item("Font.SmallCaps"))
                .AllCaps = False
                ' do not apply for nocolor template!
                If p_boolNoColor = False Then
                    .Color = get24bitColor(dictStyle_dict(strStylename).Item("Font.TextColor"))
                End If
                .Engrave = False
                .Superscript = CBool(dictStyle_dict(strStylename).Item("Font.Superscript"))
                .Subscript = CBool(dictStyle_dict(strStylename).Item("Font.Subscript"))
                .Scaling = 100
                .Kerning = 0
                .Animation = wdAnimationNone
                .Ligatures = wdLigaturesNone
                .NumberSpacing = wdNumberSpacingDefault
                .NumberForm = wdNumberFormDefault
                .StylisticSet = wdStylisticSetDefault
                .ContextualAlternates = 0
            End With
            ' these style properties do not exist in Word for MAc
            If Not IsMac Then
                docTemplate.Styles(strStylename).QuickStyle = CBool(dictStyle_dict(strStylename).Item("QuickStyle"))
                docTemplate.Styles(strStylename).Priority = dictStyle_dict(strStylename).Item("Priority")
            End If
        '''For Paragraph styles only:
        If dictStyle_dict(strStylename).Item("Type") = 1 Then
            docTemplate.Styles(strStylename).AutomaticallyUpdate = False
            With docTemplate.Styles(strStylename).ParagraphFormat
                .LeftIndent = dictStyle_dict(strStylename).Item("ParagraphFormat.LeftIndent")
                .RightIndent = dictStyle_dict(strStylename).Item("ParagraphFormat.RightIndent")
                .SpaceBefore = dictStyle_dict(strStylename).Item("ParagraphFormat.SpaceBefore")
        '        .SpaceBeforeAuto = False
                .SpaceAfter = dictStyle_dict(strStylename).Item("ParagraphFormat.SpaceAfter")
        '        .SpaceAfterAuto = False
                .LineSpacing = dictStyle_dict(strStylename).Item("ParagraphFormat.LineSpacing")
                .LineSpacingRule = dictStyle_dict(strStylename).Item("ParagraphFormat.LineSpacingRule")
                .Alignment = dictStyle_dict(strStylename).Item("ParagraphFormat.Alignment")
                .WidowControl = CBool(dictStyle_dict(strStylename).Item("ParagraphFormat.WidowControl"))
                .KeepWithNext = False
                .KeepTogether = False
                .PageBreakBefore = CBool(dictStyle_dict(strStylename).Item("ParagraphFormat.PageBreakBefore"))
                .NoLineNumber = False
                .Hyphenation = True
                .FirstLineIndent = dictStyle_dict(strStylename).Item("ParagraphFormat.FirstLineIndent")
                .OutlineLevel = dictStyle_dict(strStylename).Item("ParagraphFormat.OutlineLevel")
                .CharacterUnitLeftIndent = 0
                .CharacterUnitRightIndent = 0
                .CharacterUnitFirstLineIndent = 0
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                If Not IsMac Then
                    .MirrorIndents = False
                    .TextboxTightWrap = wdTightNone
                    .CollapsedByDefault = False
                End If
            End With
            docTemplate.Styles(strStylename).NoSpaceBetweenParagraphsOfSameStyle = _
            dictStyle_dict(strStylename).Item("NoSpaceBetweenParagraphsOfSameStyle")
            docTemplate.Styles(strStylename).ParagraphFormat.TabStops.ClearAll
            ' do not apply for nocolor template!
            If p_boolNoColor = False Then
                '' Borders and Shading Paragraph level properties
                With docTemplate.Styles(strStylename).ParagraphFormat
                    With .Shading
                        .Texture = wdTextureNone
                        .ForegroundPatternColor = wdColorAutomatic
                        .BackgroundPatternColor = get24bitColor(dictStyle_dict(strStylename).Item("ParagraphFormat.Shading.BackgroundPatternColor"))
                    End With
                    For n = -1 To -4 Step -1
                        With .Borders(n)
                            If dictStyle_dict(strStylename).Item("Borders(" & n & ")") = True Then
                                .LineStyle = dictStyle_dict(strStylename).Item("Borders(" & n & ").LineStyle")
                                .lineWidth = Val(dictStyle_dict(strStylename).Item("Borders(" & n & ").LineWidth"))
                                .Color = get24bitColor(dictStyle_dict(strStylename).Item("Borders(" & n & ").Color"))
                            End If
                        End With
                    Next n
                    With .Borders
                        .DistanceFromTop = dictStyle_dict(strStylename).Item("Borders.DistanceFromTop")
                        .DistanceFromLeft = dictStyle_dict(strStylename).Item("Borders.DistanceFromLeft")
                        .DistanceFromBottom = dictStyle_dict(strStylename).Item("Borders.DistanceFromBottom")
                        .DistanceFromRight = dictStyle_dict(strStylename).Item("Borders.DistanceFromRight")
                        .Shadow = False
                    End With
                End With
            End If
            docTemplate.Styles(strStylename).Frame.Delete
            'Add bullets or checklist symbols by linking to our ListTemplates:
            If dictStyle_dict(strStylename).Item("Bullet") = True Then
                docTemplate.Styles(strStylename).LinkToListTemplate _
                ListTemplate:=objBulletLT, ListLevelNumber:=1
            ElseIf dictStyle_dict(strStylename).Item("Checklist") = True Then
                docTemplate.Styles(strStylename).LinkToListTemplate _
                ListTemplate:=objChecklistLT, ListLevelNumber:=1
            End If
        ''' For Character styles only:
        ElseIf dictStyle_dict(strStylename).Item("Type") = 2 Then
            '' Borders & Shading for Character styles
            ' do not apply for nocolor template!
            If p_boolNoColor = False Then
                With docTemplate.Styles(strStylename).Font
                    With .Shading
                        .Texture = wdTextureNone
                        .ForegroundPatternColor = wdColorAutomatic
                        .BackgroundPatternColor = get24bitColor(dictStyle_dict(strStylename).Item("Font.Shading.BackgroundPatternColor"))
                    End With
                    With .Borders(1)
                        .LineStyle = dictStyle_dict(strStylename).Item("Font.Borders(1).LineStyle")
                        .lineWidth = Val(dictStyle_dict(strStylename).Item("Font.Borders(1).LineWidth"))
                        .Color = get24bitColor(dictStyle_dict(strStylename).Item("Font.Borders(1).Color"))
                    End With
                    .Borders.Shadow = False
                End With
            End If
        End If
        docTemplate.Styles(strStylename).LanguageID = wdEnglishUS
        docTemplate.Styles(strStylename).NoProofing = CBool(dictStyle_dict(strStylename).Item("NoProofing"))
        'for some reason these all get put in as linked styles in PC version of Word... this gets rid of that
        'As per: https://answers.microsoft.com/en-us/msoffice/forum/msoffice_word-mso_winother/how-do-i-get-rid-of-linked-styles-in-our-templates/4531bec8-1ab7-4170-bf2d-a7433fea9d5f
        If Not IsMac Then
            docTemplate.Styles(strStylename).LinkStyle = docTemplate.Styles("Normal")
            If styleExists(strStylename & " Char", docTemplate) = True Then
                docTemplate.Styles(strStylename & " Char").Delete
            End If
        End If
        '''Add Keybindings where there are values entered
        If dictStyle_dict(strStylename).Item("shortcut_keys_enabled_tf") = True Then
            If Not dictStyle_dict(strStylename).Item("shortcut_keys__letter") = vbNullString Then
                Call CreateKeybindings(dictStyle_dict, strStylename, docTemplate)
            End If
         End If
        '''Save & Reset scratch disk
        If a Mod lngIncrement = 0 Or a = docTemplate.Styles.Count Then
            docTemplate.UndoClear
            docTemplate.Save
        End If
    Next a
    'Debug.Print "Styles created!"

    '''Cycle through again to update styles properties dependent on other styles being present:
    For c = 0 To dictStyle_dict.Count - 1
        strStylename_C = dictStyle_dict.Keys(c)
        If dictStyle_dict(strStylename_C).Item("Type") = 1 Then
            docTemplate.Styles(strStylename_C).NextParagraphStyle = dictStyle_dict(strStylename_C).Item("NextParagraphStyle")
        End If
    Next c

    docTemplate.UndoClear
    docTemplate.Save

    '''Create styled demo paragraph
    For b = 0 To dictStyle_dict.Count - 1
        strStylename_B = dictStyle_dict.Keys(b)

        docTemplate.Paragraphs.Add
        'For paragraph styles
        If dictStyle_dict(strStylename_B).Item("Type") = 1 Then
            'clear character styles from previous para
            docTemplate.Paragraphs.Last.Range.Style = "Default Paragraph Font"
            'set character style
            docTemplate.Paragraphs.Last.Style = strStylename_B
        'For character styles
        ElseIf dictStyle_dict(strStylename_B).Item("Type") = 2 Then
            'set a generic paragraph style (prefer Text-standard if present)
            If styleExists("Text - Standard (tx)", docTemplate) = True Then
                docTemplate.Paragraphs.Last.Style = "Text - Standard (tx)"
            Else
                docTemplate.Paragraphs.Last.Style = "Normal"
            End If
            'apply character style
            docTemplate.Paragraphs.Last.Range.Style = strStylename_B
        End If
        'write style name
        docTemplate.Content.InsertAfter strStylename_B

        '''Save & Reset scratch disk periodically
        If b Mod lngIncrement = 0 Or b = docTemplate.Styles.Count Then
            docTemplate.UndoClear
            docTemplate.Save
        End If
    Next
    Debug.Print "Demo paragraphs written & styled!"
    
    '''Lower priority of built in styles to 9
    Call LowerPriorityBuiltInStyles(docTemplate)
    
    docTemplate.Save
    Application.ScreenUpdating = True

End Function

Function IsMac() As Boolean
    ' Test if were running on a Mac
    #If Mac Then
        IsMac = True
    #End If
End Function

Private Function styleExists(ByVal strStyleToTest As String, ByVal docToTest As Word.Document) As Boolean
    ' Quick Function to check if style exists, self-explanatory ;)
    Dim objTestStyle As Word.Style
    On Error Resume Next
    Set objTestStyle = docToTest.Styles(strStyleToTest)
    styleExists = Not objTestStyle Is Nothing
End Function


Private Function get24bitColor(rgbString As String)
    ' Take an rgb color assignation as string ("r,g,b") and return a 24bit color int
    Dim arrRGBstring() As String
    Dim lngBitColor As Long
    Dim enumBitColor As WdColor
    
    enumBitColor = wdColorAutomatic
    
    ' If the rgbstring is "auto" enumeration, return the same thing with proper type
    If rgbString = "wdColorAutomatic" Then
        get24bitColor = enumBitColor
    ' Else do the conversion
    Else
        arrRGBstring = Split(rgbString, ",")
        lngBitColor = RGB(CInt(arrRGBstring(0)), CInt(arrRGBstring(1)), CInt(arrRGBstring(2)))
        get24bitColor = lngBitColor
    End If
        
End Function



Private Function getRGB(colorLong As Long) As String
    ' Take a 24bitColor Lng int and convert to rgb values ("r,g,b") as string
    Dim strRGBvalue As String
    ReDim arrRGBsplit(1 To 3) As Variant
    
    ' This is the max value that I can convert - MS allows some longer ones!?
    If Abs(colorLong) > 16777216 Then
        strRGBvalue = "conversion error: int too large"
    ' This int "-16777216" is the Word value for "automatic" color
    ElseIf colorLong = -16777216 Then
        strRGBvalue = "wdColorAutomatic"
    ' Do the conversion
    Else
        colorLong = Abs(colorLong)
        ' Treating Color as a 24 bit number
        arrRGBsplit(1) = colorLong Mod 256          ' Red value: left most 8 bits
        arrRGBsplit(2) = colorLong \ 256 Mod 256    ' Green value: middle 8 bits
        arrRGBsplit(3) = colorLong \ 65536 Mod 256  ' Blue value: right most 8 bits
        strRGBvalue = Join(arrRGBsplit, ",")
        If Not get24bitColor(strRGBvalue) = colorLong Then
            strRGBvalue = "conversion error: reverse lookup failed"
        End If
    End If
    
    getRGB = strRGBvalue

End Function



Private Function CreateListTemplate(LTname As String, numFormat As String, fontName As String) As ListTemplate
    ' create List Template, so we can attach list styles to them; this is how
    ' we ensure that a list style is preceded by a bullet or checkmark
    Dim objListTemplate As ListTemplate
    Dim blnLTExists As Boolean
    
    blnLTExists = False
    
    '''Check for existence of templates
    For Each objListTemplate In ActiveDocument.ListTemplates
        If objListTemplate.Name = LTname Then
            blnLTExists = True
            Exit For
        End If
    Next objListTemplate
    
    '''Create List Templates
    If blnLTExists = False Then
        Set objListTemplate = ActiveDocument.ListTemplates.Add _
        (OutlineNumbered:=True, Name:="" & LTname & "")
    ElseIf blnLTExists = True Then
        Set objListTemplate = ActiveDocument.ListTemplates(LTname)
    End If
    
    '''Update settings for list templates:
    With objListTemplate.ListLevels(1)
        .NumberFormat = ChrW(numFormat)
        .NumberStyle = wdListNumberStyleBullet
        .TextPosition = CentimetersToPoints(0.75)
        .TabPosition = CentimetersToPoints(1.25)
        With .Font
            .Name = fontName
        End With
    End With
    
    Set CreateListTemplate = objListTemplate

End Function

Private Sub LowerPriorityBuiltInStyles(docTemplate As Document)
Dim q As Long

'''Set Priority for all built-in styles to 9
For q = 1 To docTemplate.Styles.Count
If docTemplate.Styles(q).BuiltIn = True Then
    docTemplate.Styles(q).Priority = 90
End If
Next q

docTemplate.UndoClear

End Sub

' ===== AddVersionNumber ======================================================
' Reads version number from text file, adds to template as doc property.
' Objects passed By Ref by default so will update same object as is passed.

Private Sub AddVersionNumber(docNewTemplate As Document)
  Dim strVersionFileFullPath As String
  Dim strVersionNumber As String
  
  ' version file is same name, same directory, different extension.
  strVersionFileFullPath = VBA.Replace(docNewTemplate.FullName, ".dotx", ".txt")
  strVersionNumber = localReadTextFile(Path:=strVersionFileFullPath)
  docNewTemplate.CustomDocumentProperties.Add Name:="Version", LinkToContent:=False, _
    Type:=msoPropertyTypeString, Value:=strVersionNumber
End Sub
