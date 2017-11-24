# welcome to my Excel Macro Page
 It's just some Excel macros I find useful

Copy and Paste into a new module in VB


### Markdown

Markdown is a lightweight and easy-to-use syntax for styling your writing. It includes conventions for

```markdown

Sub Delimit()
' Delimit Macro - Will 

    ActiveSheet.Paste
    Range("B1").Select
End Sub


Sub Fill_Blank_Cells()
'Can't really remember - probably fills blank cells!

Selection.SpecialCells(xlCellTypeBlanks).Select
Selection.FormulaR1C1 = "=R[-1]C"
End Sub


Sub Format_PTFields()
'Macro goal: allow users to quickly choose the format to apply to pivot table fields
'Code modified from Dick Kusleika's code at:
'http://www.dailydoseofexcel.com/archives/2010/06/18/formatting-pivot-tables/

    Dim pf As PivotField
    Dim FormatChoice As String 'allows you to dynamically select the format
   Dim QuestionString As String

    On Error GoTo HandleErr
 

    If TypeName(Selection) = "Range" Then Set pf = ActiveCell.PivotField

    'Consolidates the question blurb to a variable
   QuestionString = "Apply which format to this pivot field?" & vbCrLf & _
                "    '0': numbers with 0 digits after the decimals" & vbCrLf & _
                "    '1': numbers with 1 digit after the decimals" & vbCrLf & _
                "    'd': dollars (no cents)" & vbCrLf & _
                "    'c': dollars and cents"

    'Ask the user what format to apply
   FormatChoice = InputBox(QuestionString)

    'based on the FormatChoice, format the selected pivot field
   Select Case FormatChoice
        Case 0      'shows numbers with 0 digits after the decimal
           pf.NumberFormat = "#,##0"

        Case 1      'shows numbers with 1 digit after the decimal
          pf.NumberFormat = "#,##0.0"

        Case "d"    'shows dollars (no cents)
           pf.NumberFormat = "$#,##0"

        Case "c"    'Shows dollars and cents
           pf.NumberFormat = "$#,##0.00"
    End Select

ExitSub:
    Exit Sub

HandleErr:
    If Err = 1004 Then
        MsgBox ("This macro only does something useful if you are " & vbCrLf & _
                "in a pivot table value field.  Exiting macro.")
    Else
        MsgBox "Unexpected Error: " & Err & Err.Description
    End If

    GoTo ExitSub

End Sub

Sub SelectAdjacentCol()
' Select empty cells vertically next to partially filled column
' Keyboard Shortcut: Ctrl+m

    Dim rAdjacent As Range

    If TypeName(Selection) = "Range" Then
        If Selection.Column > 1 Then
            If Selection.Cells.Count = 1 Then
                If Not IsEmpty(Selection.Offset(0, -1).Value) Then
                    With Selection.Offset(0, -1)
                        Set rAdjacent = .Parent.Range(.Cells(1), .End(xlDown))
                    End With

                    Selection.Resize(rAdjacent.Cells.Count).Select
                End If
            End If
        End If
    End If

End Sub


Sub format()
' format Macro
' Keyboard Shortcut: Ctrl+w

    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    With Selection.Font
        .Name = "Tahoma"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Tahoma"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
End Sub


Sub Adjust_cols()
' Adjust_cols Macro
' Selects all columns with content and resizes to longest content
' Keyboard Shortcut: Ctrl+j

    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    ActiveCell.Columns("A:A").EntireColumn.EntireColumn.AutoFit
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Columns.AutoFit
    Range("A1").Select
End Sub

Sub Header()
' Header Macro'
' Keyboard Shortcut: Ctrl+h

    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    With Selection.Font
        .Name = "Tahoma"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Columns.AutoFit
    Selection.End(xlToLeft).Select
End Sub

Sub delete_sheet()
' delete_sheet Macro
' deletes current sheet
' Keyboard Shortcut: Ctrl+g

    ActiveWindow.SelectedSheets.Delete
End Sub

Sub Clear_Range_End()
' Clear_Range_End Macro
' Keyboard Shortcut: Ctrl+k

    ActiveWorkbook.Save
End Sub

    © 2017 GitHub, Inc.
    Terms
    Privacy
    Security
    Status
    Help

    Contact GitHub
    API
    Training
    Shop
    Blog
    About





# Header 1
## Header 2
### Header 3

- Bulleted
- List

1. Numbered
2. List

**Bold** and _Italic_ and `Code` text

[Curt's macros](https://github.com/curtlymartin/Excel/blob/master/macro_list.vb) and ![Image](src)
```

For more details see [GitHub Flavored Markdown](https://guides.github.com/features/mastering-markdown/).
[Curt's macros](https://github.com/curtlymartin/Excel/blob/master/macro_list.vb).


### Jekyll Themes

Your Pages site will use the layout and styles from the Jekyll theme you have selected in your [repository settings](https://github.com/curtlymartin/Excel/settings). The name of this theme is saved in the Jekyll `_config.yml` configuration file.

### Support or Contact

Having trouble with Pages? Check out our [documentation](https://help.github.com/categories/github-pages-basics/) or [contact support](https://github.com/contact) and we’ll help you sort it out.
