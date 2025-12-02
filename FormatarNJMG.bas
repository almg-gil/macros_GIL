REM  *****  BASIC  *****

sub FormataNJMG
rem ----------------------------------------------------------------------
rem define variables
dim document as object
dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:GoToStartOfDoc", "", 0, Array())

rem ----------------------------------------------------------------------
dim args1(18) as new com.sun.star.beans.PropertyValue
args1(0).Name = "SearchItem.StyleFamily"
args1(0).Value = 2
args1(1).Name = "SearchItem.CellType"
args1(1).Value = 0
args1(2).Name = "SearchItem.RowDirection"
args1(2).Value = true
args1(3).Name = "SearchItem.AllTables"
args1(3).Value = false
args1(4).Name = "SearchItem.Backward"
args1(4).Value = false
args1(5).Name = "SearchItem.Pattern"
args1(5).Value = false
args1(6).Name = "SearchItem.Content"
args1(6).Value = false
args1(7).Name = "SearchItem.AsianOptions"
args1(7).Value = false
args1(8).Name = "SearchItem.AlgorithmType"
args1(8).Value = 1
args1(9).Name = "SearchItem.SearchFlags"
args1(9).Value = 65536
args1(10).Name = "SearchItem.SearchString"
args1(10).Value = "   *"
args1(11).Name = "SearchItem.ReplaceString"
args1(11).Value = " "
args1(12).Name = "SearchItem.Locale"
args1(12).Value = 255
args1(13).Name = "SearchItem.ChangedChars"
args1(13).Value = 2
args1(14).Name = "SearchItem.DeletedChars"
args1(14).Value = 2
args1(15).Name = "SearchItem.InsertedChars"
args1(15).Value = 2
args1(16).Name = "SearchItem.TransliterateFlags"
args1(16).Value = 1280
args1(17).Name = "SearchItem.Command"
args1(17).Value = 3
args1(18).Name = "Quiet"
args1(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1())

rem ----------------------------------------------------------------------
rem dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, Array())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:SelectAll", "", 0, Array())

rem ----------------------------------------------------------------------
dim args2(1) as new com.sun.star.beans.PropertyValue
args2(0).Name = "Template"
args2(0).Value = "Estilo padrão"
args2(1).Name = "Family"
args2(1).Value = 2

dispatcher.executeDispatch(document, ".uno:StyleApply", "", 0, args2())
rem ----------------------------------------------------------------------
dim args4(2) as new com.sun.star.beans.PropertyValue
args4(0).Name = "Alignment.ParagraphAdjustment"
args4(0).Value = 2
args4(1).Name = "Alignment.LastLineAdjustment"
args4(1).Value = 0
args4(2).Name = "Alignment.ExpandSingleWord"
args4(2).Value = false

dispatcher.executeDispatch(document, ".uno:Alignment", "", 0, args4())

rem ----------------------------------------------------------------------
dim args5(7) as new com.sun.star.beans.PropertyValue
args5(0).Name = "LeftRightMargin.LeftMargin"
args5(0).Value = 0
args5(1).Name = "LeftRightMargin.TextLeftMargin"
args5(1).Value = 0
args5(2).Name = "LeftRightMargin.RightMargin"
args5(2).Value = 0
args5(3).Name = "LeftRightMargin.LeftRelMargin"
args5(3).Value = 100
args5(4).Name = "LeftRightMargin.RightRelMargin"
args5(4).Value = 100
args5(5).Name = "LeftRightMargin.FirstLineIndent"
args5(5).Value = 2500
args5(6).Name = "LeftRightMargin.FirstLineRelIdent"
args5(6).Value = 100
args5(7).Name = "LeftRightMargin.AutoFirst"
args5(7).Value = false

dispatcher.executeDispatch(document, ".uno:LeftRightMargin", "", 0, args5())

rem ----------------------------------------------------------------------
dim args6(4) as new com.sun.star.beans.PropertyValue
args6(0).Name = "TopBottomMargin.TopMargin"
args6(0).Value = 0
args6(1).Name = "TopBottomMargin.BottomMargin"
args6(1).Value = 0
args6(2).Name = "TopBottomMargin.ContextMargin"
args6(2).Value = false
args6(3).Name = "TopBottomMargin.TopRelMargin"
args6(3).Value = 100
args6(4).Name = "TopBottomMargin.BottomRelMargin"
args6(4).Value = 100

dispatcher.executeDispatch(document, ".uno:TopBottomMargin", "", 0, args6())

rem ----------------------------------------------------------------------
dim args7(4) as new com.sun.star.beans.PropertyValue
args7(0).Name = "CharFontName.StyleName"
args7(0).Value = ""
args7(1).Name = "CharFontName.Pitch"
args7(1).Value = 1
args7(2).Name = "CharFontName.CharSet"
args7(2).Value = -1
args7(3).Name = "CharFontName.Family"
args7(3).Value = 2
args7(4).Name = "CharFontName.FamilyName"
args7(4).Value = "Times New Roman"

dispatcher.executeDispatch(document, ".uno:CharFontName", "", 0, args7())

rem ----------------------------------------------------------------------
dim args8(2) as new com.sun.star.beans.PropertyValue
args8(0).Name = "FontHeight.Height"
args8(0).Value = 12
args8(1).Name = "FontHeight.Prop"
args8(1).Value = 100
args8(2).Name = "FontHeight.Diff"
args8(2).Value = 0

dispatcher.executeDispatch(document, ".uno:FontHeight", "", 0, args8())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:GoToEndOfDoc", "", 0, Array())

rem ----------------------------------------------------------------------
rem dispatcher.executeDispatch(document, ".uno:InsertPara", "", 0, Array())

rem ----------------------------------------------------------------------
rem dispatcher.executeDispatch(document, ".uno:InsertPara", "", 0, Array())

rem ----------------------------------------------------------------------
dim args11(0) as new com.sun.star.beans.PropertyValue
args11(0).Name = "Text"
args11(0).Value = "============================================="

rem dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args11())

rem ----------------------------------------------------------------------
rem dispatcher.executeDispatch(document, ".uno:InsertPara", "", 0, Array())

rem ----------------------------------------------------------------------
dim args13(0) as new com.sun.star.beans.PropertyValue
args13(0).Name = "Text"
args13(0).Value = "Data da última atualização: "

rem dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args13())

rem ----------------------------------------------------------------------
dim args14(5) as new com.sun.star.beans.PropertyValue
args14(0).Name = "Type"
args14(0).Value = 0
args14(1).Name = "SubType"
args14(1).Value = 0
args14(2).Name = "Name"
args14(2).Value = ""
args14(3).Name = "Content"
args14(3).Value = "0"
args14(4).Name = "Format"
args14(4).Value = 7
args14(5).Name = "Separator"
args14(5).Value = " "

rem dispatcher.executeDispatch(document, ".uno:InsertField", "", 0, args14())

rem ----------------------------------------------------------------------
dim args15(0) as new com.sun.star.beans.PropertyValue
args15(0).Name = "Text"
args15(0).Value = Format(Now(), "DD/MM/YYYY") & "."

rem dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args15())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:GoToStartOfDoc", "", 0, Array())

rem ----------------------------------------------------------------------
dim args16(7) as new com.sun.star.beans.PropertyValue
args16(0).Name = "LeftRightMargin.LeftMargin"
args16(0).Value = 7620
args16(1).Name = "LeftRightMargin.TextLeftMargin"
args16(1).Value = 7620
args16(2).Name = "LeftRightMargin.RightMargin"
args16(2).Value = 0
args16(3).Name = "LeftRightMargin.LeftRelMargin"
args16(3).Value = 100
args16(4).Name = "LeftRightMargin.RightRelMargin"
args16(4).Value = 100
args16(5).Name = "LeftRightMargin.FirstLineIndent"
args16(5).Value = 0
args16(6).Name = "LeftRightMargin.FirstLineRelIdent"
args16(6).Value = 100
args16(7).Name = "LeftRightMargin.AutoFirst"
args16(7).Value = false

dispatcher.executeDispatch(document, ".uno:LeftRightMargin", "", 0, args16())

rem ----------------------------------------------------------------------
dim args17(18) as new com.sun.star.beans.PropertyValue
args17(0).Name = "SearchItem.StyleFamily"
args17(0).Value = 2
args17(1).Name = "SearchItem.CellType"
args17(1).Value = 0
args17(2).Name = "SearchItem.RowDirection"
args17(2).Value = true
args17(3).Name = "SearchItem.AllTables"
args17(3).Value = false
args17(4).Name = "SearchItem.Backward"
args17(4).Value = false
args17(5).Name = "SearchItem.Pattern"
args17(5).Value = false
args17(6).Name = "SearchItem.Content"
args17(6).Value = false
args17(7).Name = "SearchItem.AsianOptions"
args17(7).Value = false
args17(8).Name = "SearchItem.AlgorithmType"
args17(8).Value = 1
args17(9).Name = "SearchItem.SearchFlags"
args17(9).Value = 65536
args17(12).Name = "SearchItem.Locale"
args17(12).Value = 255
args17(13).Name = "SearchItem.ChangedChars"
args17(13).Value = 2
args17(14).Name = "SearchItem.DeletedChars"
args17(14).Value = 2
args17(15).Name = "SearchItem.InsertedChars"
args17(15).Value = 2
args17(16).Name = "SearchItem.TransliterateFlags"
args17(16).Value = 1280
args17(17).Name = "SearchItem.Command"
args17(17).Value = 3
args17(18).Name = "Quiet"
args17(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args17())

rem ----------------------------------------------------------------------
dim args18(18) as new com.sun.star.beans.PropertyValue
args18(0).Name = "SearchItem.StyleFamily"
args18(0).Value = 2
args18(1).Name = "SearchItem.CellType"
args18(1).Value = 0
args18(2).Name = "SearchItem.RowDirection"
args18(2).Value = true
args18(3).Name = "SearchItem.AllTables"
args18(3).Value = false
args18(4).Name = "SearchItem.Backward"
args18(4).Value = false
args18(5).Name = "SearchItem.Pattern"
args18(5).Value = false
args18(6).Name = "SearchItem.Content"
args18(6).Value = false
args18(7).Name = "SearchItem.AsianOptions"
args18(7).Value = false
args18(8).Name = "SearchItem.AlgorithmType"
args18(8).Value = 1
args18(9).Name = "SearchItem.SearchFlags"
args18(9).Value = 65536
args18(10).Name = "SearchItem.SearchString"
args18(10).Value = "^\s\s*"
args18(11).Name = "SearchItem.ReplaceString"
args18(11).Value = ""
args18(12).Name = "SearchItem.Locale"
args18(12).Value = 255
args18(13).Name = "SearchItem.ChangedChars"
args18(13).Value = 2
args18(14).Name = "SearchItem.DeletedChars"
args18(14).Value = 2
args18(15).Name = "SearchItem.InsertedChars"
args18(15).Value = 2
args18(16).Name = "SearchItem.TransliterateFlags"
args18(16).Value = 1280
args18(17).Name = "SearchItem.Command"
args18(17).Value = 3
args18(18).Name = "Quiet"
args18(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args18())

rem ----------------------------------------------------------------------
dim args19(18) as new com.sun.star.beans.PropertyValue
args19(0).Name = "SearchItem.StyleFamily"
args19(0).Value = 2
args19(1).Name = "SearchItem.CellType"
args19(1).Value = 0
args19(2).Name = "SearchItem.RowDirection"
args19(2).Value = true
args19(3).Name = "SearchItem.AllTables"
args19(3).Value = false
args19(4).Name = "SearchItem.Backward"
args19(4).Value = false
args19(5).Name = "SearchItem.Pattern"
args19(5).Value = false
args19(6).Name = "SearchItem.Content"
args19(6).Value = false
args19(7).Name = "SearchItem.AsianOptions"
args19(7).Value = false
args19(8).Name = "SearchItem.AlgorithmType"
args19(8).Value = 1
args19(9).Name = "SearchItem.SearchFlags"
args19(9).Value = 65536
args19(10).Name = "SearchItem.SearchString"
args19(10).Value = "\t"
args19(11).Name = "SearchItem.ReplaceString"
args19(11).Value = " "
args19(12).Name = "SearchItem.Locale"
args19(12).Value = 255
args19(13).Name = "SearchItem.ChangedChars"
args19(13).Value = 2
args19(14).Name = "SearchItem.DeletedChars"
args19(14).Value = 2
args19(15).Name = "SearchItem.InsertedChars"
args19(15).Value = 2
args19(16).Name = "SearchItem.TransliterateFlags"
args19(16).Value = 1280
args19(17).Name = "SearchItem.Command"
args19(17).Value = 3
args19(18).Name = "Quiet"
args19(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args19())

rem ----------------------------------------------------------------------
dim args20(18) as new com.sun.star.beans.PropertyValue
args20(0).Name = "SearchItem.StyleFamily"
args20(0).Value = 2
args20(1).Name = "SearchItem.CellType"
args20(1).Value = 0
args20(2).Name = "SearchItem.RowDirection"
args20(2).Value = true
args20(3).Name = "SearchItem.AllTables"
args20(3).Value = false
args20(4).Name = "SearchItem.Backward"
args20(4).Value = false
args20(5).Name = "SearchItem.Pattern"
args20(5).Value = false
args20(6).Name = "SearchItem.Content"
args20(6).Value = false
args20(7).Name = "SearchItem.AsianOptions"
args20(7).Value = false
args20(8).Name = "SearchItem.AlgorithmType"
args20(8).Value = 0
args20(9).Name = "SearchItem.SearchFlags"
args20(9).Value = 65536
args20(10).Name = "SearchItem.SearchString"
args20(10).Value = " - "
args20(11).Name = "SearchItem.ReplaceString"
args20(11).Value = " – "
args20(12).Name = "SearchItem.Locale"
args20(12).Value = 255
args20(13).Name = "SearchItem.ChangedChars"
args20(13).Value = 2
args20(14).Name = "SearchItem.DeletedChars"
args20(14).Value = 2
args20(15).Name = "SearchItem.InsertedChars"
args20(15).Value = 2
args20(16).Name = "SearchItem.TransliterateFlags"
args20(16).Value = 1280
args20(17).Name = "SearchItem.Command"
args20(17).Value = 3
args20(18).Name = "Quiet"
args20(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args20())


rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:GoToStartOfDoc", "", 0, Array())

rem ----------------------------------------------------------------------
rem dispatcher.executeDispatch(document, ".uno:SpellingAndGrammarDialog", "", 0, Array())


end sub



sub Main
end sub