REM  *****  BASIC *****

Sub MacroNJMG

    Dim document            As Object
    Dim dispatcher          As Object
    Dim cursor              As Object
    Dim oText               As Object
    Dim oEnum               As Object
    Dim parags()            As Object
    Dim count               As Integer
    
    ' Search args
    Dim args1(18)           As New com.sun.star.beans.PropertyValue
    Dim args18(18)          As New com.sun.star.beans.PropertyValue
    Dim args2(1)            As New com.sun.star.beans.PropertyValue
    Dim args4(2)            As New com.sun.star.beans.PropertyValue
    Dim args5(7)            As New com.sun.star.beans.PropertyValue
    Dim args6(4)            As New com.sun.star.beans.PropertyValue
    Dim args7(4)            As New com.sun.star.beans.PropertyValue
    Dim args8(2)            As New com.sun.star.beans.PropertyValue

    Dim tables              As Object
    Dim table               As Object
    Dim cellNames()         As String
    Dim cell                As Object
    Dim paraEnum            As Object
    Dim parEnum             As Object
    Dim enumItalico         As Object
    Dim enumItalicoFinal    As Object  ' para reaplicar caput e stricto sensu no fim
    Dim pFinal              As Object
    Dim cursorFinal         As Object

    ' Substitution blocks
    Dim enumSubst           As Object
    Dim p                   As Object
    Dim pSub                As Object
    Dim textoPar            As String
    Dim textoSub            As String
    Dim iSub                As Integer
    Dim i                   As Integer
    Dim k                   As Integer
    Dim pos                 As Long
    Dim romanos()           As String

    ' Para normalizar linhas em branco
    Dim allParas()          As Object
    Dim idx                 As Integer
    Dim j                   As Long
    Dim s                   As String
    Dim isTarget            As Boolean
    Dim blankCount          As Integer

    ' --- Declaracoes para bloco 10 (Belo Horizonte, aos â€¦;â€) ---
    Dim dateParas()         As Object
    Dim dpEnum              As Object
    Dim hp2                 As Object
    Dim dpIdx               As Integer
    Dim blanksAbove         As Integer
    Dim blanksBelow         As Integer
    Dim sUP                 As String

    ' Inicializacao
    document   = ThisComponent.CurrentController.Frame
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dispatcher.executeDispatch(document, ".uno:GoToStartOfDoc", "", 0, Array())

    ' =================================================================================================================================
    ' 1) Normaliza EXATAMENTE UMA linha em branco antes de:
    '    â€¢ â€œO GOVERNADOR DO ESTADO DE MINAS GERAIS,â€
    '    â€¢ â€œDECRETA:â€
    '    â€¢ â€œArt.â€ seguido de nÃºmero
    ' =================================================================================================================================

    idx = 0
    oEnum = ThisComponent.Text.createEnumeration()
    Do While oEnum.hasMoreElements()
        Set p = oEnum.nextElement()
        If p.supportsService("com.sun.star.text.Paragraph") Then
            ReDim Preserve allParas(idx)
            Set allParas(idx) = p
            idx = idx + 1
        End If
    Loop

    For i = UBound(allParas) To 0 Step -1
        s = Trim(allParas(i).getString())
        isTarget = False
        If s <> "" Then
            If Left(s, Len("O GOVERNADOR DO ESTADO DE MINAS GERAIS,")) = "O GOVERNADOR DO ESTADO DE MINAS GERAIS," Then isTarget = True
			If Not isTarget And Left(s, Len("O PRESIDENTE DO ESTADO DE MINAS GERAIS,")) = "O PRESIDENTE DO ESTADO DE MINAS GERAIS," Then isTarget = True
			If Not isTarget And Left(s, Len("O povo do Estado de Minas Gerais")) = "O povo do Estado de Minas Gerais" Then isTarget = True
			If Not isTarget And Left(s, Len("A Mesa da Assembléia Legislativa do Estado de Minas Gerais")) = "A Mesa da Assembléia Legislativa do Estado de Minas Gerais" Then isTarget = True
			If Not isTarget And Left(s, Len("A Mesa da Assembleia Legislativa do Estado de Minas Gerais")) = "A Mesa da Assembleia Legislativa do Estado de Minas Gerais" Then isTarget = True
            If Not isTarget And Left(s, Len("DECRETA:")) = "DECRETA:" Then isTarget = True
            If Not isTarget And Left(s, 4) = "Art." Then
                If Len(s) >= 6 And Mid(s,5,1)=" " And IsNumeric(Mid(s,6,1)) Then isTarget = True
            End If
            If isTarget Then
                blankCount = 0
                j = i - 1
                While j >= 0 And Trim(allParas(j).getString()) = ""
                    blankCount = blankCount + 1
                    j = j - 1
                Wend
                Select Case blankCount
                    Case 0
                        Set cursor = ThisComponent.Text.createTextCursorByRange(allParas(i).Anchor.Start)
                        ThisComponent.Text.insertControlCharacter(cursor, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
                    Case Is > 1
                        For k = j + 2 To i - 1
                            ThisComponent.Text.removeTextContent(allParas(k))
                        Next k
                End Select
            End If
        End If
    Next i

    ' =================================================================================================================================
    ' 2) Remove mÃºltiplos espaÃ§os, tabs e hÃ­fen genÃ©rico
    ' =================================================================================================================================

    args1(0).Name = "SearchItem.StyleFamily":        args1(0).Value = 2
    args1(1).Name = "SearchItem.CellType":           args1(1).Value = 0
    args1(2).Name = "SearchItem.RowDirection":       args1(2).Value = True
    args1(3).Name = "SearchItem.AllTables":          args1(3).Value = False
    args1(4).Name = "SearchItem.Backward":           args1(4).Value = False
    args1(5).Name = "SearchItem.Pattern":            args1(5).Value = False
    args1(6).Name = "SearchItem.Content":            args1(6).Value = False
    args1(7).Name = "SearchItem.AsianOptions":       args1(7).Value = False
    args1(8).Name = "SearchItem.AlgorithmType":      args1(8).Value = 1
    args1(9).Name = "SearchItem.SearchFlags":        args1(9).Value = 65536
    args1(10).Name = "SearchItem.SearchString":      args1(10).Value = "   *"
    args1(11).Name = "SearchItem.ReplaceString":     args1(11).Value = " "
    args1(12).Name = "SearchItem.Locale":            args1(12).Value = 255
    args1(13).Name = "SearchItem.ChangedChars":      args1(13).Value = 2
    args1(14).Name = "SearchItem.DeletedChars":      args1(14).Value = 2
    args1(15).Name = "SearchItem.InsertedChars":     args1(15).Value = 2
    args1(16).Name = "SearchItem.TransliterateFlags":args1(16).Value = 1280
    args1(17).Name = "SearchItem.Command":           args1(17).Value = 3
    args1(18).Name = "Quiet":                        args1(18).Value = True
    dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1())

    For k = 0 To 18
        args18(k) = args1(k)
    Next k
    args18(10).Value = "^\s\s*": args18(11).Value = "":  dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args18())
    args18(10).Value = "\t":     args18(11).Value = " ": dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args18())
    args18(10).Value = " - ":    args18(11).Value = " â€“ ": args18(8).Value = 0
    dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args18())

    ' =================================================================================================================================
    ' 3) Estilo padrÃ£o, alinhamento, margens, fonte e tabelas
    ' =================================================================================================================================

    dispatcher.executeDispatch(document, ".uno:SelectAll", "", 0, Array())
    args2(0).Name = "Template": args2(0).Value = "Estilo padrÃ£o"
    args2(1).Name = "Family":   args2(1).Value = 2
    dispatcher.executeDispatch(document, ".uno:StyleApply", "", 0, args2())

    args4(0).Name = "Alignment.ParagraphAdjustment": args4(0).Value = 2
    args4(1).Name = "Alignment.LastLineAdjustment":  args4(1).Value = 0
    args4(2).Name = "Alignment.ExpandSingleWord":    args4(2).Value = False
    dispatcher.executeDispatch(document, ".uno:Alignment", "", 0, args4())

    args5(0).Name = "LeftRightMargin.LeftMargin":        args5(0).Value = 0
    args5(1).Name = "LeftRightMargin.TextLeftMargin":    args5(1).Value = 0
    args5(2).Name = "LeftRightMargin.RightMargin":       args5(2).Value = 0
    args5(3).Name = "LeftRightMargin.LeftRelMargin":     args5(3).Value = 100
    args5(4).Name = "LeftRightMargin.RightRelMargin":    args5(4).Value = 100
    args5(5).Name = "LeftRightMargin.FirstLineIndent":   args5(5).Value = 2500
    args5(6).Name = "LeftRightMargin.FirstLineRelIdent": args5(6).Value = 100
    args5(7).Name = "LeftRightMargin.AutoFirst":         args5(7).Value = False
    dispatcher.executeDispatch(document, ".uno:LeftRightMargin", "", 0, args5())

    args6(0).Name = "TopBottomMargin.TopMargin":       args6(0).Value = 0
    args6(1).Name = "TopBottomMargin.BottomMargin":    args6(1).Value = 0
    args6(2).Name = "TopBottomMargin.ContextMargin":   args6(2).Value = False
    args6(3).Name = "TopBottomMargin.TopRelMargin":    args6(3).Value = 100
    args6(4).Name = "TopBottomMargin.BottomRelMargin": args6(4).Value = 100
    dispatcher.executeDispatch(document, ".uno:TopBottomMargin", "", 0, args6())

    args7(0).Name = "CharFontName.FamilyName": args7(0).Value = "Times New Roman"
    dispatcher.executeDispatch(document, ".uno:CharFontName", "", 0, args7())
    args8(0).Name = "FontHeight.Height":      args8(0).Value = 12
    dispatcher.executeDispatch(document, ".uno:FontHeight", "", 0, args8())

    tables = ThisComponent.TextTables
    For i = 0 To tables.getCount() - 1
        Set table = tables.getByIndex(i)
        table.HoriOrient = com.sun.star.text.HoriOrientation.NONE
        table.LeftMargin  = 0
        table.RightMargin = 0
        table.setPropertyValue("Width", 10000)

        cellNames = table.getCellNames()
        For k = 0 To UBound(cellNames)
            Set cell = table.getCellByName(cellNames(k))
            cell.VertOrient = com.sun.star.text.VertOrientation.CENTER
            paraEnum = cell.Text.createEnumeration()
            Do While paraEnum.hasMoreElements()
                Set p = paraEnum.nextElement()
                p.CharFontName        = "Times New Roman"
                p.CharHeight          = 10
                p.ParaTopMargin       = 0
                p.ParaBottomMargin    = 0
                p.ParaLeftMargin      = 0
                p.ParaRightMargin     = 0
                p.ParaFirstLineIndent = 0
                p.ParaAdjust          = com.sun.star.style.ParagraphAdjust.CENTER
                Dim spacing As New com.sun.star.style.LineSpacing
                spacing.Mode = 0
                p.ParaLineSpacing     = spacing
            Loop
        Next k
    Next i

    ' =================================================================================================================================
    ' 4) Substitui â€œÂ°â€ por â€œÂºâ€ em Art. e Â§
    ' =================================================================================================================================

    parEnum = ThisComponent.Text.createEnumeration()
    Do While parEnum.hasMoreElements()
        Set p = parEnum.nextElement()
        If p.supportsService("com.sun.star.text.Paragraph") Then
            textoPar = p.getString()
            If InStr(textoPar, "Art. ") > 0 Or InStr(textoPar, "Â§ ") > 0 Then
                textoPar = Replace(textoPar, "Â°", "Âº")
                p.setString(textoPar)
            End If
        End If
    Loop

    ' =================================================================================================================================
    ' 5) Indenta primeiro parÃ¡grafo
    ' =================================================================================================================================

    Set p = ThisComponent.Text.createEnumeration().nextElement()
    If p.supportsService("com.sun.star.text.Paragraph") Then
        p.ParaLeftMargin      = 7620
        p.ParaFirstLineIndent = 0
    End If

    ' =================================================================================================================================
    ' 6) ItÃ¡lico em â€œcaputâ€ em todas as ocorrÃªncias
    ' =================================================================================================================================

    enumItalico = ThisComponent.Text.createEnumeration()
    Do While enumItalico.hasMoreElements()
        Set p = enumItalico.nextElement()
        If p.supportsService("com.sun.star.text.Paragraph") Then
            textoPar = p.getString()
            pos = 1
            Do
                pos = InStr(pos, textoPar, "caput", 1)
                If pos > 0 Then
                    Set cursor = ThisComponent.Text.createTextCursorByRange(p.getStart())
                    cursor.goRight(pos-1, False)
                    cursor.goRight(Len("caput"), True)
                    cursor.CharPosture = com.sun.star.awt.FontSlant.ITALIC
                    pos = pos + Len("caput")
                End If
            Loop While pos > 0
        End If
    Loop

    ' =================================================================================================================================
    ' 7) SubstituiÃ§Ã£o controlada de hÃ­fen em Art., ParÃ¡grafo Ãºnico, Â§ e incisos
    '     (cardinais e ordinais, com um espaÃ§o Ãºnico em volta do â€“ e
    '      garantir um espaÃ§o antes e depois caso nÃ£o existam)
    ' =================================================================================================================================

	Call FixHeadingDash("^\s*(Art\.\s*\d{1,3}(?:º)?)\s*(?:[.\-–—:\)]\s*|\s+)(.*)$")
	' Ex.:  Art. 1º -  ,  Art. 1º.  ,  Art. 10   → Art. 1º – …

	Call FixHeadingDash("^\s*(Par[aá]grafo\s+único|§\s*\d{1,3}(?:º)?)\s*(?:[.\-–—:\)]\s*|\s+)(.*)$")
	' Ex.:  Parágrafo único -  ,  § 1º.  ,  § 10  → Parágrafo único – … / § 1º – …

	Call FixHeadingDash("^\s*([IVXLCDM]{1,6})(?=\s*(?:-|\s|$))(?!\s*–)(?:-|\s*-\s*|\s+)\s*(.*)$", "– \2")
	' Ex.:  I -  ,  II  ,  I)  → I – … / II – …

    ' =================================================================================================================================
	' 8) Negrito quando o parágrafo INICIA com (CAIXA ALTA):
	'    • "O GOVERNADOR DO ESTADO DE MINAS GERAIS,"
	'    • "O PRESIDENTE DO ESTADO DE MINAS GERAIS,"
	'    • "DECRETA:"
	' =================================================================================================================================

	Dim enumBold    As Object
	Dim pBold       As Object
	Dim cursorBold  As Object
	Dim sLine       As String
	Dim sTrim       As String
	Dim lead        As Integer
	Dim prefix      As String

	enumBold = ThisComponent.Text.createEnumeration()
	Do While enumBold.hasMoreElements()
		Set pBold = enumBold.nextElement()
		If pBold.supportsService("com.sun.star.text.Paragraph") Then
			sLine = pBold.getString()
			lead  = Len(sLine) - Len(LTrim$(sLine))   ' nº de espaços à esquerda
			sTrim = LTrim$(sLine)                      ' texto sem espaços à esquerda

			' GOVERNADOR
			prefix = "O GOVERNADOR DO ESTADO DE MINAS GERAIS,"
			If Left$(sTrim, Len(prefix)) = prefix Then
				Set cursorBold = ThisComponent.Text.createTextCursorByRange(pBold.getStart())
				cursorBold.goRight(lead + Len(prefix), True)
				cursorBold.CharWeight = com.sun.star.awt.FontWeight.BOLD
			End If

			' PRESIDENTE
			prefix = "O PRESIDENTE DO ESTADO DE MINAS GERAIS,"
			If Left$(sTrim, Len(prefix)) = prefix Then
				Set cursorBold = ThisComponent.Text.createTextCursorByRange(pBold.getStart())
				cursorBold.goRight(lead + Len(prefix), True)
				cursorBold.CharWeight = com.sun.star.awt.FontWeight.BOLD
			End If

			' DECRETA:
			prefix = "DECRETA:"
			If Left$(sTrim, Len(prefix)) = prefix Then
				Set cursorBold = ThisComponent.Text.createTextCursorByRange(pBold.getStart())
				cursorBold.goRight(lead + Len(prefix), True)
				cursorBold.CharWeight = com.sun.star.awt.FontWeight.BOLD
			End If
		End If
	Loop

    ' =================================================================================================================================
    ' 9) Títulos/seções: centralizar e garantir 1 linha em branco antes e
    '    garantir EXATAMENTE UMA linha em branco antes
    '    (case-insensitive; â€œAnexoâ€ detecta qualquer texto iniciando por â€œAnexoâ€)
    '    E centralizar tambÃ©m o parÃ¡grafo imediatamente abaixo
    ' =================================================================================================================================

	Dim headings()  As Object
	Dim headEnum    As Object
	Dim hp          As Object
	Dim hidx        As Integer
	Dim blanks      As Integer
	Dim prevI       As Long
	Dim kw          As Variant
	Dim match       As Boolean
	Dim sNorm       As String
	Dim keywords    As Variant

	' Palavras-alvo (normalizadas, sem acentos/cedilha)
	keywords = Array("LIVRO","TITULO","CAPITULO","SECAO","SUBSECAO","ANEXO","TABELA")

	hidx = 0
	Set headEnum = ThisComponent.Text.createEnumeration()
	Do While headEnum.hasMoreElements()
		Set hp = headEnum.nextElement()
		If hp.supportsService("com.sun.star.text.Paragraph") Then
			ReDim Preserve headings(hidx)
			Set headings(hidx) = hp
			hidx = hidx + 1
		End If
	Loop

	For i = UBound(headings) To 0 Step -1
		s = Trim(headings(i).getString())
		sNorm = NormalizeHeadingText(s)  ' remove aspas iniciais, UCase e diacríticos

		match = False
		For Each kw In keywords
			' Casa no início do parágrafo (ex.: LIVRO I, "título,", ANEXO I, etc.)
			If Left(sNorm, Len(kw)) = kw Then
				match = True
				Exit For            ' sai do For Each kw
			End If
		Next kw

		If match Then
			' --- garantir EXATAMENTE UMA linha em branco antes ---
			blanks = 0
			prevI = i - 1
			While prevI >= 0 And Trim(headings(prevI).getString()) = ""
				blanks = blanks + 1
				prevI = prevI - 1
			Wend
			Select Case blanks
				Case 0
					Set cursor = ThisComponent.Text.createTextCursorByRange(headings(i).Anchor.Start)
					ThisComponent.Text.insertControlCharacter(cursor, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
				Case Is > 1
					For k = prevI + 2 To i - 1
						ThisComponent.Text.removeTextContent(headings(k))
					Next k
			End Select

			' --- centraliza o cabeçalho ---
			With headings(i)
				.ParaLeftMargin      = 0
				.ParaFirstLineIndent = 0
				.ParaAdjust          = com.sun.star.style.ParagraphAdjust.CENTER
			End With
			
			' --- centraliza o PRIMEIRO parágrafo NÃO VAZIO abaixo ---
			Dim nextIdx As Long : nextIdx = i + 1
			Do While nextIdx <= UBound(headings) And Trim(headings(nextIdx).getString()) = ""
				nextIdx = nextIdx + 1
			Loop
			If nextIdx <= UBound(headings) Then
				With headings(nextIdx)
					.ParaLeftMargin      = 0
					.ParaFirstLineIndent = 0
					.ParaAdjust          = com.sun.star.style.ParagraphAdjust.CENTER
				End With
			End If
		End If
	Next i

    ' =================================================================================================================================
    ' 10) Data â€œBelo Horizonte, aos nn de maio de nnnn;â€
    '     â€“ exata um blank antes e um blank depois (idempotente)
    ' =================================================================================================================================

    dpIdx = 0
    Set dpEnum = ThisComponent.Text.createEnumeration()
    Do While dpEnum.hasMoreElements()
        Set hp2 = dpEnum.nextElement()
        If hp2.supportsService("com.sun.star.text.Paragraph") Then
            ReDim Preserve dateParas(dpIdx)
            Set dateParas(dpIdx) = hp2
            dpIdx = dpIdx + 1
        End If
    Loop

    For i = UBound(dateParas) To 0 Step -1
        s = Trim(dateParas(i).getString())
        sUP = UCase(s)
        If sUP Like "BELO HORIZONTE, AOS ## DE MAIO DE ####;*" _
           Or sUP Like "BELO HORIZONTE, AOS # DE MAIO DE ####;*" Then

            blanksAbove = 0
            j = i - 1
            While j >= 0 And Trim(dateParas(j).getString()) = ""
                blanksAbove = blanksAbove + 1
                j = j - 1
            Wend
            Select Case blanksAbove
                Case 0
                    Set cursor = ThisComponent.Text.createTextCursorByRange(dateParas(i).Anchor.Start)
                    ThisComponent.Text.insertControlCharacter(cursor, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
                Case Is > 1
                    For k = j + 2 To i - 1
                        ThisComponent.Text.removeTextContent(dateParas(k))
                    Next k
            End Select

            blanksBelow = 0
            j = i + 1
            While j <= UBound(dateParas) And Trim(dateParas(j).getString()) = ""
                blanksBelow = blanksBelow + 1
                j = j + 1
            Wend
            Select Case blanksBelow
                Case 0
                    Set cursor = ThisComponent.Text.createTextCursorByRange(dateParas(i).Anchor.End)
                    ThisComponent.Text.insertControlCharacter(cursor, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
                Case Is > 1
                    For k = i + 2 To j - 1
                        ThisComponent.Text.removeTextContent(dateParas(k))
                    Next k
            End Select

        End If
    Next

    ' =================================================================================================================================
    ' 11) Reaplicar itÃ¡lico em â€œcaputâ€ e em â€œstricto sensuâ€ apÃ³s todas as alteraÃ§Ãµes
    ' =================================================================================================================================

    enumItalicoFinal = ThisComponent.Text.createEnumeration()
    Do While enumItalicoFinal.hasMoreElements()
        Set pFinal = enumItalicoFinal.nextElement()
        If pFinal.supportsService("com.sun.star.text.Paragraph") Then
            textoPar = pFinal.getString()
            ' caput
            pos = 1
            Do
                pos = InStr(pos, textoPar, "caput", 1)
                If pos > 0 Then
                    Set cursorFinal = ThisComponent.Text.createTextCursorByRange(pFinal.getStart())
                    cursorFinal.goRight(pos-1, False)
                    cursorFinal.goRight(Len("caput"), True)
                    cursorFinal.CharPosture = com.sun.star.awt.FontSlant.ITALIC
                    pos = pos + Len("caput")
                End If
            Loop While pos > 0
            ' stricto sensu
            pos = 1
            Do
                pos = InStr(pos, textoPar, "stricto sensu", 1)
                If pos > 0 Then
                    Set cursorFinal = ThisComponent.Text.createTextCursorByRange(pFinal.getStart())
                    cursorFinal.goRight(pos-1, False)
                    cursorFinal.goRight(Len("stricto sensu"), True)
                    cursorFinal.CharPosture = com.sun.star.awt.FontSlant.ITALIC
                    pos = pos + Len("stricto sensu")
                End If
            Loop While pos > 0
        End If
    Loop
	
	' =================================================================================================================================
	' 12) Local + data (qualquer mês): garantir EXATAMENTE 1 linha
	'  	  em branco antes e 1 depois (idempotente)
	' =================================================================================================================================
	Dim dateParas2()  As Object
	Dim dpEnum2       As Object
	Dim hp3           As Object
	Dim dpIdx2        As Integer

	dpIdx2 = 0
	Set dpEnum2 = ThisComponent.Text.createEnumeration()
	Do While dpEnum2.hasMoreElements()
		Set hp3 = dpEnum2.nextElement()
		If hp3.supportsService("com.sun.star.text.Paragraph") Then
			ReDim Preserve dateParas2(dpIdx2)
			Set dateParas2(dpIdx2) = hp3
			dpIdx2 = dpIdx2 + 1
		End If
	Loop

	For i = UBound(dateParas2) To 0 Step -1
		s = Trim(dateParas2(i).getString())
		If IsDateLineLocalData(s) Then
			' — acima —
			blanksAbove = 0
			j = i - 1
			While j >= 0 And Trim(dateParas2(j).getString()) = ""
				blanksAbove = blanksAbove + 1
				j = j - 1
			Wend
			Select Case blanksAbove
				Case 0
					Set cursor = ThisComponent.Text.createTextCursorByRange(dateParas2(i).Anchor.Start)
					ThisComponent.Text.insertControlCharacter(cursor, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
				Case Is > 1
					For k = j + 2 To i - 1
						ThisComponent.Text.removeTextContent(dateParas2(k))
					Next k
			End Select

			' — abaixo —
			blanksBelow = 0
			j = i + 1
			While j <= UBound(dateParas2) And Trim(dateParas2(j).getString()) = ""
				blanksBelow = blanksBelow + 1
				j = j + 1
			Wend
			Select Case blanksBelow
				Case 0
					Set cursor = ThisComponent.Text.createTextCursorByRange(dateParas2(i).Anchor.End)
					ThisComponent.Text.insertControlCharacter(cursor, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
				Case Is > 1
					For k = i + 2 To j - 1
						ThisComponent.Text.removeTextContent(dateParas2(k))
					Next k
			End Select
		End If
	Next i
	
	' =================================================================================================================================
	' 13) Parágrafo composto APENAS por aspa final + ponto (". ou ”.)
	'     → alinhar à direita (não altera o texto; idempotente)
	' =================================================================================================================================

	Dim enumQuote  As Object
	Dim pQ         As Object
	Dim lineTxt    As String
	Dim st         As String
	Dim ns         As String

	Set enumQuote = ThisComponent.Text.createEnumeration()
	Do While enumQuote.hasMoreElements()
		Set pQ = enumQuote.nextElement()
		If pQ.supportsService("com.sun.star.text.Paragraph") Then
			lineTxt = pQ.getString()
			st = Trim(lineTxt)

			' Remove espaços comuns e NBSP só para comparar (não altera o conteúdo real)
			ns = Replace(st, " ", "")
			ns = Replace(ns, Chr$(160), "")  ' NBSP

			' Aceita aspa ASCII (") e aspa tipográfica de FECHAMENTO (”)
			If ns = Chr$(34) & "." Or ns = "”." Then
				pQ.ParaAdjust = com.sun.star.style.ParagraphAdjust.RIGHT
			End If
		End If
	Loop	

    ' =================================================================================================================================
    ' 14) Alerta de conclusÃ£o
    ' =================================================================================================================================

		MsgBox "A macro foi executada com sucesso!", 64, "Tudo pronto!"
	End Sub

    ' =================================================================================================================================
    ' A)Helper: centraliza cabeçalhos e o parágrafo imediatamente abaixo (regex)
    ' =================================================================================================================================

	Sub CentralizarTitulosEProximo_Regex()
		Dim oDoc As Object, oSearchDesc As Object, oFound As Object
		Dim oCur As Object, oNext As Object
		On Error GoTo EH

		oDoc = ThisComponent

		oSearchDesc = oDoc.createSearchDescriptor()
		With oSearchDesc
			' Início de parágrafo, tolera aspas/“lixo” de OCR e espaços antes do termo:
			' Título/Capítulo/Seção/Subseção/Anexo/Tabela (maiús/minús, com ou sem acento/cedilha)
			.SearchString = "^\s*[""“”«»„‟]?\s*(T[íi]tulo|Cap[íi]tulo|Se[cç]ão|Subse[cç]ão|Anexo|Tabela)\b.*"
			.SearchRegularExpression = True
			.SearchCaseSensitive = False
		End With

		oFound = oDoc.findFirst(oSearchDesc)
		Do While Not IsNull(oFound)
			' Centraliza o cabeçalho encontrado
			oCur = oDoc.Text.createTextCursorByRange(oFound)
			oCur.ParaAdjust = com.sun.star.style.ParagraphAdjust.CENTER

			' Centraliza o parágrafo imediatamente abaixo (se existir)
			oNext = oDoc.Text.createTextCursorByRange(oFound.End)
			oNext.gotoEndOfParagraph(False)
			If oNext.goRight(1, False) Then
				oNext.gotoEndOfParagraph(True)
				oNext.ParaAdjust = com.sun.star.style.ParagraphAdjust.CENTER
			End If

			oFound = oDoc.findNext(oFound.End, oSearchDesc)
		Loop
		Exit Sub
	EH:
		MsgBox "Erro em CentralizarTitulosEProximo_Regex: " & Err & " - " & Error$, 16, "Macro"	
	End Sub

    ' =================================================================================================================================
    ' B) Remove aspas iniciais (ASCII e tipográficas), sobe para maiúsculas e tira acentos/ç.
    ' =================================================================================================================================

	Function NormalizeHeadingText(ByVal s As String) As String
		s = Trim(s)
		Do While Len(s) > 0 And IsLeadingQuote(Left$(s,1))
			s = Mid$(s, 2)
			s = LTrim$(s)
		Loop
		s = UCase$(s)
		s = RemoveDiacritics(s)
		NormalizeHeadingText = s
	End Function

	Function IsLeadingQuote(ch As String) As Boolean
		' Aspas comuns e tipográficas de abertura (inclui ASCII ")
		Dim q As String
		q = """" & "“”«»„‟‚‘’‹›"
		IsLeadingQuote = (InStr(q, ch) > 0)
	End Function

	Function RemoveDiacritics(ByVal s As String) As String
		' Considera entrada já em UPPER
		s = Replace(s, "Á", "A")
		s = Replace(s, "À", "A")
		s = Replace(s, "Â", "A")
		s = Replace(s, "Ã", "A")
		s = Replace(s, "Ä", "A")
		s = Replace(s, "É", "E")
		s = Replace(s, "È", "E")
		s = Replace(s, "Ê", "E")
		s = Replace(s, "Ë", "E")
		s = Replace(s, "Í", "I")
		s = Replace(s, "Ì", "I")
		s = Replace(s, "Î", "I")
		s = Replace(s, "Ï", "I")
		s = Replace(s, "Ó", "O")
		s = Replace(s, "Ò", "O")
		s = Replace(s, "Ô", "O")
		s = Replace(s, "Õ", "O")
		s = Replace(s, "Ö", "O")
		s = Replace(s, "Ú", "U")
		s = Replace(s, "Ù", "U")
		s = Replace(s, "Û", "U")
		s = Replace(s, "Ü", "U")
		s = Replace(s, "Ç", "C")
		RemoveDiacritics = s
	End Function
	Sub FixHeadingDash(ByVal pattern As String)
		Dim oDesc As Object
		oDesc = ThisComponent.createReplaceDescriptor()
		With oDesc
			.SearchRegularExpression = True
			.SearchCaseSensitive = False
			.SearchString = pattern
			.ReplaceString = "$1 – $2"   ' meia-risca com um espaço antes/depois
		End With
		ThisComponent.replaceAll(oDesc)
	End Sub

    ' =================================================================================================================================
	' C) Detecta “<local>, aos <dia> de <mês> de <ano>; ...”
	' =================================================================================================================================

	Function IsDateLineLocalData(ByVal s As String) As Boolean
		Dim su As String, months As Variant, m As Variant, pos As Long, tail As String

		su = RemoveDiacritics(UCase(Trim(s)))  ' normaliza: caixa alta, sem acentos/ç

		' Exige ponto-e-vírgula (parte comemorativa) e a sequência ", AOS "
		If InStr(su, ";") = 0 Then IsDateLineLocalData = False: Exit Function
		pos = InStr(su, ", AOS ")
		If pos = 0 Then IsDateLineLocalData = False: Exit Function

		' Parte após ", AOS "
		tail = Mid$(su, pos + Len(", AOS "))

		' Meses aceitos (normalizados)
		months = Array("JANEIRO","FEVEREIRO","MARCO","ABRIL","MAIO","JUNHO","JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO")

		' Checa “# DE <MÊS> DE ####;*” (dia de 1 ou 2 dígitos)
		For Each m In months
			If tail Like "# DE " & m & " DE ####;*" Or tail Like "## DE " & m & " DE ####;*" Then
				IsDateLineLocalData = True
				Exit Function
			End If
		Next m

		IsDateLineLocalData = False

	End Function
