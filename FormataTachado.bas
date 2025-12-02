REM  *****  BASIC  *****

Sub FormatarTachado
    Dim oDoc As Object
    oDoc = ThisComponent
    Dim oTexto As Object
    oTexto = oDoc.Text

    ' 1. Tachar conteúdo fora das tabelas
    Dim oParagrafos As Object
    oParagrafos = oTexto.createEnumeration()

    While oParagrafos.hasMoreElements()
        Dim oParagrafo As Object
        oParagrafo = oParagrafos.nextElement()

        If oParagrafo.supportsService("com.sun.star.text.Paragraph") Then
            TacharParagrafoSeNecessario oParagrafo
        End If
    Wend

    ' 2. Tachar conteúdo dentro das tabelas
    Dim oTables As Object
    oTables = oDoc.TextTables
    Dim i As Integer

    For i = 0 To oTables.Count - 1
        Dim oTable As Object
        oTable = oTables.getByIndex(i)

        Dim j As Integer, k As Integer
        For j = 0 To oTable.Rows.Count - 1
            For k = 0 To oTable.Columns.Count - 1
                Dim cellName As String
                cellName = Chr(65 + k) & (j + 1)
                Dim oCell As Object
                oCell = oTable.getCellByName(cellName)

                Dim oEnum As Object
                oEnum = oCell.createEnumeration()
                While oEnum.hasMoreElements()
                    Dim oElement As Object
                    oElement = oEnum.nextElement()
                    If oElement.supportsService("com.sun.star.text.Paragraph") Then
                        TacharParagrafoSeNecessario oElement
                    End If
                Wend
            Next k
        Next j
    Next i
End Sub

Sub TacharParagrafoSeNecessario(oParagrafo As Object)
    Dim sTexto As String
    sTexto = oParagrafo.getString()

    ' Remove espaços e tabulações do início e fim
    Dim sClean As String
    sClean = Trim(Replace(sTexto, Chr(9), ""))

    Dim bTachar As Boolean
    bTachar = True ' Por padrão, tachar tudo

    ' Exceções
    If (sClean = String(Len(sClean), "=")) _
        Or (InStr(UCase(sClean), "DATA DA ÚLTIMA ATUALIZAÇÃO:") > 0) Then
        bTachar = False
    ElseIf (Left(sClean,1) = "(" And Right(sClean,1) = ")") Then
        ' Parágrafo todo entre parênteses
        Dim conteudoInterno As String
        conteudoInterno = LCase(Trim(Mid(sClean, 2, Len(sClean) - 2)))

        ' ⚠️ Nova verificação aprimorada:
        If Left(conteudoInterno, 16) = "a que se refere" _
            Or Left(conteudoInterno, 17) = "à que se refere" Then
            bTachar = True ' Força o tachado
        Else
            bTachar = False ' Deixa sem tachado
        End If
    End If

    ' Aplicar tachado se necessário
    If bTachar Then
        Dim oCursor As Object
        oCursor = oParagrafo.getText().createTextCursorByRange(oParagrafo.getStart())
        oCursor.gotoRange(oParagrafo.getEnd(), True)
        oCursor.CharStrikeout = com.sun.star.awt.FontStrikeout.SINGLE
    End If
End Sub
