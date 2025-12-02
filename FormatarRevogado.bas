REM  *****  BASIC  *****

Sub FormatarRevogado
    Dim oDoc As Object
    Dim oText As Object
    Dim oEnum As Object
    Dim oPara As Object
    Dim listaParagrafos() As Object
    Dim i As Long, total As Long

    oDoc = ThisComponent
    oText = oDoc.Text
    oEnum = oText.createEnumeration

    ' Coleta todos os parágrafos
    Do While oEnum.hasMoreElements
        oPara = oEnum.nextElement
        If oPara.supportsService("com.sun.star.text.Paragraph") Then
            total = total + 1
            ReDim Preserve listaParagrafos(total - 1)
            listaParagrafos(total - 1) = oPara
        End If
    Loop

    Dim iStart As Long, iEnd As Long
    Dim achouRevogado As Boolean
    Dim achouAspaAbertura As Boolean
    Dim achouAspaFechamento As Boolean
    Dim notaRevogacao As String

    For i = 0 To total - 1
        Dim sText As String
        sText = listaParagrafos(i).getString()

        ' Verifica se encontrou "Dispositivo revogado:"
        If InStr(sText, "Dispositivo revogado:") > 0 Then
            achouRevogado = True

            ' Extrai e remove completamente a linha anterior com a nota entre parênteses
            notaRevogacao = ""
            If i > 0 Then
                Dim linhaAnterior As String
                linhaAnterior = Trim(listaParagrafos(i - 1).getString())
                Dim posA As Long, posB As Long
                posA = InStr(linhaAnterior, "(")
                posB = InStrRevCompat(linhaAnterior, ")")
                If posA > 0 And posB > posA Then
                    notaRevogacao = Mid(linhaAnterior, posA, posB - posA + 1)
                End If

                ' Remove completamente o parágrafo da linha anterior (elimina linha em branco acima da nota)
                Dim oCursorNota As Object
                oCursorNota = oText.createTextCursorByRange(listaParagrafos(i - 1).getStart())
                oCursorNota.gotoRange(listaParagrafos(i - 1).getEnd(), True)
                oText.removeTextContent(oCursorNota.TextParagraph)
            End If

            ' Remove a linha "Dispositivo revogado:"
            Dim oCursorRemover As Object
            oCursorRemover = oText.createTextCursorByRange(listaParagrafos(i).getStart())
            oCursorRemover.gotoRange(listaParagrafos(i).getEnd(), True)
            oText.removeTextContent(oCursorRemover.TextParagraph)

            ' Aplica tachado entre aspas (multilinha)
            iStart = i + 1
            achouAspaAbertura = False
            achouAspaFechamento = False

            For iEnd = iStart To total - 1
                If iEnd >= UBound(listaParagrafos) + 1 Then Exit For

                Dim sLinha As String
                sLinha = listaParagrafos(iEnd).getString()

                ' Aspa de abertura
                If Not achouAspaAbertura Then
                    If InStr(sLinha, "“") > 0 Or InStr(sLinha, """") > 0 Then
                        achouAspaAbertura = True
                        sLinha = Replace(sLinha, "“", "")
                        sLinha = Replace(sLinha, """", "", , 1)
                        listaParagrafos(iEnd).setString(sLinha)
                    End If
                End If

                ' Tachar
                If achouAspaAbertura Then
                    Dim oCursor As Object
                    oCursor = oDoc.Text.createTextCursorByRange(listaParagrafos(iEnd).getStart())
                    oCursor.gotoRange(listaParagrafos(iEnd).getEnd(), True)
                    oCursor.CharStrikeout = com.sun.star.awt.FontStrikeout.SINGLE
                End If

                ' Aspa de fechamento
                If achouAspaAbertura Then
                    If InStr(sLinha, "”") > 0 Or (InStrRevCompat(sLinha, """") > InStr(sLinha, """")) Then
                        achouAspaFechamento = True
                        sLinha = listaParagrafos(iEnd).getString()
                        sLinha = Replace(sLinha, "”", "")
                        sLinha = ReplaceRev(sLinha, """", "")
                        listaParagrafos(iEnd).setString(sLinha)
                        Exit For
                    End If
                End If
            Next iEnd

            ' Insere a nota SEM tachado após o trecho revogado
            If achouAspaAbertura And achouAspaFechamento And notaRevogacao <> "" Then
                Dim oInsertPos As Object
                oInsertPos = listaParagrafos(iEnd).getEnd()

                ' Insere quebra de linha com cursor limpo
                oText.insertControlCharacter(oInsertPos, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, False)
                Dim novoCursor As Object
                novoCursor = oText.createTextCursorByRange(oInsertPos)
                novoCursor.CharStrikeout = com.sun.star.awt.FontStrikeout.NONE
                oText.insertString(oInsertPos, notaRevogacao, False)
            End If

            achouRevogado = False
            achouAspaAbertura = False
            achouAspaFechamento = False
        End If
    Next i
End Sub

Function InStrRevCompat(sTexto As String, sProcurado As String) As Long
    Dim i As Long
    For i = Len(sTexto) - Len(sProcurado) + 1 To 1 Step -1
        If Mid(sTexto, i, Len(sProcurado)) = sProcurado Then
            InStrRevCompat = i
            Exit Function
        End If
    Next i
    InStrRevCompat = 0
End Function

Function ReplaceRev(sTexto As String, sProcurado As String, sSubstituto As String) As String
    Dim pos As Long
    pos = InStrRevCompat(sTexto, sProcurado)
    If pos > 0 Then
        ReplaceRev = Left(sTexto, pos - 1) & sSubstituto & Mid(sTexto, pos + Len(sProcurado))
    Else
        ReplaceRev = sTexto
    End If
End Function
