Sub GerarPDFsPorData()
    Dim wsAmostras As Worksheet
    Dim wsLaudo As Worksheet
    Dim ultimaLinha As Long, i As Long
    Dim dataBusca As Date
    Dim dataLinha As Date
    Dim amostra As String, especie As String
    Dim exameDireto As String, metodoSheather As String
    Dim nomeArquivo As String, pastaDestino As String
    Dim contador As Integer
    Dim entradaData As String

    Set wsAmostras = ThisWorkbook.Sheets("Amostras")
    Set wsLaudo = ThisWorkbook.Sheets("Laudo")

    ' Solicita a data desejada
    entradaData = InputBox("Digite a data de recebimento (dd/mm/aaaa):", "Filtrar Amostras por Data")

    If entradaData = "" Then
        MsgBox "Data não informada. Operação cancelada.", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErroData
    dataBusca = CDate(entradaData)
    On Error GoTo 0

    ultimaLinha = wsAmostras.Cells(wsAmostras.Rows.Count, 3).End(xlUp).Row
    pastaDestino = ThisWorkbook.Path
    contador = 0

    If pastaDestino = "" Then
        MsgBox "Salve o arquivo antes de gerar os PDFs.", vbExclamation
        Exit Sub
    End If

    ' Percorre as linhas da planilha Amostras
    For i = 2 To ultimaLinha
        If IsDate(wsAmostras.Cells(i, 3).Value) Then
            dataLinha = CDate(wsAmostras.Cells(i, 3).Value)

            If dataLinha = dataBusca Then
                ' Coleta os dados da linha atual
                amostra = wsAmostras.Cells(i, 1).Value           ' Coluna A
                especie = wsAmostras.Cells(i, 2).Value           ' Coluna B
                exameDireto = wsAmostras.Cells(i, 4).Value       ' Coluna D
                metodoSheather = wsAmostras.Cells(i, 5).Value    ' Coluna E

                ' Preenche a planilha Laudo
                With wsLaudo
                    .Range("C35").Value = amostra
                    .Range("C37").Value = especie
                    .Range("C39").Value = Format(dataLinha, "dd/mm/yyyy")
                    .Range("C43").Value = exameDireto
                    .Range("C45").Value = metodoSheather

                    ' Define área de impressão correta para evitar página em branco
                    .PageSetup.PrintArea = "A1:E49"
                    With .PageSetup
                        .Zoom = False
                        .FitToPagesWide = 1
                        .FitToPagesTall = 1
                        .Orientation = xlPortrait
                    End With
                End With

                ' Define o nome do PDF
                nomeArquivo = pastaDestino & "\Laudo_" & LimparNomeArquivo(amostra) & "_" & Format(dataLinha, "yyyymmdd") & ".pdf"

                ' Exporta para PDF
                wsLaudo.ExportAsFixedFormat Type:=xlTypePDF, Filename:=nomeArquivo, Quality:=xlQualityStandard

                contador = contador + 1
            End If
        End If
    Next i

    If contador > 0 Then
        MsgBox contador & " laudo(s) gerado(s) com sucesso na pasta:" & vbCrLf & pastaDestino, vbInformation
    Else
        MsgBox "Nenhuma amostra encontrada com a data " & Format(dataBusca, "dd/mm/yyyy"), vbExclamation
    End If

    Exit Sub

ErroData:
    MsgBox "Formato de data inválido. Use o formato dd/mm/aaaa.", vbCritical
End Sub

Function LimparNomeArquivo(texto As String) As String
    Dim caracteresInvalidos As Variant
    Dim c As Variant

    caracteresInvalidos = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For Each c In caracteresInvalidos
        texto = Replace(texto, c, "_")
    Next

    LimparNomeArquivo = texto
End Function
