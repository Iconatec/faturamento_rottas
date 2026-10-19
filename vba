Sub apura_rottas()

Dim ws As Worksheet
Dim ws2 As Worksheet
Dim ws3 As Worksheet
Dim ws4 As Worksheet
Dim ws5 As Worksheet
Dim ws6 As Worksheet
Dim ws7 As Worksheet
Dim ws8 As Worksheet

Set ws = ActiveWorkbook.Sheets("Base")
Set ws2 = ActiveWorkbook.Sheets("Empreendimento")
Set ws3 = ActiveWorkbook.Sheets("Rendas")
Set ws4 = ActiveWorkbook.Sheets("Empreendimento")
Set ws5 = Nothing
Set ws6 = ActiveWorkbook.Sheets("CEF")
Set ws7 = ActiveWorkbook.Sheets("ITAU")
Set ws8 = ActiveWorkbook.Sheets("GRAFENO")

Dim i As Long
Dim j As Long
Dim ultimaLinha As Long
Dim ultimaLinhaCEF As Long
Dim ultimaLinhaITAU As Long
Dim ultimaLinhaGRAFENO As Long
Dim ultimaLinhaApoio As Long
Dim ultimaLinhaCheck As Long
Dim ultimacef As Long
Dim contador As Long
Dim linha As Long

Dim emp As String
Dim empAtual As String
Dim txt As String
Dim pos As Long
Dim total As Long
Dim chave As String
Dim historicoGrafeno As String

Dim dict As Object
Dim dictDatas As Object
Dim dictCEFChaves As Object
Dim dictITAUChaves As Object
Dim dictGRAFENOChaves As Object

Dim dataVal As Variant
Dim valor As Double
Dim dataNum As Long

Dim inicio As Double
Dim fim As Double
Dim tempo As Double

Dim histItau As String
Dim histApoio As String
Dim prefixoItau As String
Dim prefixoApoio As String
Dim achou As Boolean

Dim kEmp As Variant
Dim kData As Variant
Dim arrEmp As Object
Dim arrDatas As Object

Dim ultimaLinhaEmpI As Long
Dim empresaLista As String
Dim dictEmpresasLista As Object
Dim chaveITAU As Variant
Dim partesChave As Variant

Dim resultadoEmp As Variant
Dim resultadoRenda As Variant

inicio = Timer
Application.ScreenUpdating = False

With ws
    .Cells(1, "E").Value = "Data"
    .Cells(1, "Q").Value = "Valor"
    .Columns("A:B").Insert
    .Cells(1, "A").Value = "Ordem"
    .Cells(1, "B").Value = "Empreendimento"
End With

'Fazer o proc do empreendimento com a filial
ultimaLinha = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row - 2
contador = 1

For i = 2 To ultimaLinha
    ws.Cells(i, "A").Value = contador
    contador = contador + 1

    If Left(ws.Cells(i, "C").Value, 14) = "Empreendimento" Then
        ws.Cells(i, "B").Value = Application.XLookup(ws.Cells(i, "C").Value, Worksheets("Empreendimento").Columns("A"), _
                                                     Worksheets("Empreendimento").Columns("B"), "VERIFICAR")
    Else
        ws.Cells(i, "B").Value = ws.Cells(i - 1, "B").Value
    End If
Next i

ws.Cells(2, "B").Value = ""
ws.Columns("D:E").Insert
    
'Contagem de empreendimentos em branco
On Error Resume Next
    totalEmpreendimentos = Application.WorksheetFunction.CountIf(ws.Range("B2:B" & ultimaLinha), "VERIFICAR")

If totalEmpreendimentos > 1 Then
    For i = 2 To ultimaLinha
        If ws.Cells(i, 2).Value = "VERIFICAR" And ws.Cells(i, "J").Value <> "" Then
            ws.Range("A" & i & ":AA" & i).Interior.Color = RGB(251, 226, 213)
        End If
    Next i
End If

On Error GoTo 0

'Corrigir os nomes, fazer o proc da % do RET e fazer o proc do CNPJ RET
For i = 2 To ultimaLinha
    txt = ws.Cells(i, "C").Value

    If InStr(txt, "-") > 0 Then
        pos = InStrRev(txt, "-")
        ws.Cells(i, "E").Value = Trim(Mid(txt, pos + 1))
        ws.Cells(i, "E").Value = Replace(ws.Cells(i, "E").Value, "-", "")
        ws.Cells(i, "E").Value = Replace(ws.Cells(i, "E").Value, "  ", " ")
        ws.Cells(i, "D").Value = Replace(ws.Cells(i, "E").Value, " ", "")

        ws.Cells(i, "Z").NumberFormat = "General"
        ws.Cells(i, "Z").Formula2 = "=XLOOKUP(D" & i & ",Rendas!B:B,Rendas!E:E,""VERIFICAR"")"
        ws.Cells(i, "Z").Calculate
        ws.Cells(i, "Z").Value = ws.Cells(i, "Z").Value
        ws.Cells(i, "Z").NumberFormat = "0%"
        
        'Checar CNPJ RET
        ws.Cells(i, "AA").Formula2 = "=XLOOKUP(B" & i & ",Empreendimento!B:B,Empreendimento!D:D,""SEM RET"")"
        ws.Cells(i, "AA").Value = ws.Cells(i, "AA").Value
        
    End If
Next i

'Verificar os empreendimentos de alto padrão
For i = 2 To ultimaLinha
    If Trim(UCase(ws.Cells(i, "Z").Value)) = "VERIFICAR" Then
        ws.Cells(i, "Z").NumberFormat = "General"
        ws.Cells(i, "Z").Formula2 = "=XLOOKUP(B" & i & ",Empreendimento!B:B,Empreendimento!C:C,""VERIFICAR"")"
        ws.Cells(i, "Z").Calculate
        ws.Cells(i, "Z").Value = ws.Cells(i, "Z").Value

        If Trim(CStr(ws.Cells(i, "Z").Value)) = "" Then
            ws.Cells(i, "Z").Value = "VERIFICAR"
        End If

        ws.Cells(i, "Z").NumberFormat = "0%"
    End If
Next i

'Verificação dos termos
ws.Columns("C").Insert

'Preencher os termos
For i = 2 To ultimaLinha
    If InStr(ws.Cells(i, "D").Value, "Termo") > 0 Or InStr(ws.Cells(i, "D").Value, "termo") > 0 Then
        ws.Cells(i, "C").Value = ws.Cells(i, "D").Value
    Else
        ws.Cells(i, "C").Value = ws.Cells(i - 1, "C").Value
    End If
Next i

    'Pintar as linhas dos termos Documentação, IPTU Saltinho e Juros de obra
    For i = 2 To ultimaLinha
        If ws.Cells(i, "C").Value = "Termo: Documentação" Or _
           ws.Cells(i, "C").Value = "Termo: Juros de Obra" Or _
           ws.Cells(i, "C").Value = "Termo: IPTU Saltinho" Then
            ws.Rows(i).Interior.Color = RGB(173, 216, 230)
            ws.Cells(i, "AA").Value = "Termos juro e doc"
        End If
    Next i

    'Criar a aba de conferência
    Set dict = CreateObject("Scripting.Dictionary")
    ultimaLinha = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row

    'Montar dicionário do boletim
    For i = 2 To ultimaLinha
        If IsDate(ws.Cells(i, "J").Value) Then
            emp = Trim(UCase(ws.Cells(i, "B").Value))
            dataNum = CLng(Int(CDate(ws.Cells(i, "J").Value)))
            valor = ws.Cells(i, "V").Value

            If Not dict.Exists(emp) Then
                Set dictDatas = CreateObject("Scripting.Dictionary")
                dict.Add emp, dictDatas
            End If

            Set dictDatas = dict(emp)

            If dictDatas.Exists(dataNum) Then
                dictDatas(dataNum) = dictDatas(dataNum) + valor
            Else
                dictDatas.Add dataNum, valor
            End If
        End If
    Next i

    'Criar / limpar aba CHECK
    On Error Resume Next
    Set ws5 = Worksheets("CHECK")
    On Error GoTo 0

    If ws5 Is Nothing Then
        Set ws5 = Worksheets.Add
        ws5.Name = "CHECK"
    Else
        ws5.Cells.Clear
    End If

    'Saída dos dados do boletim
    linha = 2
    Set arrEmp = CreateObject("System.Collections.ArrayList")

    For Each kEmp In dict.Keys
        arrEmp.Add kEmp
    Next kEmp

    arrEmp.Sort

    For Each kEmp In arrEmp
        Set dictDatas = dict(kEmp)
        Set arrDatas = CreateObject("System.Collections.ArrayList")

        For Each kData In dictDatas.Keys
            arrDatas.Add kData
        Next kData

        arrDatas.Sort

        For Each kData In arrDatas
            ws5.Cells(linha, "A").Value = kEmp
            ws5.Cells(linha, "B").Value = CDate(kData)
            ws5.Cells(linha, "C").Value = dictDatas(kData)
            linha = linha + 1
        Next kData
    Next kEmp

    'Verificar se algum cliente ficou sem % e sem RET
    ultimaLinha = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row - 2
    ws.Cells(1, "AA").Value = "Ret"

    On Error Resume Next
    total = Application.WorksheetFunction.CountIf(ws.Range("AA2:AA" & ultimaLinha), "VERIFICAR")
    totalRET = Application.WorksheetFunction.CountIf(ws.Range("AB2:AB" & ultimaLinha), "0")
    If total > 1 Then
        ws.Range("A1:AA" & ultimaLinha).AutoFilter Field:=27, Criteria1:="VERIFICAR"
    End If
    On Error GoTo 0

    'Preencher status da guia ITAU com base no histórico
    ultimaLinhaITAU = ws7.Cells(ws7.Rows.Count, "D").End(xlUp).Row
    ultimaLinhaApoio = Worksheets("Apoiocef").Cells(Worksheets("Apoiocef").Rows.Count, "L").End(xlUp).Row

    For i = 2 To ultimaLinhaITAU
        histItau = Trim(UCase(ws7.Cells(i, "D").Value))
        achou = False

        If histItau <> "" Then
            prefixoItau = Split(Replace(Replace(histItau, ".", " "), "-", " "), " ")(0)

            For j = 2 To ultimaLinhaApoio
                histApoio = Trim(UCase(Worksheets("Apoiocef").Cells(j, "L").Value))

                If histApoio <> "" Then
                    prefixoApoio = Split(Replace(Replace(histApoio, ".", " "), "-", " "), " ")(0)

                    If prefixoItau = prefixoApoio Then
                        ws7.Cells(i, "I").Value = Worksheets("Apoiocef").Cells(j, "M").Value
                        achou = True
                        Exit For
                    End If
                End If
            Next j
        End If

        If Not achou Then
            ws7.Cells(i, "I").Value = "VERIFICAR"
        End If
    Next i

    'Corrigir o nome dos extratos na coluna A
    ultimacef = ws6.Cells(ws6.Rows.Count, "D").End(xlUp).Row

    For i = 2 To ultimacef
        ws6.Cells(i, "A").Value = Replace(ws6.Cells(i, "A").Value, ".pdf", "")

        If ws6.Cells(i, "J").Value = "Débito" Then
            ws6.Cells(i, "M").Value = "nok"
        Else
            ws6.Cells(i, "M").Value = Application.XLookup(ws6.Cells(i, "G").Value, Worksheets("Apoiocef").Columns("E"), _
                                                          Worksheets("Apoiocef").Columns("F"), "VERIFICAR")
        End If

        ws6.Cells(i, "B").Value = Application.XLookup(ws6.Cells(i, "A").Value, Worksheets("Apoiocef").Columns("A"), _
                                                      Worksheets("Apoiocef").Columns("C"), "VERIFICAR")
    Next i

    'Montar dicionário consolidado CEF por chave composta
    ultimaLinhaCEF = ws6.Cells(ws6.Rows.Count, "B").End(xlUp).Row
    Set dictCEFChaves = CreateObject("Scripting.Dictionary")

    For i = 2 To ultimaLinhaCEF
        If LCase(Trim(ws6.Cells(i, "M").Value)) = "ok" And IsDate(ws6.Cells(i, "E").Value) Then
            emp = Trim(UCase(ws6.Cells(i, "B").Value))
            dataNum = CLng(Int(CDate(ws6.Cells(i, "E").Value)))
            valor = ws6.Cells(i, "H").Value

            chave = emp & "|" & dataNum

            If dictCEFChaves.Exists(chave) Then
                dictCEFChaves(chave) = dictCEFChaves(chave) + valor
            Else
                dictCEFChaves.Add chave, valor
            End If
        End If
    Next i

    'Montar dicionário consolidado ITAU por chave composta
    ultimaLinhaITAU = ws7.Cells(ws7.Rows.Count, "B").End(xlUp).Row
    Set dictITAUChaves = CreateObject("Scripting.Dictionary")

    For i = 2 To ultimaLinhaITAU
        If LCase(Trim(ws7.Cells(i, "I").Value)) = "ok" _
           And IsDate(ws7.Cells(i, "C").Value) _
           And IsNumeric(ws7.Cells(i, "G").Value) Then

            emp = Trim(UCase(ws7.Cells(i, "B").Value))
            dataNum = CLng(Int(CDate(ws7.Cells(i, "C").Value)))
            valor = CDbl(ws7.Cells(i, "G").Value)

            chave = emp & "|" & dataNum

            If dictITAUChaves.Exists(chave) Then
                dictITAUChaves(chave) = dictITAUChaves(chave) + valor
            Else
                dictITAUChaves.Add chave, valor
            End If
        End If
    Next i

    'Montar dicionário consolidado GRAFENO por chave composta
    ultimaLinhaGRAFENO = ws8.Cells(ws8.Rows.Count, "A").End(xlUp).Row
    Set dictGRAFENOChaves = CreateObject("Scripting.Dictionary")

    For i = 2 To ultimaLinhaGRAFENO
        historicoGrafeno = Trim(UCase(ws8.Cells(i, "C").Value))

        If IsDate(ws8.Cells(i, "B").Value) _
           And Left(historicoGrafeno, 11) = "RECEBIMENTO" _
           And IsNumeric(ws8.Cells(i, "H").Value) Then

            emp = Trim(UCase(ws8.Cells(i, "A").Value))
            dataNum = CLng(Int(CDate(ws8.Cells(i, "B").Value)))
            valor = CDbl(ws8.Cells(i, "H").Value)

            chave = emp & "|" & dataNum

            If dictGRAFENOChaves.Exists(chave) Then
                dictGRAFENOChaves(chave) = dictGRAFENOChaves(chave) + valor
            Else
                dictGRAFENOChaves.Add chave, valor
            End If
        End If
    Next i

    'Preencher CHECK com CEF, ITAU, GRAFENO e Dif = Boletim - CEF - ITAU - GRAFENO
    ultimaLinhaCheck = ws5.Cells(ws5.Rows.Count, "A").End(xlUp).Row

    For i = 2 To ultimaLinhaCheck
        empAtual = Trim(UCase(ws5.Cells(i, "A").Value))

        If IsDate(ws5.Cells(i, "B").Value) Then
            dataNum = CLng(Int(CDate(ws5.Cells(i, "B").Value)))
            chave = empAtual & "|" & dataNum

            If dictCEFChaves.Exists(chave) Then
                ws5.Cells(i, "D").Value = dictCEFChaves(chave)
            Else
                ws5.Cells(i, "D").Value = 0
            End If

            If dictITAUChaves.Exists(chave) Then
                ws5.Cells(i, "E").Value = dictITAUChaves(chave)
            Else
                ws5.Cells(i, "E").Value = 0
            End If

            If dictGRAFENOChaves.Exists(chave) Then
                ws5.Cells(i, "F").Value = dictGRAFENOChaves(chave)
            Else
                ws5.Cells(i, "F").Value = 0
            End If

            ws5.Cells(i, "G").Value = ws5.Cells(i, "C").Value - ws5.Cells(i, "D").Value - ws5.Cells(i, "E").Value - ws5.Cells(i, "F").Value
        End If
    Next i

'Adicionar na CHECK as movimentações referentes as quadras
'considerando apenas movimentações da ITAU
ultimaLinhaEmpI = ws2.Cells(ws2.Rows.Count, "I").End(xlUp).Row
Set dictEmpresasLista = CreateObject("Scripting.Dictionary")

'Guardar a lista de empresas da coluna I
For i = 2 To ultimaLinhaEmpI
    empresaLista = Trim(UCase(ws2.Cells(i, "I").Value))
    
    If empresaLista <> "" Then
        If Not dictEmpresasLista.Exists(empresaLista) Then
            dictEmpresasLista.Add empresaLista, True
        End If
    End If
Next i

'Descobrir a próxima linha livre da CHECK
linha = ws5.Cells(ws5.Rows.Count, "A").End(xlUp).Row + 1

'Percorrer o dicionário da ITAU e lançar apenas empresas da lista da coluna I
For Each chaveITAU In dictITAUChaves.Keys
    partesChave = Split(chaveITAU, "|")
    
    If UBound(partesChave) = 1 Then
        empresaLista = Trim(UCase(partesChave(0)))
        
        If dictEmpresasLista.Exists(empresaLista) Then
            ws5.Cells(linha, "A").Value = empresaLista
            ws5.Cells(linha, "B").Value = CDate(CLng(partesChave(1)))
            ws5.Cells(linha, "C").Value = 0
            ws5.Cells(linha, "D").Value = 0
            ws5.Cells(linha, "E").Value = dictITAUChaves(chaveITAU)
            ws5.Cells(linha, "F").Value = 0
            ws5.Cells(linha, "G").Value = ""
            linha = linha + 1
        End If
    End If
Next chaveITAU

    'Formatação final CHECK
    With ws5
        .Cells(1, "A").Value = "Empreendimento"
        .Cells(1, "B").Value = "Data"
        .Cells(1, "C").Value = "Valor boletim"
        .Cells(1, "D").Value = "CEF"
        .Cells(1, "E").Value = "ITAU"
        .Cells(1, "F").Value = "GRAFENO"
        .Cells(1, "G").Value = "Dif"

        .Rows(1).Font.Bold = True
        .Columns("B").NumberFormat = "dd/mm/yyyy"
        .Columns("C:G").NumberFormat = "#,##0.00"
        .Columns("A:G").AutoFit

        If .AutoFilterMode Then .AutoFilterMode = False
        .Range("A1:G" & .Cells(.Rows.Count, "A").End(xlUp).Row).AutoFilter
    End With

    ws5.Activate
    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With

    Application.ScreenUpdating = True
    fim = Timer
    tempo = (fim - inicio) / 60.2
    
    
    MsgBox "Processo concluído em " & Round(tempo) & " minutos." & vbCrLf & _
           "Há " & total & " clientes sem a % do RET" & vbCrLf & _
           "Há " & totalEmpreendimentos & " linhas sem o empreendimento/bloco" & vbCrLf & _
           "Há " & totalRET & " valores recebidos e sem CNPJ RET na aba Empreendimentos", vbExclamation, "Resumo de inconsistências"
           
           
End Sub
