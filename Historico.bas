Sub historicoBTC_BRL(intervalo As String)
    Dim xmlHttp As Object
    Dim resposta As String
    Dim dados() As String
    Dim linha As Integer
    Dim coluna As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("HISTORICO_BTC") ' Substitua "NomeDaSuaPlanilha" pelo nome real da sua planilha


    
    ' Mapear intervalos para descrições mais amigáveis
    Select Case intervalo
        Case "1m"
            descricaoIntervalo = "Cotação - 1 Minuto"
        Case "3m"
            descricaoIntervalo = "Cotação - 3 Minutos"
        Case "5m"
            descricaoIntervalo = "Cotação - 5 Minutos"
        Case "15m"
            descricaoIntervalo = "Cotação - 15 Minutos"
        Case "30m"
            descricaoIntervalo = "Cotação - 30 Minutos"
        Case "1h"
            descricaoIntervalo = "Cotação - 1 Hora"
        Case "2h"
            descricaoIntervalo = "Cotação - 2 Hora"
        Case "4h"
            descricaoIntervalo = "Cotação - 4 Hora"
        Case "6h"
            descricaoIntervalo = "Cotação - 6 Hora"
        Case "8h"
            descricaoIntervalo = "Cotação - 8 Hora"
        Case "12h"
            descricaoIntervalo = "Cotação - 12 Hora"
        Case "1d"
            descricaoIntervalo = "Cotação - 1 Dia"
        Case "3d"
            descricaoIntervalo = "Cotação - 3 Dias"
        Case "1w"
            descricaoIntervalo = "Cotação - 1 Semana"
        Case "1M"
            descricaoIntervalo = "Cotação - 1 Mês"
        Case Else
            descricaoIntervalo = "Cotação - " & intervalo
    End Select
    
    ' Preencher a célula A1 com base no intervalo escolhido
    Range("A1").Value = descricaoIntervalo
    
    
    
    ' Pega a moeda da célula C1
    Dim moeda As String
    moeda = Range("C1").Value
    
    ' Verifica se a célula C1 está vazia
    If moeda = "" Then
        MsgBox "Por favor, defina a moeda na célula C1."
        Exit Sub
    End If
    
    ' Define a URL com base na moeda e intervalo
    Dim url As String
    url = "https://api.binance.com/api/v3/klines?symbol=" & moeda & "&interval=" & intervalo & "&limit=18"
    
    ' Crie uma instância do objeto XMLHTTP
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' Abra a URL
    xmlHttp.Open "GET", url, False
    xmlHttp.send
    
    ' Obtenha a resposta da requisição
    resposta = xmlHttp.responseText
    
    

    ' Inserir títulos na primeira linha
    ws.Cells(2, 2).Value = "Data Inicial"
    ws.Cells(2, 3).Value = "Abertura"
    ws.Cells(2, 4).Value = "Máxima"
    ws.Cells(2, 5).Value = "Mínima"
    ws.Cells(2, 6).Value = "Fechamento"
    ws.Cells(2, 7).Value = "Data Final"

    
    
    ' Divida os dados em linhas
    dados = Split(resposta, "],[")
    
    ' Comece a escrever os dados nas colunas B, C, D, E, F e H (ou ajuste conforme necessário)
    linha = 3
    
    For Each linhaDados In dados
        ' Remova os colchetes iniciais e finais
        linhaDados = Replace(linhaDados, "[", "")
        linhaDados = Replace(linhaDados, "]", "")
        
        ' Divida os valores em colunas
        valores = Split(linhaDados, ",")
        
        ' Escreva os valores nas colunas especificadas
        Cells(linha, 2).Value = CDbl(valores(0)) / 86400000 + 25569 - (3 / 24) ' Converta a primeira coluna para data
        Cells(linha, 3).Value = Replace(valores(1), """", "") ' Remova as aspas duplas
        Cells(linha, 4).Value = Replace(valores(2), """", "") ' Remova as aspas duplas
        Cells(linha, 5).Value = Replace(valores(3), """", "") ' Remova as aspas duplas
        Cells(linha, 6).Value = Replace(valores(4), """", "") ' Remova as aspas duplas
        Cells(linha, 7).Value = CDbl(valores(6)) / 86400000 + 25569 - (3 / 24) ' Converta a última coluna para data
        
        ' Vá para a próxima linha
        linha = linha + 1
    Next linhaDados
End Sub
