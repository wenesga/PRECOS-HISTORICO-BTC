Sub ObterHistoricoBTC()
    ' Define a moeda para BTCBRL
    Range("C1").Value = "BTCBRL"
    ' Chama a sub-rotina para obter o histórico com intervalo de 1 Dia
    Call historicoBTC_BRL("1d")
End Sub



Sub ObterHistoricoUSD()
    ' Define a moeda para BTCUSD
    Range("C1").Value = "BTCUSDT"
    ' Chama a sub-rotina para obter o histórico com intervalo de 1 Dia
    Call historicoBTC_BRL("1d")
End Sub



' Botao historico 1 minutos
Sub historicoBTC_1mi()
    Call historicoBTC_BRL("1m")
End Sub


' Botao historico 3 minutos
Sub historicoBTC_3m()
    Call historicoBTC_BRL("3m")
End Sub


' Botao historico 5 minutos
Sub historicoBTC_5m()
    Call historicoBTC_BRL("5m")
End Sub


' Botao historico 15 minutos
Sub historicoBTC_15m()
    Call historicoBTC_BRL("15m")
End Sub


' Botao historico 30 minutos
Sub historicoBTC_30m()
    Call historicoBTC_BRL("30m")
End Sub


' Botao historico 1 hora
Sub historicoBTC_1h()
    Call historicoBTC_BRL("1h")
End Sub


' Botao historico 2 horas
Sub historicoBTC_2h()
    Call historicoBTC_BRL("2h")
End Sub


' Botao historico 4 horas
Sub historicoBTC_4h()
    Call historicoBTC_BRL("4h")
End Sub


' Botao historico 6 horas
Sub historicoBTC_6h()
    Call historicoBTC_BRL("6h")
End Sub


' Botao historico 8 horas
Sub historicoBTC_8h()
    Call historicoBTC_BRL("8h")
End Sub

' Botao historico 12 horas
Sub historicoBTC_12h()
    Call historicoBTC_BRL("12h")
End Sub

' Botao historico 1 dia
Sub historicoBTC_1d()
    Call historicoBTC_BRL("1d")
End Sub


' Botao historico 3 dias
Sub historicoBTC_3d()
    Call historicoBTC_BRL("3d")
End Sub


' Botao historico 1 semana
Sub historicoBTC_1w()
    Call historicoBTC_BRL("1w")
End Sub


' Botao historico 1 mês
Sub historicoBTC_1M()
    Call historicoBTC_BRL("1M")
End Sub
