Attribute VB_Name = "M�dulo_AG"
Type individuos
    medias_moveis(1) As Integer
    res_back_test As Double
End Type

Public matriz_populacao() As individuos, matriz_resultados() As Double

Sub run_ag(instrumento As String, num_iteracoes As Integer, arq_cotacao As String, tam_populacao As Integer, percent_sobrevivencia As Double, percent_mutacao As Double, Optional carregar_pop_ini As Boolean)
    Dim inicio As Date, duracao As Date
    inicio = Now
    ReDim matriz_populacao(tam_populacao - 1)
    If carregar_pop_ini Then
        Call get_populacao_inicial
    Else
        Call gerar_populacao_inicial
    End If
    
    For i = 0 To num_iteracoes
        For i2 = 0 To UBound(matriz_populacao)
            matriz_populacao(i2).res_back_test = run_back_test(arq_cotacao, instrumento, matriz_populacao(i2).medias_moveis(0), matriz_populacao(i2).medias_moveis(1))
        Next
        Call selecionar_resultados(percent_sobrevivencia)
        Call recombinacao(percent_sobrevivencia)
        Call mutacao(percent_mutacao)
    Next
    
    For i = 0 To UBound(matriz_populacao)
        matriz_populacao(i).res_back_test = run_back_test(arq_cotacao, instrumento, matriz_populacao(i).medias_moveis(0), matriz_populacao(i).medias_moveis(1))
    Next
    duracao = Now - inicio
    Call salvar_resultado(instrumento, num_iteracoes, tam_populacao, duracao)

End Sub

Sub gerar_populacao_inicial()
    Dim MM1 As Integer, MM2 As Integer, MM_temp As Integer
    For i = 0 To UBound(matriz_populacao)
        Randomize
        MM1 = Int(35 * Rnd) + 1
        MM2 = Int(50 * Rnd) + 1
        If MM1 > MM2 Then
            MM_temp = MM1
            MM1 = MM2
            MM2 = MM_temp
        End If
        matriz_populacao(i).medias_moveis(0) = MM1
        matriz_populacao(i).medias_moveis(1) = MM2
    Next
End Sub

Sub selecionar_resultados(percent As Double)
    ReDim matriz_resultados(UBound(matriz_populacao), 1)
    For i = 0 To UBound(matriz_populacao)
        matriz_resultados(i, 0) = matriz_populacao(i).res_back_test
        matriz_resultados(i, 1) = i
    Next
    reordenar_matriz (percent)
End Sub

Sub get_populacao_inicial()

End Sub

Sub mutacao(percent)

End Sub

Sub salvar_resultado(instr As String, iteracoes As Integer, populacao As Integer, tempo As Date)

End Sub

Sub recombinacao(percent_sobrevivencia As Double)
    Dim matriz_temp() As individuos
    
    ReDim matriz_temp(UBound(matriz_populacao))
    For i = 0 To UBound(matriz_populacao)
        matriz_temp(i).medias_moveis(0) = matriz_populacao(get_selecao_por_aptidao(percent_sobrevivencia)).medias_moveis(0)
        matriz_temp(i).medias_moveis(1) = matriz_populacao(get_selecao_por_aptidao(percent_sobrevivencia)).medias_moveis(1)
    Next
    
    For i = 0 To UBound(matriz_populacao)
        matriz_populacao(i).medias_moveis(0) = matriz_temp(i).medias_moveis(0)
        matriz_populacao(i).medias_moveis(1) = matriz_temp(i).medias_moveis(1)
        matriz_populacao(i).res_back_test = matriz_temp(i).res_back_test
    Next
End Sub

Sub reordenar_matriz(persistencia)
    Dim valor_temp As Double, indice_temp As Double, inconsistencia As Boolean, maior_valor As Double, qtd_ent As Integer, qtd_persistente As Integer, matriz_temp() As individuos, linha_preenchida As Boolean
    
    ReDim matriz_temp(UBound(matriz_populacao))
    qtd_ent = UBound(matriz_populacao) - 1
    qtd_persistente = Int((qtd_ent + 1) * persistencia)
    inconsistencia = True
    
    While inconsistencia
        inconsistencia = False
        For i = 0 To qtd_ent
            If matriz_resultados(i + 1, 0) > matriz_resultados(i, 0) Then
                valor_temp = matriz_resultados(i, 0)
                indice_temp = matriz_resultados(i, 1)
                matriz_resultados(i, 0) = matriz_resultados(i + 1, 0)
                matriz_resultados(i, 1) = matriz_resultados(i + 1, 1)
                matriz_resultados(i + 1, 0) = valor_temp
                matriz_resultados(i + 1, 1) = indice_temp
            End If
            If i = 0 Then
                maior_valor = matriz_resultados(i, 0)
            End If
            If matriz_resultados(i, 0) > maior_valor Then
                inconsistencia = True
            End If
        Next
    Wend
    For i = 0 To qtd_persistente
        linha_preenchida = False
        For i2 = i To UBound(matriz_resultados)
            If matriz_resultados(i2, 1) > 0 And Not linha_preenchida Then
                matriz_temp(i).medias_moveis(0) = matriz_populacao(matriz_resultados(i, 1)).medias_moveis(0)
                matriz_temp(i).medias_moveis(1) = matriz_populacao(matriz_resultados(i, 1)).medias_moveis(1)
                matriz_temp(i).res_back_test = matriz_populacao(matriz_resultados(i, 1)).res_back_test
                linha_preenchida = True
            End If
            If i2 = UBound(matriz_resultados) And Not linha_preenchida Then
                matriz_temp(i).medias_moveis(0) = 0
                matriz_temp(i).medias_moveis(1) = 0
                matriz_temp(i).res_back_test = 0
                linha_preenchida = True
            End If
        Next
    Next
    For i = 0 To UBound(matriz_populacao)
        matriz_populacao(i).medias_moveis(0) = matriz_temp(i).medias_moveis(0)
        matriz_populacao(i).medias_moveis(1) = matriz_temp(i).medias_moveis(1)
        matriz_populacao(i).res_back_test = matriz_temp(i).res_back_test
    Next
End Sub

Function get_selecao_por_aptidao(percent_sobrevivencia As Double) As Integer
    Dim possibilidade_descendencia() As Double, tam_vetor As Integer, soma_res_backtest As Double, random_var As Double
    
    tam_vetor = Int(UBound(matriz_populacao) * percent_sobrevivencia)
    ReDim possibilidade_descendencia(tam_vetor)
    
    For i = 0 To tam_vetor
        soma_res_backtest = soma_res_backtest + matriz_populacao(i).res_back_test
    Next
    
    For i = 0 To tam_vetor
        If i = 0 Then
            possibilidade_descendencia(i) = matriz_populacao(i).res_back_test / soma_res_backtest
        Else
            possibilidade_descendencia(i) = possibilidade_descendencia(i - 1) + matriz_populacao(i).res_back_test / soma_res_backtest
        End If
    Next
    
    Randomize
    random_var = Rnd
    
    For i = 0 To tam_vetor
        If random_var <= possibilidade_descendencia(i) Then
            get_selecao_por_aptidao = i
            Exit Function
        End If
    Next
    
End Function
