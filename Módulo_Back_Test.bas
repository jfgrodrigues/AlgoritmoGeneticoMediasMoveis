Attribute VB_Name = "Módulo_Back_Test"
Public matriz_estrategia(10000, 7) As String, linha_matriz As Integer, position As String, P_L As Double, entrada As Boolean, saida As Boolean, entrada_comprado As Boolean, prc_ent As Double, prc_out As Double


Sub pegar_dados(caminho As String, papel As String, MM1 As Integer, MM2 As Integer)
    Dim raw_row As String, papel_normalizado As String
    
    P_L = 0
    limpar_matriz
    entrada = False
    saida = False
    entrada_comprado = False
    
    Set ObjFSO = CreateObject("Scripting.FileSystemObject")
    Set ObjFile = ObjFSO.OpenTextFile(caminho, 1)
    
    While Not ObjFile.AtEndOfStream
        raw_row = ObjFile.ReadLine
        papel_normalizado = get_instr_normalizado(papel)
        If Left(raw_row, 2) = "01" And Right(Left(raw_row, 24), 12) = papel_normalizado Then
            matriz_estrategia(linha_matriz, 0) = get_price(Right(Left(raw_row, 69), 11))
            matriz_estrategia(linha_matriz, 1) = get_price(Right(Left(raw_row, 121), 11))
            matriz_estrategia(linha_matriz, 2) = get_MM(MM1)
            matriz_estrategia(linha_matriz, 3) = get_MM(MM2)
            If linha_matriz = 0 Then
                matriz_estrategia(0, 4) = "-"
                matriz_estrategia(0, 5) = "-"
                matriz_estrategia(0, 6) = "-"
                matriz_estrategia(0, 7) = "0"
            Else
                matriz_estrategia(linha_matriz, 4) = get_position
                matriz_estrategia(linha_matriz, 5) = get_action
                matriz_estrategia(linha_matriz - 1, 6) = get_exec_price
            End If
            
            linha_matriz = linha_matriz + 1
        End If
    Wend
    linha_matriz = 0
    Sheets("resultado").Range("A2:H1000") = matriz_estrategia
    Sheets("resultado").Range("K3") = P_L
End Sub

Function get_price(raw_price) As String
    Dim prc_normalizado As String, prc_saida As Double
    prc_normalizado = raw_price
    While Left(prc_normalizado, 1) = "0"
        prc_normalizado = Right(prc_normalizado, Len(prc_normalizado) - 1)
    Wend
    prc_saida = CInt(prc_normalizado) / 100
    get_price = CStr(prc_saida)
End Function

Function get_MM(periodos As Integer) As String
    Dim MM As Double
    
    If periodos <= (linha_matriz + 1) Then
        For i = (linha_matriz + 1 - periodos) To linha_matriz
            MM = MM + CDbl(matriz_estrategia(i, 1))
        Next
        get_MM = CStr(MM / periodos)
    End If
End Function

Function get_action() As String
    If matriz_estrategia(linha_matriz - 1, 2) <> "" And matriz_estrategia(linha_matriz - 1, 3) <> "" Then
        If matriz_estrategia(linha_matriz, 4) = "-" Then
            If CDbl(matriz_estrategia(linha_matriz, 2)) > CDbl(matriz_estrategia(linha_matriz, 3)) Then
                get_action = "C"
            Else
                get_action = "V"
            End If
        ElseIf CDbl(matriz_estrategia(linha_matriz - 1, 2)) <= CDbl(matriz_estrategia(linha_matriz - 1, 3)) And CDbl(matriz_estrategia(linha_matriz, 2)) > CDbl(matriz_estrategia(linha_matriz, 3)) Then
            get_action = "C"
        ElseIf CDbl(matriz_estrategia(linha_matriz - 1, 2)) >= CDbl(matriz_estrategia(linha_matriz - 1, 3)) And CDbl(matriz_estrategia(linha_matriz, 2)) < CDbl(matriz_estrategia(linha_matriz, 3)) Then
            get_action = "V"
        Else
            get_action = "-"
        End If
    Else
        get_action = "-"
    End If
End Function

Function get_position() As String
    If matriz_estrategia(linha_matriz - 1, 4) = "-" Then
        get_position = matriz_estrategia(linha_matriz - 1, 5)
    Else
        If matriz_estrategia(linha_matriz - 1, 5) <> matriz_estrategia(linha_matriz - 1, 4) And matriz_estrategia(linha_matriz - 1, 5) <> "-" Then
            get_position = "-"
        Else
            get_position = matriz_estrategia((linha_matriz - 1), 4)
        End If
    End If
    position = get_position
        
End Function

Function get_instr_normalizado(raw_instr As String) As String
    Dim saida As String
    saida = raw_instr
    While Len(saida) < 12
        saida = saida & " "
    Wend
    get_instr_normalizado = saida
End Function

Function get_exec_price() As String
    If matriz_estrategia(linha_matriz - 1, 5) <> "-" Then
        get_exec_price = matriz_estrategia(linha_matriz, 0)
        ajustar_precos_entrada_saida
    End If
End Function

Sub ajustar_p_l(preco_entrada, preco_saida, entr_c As Boolean)
    If entr_c Then
        P_L = P_L + 100 * (preco_saida - preco_entrada)
    Else
        P_L = P_L + 100 * (preco_entrada - preco_saida)
    End If
    entrada = False
    saida = False
    entrada_comprado = False
    'prc_ent = 0
    'prc_out = 0
    matriz_estrategia(linha_matriz - 1, 7) = P_L
End Sub

Sub ajustar_precos_entrada_saida()
    If Not entrada And Not saida Then
        entrada = True
        prc_ent = CDbl(matriz_estrategia(linha_matriz, 0))
        If position = "C" Then
            entrada_comprado = True
        End If
    Else
        saida = True
        prc_out = CDbl(matriz_estrategia(linha_matriz, 0))
        Call ajustar_p_l(prc_ent, prc_out, entrada_comprado)
    End If
End Sub


Function run_back_test(caminho As String, papel As String, MM1 As Integer, MM2 As Integer) As Double
    Dim raw_row As String, papel_normalizado As String
    
    P_L = 0
    limpar_matriz
    entrada = False
    saida = False
    entrada_comprado = False
    
    Set ObjFSO = CreateObject("Scripting.FileSystemObject")
    Set ObjFile = ObjFSO.OpenTextFile(caminho, 1)
    
    While Not ObjFile.AtEndOfStream
        raw_row = ObjFile.ReadLine
        papel_normalizado = get_instr_normalizado(papel)
        If Left(raw_row, 2) = "01" And Right(Left(raw_row, 24), 12) = papel_normalizado Then
            matriz_estrategia(linha_matriz, 0) = get_price(Right(Left(raw_row, 69), 11))
            matriz_estrategia(linha_matriz, 1) = get_price(Right(Left(raw_row, 121), 11))
            matriz_estrategia(linha_matriz, 2) = get_MM(MM1)
            matriz_estrategia(linha_matriz, 3) = get_MM(MM2)
            If linha_matriz = 0 Then
                matriz_estrategia(0, 4) = "-"
                matriz_estrategia(0, 5) = "-"
                matriz_estrategia(0, 6) = "-"
                matriz_estrategia(0, 7) = "0"
            Else
                matriz_estrategia(linha_matriz, 4) = get_position
                matriz_estrategia(linha_matriz, 5) = get_action
                matriz_estrategia(linha_matriz - 1, 6) = get_exec_price
            End If
            
            linha_matriz = linha_matriz + 1
        End If
    Wend
    linha_matriz = 0
    run_back_test = P_L
End Function

Sub limpar_matriz()
    For i = 0 To 10000
        For i2 = 0 To 7
            'If i2 = 1 Then
            '    matriz_estrategia(i, i2) = "0"
            'Else
                matriz_estrategia(i, i2) = ""
            'End If
        Next
    Next
End Sub
