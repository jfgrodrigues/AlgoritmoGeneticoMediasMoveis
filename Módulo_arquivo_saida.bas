Attribute VB_Name = "Módulo_arquivo_saida"
Sub salvar_arquivo(caminho As String, papel As String, num_individuos As String, num_iteracoes As String)
Attribute salvar_arquivo.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim hoje As String
    
    hoje = Format(Now(), "YYYY_MM_DD-HHMMSS")
    Sheets("arq_saida").Select
    Sheets("arq_saida").Copy
    ActiveWorkbook.SaveAs Filename:=caminho & "res_AG_" & hoje & "_" & papel & "_" & num_individuos & "_" & num_iteracoes & ".csv", _
        FileFormat:=xlCSVUTF8, CreateBackup:=False
End Sub
