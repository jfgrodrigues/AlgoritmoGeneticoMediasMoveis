Attribute VB_Name = "Módulo_Random"
Function get_random() As Double
    Dim agora_s As String
    
    agora_s = Right(Format(Timer, "#0.00"), 2)
    get_random = CDbl(agora_s) / 100
End Function

