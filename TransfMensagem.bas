Attribute VB_Name = "TransfMensagem"
Option Explicit

Function Mensagem_Transformada_Bits(Texto As String, _
Num_Blocos As Long) As String

    Dim num_caracts_texto As Integer
    Dim num_bits_texto As Variant
    Dim K As Integer
    Dim num_bits_texto_binario As String
    
    num_caracts_texto = Len(Texto)
    num_bits_texto = 8 * num_caracts_texto
        
    num_bits_texto_binario = _
    DecParaBin(num_bits_texto, 64)
    
    Mensagem_Transformada_Bits = ""
    Dim Contador As Integer
    Dim Letra As String
    Dim BinarioProcurado As String
    For Contador = 1 To num_caracts_texto
        Letra = Mid(Texto, Contador, 1)
        BinarioProcurado = DecParaBin(Asc(Letra), 8)
        Mensagem_Transformada_Bits = _
        Mensagem_Transformada_Bits & BinarioProcurado
    Next Contador
    
    Mensagem_Transformada_Bits = _
    Mensagem_Transformada_Bits & "1"
    
    K = _
    RestoDiv(512 * Num_Blocos - 64, 512 * Num_Blocos) - _
    (num_bits_texto + 1)
    
    For Contador = 1 To K
        Mensagem_Transformada_Bits = _
        Mensagem_Transformada_Bits & "0"
    Next Contador
    
    Mensagem_Transformada_Bits = _
    Mensagem_Transformada_Bits & num_bits_texto_binario
    
End Function
