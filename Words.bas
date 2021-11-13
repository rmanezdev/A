Attribute VB_Name = "Words"
Option Explicit

Function GerarWords(Mensagem_em_Bits As String, _
NumDeBlocos As Long) As Variant

    ReDim W(0 To NumDeBlocos - 1, 0 To 63) As String
    ReDim W_Hexa(0 To NumDeBlocos - 1, 0 To 63) As String
    Dim c As Integer
    
    c = 0
        
    Dim i As Integer
    Dim j As Integer
    For i = 0 To NumDeBlocos - 1
        For j = 0 To 15
            W(i, j) = Mid(Mensagem_em_Bits, c * 32 + 1, 32)
            W_Hexa(i, j) = BinParaHexa(W(i, j), 8)
            c = c + 1
        Next j
    Next i
    
    Dim sigma_zero As String
    Dim sigma_um As String
    For i = 0 To NumDeBlocos - 1
        For j = 16 To 63
            sigma_um = sigma_um_Function(W_Hexa(i, j - 2))
            sigma_zero = sigma_zero_Function(W_Hexa(i, j - 15))
            W_Hexa(i, j) = _
            W_Function(sigma_um, W_Hexa(i, j - 7), _
            sigma_zero, W_Hexa(i, j - 16))
        Next j
    Next i
    
    GerarWords = W_Hexa

End Function


Function sigma_zero_Function(Valor As String) As String


    Valor = HexaParaBin(Valor, 32)
    
    Dim ResultadoRotation1 As String
    Dim ResultadoRotation2 As String
    Dim ResultadoShift As String
    Dim ResultadoXOR1 As String
    
    ResultadoRotation1 = right_rotation(Valor, 7)
    
    ResultadoRotation2 = right_rotation(Valor, 18)
    
    ResultadoShift = right_shift(Valor, 3)
    
    ResultadoXOR1 = _
    bitwise_XOR(ResultadoRotation1, ResultadoRotation2)
    
    sigma_zero_Function = _
    bitwise_XOR(ResultadoXOR1, ResultadoShift)
    
    sigma_zero_Function = _
    BinParaHexa(sigma_zero_Function, 8)
    
    Valor = BinParaHexa(Valor, 8)

End Function


Function sigma_um_Function(Valor As String) As String

    Valor = HexaParaBin(Valor, 32)
    
    Dim ResultadoRotation1 As String
    Dim ResultadoRotation2 As String
    Dim ResultadoShift As String
    Dim ResultadoXOR1 As String
    
    ResultadoRotation1 = right_rotation(Valor, 17)
    
    ResultadoRotation2 = right_rotation(Valor, 19)
    
    ResultadoShift = right_shift(Valor, 10)
    
    ResultadoXOR1 = _
    bitwise_XOR(ResultadoRotation1, ResultadoRotation2)
    
    sigma_um_Function = _
    bitwise_XOR(ResultadoXOR1, ResultadoShift)
    
    sigma_um_Function = _
    BinParaHexa(sigma_um_Function, 8)
        
    Valor = BinParaHexa(Valor, 8)

End Function


Function W_Function(valor_1 As String, valor_2 As String, _
valor_3 As String, valor_4 As String) As String

    Dim valor_1_funcao As Double
    Dim valor_2_funcao As Double
    Dim valor_3_funcao As Double
    Dim valor_4_funcao As Double

    valor_1_funcao = HexaParaDec(valor_1)
    valor_2_funcao = HexaParaDec(valor_2)
    valor_3_funcao = HexaParaDec(valor_3)
    valor_4_funcao = HexaParaDec(valor_4)
    
    Dim Soma As Variant
    Soma = valor_1_funcao + valor_2_funcao + _
    valor_3_funcao + valor_4_funcao
    
    Dim c As Variant
    c = 2 ^ 32
    
    Dim mod_2_32_dec As Variant
    mod_2_32_dec = RestoDiv(Soma, c)
        
    W_Function = DecParaHex(mod_2_32_dec, 8)

End Function

