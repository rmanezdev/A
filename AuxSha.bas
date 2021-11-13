Attribute VB_Name = "AuxSha"
Option Explicit

Function Ch_Function(valor_1 As String, valor_2 As String, _
valor_3 As String) As String

    valor_1 = HexaParaBin(valor_1, 32)
    valor_2 = HexaParaBin(valor_2, 32)
    valor_3 = HexaParaBin(valor_3, 32)
    
    Dim ResultadoAND1 As String
    Dim ResultadoComplement As String
    Dim ResultadoAND2 As String
    
    ResultadoAND1 = bitwise_AND(valor_1, valor_2)
    
    ResultadoComplement = bitwise_complement(valor_1)
    
    ResultadoAND2 = bitwise_AND(ResultadoComplement, valor_3)
    
    Ch_Function = bitwise_XOR(ResultadoAND1, ResultadoAND2)
    
    valor_1 = BinParaHexa(valor_1)
    valor_2 = BinParaHexa(valor_2)
    valor_3 = BinParaHexa(valor_3)
    
    Ch_Function = BinParaHexa(Ch_Function)
    

End Function


Function Maj_Function(valor_1 As String, valor_2 As String, _
valor_3 As String) As String

    valor_1 = HexaParaBin(valor_1, 32)
    valor_2 = HexaParaBin(valor_2, 32)
    valor_3 = HexaParaBin(valor_3, 32)
    
    Dim ResultadoAND1 As String
    Dim ResultadoAND2 As String
    Dim ResultadoAND3 As String
    Dim ResultadoXOR1 As String
    
    ResultadoAND1 = bitwise_AND(valor_1, valor_2)
    
    ResultadoAND2 = bitwise_AND(valor_1, valor_3)
    
    ResultadoAND3 = bitwise_AND(valor_2, valor_3)
    
    ResultadoXOR1 = bitwise_XOR(ResultadoAND1, ResultadoAND2)
    
    Maj_Function = bitwise_XOR(ResultadoXOR1, ResultadoAND3)
    
    valor_1 = BinParaHexa(valor_1)
    valor_2 = BinParaHexa(valor_2)
    valor_3 = BinParaHexa(valor_3)
    
    Maj_Function = BinParaHexa(Maj_Function)
    

End Function


Function s_zero_Function(Valor As String) As String

    Valor = HexaParaBin(Valor, 32)
    
    Dim ResultadoRotation1 As String
    Dim ResultadoRotation2 As String
    Dim ResultadoRotation3 As String
    Dim ResultadoXOR1 As String
    
    ResultadoRotation1 = right_rotation(Valor, 2)
    
    ResultadoRotation2 = right_rotation(Valor, 13)
    
    ResultadoRotation3 = right_rotation(Valor, 22)
    
    ResultadoXOR1 = bitwise_XOR(ResultadoRotation1, _
    ResultadoRotation2)
    
    s_zero_Function = bitwise_XOR(ResultadoXOR1, _
    ResultadoRotation3)
    
    Valor = BinParaHexa(Valor)
    
    s_zero_Function = _
    BinParaHexa(s_zero_Function)
    

End Function


Function s_um_Function(Valor As String) As String

    Valor = HexaParaBin(Valor, 32)
    
    Dim ResultadoRotation1 As String
    Dim ResultadoRotation2 As String
    Dim ResultadoRotation3 As String
    Dim ResultadoXOR1 As String
    
    ResultadoRotation1 = right_rotation(Valor, 6)
    
    ResultadoRotation2 = right_rotation(Valor, 11)
    
    ResultadoRotation3 = right_rotation(Valor, 25)
    
    ResultadoXOR1 = bitwise_XOR(ResultadoRotation1, _
    ResultadoRotation2)
    
    s_um_Function = bitwise_XOR(ResultadoXOR1, _
    ResultadoRotation3)
    
    Valor = BinParaHexa(Valor)
    
    s_um_Function = _
    BinParaHexa(s_um_Function)
    
End Function


Function T1_Function(valor_1 As String, valor_2 As String, _
valor_3 As String, valor_4 As String, valor_5 As String) As String

    Dim valor_1_funcao As Double
    Dim valor_2_funcao As Double
    Dim valor_3_funcao As Double
    Dim valor_4_funcao As Double
    Dim valor_5_funcao As Double
    
    valor_1_funcao = HexaParaDec(valor_1)
    valor_2_funcao = HexaParaDec(valor_2)
    valor_3_funcao = HexaParaDec(valor_3)
    valor_4_funcao = HexaParaDec(valor_4)
    valor_5_funcao = HexaParaDec(valor_5)
    
    Dim Soma As Variant
    Soma = valor_1_funcao + valor_2_funcao + _
    valor_3_funcao + valor_4_funcao + valor_5_funcao
    
    Dim c As Variant
    c = 2 ^ 32
    
    Dim mod_2_32_dec As Variant
    mod_2_32_dec = RestoDiv(Soma, c)
    
    T1_Function = DecParaHex(mod_2_32_dec, 8)

End Function
