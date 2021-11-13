Attribute VB_Name = "OpBasicas"
Option Explicit

Function bitwise_AND(valor_1 As String, _
valor_2 As String) As String

    bitwise_AND = ""
    
    Dim Num_Caracts As Integer
    Num_Caracts = Len(valor_1)
    
    Dim Contador As Integer
    For Contador = 1 To Num_Caracts
        bitwise_AND = bitwise_AND & _
        WorksheetFunction.Min( _
        CInt(Mid(valor_1, Contador, 1)), _
        CInt(Mid(valor_2, Contador, 1)))
    Next Contador

End Function

Function bitwise_OR(valor_1 As String, _
valor_2 As String) As String

    bitwise_OR = ""
    
    
    Dim Num_Caracts As Integer
    Num_Caracts = Len(valor_1)
    
    Dim Contador As Integer
    For Contador = 1 To Num_Caracts
        bitwise_OR = bitwise_OR & _
        WorksheetFunction.Max( _
        CInt(Mid(valor_1, Contador, 1)), _
        CInt(Mid(valor_2, Contador, 1)))
    Next Contador

End Function

Function bitwise_XOR(valor_1 As String, _
valor_2 As String) As String

    bitwise_XOR = ""
    
    Dim Num_Caracts As Integer
    Num_Caracts = Len(valor_1)
    
    Dim bit_valor_1 As Integer
    Dim bit_valor_2 As Integer
    Dim Contador As Integer
    
    For Contador = 1 To Num_Caracts
        bit_valor_1 = CInt(Mid(valor_1, Contador, 1))
        bit_valor_2 = CInt(Mid(valor_2, Contador, 1))
        bitwise_XOR = bitwise_XOR & WorksheetFunction.Max( _
        WorksheetFunction.Min(bit_valor_1, 1 - bit_valor_2), _
        WorksheetFunction.Min(1 - bit_valor_1, bit_valor_2))
    Next Contador

End Function

Function bitwise_complement(Valor As String) As String

    Dim Num_Caracts As Integer
    Dim Contador As Integer
    
    Num_Caracts = Len(Valor)
    
    
    bitwise_complement = ""
    For Contador = 1 To Num_Caracts
        bitwise_complement = bitwise_complement & _
        CStr(1 - CInt(Mid(Valor, Contador, 1)))
    Next Contador

End Function


Function mod_2_32_addition(valor_1 As String, _
valor_2 As String) As String

    Dim valor_1_funcao As Double
    Dim valor_2_funcao As Double

    valor_1_funcao = HexaParaDec(valor_1)
    valor_2_funcao = HexaParaDec(valor_2)
    
    Dim Soma As Variant
    Soma = valor_1_funcao + valor_2_funcao
    
    Dim c As Variant
    c = 2 ^ 32
    
    Dim mod_2_32_dec As Variant
    mod_2_32_addition = RestoDiv(Soma, c)
    
    mod_2_32_addition = DecParaHex(mod_2_32_addition, 8)

End Function


Function right_shift(Valor As String, _
Num_Bits As Integer) As String

    Dim Num_Caracts As Integer
    
    Num_Caracts = Len(Valor)
    
    Dim i As Integer
    For i = 1 To Num_Bits
        Valor = "0" & Valor
    Next i
    
    right_shift = Left(Valor, Num_Caracts)
    
    Valor = Right(Valor, 32)

End Function

Function right_rotation(Valor As String, _
Num_Bits As Integer) As String

    Dim Num_Caracts As Integer
    Num_Caracts = Len(Valor)
    
    right_rotation = _
    Left(Right(Valor, Num_Bits) & Valor, Num_Caracts)

End Function
