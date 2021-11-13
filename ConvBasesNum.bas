Attribute VB_Name = "ConvBasesNum"
Option Explicit

Function DecParaBin(Dec As Variant, _
Optional NBits As Integer = 32) As String
    
    Dim Num_Decimal As Variant
    Dim R As Integer
    
    DecParaBin = ""
    Num_Decimal = Dec
    
    Do While Num_Decimal > 0
        R = RestoDiv(Num_Decimal, 2)
        Num_Decimal = Int(Num_Decimal / 2)
        DecParaBin = R & DecParaBin
    Loop
    
    Do While Len(DecParaBin) < NBits
        DecParaBin = "0" & DecParaBin
    Loop

End Function

Function RestoDiv(Numerador As Variant, Denominador As Variant) As Double

    Dim Div As Variant
    Dim DivInteiro As Variant
    
    Div = Numerador / Denominador
    DivInteiro = WorksheetFunction.RoundDown(Div, 0)

    RestoDiv = Numerador - Denominador * DivInteiro

End Function


Function BinParaDec(Bin As String) As Double

    Dim Bit As Integer
    Dim i As Integer
    Dim Num_Bits As Integer
    
    Num_Bits = Len(Bin)
    
    BinParaDec = 0
    
    Do While i < Num_Bits
        Bit = CInt(Mid(Bin, Num_Bits - i, 1))
        BinParaDec = BinParaDec + Bit * 2 ^ i
        i = i + 1
    Loop
    

End Function

Function DecParaHex(Dec As Variant, _
Optional NDigs As Integer = 8) As String
    
    DecParaHex = ""
    
    Dim R As Long
    Do While Dec > 0
        R = RestoDiv(Dec, 16)
        
        Dec = Int(Dec / 16)
        
        DecParaHex = _
        TabConv(CStr(R), "dec", "hex") & DecParaHex
    Loop
    
    Do While Len(DecParaHex) < NDigs
        DecParaHex = "0" & DecParaHex
    Loop

End Function

Function HexaParaDec(Hexa As Variant) As Double

    HexaParaDec = 0

    Dim NDigs As Integer
    NDigs = Len(Hexa)
    
    Dim NumMult As Integer
    
    Dim Digito As String
    Dim i As Integer
    For i = 0 To NDigs - 1
        Digito = Mid(Hexa, NDigs - i, 1)
        NumMult = TabConv(Digito, "hex", "dec")
        HexaParaDec = HexaParaDec + NumMult * 16 ^ i
    Next i
    
End Function

Function HexaParaBin(Hexa As String, _
Optional NBits As Integer = 32) As String
    
    Dim NDigs As Integer
    NDigs = Len(Hexa)
    
    Dim Digito As String
    Dim i As Integer
    For i = 1 To NDigs
        Digito = Mid(Hexa, i, 1)
        HexaParaBin = _
        HexaParaBin & TabConv(Digito, "hex", "bin")
    Next i
    
    Do While Len(HexaParaBin) < NBits
        HexaParaBin = "0" & HexaParaBin
    Loop

End Function

Function BinParaHexa(Bin As String, _
Optional NDigsFinal As Integer = 8) As String
    
    Dim NDigs As Integer
    NDigs = Len(Bin)
    
    Dim R As Integer
    R = RestoDiv(NDigs, 4)
    Do While 4 - R > 0 And R <> 0
        Bin = "0" & Bin
        R = R + 1
    Loop
    
    Dim Bits As String
    Dim i As Integer
    For i = 1 To NDigs Step i + 4
        Bits = Mid(Bin, i, 4)
        BinParaHexa = _
        BinParaHexa & TabConv(Bits, "bin", "hex")
    Next i
    
    Do While Len(BinParaHexa) < NDigsFinal
        BinParaHexa = "0" & BinParaHexa
    Loop

End Function
