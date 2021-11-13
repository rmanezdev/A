Attribute VB_Name = "Tabelas"
Option Explicit

Function TabConv(Valor As String, De As String, _
Para As String) As Variant
    
    Dim TabelaHexa As Variant
    TabelaHexa = Array( _
    Array("0", "1", "2", "3", "4", "5", "6", "7", _
    "8", "9", "a", "b", "c", "d", "e", "f"), _
    Array("0", "1", "2", "3", "4", "5", "6", "7", _
    "8", "9", "10", "11", "12", "13", "14", "15"), _
    Array("0000", "0001", "0010", "0011", "0100", "0101", _
    "0110", "0111", "1000", "1001", "1010", "1011", "1100", _
    "1101", "1110", "1111"))
    
    Dim iProc As Integer
    Dim iPara As Integer
    
    Dim i As Integer
    If De = "hex" Then
        iProc = 0
        If Para = "dec" Then
            iPara = 1
        ElseIf Para = "bin" Then
            iPara = 2
        End If
    ElseIf De = "dec" Then
        iProc = 1
        If Para = "hex" Then
            iPara = 0
        ElseIf Para = "bin" Then
            iPara = 2
        End If
    ElseIf De = "bin" Then
        iProc = 2
        If Para = "hex" Then
            iPara = 0
        ElseIf Para = "dec" Then
            iPara = 1
        End If
    End If
    
    For i = 0 To UBound(TabelaHexa(iProc))
        If TabelaHexa(iProc)(i) = Valor Then
            TabConv = TabelaHexa(iPara)(i)
            Exit Function
        End If
    Next i
    

End Function
