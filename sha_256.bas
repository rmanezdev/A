Attribute VB_Name = "sha_256"
Option Explicit

Function HASH_SHA_256(Optional MensagemTexto As String = "", _
                      Optional Binario As String = "", _
                      Optional ConstH As String = "") As String
    
    If Binario = "" Then
        Dim NumDeBlocos As Long
        NumDeBlocos = WorksheetFunction.RoundUp( _
        0.015625 * Len(MensagemTexto) + 0.140625, 0)
        
        Dim Mensagem_em_Bits As String
        Mensagem_em_Bits = _
        Mensagem_Transformada_Bits(MensagemTexto, NumDeBlocos)
    Else
        Mensagem_em_Bits = Binario
        NumDeBlocos = Int(Len(Mensagem_em_Bits) / 512)
    End If
    
    
    Dim Words As Variant
    Words = GerarWords(Mensagem_em_Bits, NumDeBlocos)
    
    
    Dim Primos As Variant
    Primos = _
    Array(2, 3, 5, 7, 11, 13, 17, 19, _
          23, 29, 31, 37, 41, 43, 47, 53, _
          59, 61, 67, 71, 73, 79, 83, 89, _
          97, 101, 103, 107, 109, 113, 127, 131, _
          137, 139, 149, 151, 157, 163, 167, 173, _
          179, 181, 191, 193, 197, 199, 211, 223, _
          227, 229, 233, 239, 241, 251, 257, 263, _
          269, 271, 277, 281, 283, 293, 307, 311)
    
    ReDim Hashes(0 To 7, 0 To NumDeBlocos) As String
    If ConstH = "" Then
        Dim VarAux As Variant
        Dim VarAuxInt As Variant
        Dim Cont As Integer
        For Cont = 0 To 7
            VarAux = Primos(Cont) ^ (1 / 2)
            VarAuxInt = _
            WorksheetFunction.RoundDown(Primos(Cont) ^ (1 / 2), 0)
            VarAux = VarAux - VarAuxInt
            VarAux = _
            WorksheetFunction.RoundDown(VarAux * 2 ^ 32, 0)
            Hashes(Cont, 0) = DecParaHex(VarAux)
        Next Cont
    Else
        For Cont = 0 To 7
            Hashes(Cont, 0) = Mid(ConstH, Cont * 8 + 1, 8)
        Next Cont
    End If
    
    Dim K(0 To 63) As String
    For Cont = 0 To 63
        VarAux = Primos(Cont) ^ (1 / 3)
        VarAuxInt = _
        WorksheetFunction.RoundDown(Primos(Cont) ^ (1 / 3), 0)
        VarAux = VarAux - VarAuxInt
        VarAux = WorksheetFunction.RoundDown(VarAux * 2 ^ 32, 0)
        K(Cont) = DecParaHex(VarAux)
    Next Cont
    
    Dim a As String
    Dim b As String
    Dim c As String
    Dim d As String
    Dim e As String
    Dim f As String
    Dim g As String
    Dim H As String
    
    Dim i As Integer
    Dim j As Integer

    For i = 0 To NumDeBlocos - 1
        a = Hashes(0, i)
        b = Hashes(1, i)
        c = Hashes(2, i)
        d = Hashes(3, i)
        e = Hashes(4, i)
        f = Hashes(5, i)
        g = Hashes(6, i)
        H = Hashes(7, i)
        
        Dim Ch As String
        Dim Maj As String
        
        Dim s_zero As String
        Dim s_um As String
        
        Dim T1 As String
        Dim T2 As String
        
        For j = 0 To 63
            Ch = Ch_Function(e, f, g)
            Maj = Maj_Function(a, b, c)
            s_zero = s_zero_Function(a)
            s_um = s_um_Function(e)
            T1 = T1_Function(H, s_um, Ch, _
            CStr(K(j)), CStr(Words(i, j)))
            
            T2 = mod_2_32_addition(s_zero, Maj)
            
            H = g
            g = f
            f = e
            e = mod_2_32_addition(d, T1)
            d = c
            c = b
            b = a
            a = mod_2_32_addition(T1, T2)
        Next j
        Hashes(0, i + 1) = mod_2_32_addition(a, Hashes(0, i))
        Hashes(1, i + 1) = mod_2_32_addition(b, Hashes(1, i))
        Hashes(2, i + 1) = mod_2_32_addition(c, Hashes(2, i))
        Hashes(3, i + 1) = mod_2_32_addition(d, Hashes(3, i))
        Hashes(4, i + 1) = mod_2_32_addition(e, Hashes(4, i))
        Hashes(5, i + 1) = mod_2_32_addition(f, Hashes(5, i))
        Hashes(6, i + 1) = mod_2_32_addition(g, Hashes(6, i))
        Hashes(7, i + 1) = mod_2_32_addition(H, Hashes(7, i))
        
    Next i
    
    Dim HashFinal As String
    HashFinal = ""
    For j = 0 To 7
        HashFinal = HashFinal & LCase(Hashes(j, i))
    Next j
        
    HASH_SHA_256 = HashFinal
                
End Function
