Attribute VB_Name = "y"
Function CompactarCodigo(codigo As String) As String
    Dim valor As Double
    valor = CDbl(codigo)
    
    ' Converter para base 36
    CompactarCodigo = ConverterParaBase(valor, 36)
End Function

Function ConverterParaBase(valor As Double, base As Integer) As String
    Dim resultado As String
    Dim quociente As Double
    Dim resto As Integer
    
    Do While valor > 0
        quociente = Int(valor / base)
        resto = valor - quociente * base
        resultado = DigitoParaCaractere(resto) & resultado
        valor = quociente
    Loop
    
    ConverterParaBase = resultado
End Function

Function DigitoParaCaractere(digito As Integer) As String
    If digito >= 0 And digito <= 9 Then
        DigitoParaCaractere = CStr(digito)
    ElseIf digito >= 10 And digito <= 35 Then
        DigitoParaCaractere = Chr(digito + 55)
    Else
        DigitoParaCaractere = ""
    End If
End Function

Function DescompactarCodigo(codigoCompactado As String) As String
    Dim valor As Double
    valor = ConverterDeBase(codigoCompactado, 36)
    
    ' Converter para formato de 12 dígitos numéricos
    DescompactarCodigo = Format$(valor, "000000000000")
End Function

Function ConverterDeBase(codigo As String, base As Integer) As Double
    Dim resultado As Double
    Dim i As Integer
    Dim caractere As String
    
    For i = 1 To Len(codigo)
        caractere = Mid(codigo, i, 1)
        resultado = resultado * base + CaractereParaDigito(caractere)
    Next i
    
    ConverterDeBase = resultado
End Function

Function CaractereParaDigito(caractere As String) As Integer
    If IsNumeric(caractere) Then
        CaractereParaDigito = CInt(caractere)
    ElseIf caractere >= "A" And caractere <= "Z" Then
        CaractereParaDigito = Asc(caractere) - 55
    Else
        CaractereParaDigito = -1
    End If
End Function
