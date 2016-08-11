Attribute VB_Name = "Module2"

Sub ListaPalabras()
Dim coinc(25084)
Dim conteo(25084)
CualVamos = 0
ColumnaALeer = 2 ' En este caso es B


    LastD = Range("B1").End(xlDown).Row
    For a = 1 To LastD
        variable = Cells(a, ColumnaALeer).Value
        palabras = Split(variable)
        For Each bito In palabras
            If bit = " " Then GoTo skipo
            Si = YaEsta(coinc, CualVamos, bito)
            Select Case Si
                Case Is > 0: conteo(Si) = conteo(Si) + 1: GoTo skipo
                Case Else: CualVamos = CualVamos + 1: coinc(CualVamos) = bito: conteo(CualVamos) = 1
            End Select
skipo:
        Next bito
        
    Next a

    Open "C:\Users\franro04\Documents\VBA\PalCads.csv" For Output As #1
        For a = 1 To CualVamos
            linea = coinc(a) & "," & conteo(a)
            Print #1, linea
        Next a
    Close #1

End Sub
Function YaEsta(coinc, CualVamos, bito)
valor = 0
    For b = 1 To CualVamos
        If bito = coinc(b) Then
                valor = b
                Exit For
        End If
    Next b
YaEsta = valor
End Function

Sub moreless()

ColumnaALeer = 7 'Columna G : Nome
CualColumna = 61 'Columna BI

    LastD = Range("B1").End(xlDown).Row
        For a = 1 To LastD
            variable = Cells(a, ColumnaALeer).Value
            variable = FiltroComunes(variable)
            variable = LCase(variable)
            calif = evalua(variable)
            califped = Split(calif, ":")
            Z = 0
            For Each comcad In califped
               If Z = 0 Then NCadena = comcad Else NPuntos = Val(comcad)
               Z = Z + 1
            Next comcad
            If NPuntos > 0 Then
                Cells(a, CualColumna).Value = NCadena
                Cells(a, CualColumna + 1).Value = NPuntos
            End If

        Next a


End Sub
Function evalua(variable)

Dim cadena(1000)
Dim coincidencias(1000)
For a = 1 To 1000
        coincidencias(a) = 0
Next a

' Extract chain names from file
Open "C:\Users\franro04\Documents\VBA\Cadenas.csv" For Input As #2
linea = 0
Do While Not EOF(2)
    linea = linea + 1
    Input #2, cadena(linea)
    cadena(linea) = LCase(cadena(linea))
Loop
Close #2



'Now evaluate every bit against chains
largo = Len(variable)
palabras = Split(variable)
marcador = ""
    For Each bito In palabras
        If Len(bito) < 5 Then
            result = igual(bito, cadena, linea)
        Else
            result = contiene(bito, cadena, linea)
        End If
        If result <> "" And marcador = "" Then
            marcador = result
        ElseIf result <> "" And marcador <> "" Then
            marcador = marcador & "-" & result
        End If
    Next bito

'Resumen de puntos
componentes = Split(marcador, "-")
cuentalos = 0
   For Each pedazo In componentes
       cuentalos = cuentalos + 1
       lugar = InStr(1, pedazo, ":")
       numstr = Left(pedazo, lugar - 1)
       cadenita = Val(numstr)
       puntosstr = Right(pedazo, Len(pedazo) - lugar)
       puntos = Val(puntosstr)
       coincidencias(cadenita) = coincidencias(cadenita) + puntos
   Next pedazo

Record = 0
    For a = 1 To 1000
        suma = coincidencias(a)
        If suma > Record Then
            Record = suma
            NumCadena = a
        End If
    Next a

If Record > 0 Then
    devuelve = Str(NumCadena) & ":" & Str(Record)
    evalua = devuelve
Else
    evalua = "0:0"
End If
End Function


Function igual(bito, cadena, linea)
    Resultado = ""
    For a = 1 To linea
        If cadena(a) = bito Then
            Resultado = Resultado & Str(a) & ":" & Len(bito) & "-"
            MsgBox (Resultado)
        Else
            Resultado = ""
        End If
    Next a
If Len(Resultado) > 1 And Right(Resultado, 1) = "-" Then Resultado = Left(Resultado, Len(Resultado) - 1)
igual = Resultado
End Function


Function contiene(bito, cadena, linea)
Dim regEx1 As New RegExp

With regEx1
    .Global = True
    .MultiLine = True
    .IgnoreCase = True
End With

    Resultado = ""
    For a = 1 To linea
        For b = 1 To Len(bito) - 1
            nuevobito = Left(bito, b) & "(.?)" & Right(bito, Len(bito) - b)
            regEx1.Pattern = nuevobito
            If regEx1.Test(cadena(a)) Then
                Resultado = Resultado & Str(a) & ":" & Len(bito) & "-"
                MsgBox (Resultado)
            Else
                Resultado = ""
            End If
        Next b
    Next a
If Len(Resultado) > 1 And Right(Resultado, 1) = "-" Then Resultado = Left(Resultado, Len(Resultado) - 1)
contiene = Resultado
End Function


Function FiltroComunes(cadena)
cadena = LCase(cadena)
Dim Palabra(200)
    Open "C:\Users\franro04\Documents\VBA\PalsSolo.csv" For Input As #1
    nume = 0
    Do While Not EOF(1)
        nume = nume + 1
        Input #1, Palabra(nume)
        Palabra(nume) = LCase(Palabra(nume))
        If Len(Palabra(nume)) < 4 Then Palabra(nume) = " " & Palabra(nume)
        Palabra(nume) = Palabra(nume) & " "
    Loop
    Close #1
    For a = 1 To nume
        cont = InStr(1, cadena, Palabra(a))
        If cont <> 0 Then
            partida = Split(cadena, Palabra(a))
            nueva = ""
            For Each biit In partida
                nueva = nueva & biit & " "
            Next biit
            cadena = Trim(nueva)
        End If
    Next a
    FiltroComunes = cadena
End Function

Function barely()




End Function
