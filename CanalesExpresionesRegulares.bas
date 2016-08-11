Attribute VB_Name = "Module1"
Function TodasJuntas()

' Escoger los módulos necesarios
' Revisar para cada caso los números de las columnas
Autos = 1
Trad = 1
Pad = 1
Bar = 1
Cadeia = 0
CadYAuto = 1
CashCarry = 1
Farma = 1

CualColumna = 6 'AJ=36 Dónde insertar los factores
ColumnaALeer = 4 'Dónde están los valores a ser evaluados


'Autoservicios
ClaveBuscada = "Autoservicio" 'Poner aquí factor a ser insertado
patron1 = "sadflkjsdlfksja" 'Aquí va la primera expresión regular
patron2 = "^AS(.*)CK" 'Aquí va la segunda expresión regular
patron3 = "[^\+]Conveniencia[^\+]" 'Aquí va la tercera expresión regular
patron4 = "asdfsdf"

If Autos Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)

'Para encontrar Tradicional hace falta dos ciclos:
'1er CICLO
ClaveBuscada = "Tradicional" 'Poner aquí factor a ser insertado
patron1 = "radicional(\d | \D)*" 'Aquí va la primera expresión regular
patron2 = "^TD (\d | \D)*" 'Aquí va la segunda expresión regular
patron3 = "^TD.(\d | \D)*" 'Aquí va la tercera expresión regular
patron4 = "RADICIONAL"

If Trad Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)

'2do CICLO
ClaveBuscada = "Tradicional" 'Poner aquí factor a ser insertado
patron1 = "^TRAD.(\d | \D)*" 'Aquí va la primera expresión regular
patron2 = "^TRAD (\d | \D)*" 'Aquí va la segunda expresión regular
patron3 = "^asdasdasd(\d | \D)*" 'Aquí va la tercera expresión regular
patron4 = "aaasdasd"

If Trad Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)

'Para Panaderías es necesario dos ciclos:
'1er CICLO
ClaveBuscada = "Padaria" 'Poner aquí factor a ser insertado
patron1 = "^PADARIA(\d | \D)*" 'Aquí va la primera expresión regular
patron2 = "^PAD (\d | \D)*" 'Aquí va la segunda expresión regular
patron3 = "^PAD." 'Aquí va la tercera expresión regular
patron4 = "^PD."

If Pad Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)

'2do CICLO
ClaveBuscada = "Padaria" 'Poner aquí factor a ser insertado
patron1 = "^PD.(\d | \D)*" 'Aquí va la primera expresión regular
patron2 = "^PD (\d | \D)*" 'Aquí va la segunda expresión regular
patron3 = "asdasdasd" 'Aquí va la tercera expresión regular
patron4 = "lksjdfl"

If Pad Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)


'Para bares y Horecas juntos (2 ciclos):
'1er CICLO
ClaveBuscada = "Bar" 'Poner aquí factor a ser insertado
patron1 = "^BAR.(\d | \D)*" 'Aquí va la primera expresión regular
patron2 = "^BAR (\d | \D)*" 'Aquí va la segunda expresión regular
patron3 = "^HRCN(\d | \D)*" 'Aquí va la tercera expresión regular
patron4 = "lksjdfl"

If Bar Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)

'2do CICLO
ClaveBuscada = "Bar" 'Poner aquí factor a ser insertado
patron1 = "^PMIX.(\d | \D)*" 'Aquí va la primera expresión regular
patron2 = "^PMIX (\d | \D)*" 'Aquí va la segunda expresión regular
patron3 = "^asdjhkas(\d | \D)*" 'Aquí va la tercera expresión regular
patron4 = "lksjdfl"

If Bar Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)

'Para la mayoría de las cadenas funciona:
ClaveBuscada = "Cadena" 'Poner aquí factor a ser insertado
patron1 = "[^(AS)]CK(\d | \D)*" 'Aquí va la primera expresión regular
patron2 = "^asdasda (\d | \D)*" 'Aquí va la segunda expresión regular
patron3 = "^kalsdkj(\d | \D)*" 'Aquí va la tercera expresión regular
patron4 = "lksjdfl"

If Cadeia Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)

'Sirve para encontrar cadenas mezcladas con autoservicio:
ClaveBuscada = "Autoservicio" 'Poner aquí factor a ser insertado
patron1 = "(\d)-(\d)" 'Aquí va la primera expresión regular
patron2 = "^AS (\d | \D)*" 'Aquí va la segunda expresión regular
patron3 = "onveniencia(\d | \D)*" 'Aquí va la tercera expresión regular
patron4 = "ONVENIENCIA(\d | \D)*" 'Aquí v

If CadYAuto Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)

'Para Cash & Carry:
ClaveBuscada = "Cash & Carry" 'Poner aquí factor a ser insertado
patron1 = "ASH &(\d | \D)*" 'Aquí va la primera expresión regular
patron2 = "^asdasda (\d | \D)*" 'Aquí va la segunda expresión regular
patron3 = "^kalsdkj(\d | \D)*" 'Aquí va la tercera expresión regular
patron4 = "lksjdfl" 'Aquí va la tercera expresión regular'

If CashCarry Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)

'Para Farmacias:
ClaveBuscada = "Farma" 'Poner aquí factor a ser insertado
patron1 = "FARMA(\d | \D)*" 'Aquí va la primera expresión regular
patron2 = "DROGA (\d | \D)*" 'Aquí va la segunda expresión regular
patron3 = "^kalsdkj(\d | \D)*" 'Aquí va la tercera expresión regular
patron4 = "lksjdfl"

If Farma Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)


End Function

Function Encuentra(ColIns, ColLeer, Clave, pat1, pat2, pat3, pat4)

CualColumna = ColIns 'AJ=36 Dónde insertar los factores
ColumnaALeer = ColLeer 'Dónde están los valores a ser evaluados
ClaveBuscada = Clave 'Poner aquí factor a ser insertado
patron1 = pat1 'Aquí va la primera expresión regular
patron2 = pat2 'Aquí va la segunda expresión regular
patron3 = pat3 'Aquí va la tercera expresión regular
patron4 = pat4  'Aquí va la tercera expresión regular'   'Aquí va la tercera expresión regular'   Combinaciones posibles
' "[^\+]*CE(.*)\+(.*)RN[^\+]*", "[^\+]*RN(.*)\+(.*)CE[^\+]*",
' "[^\+]*(CEAR|Cear|cear)(.*)\+(.*)(NORTE|norte|Norte)[^\+]*"
' "(CEAR|Cear|cear)" Utilizar varias formas de escribir lo mismo
' Omitir caracteres especiales o usar "?" "Maranh?o"
' \d any digit
' \D any NON-digit
' \( paréntesis
' \n la expresión anterior equis número de veces
' \w cualquier caracter de palabra [a-zA-Z_0-9]
' [xyz] cualquier caracter de estos (OR)
' [^xyz] ninguno de estos
' \s cualquier caracter de espacio
' ? cero o una vez. Equivale a {0,1}
' * cero o muchas. Equivale a {0,}
' + una o muchas. Equivale a {1,}
'


Dim regEx1 As New RegExp
Dim regEx2 As New RegExp
Dim regEx3 As New RegExp
Dim regEx4 As New RegExp
With regEx1
    .Global = True
    .MultiLine = True
    .IgnoreCase = True
    .Pattern = patron1
End With
With regEx2
    .Global = True
    .MultiLine = True
    .IgnoreCase = True
    .Pattern = patron2
End With
With regEx3
    .Global = True
    .MultiLine = True
    .IgnoreCase = True
    .Pattern = patron3
End With
With regEx4
    .Global = True
    .MultiLine = True
    .IgnoreCase = True
    .Pattern = patron4
End With

LastD = Range("D1").End(xlDown).Row

For a = 1 To LastD
    variable = Cells(a, ColumnaALeer).Value
    If (regEx1.Test(variable) And (regEx2.Test(variable) And (regEx3.Test(variable) And regEx4.Test(variable)))) And Cells(a, CualColumna).Value <> "" Then
        MsgBox ("Error en tu código!!!")
        End
    End If
    If regEx1.Test(variable) Then Cells(a, CualColumna).Value = ClaveBuscada
    If regEx2.Test(variable) Then Cells(a, CualColumna).Value = ClaveBuscada
    If regEx3.Test(variable) Then Cells(a, CualColumna).Value = ClaveBuscada
    If regEx4.Test(variable) Then Cells(a, CualColumna).Value = ClaveBuscada

Next



End Function

Function PasaCapitais()
Dim RE(84) '84 = Número de registros en los archivos
Dim Nomes(84)
Open "Nomes.txt" For Input As #1
Open "NomesRE.txt" For Input As #2
i = 0
Do While Not EOF(1)
    i = i + 1
    Input #1, Nomes(i)
    Input #2, RE(i)
Loop
Close #1

End Function

Function ModifQuiebra(pat1, pat2, pat3, pat4, identificador)

CualColumna = 7 'AJ=36 Dónde insertar los factores
ColumnaALeer = 4 'Dónde están los valores a ser evaluados
ClaveBuscada = identificador 'Poner aquí factor a ser insertado
patron1 = pat1 'Aquí va la primera expresión regular
patron2 = pat2 'Aquí va la segunda expresión regular
patron3 = pat3 'Aquí va la tercera expresión regular
patron4 = pat4 'Aquí va la tercera expresión regular'   Combinaciones posibles
' "[^\+]*CE(.*)\+(.*)RN[^\+]*", "[^\+]*RN(.*)\+(.*)CE[^\+]*",
' "[^\+]*(CEAR|Cear|cear)(.*)\+(.*)(NORTE|norte|Norte)[^\+]*"
' "(CEAR|Cear|cear)" Utilizar varias formas de escribir lo mismo
' Omitir caracteres especiales o usar "?" "Maranh?o"
' \d any digit
' \D any NON-digit
' \( paréntesis
' \n la expresión anterior equis número de veces
' \w cualquier caracter de palabra [a-zA-Z_0-9]
' [xyz] cualquier caracter de estos (OR)
' [^xyz] ninguno de estos
' \s cualquier caracter de espacio
' ? cero o una vez. Equivale a {0,1}
' * cero o muchas. Equivale a {0,}
' + una o muchas. Equivale a {1,}
'


Dim regEx1 As New RegExp
Dim regEx2 As New RegExp
Dim regEx3 As New RegExp
Dim regEx4 As New RegExp
With regEx1
    .Global = True
    .MultiLine = True
    .IgnoreCase = True
    .Pattern = patron1
End With
With regEx2
    .Global = True
    .MultiLine = True
    .IgnoreCase = True
    .Pattern = patron2
End With
With regEx3
    .Global = True
    .MultiLine = True
    .IgnoreCase = True
    .Pattern = patron3
End With
With regEx4
    .Global = True
    .MultiLine = True
    .IgnoreCase = True
    .Pattern = patron4
End With

LastD = Range("D1").End(xlDown).Row

For a = 1 To LastD
    variable = Cells(a, ColumnaALeer).Value
    If (regEx1.Test(variable) And (regEx2.Test(variable) And (regEx3.Test(variable) And regEx4.Test(variable)))) And Cells(a, CualColumna).Value <> "" Then
        MsgBox ("Error en tu código!!!")
        End
    End If
    If regEx1.Test(variable) Then Cells(a, CualColumna).Value = ClaveBuscada
    If regEx2.Test(variable) Then Cells(a, CualColumna).Value = ClaveBuscada
    If regEx3.Test(variable) Then Cells(a, CualColumna).Value = ClaveBuscada
    If regEx4.Test(variable) Then Cells(a, CualColumna).Value = ClaveBuscada

Next



End Function


