Attribute VB_Name = "Module1"
Function TodasJuntas()

' Escoger los m�dulos necesarios
' Revisar para cada caso los n�meros de las columnas
Autos = 1
Trad = 1
Pad = 1
Bar = 1
Cadeia = 0
CadYAuto = 1
CashCarry = 1
Farma = 1

CualColumna = 6 'AJ=36 D�nde insertar los factores
ColumnaALeer = 4 'D�nde est�n los valores a ser evaluados


'Autoservicios
ClaveBuscada = "Autoservicio" 'Poner aqu� factor a ser insertado
patron1 = "sadflkjsdlfksja" 'Aqu� va la primera expresi�n regular
patron2 = "^AS(.*)CK" 'Aqu� va la segunda expresi�n regular
patron3 = "[^\+]Conveniencia[^\+]" 'Aqu� va la tercera expresi�n regular
patron4 = "asdfsdf"

If Autos Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)

'Para encontrar Tradicional hace falta dos ciclos:
'1er CICLO
ClaveBuscada = "Tradicional" 'Poner aqu� factor a ser insertado
patron1 = "radicional(\d | \D)*" 'Aqu� va la primera expresi�n regular
patron2 = "^TD (\d | \D)*" 'Aqu� va la segunda expresi�n regular
patron3 = "^TD.(\d | \D)*" 'Aqu� va la tercera expresi�n regular
patron4 = "RADICIONAL"

If Trad Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)

'2do CICLO
ClaveBuscada = "Tradicional" 'Poner aqu� factor a ser insertado
patron1 = "^TRAD.(\d | \D)*" 'Aqu� va la primera expresi�n regular
patron2 = "^TRAD (\d | \D)*" 'Aqu� va la segunda expresi�n regular
patron3 = "^asdasdasd(\d | \D)*" 'Aqu� va la tercera expresi�n regular
patron4 = "aaasdasd"

If Trad Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)

'Para Panader�as es necesario dos ciclos:
'1er CICLO
ClaveBuscada = "Padaria" 'Poner aqu� factor a ser insertado
patron1 = "^PADARIA(\d | \D)*" 'Aqu� va la primera expresi�n regular
patron2 = "^PAD (\d | \D)*" 'Aqu� va la segunda expresi�n regular
patron3 = "^PAD." 'Aqu� va la tercera expresi�n regular
patron4 = "^PD."

If Pad Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)

'2do CICLO
ClaveBuscada = "Padaria" 'Poner aqu� factor a ser insertado
patron1 = "^PD.(\d | \D)*" 'Aqu� va la primera expresi�n regular
patron2 = "^PD (\d | \D)*" 'Aqu� va la segunda expresi�n regular
patron3 = "asdasdasd" 'Aqu� va la tercera expresi�n regular
patron4 = "lksjdfl"

If Pad Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)


'Para bares y Horecas juntos (2 ciclos):
'1er CICLO
ClaveBuscada = "Bar" 'Poner aqu� factor a ser insertado
patron1 = "^BAR.(\d | \D)*" 'Aqu� va la primera expresi�n regular
patron2 = "^BAR (\d | \D)*" 'Aqu� va la segunda expresi�n regular
patron3 = "^HRCN(\d | \D)*" 'Aqu� va la tercera expresi�n regular
patron4 = "lksjdfl"

If Bar Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)

'2do CICLO
ClaveBuscada = "Bar" 'Poner aqu� factor a ser insertado
patron1 = "^PMIX.(\d | \D)*" 'Aqu� va la primera expresi�n regular
patron2 = "^PMIX (\d | \D)*" 'Aqu� va la segunda expresi�n regular
patron3 = "^asdjhkas(\d | \D)*" 'Aqu� va la tercera expresi�n regular
patron4 = "lksjdfl"

If Bar Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)

'Para la mayor�a de las cadenas funciona:
ClaveBuscada = "Cadena" 'Poner aqu� factor a ser insertado
patron1 = "[^(AS)]CK(\d | \D)*" 'Aqu� va la primera expresi�n regular
patron2 = "^asdasda (\d | \D)*" 'Aqu� va la segunda expresi�n regular
patron3 = "^kalsdkj(\d | \D)*" 'Aqu� va la tercera expresi�n regular
patron4 = "lksjdfl"

If Cadeia Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)

'Sirve para encontrar cadenas mezcladas con autoservicio:
ClaveBuscada = "Autoservicio" 'Poner aqu� factor a ser insertado
patron1 = "(\d)-(\d)" 'Aqu� va la primera expresi�n regular
patron2 = "^AS (\d | \D)*" 'Aqu� va la segunda expresi�n regular
patron3 = "onveniencia(\d | \D)*" 'Aqu� va la tercera expresi�n regular
patron4 = "ONVENIENCIA(\d | \D)*" 'Aqu� v

If CadYAuto Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)

'Para Cash & Carry:
ClaveBuscada = "Cash & Carry" 'Poner aqu� factor a ser insertado
patron1 = "ASH &(\d | \D)*" 'Aqu� va la primera expresi�n regular
patron2 = "^asdasda (\d | \D)*" 'Aqu� va la segunda expresi�n regular
patron3 = "^kalsdkj(\d | \D)*" 'Aqu� va la tercera expresi�n regular
patron4 = "lksjdfl" 'Aqu� va la tercera expresi�n regular'

If CashCarry Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)

'Para Farmacias:
ClaveBuscada = "Farma" 'Poner aqu� factor a ser insertado
patron1 = "FARMA(\d | \D)*" 'Aqu� va la primera expresi�n regular
patron2 = "DROGA (\d | \D)*" 'Aqu� va la segunda expresi�n regular
patron3 = "^kalsdkj(\d | \D)*" 'Aqu� va la tercera expresi�n regular
patron4 = "lksjdfl"

If Farma Then Call Encuentra(CualColumna, ColumnaALeer, ClaveBuscada, patron1, patron2, patron3, patron4)


End Function

Function Encuentra(ColIns, ColLeer, Clave, pat1, pat2, pat3, pat4)

CualColumna = ColIns 'AJ=36 D�nde insertar los factores
ColumnaALeer = ColLeer 'D�nde est�n los valores a ser evaluados
ClaveBuscada = Clave 'Poner aqu� factor a ser insertado
patron1 = pat1 'Aqu� va la primera expresi�n regular
patron2 = pat2 'Aqu� va la segunda expresi�n regular
patron3 = pat3 'Aqu� va la tercera expresi�n regular
patron4 = pat4  'Aqu� va la tercera expresi�n regular'   'Aqu� va la tercera expresi�n regular'   Combinaciones posibles
' "[^\+]*CE(.*)\+(.*)RN[^\+]*", "[^\+]*RN(.*)\+(.*)CE[^\+]*",
' "[^\+]*(CEAR|Cear|cear)(.*)\+(.*)(NORTE|norte|Norte)[^\+]*"
' "(CEAR|Cear|cear)" Utilizar varias formas de escribir lo mismo
' Omitir caracteres especiales o usar "?" "Maranh?o"
' \d any digit
' \D any NON-digit
' \( par�ntesis
' \n la expresi�n anterior equis n�mero de veces
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
        MsgBox ("Error en tu c�digo!!!")
        End
    End If
    If regEx1.Test(variable) Then Cells(a, CualColumna).Value = ClaveBuscada
    If regEx2.Test(variable) Then Cells(a, CualColumna).Value = ClaveBuscada
    If regEx3.Test(variable) Then Cells(a, CualColumna).Value = ClaveBuscada
    If regEx4.Test(variable) Then Cells(a, CualColumna).Value = ClaveBuscada

Next



End Function

Function PasaCapitais()
Dim RE(84) '84 = N�mero de registros en los archivos
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

CualColumna = 7 'AJ=36 D�nde insertar los factores
ColumnaALeer = 4 'D�nde est�n los valores a ser evaluados
ClaveBuscada = identificador 'Poner aqu� factor a ser insertado
patron1 = pat1 'Aqu� va la primera expresi�n regular
patron2 = pat2 'Aqu� va la segunda expresi�n regular
patron3 = pat3 'Aqu� va la tercera expresi�n regular
patron4 = pat4 'Aqu� va la tercera expresi�n regular'   Combinaciones posibles
' "[^\+]*CE(.*)\+(.*)RN[^\+]*", "[^\+]*RN(.*)\+(.*)CE[^\+]*",
' "[^\+]*(CEAR|Cear|cear)(.*)\+(.*)(NORTE|norte|Norte)[^\+]*"
' "(CEAR|Cear|cear)" Utilizar varias formas de escribir lo mismo
' Omitir caracteres especiales o usar "?" "Maranh?o"
' \d any digit
' \D any NON-digit
' \( par�ntesis
' \n la expresi�n anterior equis n�mero de veces
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
        MsgBox ("Error en tu c�digo!!!")
        End
    End If
    If regEx1.Test(variable) Then Cells(a, CualColumna).Value = ClaveBuscada
    If regEx2.Test(variable) Then Cells(a, CualColumna).Value = ClaveBuscada
    If regEx3.Test(variable) Then Cells(a, CualColumna).Value = ClaveBuscada
    If regEx4.Test(variable) Then Cells(a, CualColumna).Value = ClaveBuscada

Next



End Function


