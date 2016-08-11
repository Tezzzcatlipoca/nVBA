Attribute VB_Name = "Module1"
Function Encuentra()

CualColumna = 7 'AJ=36 Dónde insertar los factores
ColumnaALeer = 4 'Dónde están los valores a ser evaluados
ClaveBuscada = "Tradicional" 'Poner aquí factor a ser insertado
patron1 = "radicional(\d | \D)*" 'Aquí va la primera expresión regular
patron2 = "^TD (\d | \D)*" 'Aquí va la segunda expresión regular
patron3 = "^TD.(\d | \D)*" 'Aquí va la tercera expresión regular
patron4 = "RADICIONAL" 'Aquí va la tercera expresión regular'   Combinaciones posibles
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
Total = 147 '147 = Número de registros en los archivos
Dim RE(147)
Dim Nomes(147)
Open "C:\Users\franro04\Documents\VBA\Nomes.txt" For Input As #1
Open "C:\Users\franro04\Documents\VBA\NomesRE.txt" For Input As #2 'Regular Expressions here!!
i = 0
'Read the names and regular expressions from files
Do While Not EOF(1)
    i = i + 1
    Input #1, Nomes(i)
    Input #2, RE(i)
Loop
Close #1
Close #2

'Call function and perform changes
'WARNING!!! Make sure the function is pointing to the right column.
For a = 1 To Total
    Call ModifQuiebra(RE(a), "asadasd", "asdasdasd", "hfhghg", Nomes(a))

Next

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


