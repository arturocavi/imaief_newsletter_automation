Attribute VB_Name = "AUT_IMAIEF_24_01"
Global Grafica_desytc
Global Grafica_act_sec_rank
Global Grafica_desest_rank_anu
Global Grafica_desest_rank_men
Global Grafica_con_rank
Global Grafica_man_rank
Global Grafica_act_com
Global grafica_act_sec_var
Global grafica_con_var
Global grafica_man_var

Sub Macro_inicio()
Attribute Macro_inicio.VB_ProcData.VB_Invoke_Func = "a\n14"
'Ctrl + a

Call WorksheetLoop
Call Generación_Word
End Sub

Sub WorksheetLoop()

        Dim WS_Count As Integer
        Dim I As Integer

         ' Set WS_Count equal to the number of worksheets in the active
         ' workbook.
        WS_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the loop.
        For I = 1 To WS_Count
            
            ' Insert your code here.
            ' The following line shows how to reference a sheet within
            ' the loop by displaying the worksheet name in a dialog box.
            
            'MsgBox ActiveWorkbook.Worksheets(I).Name
            
            Worksheets(I).Select
            nombre = Worksheets(I).Name
            
            If InStr(1, nombre, "VAR", vbBinaryCompare) <> 0 Then
                Call Macro_Color_Mes
            ElseIf InStr(1, nombre, "RANK", vbBinaryCompare) <> 0 Then
                Call Macro_Ranking
            ElseIf InStr(1, nombre, "COM", vbBinaryCompare) <> 0 Then
                Call Macro_Compara_Industria
            ElseIf InStr(1, nombre, "DESYTC", vbBinaryCompare) <> 0 Then
                Call Macro_Var_Serie_Des
            End If

        Next I

End Sub

Sub Macro_Color_Mes()
'Barras de mismo color para mismo mes
'Gráfica de Barras de Históricos Mensuales  con Linea de Promedio de Últimos 12 Meses
'Ctrl + p
If InStr(1, ActiveSheet.Name, "ACT SEC VAR", vbBinaryCompare) = 1 Then
    Range("C6:D6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "0.0"
    
    nombre = ActiveSheet.Name
    
    inicio = 5
    fin = inicio
    
    Do While Cells(fin, 2) <> ""
        fin = fin + 1
    Loop
    
    Range("A" & (inicio) & ":D" & (fin - 1)).Select
    
    Set grafica_act_sec_var = ActiveSheet.ChartObjects.Add(Left:=5 * 48, Width:=468.1, Top:=60, Height:=250)
        
        With grafica_act_sec_var.Chart
            .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\BO Barras Prom Editable.crtx") ' UBICACIÓN PERSONAL
            .SetSourceData Source:=Range("A" & (inicio) & ":D" & (fin - 1))
            For k = 1 To (fin - 1)
                If (k Mod 12) = ((fin - 1 - inicio) Mod 12) Then
                    .SeriesCollection(1).Points(k).Format.Fill.ForeColor.RGB = RGB(124, 135, 142)
                End If
            Next k
            .SeriesCollection(1).Points(fin - 1 - inicio).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
        End With
ElseIf InStr(1, ActiveSheet.Name, "CON VAR", vbBinaryCompare) = 1 Then
    Range("C6:D6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "0.0"
    
    nombre = ActiveSheet.Name
    
    inicio = 5
    fin = inicio
    
    Do While Cells(fin, 2) <> ""
        fin = fin + 1
    Loop
    
    Range("A" & (inicio) & ":D" & (fin - 1)).Select
    
    Set grafica_con_var = ActiveSheet.ChartObjects.Add(Left:=5 * 48, Width:=468.1, Top:=60, Height:=250)
        
        With grafica_con_var.Chart
            .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\BO Barras Prom Editable.crtx") ' UBICACIÓN PERSONAL
            .SetSourceData Source:=Range("A" & (inicio) & ":D" & (fin - 1))
            For k = 1 To (fin - 1)
                If (k Mod 12) = ((fin - 1 - inicio) Mod 12) Then
                    .SeriesCollection(1).Points(k).Format.Fill.ForeColor.RGB = RGB(124, 135, 142)
                End If
            Next k
            .SeriesCollection(1).Points(fin - 1 - inicio).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
        End With
ElseIf InStr(1, ActiveSheet.Name, "MAN VAR", vbBinaryCompare) = 1 Then
    Range("C6:D6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "0.0"
    
    nombre = ActiveSheet.Name
    
    inicio = 5
    fin = inicio
    
    Do While Cells(fin, 2) <> ""
        fin = fin + 1
    Loop
    
    Range("A" & (inicio) & ":D" & (fin - 1)).Select
    
    Set grafica_man_var = ActiveSheet.ChartObjects.Add(Left:=5 * 48, Width:=468.1, Top:=60, Height:=250)
        
        With grafica_man_var.Chart
            .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\BO Barras Prom Editable.crtx") ' UBICACIÓN PERSONAL
            .SetSourceData Source:=Range("A" & (inicio) & ":D" & (fin - 1))
            For k = 1 To (fin - 1)
                If (k Mod 12) = ((fin - 1 - inicio) Mod 12) Then
                    .SeriesCollection(1).Points(k).Format.Fill.ForeColor.RGB = RGB(124, 135, 142)
                End If
            Next k
            .SeriesCollection(1).Points(fin - 1 - inicio).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
        End With
End If
End Sub
Sub Macro_Ranking()
Attribute Macro_Ranking.VB_ProcData.VB_Invoke_Func = "g\n14"
'Macro que convierte datos en un Ranking de gráfica de barras
' ctrl + g

'Dim Grafica As ChartObject
If InStr(1, ActiveSheet.Name, "ACT SEC RANK", vbBinaryCompare) = 1 Then
    inicio = 5
    fin = inicio
    
    Do While Cells(fin, 1) <> ""
        fin = fin + 1
    Loop
    
    '
    
    '
    Range("B" & (inicio + 1) & ":B" & (fin - 1)).NumberFormat = "0.0"
    hoja = ActiveSheet.Name
    
    '
    Range("A" & inicio & ":B" & (fin - 1)).AutoFilter
    '
    
    Range("A" & (inicio + 1) & ":B" & (fin - 1)).Sort Key1:=Range("B" & (inicio + 1)), Order1:=xlAscending
    
    '
    jalisco = inicio
    nacional = inicio
    
    Do While Cells(jalisco, 1) <> "Jalisco"
        jalisco = jalisco + 1
    Loop
    jalisco = jalisco - inicio
    
    Do While Cells(nacional, 1) <> "Nacional"
        nacional = nacional + 1
    Loop
    nacional = nacional - inicio
    '
    Range("A" & (inicio + 1) & ":B" & (fin - 1)).Select
    '
    Set Grafica_act_sec_rank = ActiveSheet.ChartObjects.Add(Left:=287, Width:=463, Top:=105, Height:=448.5)
    
    With Grafica_act_sec_rank.Chart
        .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\AUT RANKING.crtx")
        .SetSourceData Source:=Range("A" & (inicio + 1) & ":B" & (fin - 1))
        .SeriesCollection(1).Points(jalisco).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
        .SeriesCollection(1).Points(nacional).Format.Fill.ForeColor.RGB = RGB(149, 104, 43)
    End With
    '
    Range("A" & inicio & ":B" & (fin - 1)).AutoFilter
    '
ElseIf InStr(1, ActiveSheet.Name, "DESEST RANK ANU", vbBinaryCompare) = 1 Then
    inicio = 5
    fin = inicio
    
    Do While Cells(fin, 1) <> ""
        fin = fin + 1
    Loop
    
    '
    
    '
    Range("B" & (inicio + 1) & ":B" & (fin - 1)).NumberFormat = "0.0"
    hoja = ActiveSheet.Name
    
    '
    Range("A" & inicio & ":B" & (fin - 1)).AutoFilter
    '
    
    Range("A" & (inicio + 1) & ":B" & (fin - 1)).Sort Key1:=Range("B" & (inicio + 1)), Order1:=xlAscending
    
    '
    jalisco = inicio
    nacional = inicio
    
    Do While Cells(jalisco, 1) <> "Jalisco"
        jalisco = jalisco + 1
    Loop
    jalisco = jalisco - inicio
    
    Do While Cells(nacional, 1) <> "Nacional"
        nacional = nacional + 1
    Loop
    nacional = nacional - inicio
    '
    Range("A" & (inicio + 1) & ":B" & (fin - 1)).Select
    '
    Set Grafica_desest_rank_anu = ActiveSheet.ChartObjects.Add(Left:=287, Width:=463, Top:=105, Height:=448.5)
    
    With Grafica_desest_rank_anu.Chart
        .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\AUT RANKING.crtx")
        .SetSourceData Source:=Range("A" & (inicio + 1) & ":B" & (fin - 1))
        .SeriesCollection(1).Points(jalisco).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
        .SeriesCollection(1).Points(nacional).Format.Fill.ForeColor.RGB = RGB(149, 104, 43)
    End With
    '
    Range("A" & inicio & ":B" & (fin - 1)).AutoFilter
    '
ElseIf InStr(1, ActiveSheet.Name, "DESEST RANK MEN", vbBinaryCompare) = 1 Then
    inicio = 5
    fin = inicio
    
    Do While Cells(fin, 1) <> ""
        fin = fin + 1
    Loop
    
    '
    
    '
    Range("B" & (inicio + 1) & ":B" & (fin - 1)).NumberFormat = "0.0"
    hoja = ActiveSheet.Name
    
    '
    Range("A" & inicio & ":B" & (fin - 1)).AutoFilter
    '
    
    Range("A" & (inicio + 1) & ":B" & (fin - 1)).Sort Key1:=Range("B" & (inicio + 1)), Order1:=xlAscending
    
    '
    jalisco = inicio
    nacional = inicio
    
    Do While Cells(jalisco, 1) <> "Jalisco"
        jalisco = jalisco + 1
    Loop
    jalisco = jalisco - inicio
    
    Do While Cells(nacional, 1) <> "Nacional"
        nacional = nacional + 1
    Loop
    nacional = nacional - inicio
    '
    Range("A" & (inicio + 1) & ":B" & (fin - 1)).Select
    '
    Set Grafica_desest_rank_men = ActiveSheet.ChartObjects.Add(Left:=287, Width:=463, Top:=105, Height:=448.5)
    
    With Grafica_desest_rank_men.Chart
        .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\AUT RANKING.crtx")
        .SetSourceData Source:=Range("A" & (inicio + 1) & ":B" & (fin - 1))
        .SeriesCollection(1).Points(jalisco).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
        .SeriesCollection(1).Points(nacional).Format.Fill.ForeColor.RGB = RGB(149, 104, 43)
    End With
    '
    Range("A" & inicio & ":B" & (fin - 1)).AutoFilter
    '
ElseIf InStr(1, ActiveSheet.Name, "CON RANK", vbBinaryCompare) = 1 Then
    inicio = 5
    fin = inicio
    
    Do While Cells(fin, 1) <> ""
        fin = fin + 1
    Loop
    
    '
    
    '
    Range("B" & (inicio + 1) & ":B" & (fin - 1)).NumberFormat = "0.0"
    hoja = ActiveSheet.Name
    
    '
    Range("A" & inicio & ":B" & (fin - 1)).AutoFilter
    '
    
    Range("A" & (inicio + 1) & ":B" & (fin - 1)).Sort Key1:=Range("B" & (inicio + 1)), Order1:=xlAscending
    
    '
    jalisco = inicio
    nacional = inicio
    
    Do While Cells(jalisco, 1) <> "Jalisco"
        jalisco = jalisco + 1
    Loop
    jalisco = jalisco - inicio
    
    Do While Cells(nacional, 1) <> "Nacional"
        nacional = nacional + 1
    Loop
    nacional = nacional - inicio
    '
    Range("A" & (inicio + 1) & ":B" & (fin - 1)).Select
    '
    Set Grafica_con_rank = ActiveSheet.ChartObjects.Add(Left:=287, Width:=463, Top:=105, Height:=448.5)
    
    With Grafica_con_rank.Chart
        .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\AUT RANKING.crtx")
        .SetSourceData Source:=Range("A" & (inicio + 1) & ":B" & (fin - 1))
        .SeriesCollection(1).Points(jalisco).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
        .SeriesCollection(1).Points(nacional).Format.Fill.ForeColor.RGB = RGB(149, 104, 43)
    End With
    '
    Range("A" & inicio & ":B" & (fin - 1)).AutoFilter
    '
ElseIf InStr(1, ActiveSheet.Name, "MAN RANK", vbBinaryCompare) = 1 Then
    inicio = 5
    fin = inicio
    
    Do While Cells(fin, 1) <> ""
        fin = fin + 1
    Loop
    
    '
    
    '
    Range("B" & (inicio + 1) & ":B" & (fin - 1)).NumberFormat = "0.0"
    hoja = ActiveSheet.Name
    
    '
    Range("A" & inicio & ":B" & (fin - 1)).AutoFilter
    '
    
    Range("A" & (inicio + 1) & ":B" & (fin - 1)).Sort Key1:=Range("B" & (inicio + 1)), Order1:=xlAscending
    
    '
    jalisco = inicio
    nacional = inicio
    
    Do While Cells(jalisco, 1) <> "Jalisco"
        jalisco = jalisco + 1
    Loop
    jalisco = jalisco - inicio
    
    Do While Cells(nacional, 1) <> "Nacional"
        nacional = nacional + 1
    Loop
    nacional = nacional - inicio
    '
    Range("A" & (inicio + 1) & ":B" & (fin - 1)).Select
    '
    Set Grafica_man_rank = ActiveSheet.ChartObjects.Add(Left:=287, Width:=463, Top:=105, Height:=448.5)
    
    With Grafica_man_rank.Chart
        .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\AUT RANKING.crtx")
        .SetSourceData Source:=Range("A" & (inicio + 1) & ":B" & (fin - 1))
        .SeriesCollection(1).Points(jalisco).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
        .SeriesCollection(1).Points(nacional).Format.Fill.ForeColor.RGB = RGB(149, 104, 43)
    End With
    '
    Range("A" & inicio & ":B" & (fin - 1)).AutoFilter
    '
End If
End Sub

Sub Macro_Var_Serie_Des()
Attribute Macro_Var_Serie_Des.VB_ProcData.VB_Invoke_Func = "t\n14"
'Macro para variación de serie desestacionalizada
'Gráfica de Barras de Históricos Mensuales  con Linea de tendencia-ciclo de
'Ctrl + t

Range("C6:D6").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.NumberFormat = "0.0"

nombre = ActiveSheet.Name

inicio = 5
fin = inicio

Do While Cells(fin, 2) <> ""
    fin = fin + 1
Loop

Range("A" & (inicio) & ":D" & (fin - 1)).Select

Set Grafica_desytc = ActiveSheet.ChartObjects.Add(Left:=5 * 48, Width:=468.1, Top:=60, Height:=250)
    
    With Grafica_desytc.Chart
        .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\BO Barras Prom Editable.crtx") ' UBICACIÓN PERSONAL
        .SetSourceData Source:=Range("A" & (inicio) & ":D" & (fin - 1))
        For k = 1 To (fin - 1)
            If (k Mod 12) = ((fin - 1 - inicio) Mod 12) Then
                .SeriesCollection(1).Points(k).Format.Fill.ForeColor.RGB = RGB(124, 135, 142)
            End If
        Next k
        .SeriesCollection(1).Points(fin - 1 - inicio).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
        .Axes(xlValue).MinimumScale = 60
        .Axes(xlValue).MaximumScale = 110
    End With



End Sub


Sub Macro_Compara_Industria()
Attribute Macro_Compara_Industria.VB_ProcData.VB_Invoke_Func = "h\n14"

'Macro para comparativo de gráfico de barras por industria
'Ctrl + h

inicio = 5
fin = inicio

Do While Cells(fin, 2) <> ""
    fin = fin + 1
Loop

Range("B" & (inicio + 1) & ":B" & (fin - 1)).NumberFormat = "0.0"

Set Grafica_act_com = ActiveSheet.ChartObjects.Add(Left:=287, Width:=463, Top:=105, Height:=250)

With Grafica_act_com.Chart
    .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\AUT IMAIEF C5.crtx")
    .SetSourceData Source:=Range("A" & (inicio + 1) & ":B" & (fin - 1))
End With


End Sub

Sub Generación_Word()
'Nombre  y ubicación de la plantilla
plantilla = "C:\Users\arturo.carrillo\Documents\IMAIEF\AUT\PLANTILLA.dotx" ' UBICACIÓN PERSONAL

'Creamos el nuevo archivo word usando la plantilla
Set aplicacion = CreateObject("Word.Application")
aplicacion.Visible = True
Set documento = aplicacion.Documents.Add(Template:=plantilla, NewTemplate:=False, DocumentType:=0)


'Cambiamos la fecha del encabezado
diahoy = Format(Day(Now), "00")
meshoy = Format(Month(Now), "00")
añohoy = Year(Now)
If Month(Now) = 1 Then
    meshoypal = "enero"
    mesbas = Format(9, "00")
    mesbaspal = "septiembre"
    añobas = Year(Now) - 1
ElseIf Month(Now) = 2 Then
    meshoypal = "febrero"
    mesbas = Format(10, "00")
    mesbaspal = "octubre"
    añobas = Year(Now) - 1
ElseIf Month(Now) = 3 Then
    meshoypal = "marzo"
    mesbas = Format(11, "00")
    mesbaspal = "noviembre"
    añobas = Year(Now) - 1
ElseIf Month(Now) = 4 Then
    meshoypal = "abril"
    mesbas = Format(12, "00")
    mesbaspal = "diciembre"
    añobas = Year(Now) - 1
ElseIf Month(Now) = 5 Then
    meshoypal = "mayo"
    mesbas = Format(1, "00")
    mesbaspal = "enero"
    añobas = Year(Now)
ElseIf Month(Now) = 6 Then
    meshoypal = "junio"
    mesbas = Format(2, "00")
    mesbaspal = "febrero"
    añobas = Year(Now)
ElseIf Month(Now) = 7 Then
    meshoypal = "julio"
    mesbas = Format(3, "00")
    mesbaspal = "marzo"
    añobas = Year(Now)
ElseIf Month(Now) = 8 Then
    meshoypal = "agosto"
    mesbas = Format(4, "00")
    mesbaspal = "abril"
    añobas = Year(Now)
ElseIf Month(Now) = 9 Then
    meshoypal = "septiembre"
    mesbas = Format(5, "00")
    mesbaspal = "mayo"
    añobas = Year(Now)
ElseIf Month(Now) = 10 Then
    meshoypal = "octubre"
    mesbas = Format(6, "00")
    mesbaspal = "junio"
    añobas = Year(Now)
ElseIf Month(Now) = 11 Then
    meshoypal = "noviembre"
    mesbas = Format(7, "00")
    mesbaspal = "julio"
    añobas = Year(Now)
ElseIf Month(Now) = 12 Then
    meshoypal = "diciembre"
    mesbas = Format(8, "00")
    mesbaspal = "agosto"
    añobas = Year(Now)
End If

'FECHAS MANUALES
'diahoy = InputBox("Ingresa el día de hoy en formato de número a dos dígitos (ej. 23):")'
'meshoy = InputBox("Ingresa el mes de hoy en formato de número a dos dígitos (ej. 10):")
'añohoy = InputBox("Ingresa el año de hoy en formato de número a cuatro dígitos (ej. 2019):")
'meshoypal = InputBox("Ingresa el mes de hoy en formato de palabra en minúsculas (ej. octubre):")
'mesbas = InputBox("Ingresa el mes de la última base de datos del INEGI (dos meses atrás) en formato de número a dos dígitos (ej. 08):")
'mesbaspal = InputBox("Ingresa el mes de la última base de datos del INEGI (dos meses atrás) en formato de palabra en minúsculas (ej. agosto):")
'añobas = InputBox("Ingresa el año de la última base de datos del INEGI (dos meses atrás) en formato de número a cuatro dígitos (ej. 2019):")


' Cambiamos los espaciados del boletín
With documento.Content
    .Style = "Espaciado principal"
End With

' Insertar título del boletín
documento.Content.insertparagraphafter

With documento.Content
    .InsertAfter Hoja11.Cells(2, 1).Value ' Título del boletín [Paragraphs(2)]
    .insertparagraphafter
End With


With documento.Paragraphs(2).Range
    .Style = "Título 1"
End With


' ACT SEC VAR
' Insertar párrafo de texto Texto1_1
With documento.Content
    .InsertAfter Hoja11.Cells(5, 2).Value ' Texto1_1 [Paragraphs(4)]
    .insertparagraphafter
End With

With documento.Paragraphs(3).Range
    .Style = "Normal"
End With

' Insertar párrafo de texto Texto1_2
With documento.Content
    .InsertAfter Hoja11.Cells(6, 2).Value ' Texto1_2 [Paragraphs(5)]
    .insertparagraphafter
End With

With documento.Paragraphs(4).Range
    .Style = "Normal"
End With

' Insertar párrafo de texto Texto1_3
With documento.Content
    .InsertAfter Hoja11.Cells(7, 2).Value ' Texto1_3 [Paragraphs(6)]
    .insertparagraphafter
End With

With documento.Paragraphs(5).Range
    .Style = "Normal"
End With

' Insertar título de gráfica ACT SEC VAR
With documento.Content
    .InsertAfter Hoja11.Cells(8, 2).Value ' Título de gráfica ACT SEC VAR [Paragraphs(7)]
    .insertparagraphafter
End With

With documento.Paragraphs(6).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica ACT SEC VAR
grafica_act_sec_var.Chart.ChartArea.Copy
documento.Paragraphs(7).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter Hoja1.Cells(2, 1).Value ' Nota [Paragraphs(9)]
    .insertparagraphafter
End With

With documento.Paragraphs(8).Range
    .Style = "Fuentes"
End With
' Insertar nota
With documento.Content
    .InsertAfter Hoja11.Cells(9, 2).Value ' Nota [Paragraphs(10)]
    .insertparagraphafter
End With

With documento.Paragraphs(9).Range
    .Style = "Fuentes"
End With
' Insertar salto de página
documento.Paragraphs(10).Range.InsertBreak Type:=7 'wdSectionBreakNextPage
documento.Content.insertparagraphafter
'*********
'ACT SEC RANK
' Insertar párrafo de texto Texto_2
With documento.Content
    .InsertAfter Hoja11.Cells(12, 2).Value ' Texto2_ [Paragraphs(13)]
    .insertparagraphafter
End With

With documento.Paragraphs(12).Range
    .Style = "Normal"
End With

' Insertar título de gráfica ACT SEC VAR
With documento.Content
    .InsertAfter Hoja11.Cells(13, 2).Value ' Título de gráfica ACT SEC VAR [Paragraphs(14)]
    .insertparagraphafter
End With

With documento.Paragraphs(13).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica ACT SEC RANK
Grafica_act_sec_rank.Chart.ChartArea.Copy
documento.Paragraphs(14).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter Hoja2.Cells(2, 1).Value ' Nota [Paragraphs(16)]
    .insertparagraphafter
End With

With documento.Paragraphs(15).Range
    .Style = "Fuentes"
End With
' Insertar nota
With documento.Content
    .InsertAfter Hoja11.Cells(14, 2).Value ' Nota [Paragraphs(17)]
    .insertparagraphafter
End With

With documento.Paragraphs(16).Range
    .Style = "Fuentes"
End With
' Insertar salto de página
documento.Paragraphs(17).Range.InsertBreak Type:=7 'wdSectionBreakNextPage
documento.Content.insertparagraphafter
'*********
'DESYTC
' Insertar párrafo de texto Texto_3
With documento.Content
    .InsertAfter Hoja11.Cells(17, 2).Value ' Texto3_ [Paragraphs(20)]
    .insertparagraphafter
End With

With documento.Paragraphs(19).Range
    .Style = "Normal"
End With

' Insertar título de gráfica DESYTC
With documento.Content
    .InsertAfter Hoja11.Cells(18, 2).Value ' Título de gráfica ACT SEC VAR [Paragraphs(21)]
    .insertparagraphafter
End With

With documento.Paragraphs(20).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica DESYTC
Grafica_desytc.Chart.ChartArea.Copy
documento.Paragraphs(21).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter Hoja3.Cells(2, 1).Value ' Nota [Paragraphs(23)]
    .insertparagraphafter
End With

With documento.Paragraphs(22).Range
    .Style = "Fuentes"
End With
' Insertar nota
With documento.Content
    .InsertAfter Hoja11.Cells(19, 2).Value ' Nota [Paragraphs(24)]
    .insertparagraphafter
End With

With documento.Paragraphs(23).Range
    .Style = "Fuentes"
End With
' Insertar salto de página
documento.Paragraphs(24).Range.InsertBreak Type:=7 'wdSectionBreakNextPage
documento.Content.insertparagraphafter
'*********
'DESEST RANK ANU
' Insertar párrafo de texto Texto_4
With documento.Content
    .InsertAfter Hoja11.Cells(22, 2).Value ' Texto4_ [Paragraphs(27)]
    .insertparagraphafter
End With

With documento.Paragraphs(26).Range
    .Style = "Normal"
End With

' Insertar título de gráfica DESEST RANK ANU
With documento.Content
    .InsertAfter Hoja11.Cells(23, 2).Value ' Título de gráfica DESEST RANK ANU [Paragraphs(28)]
    .insertparagraphafter
End With

With documento.Paragraphs(27).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica DESEST RANK ANU
Grafica_desest_rank_anu.Chart.ChartArea.Copy
documento.Paragraphs(28).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter Hoja4.Cells(2, 1).Value ' Nota [Paragraphs(30)]
    .insertparagraphafter
End With

With documento.Paragraphs(29).Range
    .Style = "Fuentes"
End With
' Insertar nota
With documento.Content
    .InsertAfter Hoja11.Cells(24, 2).Value ' Nota [Paragraphs(31)]
    .insertparagraphafter
End With

With documento.Paragraphs(30).Range
    .Style = "Fuentes"
End With
' Insertar salto de página
documento.Paragraphs(31).Range.InsertBreak Type:=7 'wdSectionBreakNextPage
documento.Content.insertparagraphafter
'*********
'DESEST RANK MEN
' Insertar párrafo de texto Texto_5
With documento.Content
    .InsertAfter Hoja11.Cells(27, 2).Value ' Texto5_ [Paragraphs(34)]
    .insertparagraphafter
End With

With documento.Paragraphs(33).Range
    .Style = "Normal"
End With

' Insertar título de gráfica DESEST RANK MEN
With documento.Content
    .InsertAfter Hoja11.Cells(28, 2).Value ' Título de gráfica DESEST RANK MEN [Paragraphs(35)]
    .insertparagraphafter
End With

With documento.Paragraphs(34).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica DESEST RANK MEN
Grafica_desest_rank_men.Chart.ChartArea.Copy
documento.Paragraphs(35).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter Hoja5.Cells(2, 1).Value ' Nota [Paragraphs(37)]
    .insertparagraphafter
End With

With documento.Paragraphs(36).Range
    .Style = "Fuentes"
End With
' Insertar nota
With documento.Content
    .InsertAfter Hoja11.Cells(29, 2).Value ' Nota [Paragraphs(38)]
    .insertparagraphafter
End With

With documento.Paragraphs(37).Range
    .Style = "Fuentes"
End With
' Insertar salto de página
documento.Paragraphs(38).Range.InsertBreak Type:=7 'wdSectionBreakNextPage
documento.Content.insertparagraphafter
'*********
'ACT COM
' Insertar párrafo de texto Texto_6
With documento.Content
    .InsertAfter Hoja11.Cells(32, 2).Value ' Texto6_ [Paragraphs(41)]
    .insertparagraphafter
End With

With documento.Paragraphs(40).Range
    .Style = "Normal"
    .ParagraphFormat.LineSpacing = 15
End With

' Insertar título de gráfica ACT COM
With documento.Content
    .InsertAfter Hoja11.Cells(33, 2).Value ' Título de gráfica ACT COM [Paragraphs(42)]
    .insertparagraphafter
End With

With documento.Paragraphs(41).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica ACT COM
Grafica_act_com.Chart.ChartArea.Copy
documento.Paragraphs(42).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter Hoja6.Cells(2, 1).Value ' Nota [Paragraphs(44)]
    .insertparagraphafter
End With

With documento.Paragraphs(43).Range
    .Style = "Fuentes"
End With
' Insertar nota
With documento.Content
    .InsertAfter Hoja11.Cells(34, 2).Value ' Nota [Paragraphs(45)]
    .insertparagraphafter
End With

With documento.Paragraphs(44).Range
    .Style = "Fuentes"
End With
' Insertar salto de página
documento.Paragraphs(45).Range.InsertBreak Type:=7 'wdSectionBreakNextPage
documento.Content.insertparagraphafter
'*********
'CON VAR
' Insertar párrafo de texto Texto_7
With documento.Content
    .InsertAfter Hoja11.Cells(37, 2).Value ' Texto7_ [Paragraphs(48)]
    .insertparagraphafter
End With

With documento.Paragraphs(47).Range
    .Style = "Normal"
End With

' Insertar título de gráfica CON VAR
With documento.Content
    .InsertAfter Hoja11.Cells(38, 2).Value ' Título de gráfica CON VAR [Paragraphs(49)]
    .insertparagraphafter
End With

With documento.Paragraphs(48).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica CON VAR
grafica_con_var.Chart.ChartArea.Copy
documento.Paragraphs(49).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter Hoja7.Cells(2, 1).Value ' Nota [Paragraphs(51)]
    .insertparagraphafter
End With

With documento.Paragraphs(50).Range
    .Style = "Fuentes"
End With
' Insertar nota
With documento.Content
    .InsertAfter Hoja11.Cells(39, 2).Value ' Nota [Paragraphs(52)]
    .insertparagraphafter
End With

With documento.Paragraphs(51).Range
    .Style = "Fuentes"
End With
' Insertar salto de página
documento.Paragraphs(52).Range.InsertBreak Type:=7 'wdSectionBreakNextPage
documento.Content.insertparagraphafter
'*********
'CON RANK
' Insertar párrafo de texto Texto_8
With documento.Content
    .InsertAfter Hoja11.Cells(42, 2).Value ' Texto8_ [Paragraphs(55)]
    .insertparagraphafter
End With

With documento.Paragraphs(54).Range
    .Style = "Normal"
End With

' Insertar título de gráfica CON RANK
With documento.Content
    .InsertAfter Hoja11.Cells(43, 2).Value ' Título de gráfica CON RANK [Paragraphs(56)]
    .insertparagraphafter
End With

With documento.Paragraphs(55).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica CON RANK
Grafica_con_rank.Chart.ChartArea.Copy
documento.Paragraphs(56).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter Hoja8.Cells(2, 1).Value ' Nota [Paragraphs(58)]
    .insertparagraphafter
End With

With documento.Paragraphs(57).Range
    .Style = "Fuentes"
End With
' Insertar nota
With documento.Content
    .InsertAfter Hoja11.Cells(44, 2).Value ' Nota [Paragraphs(59)]
    .insertparagraphafter
End With

With documento.Paragraphs(58).Range
    .Style = "Fuentes"
End With
' Insertar salto de página
documento.Paragraphs(59).Range.InsertBreak Type:=7 'wdSectionBreakNextPage
documento.Content.insertparagraphafter
'*********
'MAN VAR
' Insertar párrafo de texto Texto_9
With documento.Content
    .InsertAfter Hoja11.Cells(47, 2).Value ' Texto9_ [Paragraphs(62)]
    .insertparagraphafter
End With

With documento.Paragraphs(61).Range
    .Style = "Normal"
End With

' Insertar título de gráfica MAN VAR
With documento.Content
    .InsertAfter Hoja11.Cells(48, 2).Value ' Título de gráfica MAN VAR [Paragraphs(63)]
    .insertparagraphafter
End With

With documento.Paragraphs(62).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica MAN VAR
grafica_man_var.Chart.ChartArea.Copy
documento.Paragraphs(63).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter Hoja9.Cells(2, 1).Value ' Nota [Paragraphs(65)]
    .insertparagraphafter
End With

With documento.Paragraphs(64).Range
    .Style = "Fuentes"
End With
' Insertar nota
With documento.Content
    .InsertAfter Hoja11.Cells(49, 2).Value ' Nota [Paragraphs(66)]
    .insertparagraphafter
End With

With documento.Paragraphs(65).Range
    .Style = "Fuentes"
End With
' Insertar salto de página
documento.Paragraphs(66).Range.InsertBreak Type:=7 'wdSectionBreakNextPage
documento.Content.insertparagraphafter
'*********
'MAN RANK
' Insertar párrafo de texto Texto_10
With documento.Content
    .InsertAfter Hoja11.Cells(52, 2).Value ' Texto10_ [Paragraphs(69)]
    .insertparagraphafter
End With

With documento.Paragraphs(68).Range
    .Style = "Normal"
End With

' Insertar título de gráfica MAN RANK
With documento.Content
    .InsertAfter Hoja11.Cells(53, 2).Value ' Título de gráfica MAN RANK [Paragraphs(70)]
    .insertparagraphafter
End With

With documento.Paragraphs(69).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica MAN RANK
Grafica_man_rank.Chart.ChartArea.Copy
documento.Paragraphs(70).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter Hoja10.Cells(2, 1).Value ' Nota [Paragraphs(72)]
    .insertparagraphafter
End With

With documento.Paragraphs(71).Range
    .Style = "Fuentes"
End With
' Insertar nota
With documento.Content
    .InsertAfter Hoja11.Cells(54, 2).Value ' Nota [Paragraphs(73)]
    .insertparagraphafter
End With

With documento.Paragraphs(72).Range
    .Style = "Fuentes"
End With

' Cambiar la fecha de realización
Set cuadrofecha = documento.Sections(1).Headers(1).Shapes.AddTextbox(msoTextOrientationHorizontal, _
                  340, 35 - 7, 240, 70 / 2)
                  ' wdHeaderFooterPrimary = 1
cuadrofecha.TextFrame.TextRange.Text = "Ficha informativa, " & diahoy & " de " & meshoypal & " de " & añohoy
cuadrofecha.TextFrame.TextRange.Font.Color = RGB(98, 113, 120)
cuadrofecha.TextFrame.TextRange.Font.Underline = wdUnderlineSingle
cuadrofecha.TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphRight
cuadrofecha.Fill.ForeColor = RGB(255, 255, 255)
cuadrofecha.Line.ForeColor = RGB(255, 255, 255)

' Guardar el documento (escribir la dirección en donde se quiera guardar)
documento.SaveAs "C:\Users\arturo.carrillo\Documents\IMAIEF\" & añobas & " " & mesbas & "\Ficha informativa Indicador Mensual de la Actividad Industrial por Entidad (IMAIEF), " & mesbaspal & " " & añobas & "-" & añohoy & meshoy & diahoy & ".docx" ' UBICACIÓN PERSONAL


End Sub
