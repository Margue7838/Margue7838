Sub CrearPresentacionMarketing()
    ' Crear nueva presentación
    Dim pptApp As Object
    Dim pptPres As Object
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    Set pptPres = pptApp.Presentations.Add

    ' Diapositiva de Introducción
    CrearDiapositiva pptPres, "Introducción al Marketing", "Una visión general del concepto de marketing."

    ' Diapositivas con ideas de marketing
    Dim ideasMarketing As Variant
    ideasMarketing = Array( _
        "Investigación de mercado", _
        "Segmentación de clientes", _
        "Diseño de productos/servicios", _
        "Estrategias de publicidad", _
        "Gestión de la marca", _
        "Canales de distribución", _
        "Fijación de precios", _
        "Servicio al cliente" _
    )

    For i = 0 To UBound(ideasMarketing)
        CrearDiapositiva pptPres, "Idea " & (i + 1), ideasMarketing(i)
    Next i

    ' Diapositiva de Conclusión
    CrearDiapositiva pptPres, "Conclusión", "El marketing es esencial para el éxito comercial. ¡Gracias!"

End Sub

Sub CrearDiapositiva(pres As Object, titulo As String, contenido As String)
    ' Añadir una nueva diapositiva
    Dim slideIndex As Integer
    slideIndex = pres.Slides.Count + 1
    pres.Slides.Add slideIndex, ppLayoutText

    ' Establecer el título y el contenido
    pres.Slides(slideIndex).Shapes(1).TextFrame.TextRange.Text = titulo
    pres.Slides(slideIndex).Shapes(2).TextFrame.TextRange.Text = contenido
End Sub
