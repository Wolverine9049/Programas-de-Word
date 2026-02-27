Attribute VB_Name = "FormatWordDocument"
' ============================================================
' FormatWordDocument.bas
' Macro para reformatear documentos Word convertidos desde PDF
' ============================================================
' INSTRUCCIONES DE USO:
'   1. Abrir el documento Word a formatear
'   2. Abrir el Editor de Visual Basic (Alt+F11)
'   3. Importar este archivo: Archivo > Importar archivo
'   4. Ejecutar la macro "FormatearDocumentoDesdePDF"
' ============================================================

Option Explicit

' -------------------------------------------------------
' Constantes de configuracion corporativa
' -------------------------------------------------------
Private Const MARGIN_TOP_CM      As Double = 2.5
Private Const MARGIN_BOTTOM_CM   As Double = 2.5
Private Const MARGIN_LEFT_CM     As Double = 3.0
Private Const MARGIN_RIGHT_CM    As Double = 2.5
Private Const FONT_BODY          As String = "Calibri"
Private Const FONT_HEADING       As String = "Calibri"
Private Const FONT_SIZE_BODY     As Integer = 11
Private Const FONT_SIZE_H1       As Integer = 16
Private Const FONT_SIZE_H2       As Integer = 14
Private Const FONT_SIZE_H3       As Integer = 12
Private Const COLOR_H1           As Long = RGB(31, 73, 125)   ' Azul corporativo
Private Const COLOR_H2           As Long = RGB(31, 73, 125)
Private Const COLOR_H3           As Long = RGB(31, 73, 125)
Private Const LINE_SPACING_PT    As Double = 14                ' Interlineado (puntos)
Private Const PARA_SPACE_AFTER   As Double = 6                 ' Espacio despues parrafo

' -------------------------------------------------------
' Punto de entrada principal
' -------------------------------------------------------
Public Sub FormatearDocumentoDesdePDF()
    Dim oDocOrig   As Document
    Dim oDocNuevo  As Document
    Dim sRutaNueva As String

    ' Verificar que hay un documento abierto
    If Documents.Count = 0 Then
        MsgBox "No hay ningun documento Word abierto." & vbCrLf & _
               "Abra el documento a formatear y vuelva a ejecutar la macro.", _
               vbExclamation, "FormatWordDocument"
        Exit Sub
    End If

    Set oDocOrig = ActiveDocument

    If oDocOrig.FullName = "" Then
        MsgBox "Guarde el documento antes de ejecutar la macro.", _
               vbExclamation, "FormatWordDocument"
        Exit Sub
    End If

    ' Construir ruta del nuevo archivo
    sRutaNueva = BuscarRutaNueva(oDocOrig.FullName)

    ' Preguntar confirmacion
    If MsgBox("Se creara una copia formateada del documento:" & vbCrLf & vbCrLf & _
              sRutaNueva & vbCrLf & vbCrLf & _
              "El documento original NO sera modificado." & vbCrLf & _
              "Desea continuar?", _
              vbYesNo + vbQuestion, "FormatWordDocument") = vbNo Then
        Exit Sub
    End If

    ' Deshabilitar actualizacion de pantalla para mayor velocidad
    Application.ScreenUpdating = False

    On Error GoTo ManejadorError

    ' Crear copia del documento
    Set oDocNuevo = CrearCopia(oDocOrig, sRutaNueva)

    ' Ejecutar pasos de limpieza y formato sobre la COPIA
    With oDocNuevo
        Call LimpiarArtefactosPDF(oDocNuevo)
        Call AplicarEstilosCorporativos(oDocNuevo)
        Call DetectarAplicarTitulos(oDocNuevo)
        Call EliminarIndiceFalso(oDocNuevo)
        Call InsertarTOCAutomatico(oDocNuevo)
        Call AplicarMargenesCorporativos(oDocNuevo)
        Call AplicarEncabezadoPieDePagina(oDocNuevo)
        Call PrepararParaTraduccion(oDocNuevo)
    End With

    ' Guardar y activar el nuevo documento
    oDocNuevo.Save
    oDocNuevo.Activate

    Application.ScreenUpdating = True

    MsgBox "Documento formateado correctamente:" & vbCrLf & vbCrLf & _
           sRutaNueva & vbCrLf & vbCrLf & _
           "* Estilos de titulo aplicados" & vbCrLf & _
           "* Tabla de Contenido automatica insertada" & vbCrLf & _
           "* Formato corporativo aplicado" & vbCrLf & _
           "* Listo para edicion y traduccion" & vbCrLf & vbCrLf & _
           "Para actualizar el indice: clic derecho sobre el -> Actualizar campo", _
           vbInformation, "FormatWordDocument - Completado"
    Exit Sub

ManejadorError:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "En: " & Err.Source, vbCritical, "FormatWordDocument - Error"
End Sub

' -------------------------------------------------------
' Crea la ruta del nuevo archivo (_formateado.docx)
' -------------------------------------------------------
Private Function BuscarRutaNueva(sRutaOrig As String) As String
    Dim sBase As String
    Dim sExt  As String
    Dim sDir  As String

    sDir  = Left(sRutaOrig, InStrRev(sRutaOrig, Application.PathSeparator))
    sBase = Mid(sRutaOrig, Len(sDir) + 1)

    ' Quitar extension
    If InStrRev(sBase, ".") > 0 Then
        sExt  = Mid(sBase, InStrRev(sBase, "."))
        sBase = Left(sBase, InStrRev(sBase, ".") - 1)
    Else
        sExt = ".docx"
    End If

    ' Evitar sufijos duplicados
    If Right(sBase, 11) = "_formateado" Then
        BuscarRutaNueva = sDir & sBase & ".docx"
    Else
        BuscarRutaNueva = sDir & sBase & "_formateado.docx"
    End If
End Function

' -------------------------------------------------------
' Genera una copia del documento original y la devuelve
' El documento original permanece abierto e intacto.
' -------------------------------------------------------
Private Function CrearCopia(oOrig As Document, sRuta As String) As Document
    ' Copiar el archivo fisicamente para no modificar el original
    FileCopy oOrig.FullName, sRuta
    ' Abrir la copia
    Set CrearCopia = Documents.Open(FileName:=sRuta, AddToRecentFiles:=False)
End Function

' -------------------------------------------------------
' Limpia artefactos tipicos de conversion PDF -> Word
' -------------------------------------------------------
Private Sub LimpiarArtefactosPDF(oDoc As Document)
    Dim i As Integer

    ' --- 1. Reemplazar saltos de linea manuales (Shift+Enter) por espacios
    With oDoc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text        = "^l"          ' Salto de linea manual
        .Replacement.Text = " "
        .Forward     = True
        .Wrap        = wdFindContinue
        .MatchCase   = False
        .Execute Replace:=wdReplaceAll
    End With

    ' --- 2. Eliminar parrafos vacios multiples (mas de uno seguido)
    '        Reemplazar ^p^p^p por ^p^p (iterativo)
    For i = 1 To 5
        With oDoc.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text        = "^p^p^p"
            .Replacement.Text = "^p^p"
            .Forward     = True
            .Wrap        = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    Next i

    ' --- 3. Quitar espacios al inicio de parrafos (usando Find/Replace con comodines)
    With oDoc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = True
        .Text           = "^13 @"            ' parrafo seguido de uno o mas espacios
        .Replacement.Text = "^p"
        .Forward        = True
        .Wrap           = wdFindContinue
        .Execute Replace:=wdReplaceAll
        .MatchWildcards = False
    End With

    ' --- 4. Quitar espacios al final de parrafos
    With oDoc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = True
        .Text           = " @^13"            ' uno o mas espacios antes de fin de parrafo
        .Replacement.Text = "^p"
        .Forward        = True
        .Wrap           = wdFindContinue
        .Execute Replace:=wdReplaceAll
        .MatchWildcards = False
    End With

    ' --- 5. Comprimir espacios multiples internos (iterativo)
    For i = 1 To 5
        With oDoc.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text        = "  "
            .Replacement.Text = " "
            .Forward     = True
            .Wrap        = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    Next i

    ' --- 6. Eliminar tabulaciones excesivas convertidas desde PDF
    With oDoc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text        = "^t^t"
        .Replacement.Text = "^t"
        .Forward     = True
        .Wrap        = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With

    ' --- 7. Quitar tabulaciones iniciales de parrafos de cuerpo
    With oDoc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text        = "^p^t"
        .Replacement.Text = "^p"
        .Forward     = True
        .Wrap        = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
End Sub

' -------------------------------------------------------
' Devuelve True si el rango esta dentro de una celda de tabla
' -------------------------------------------------------
Private Function EstaEnTabla(oRng As Range) As Boolean
    On Error Resume Next
    EstaEnTabla = (oRng.Tables.Count > 0)
    On Error GoTo 0
End Function

' -------------------------------------------------------
' Aplica los estilos corporativos base al documento
' -------------------------------------------------------
Private Sub AplicarEstilosCorporativos(oDoc As Document)
    ' -- Estilo Normal (cuerpo) --
    Call ConfigurarEstilo(oDoc, "Normal", FONT_BODY, FONT_SIZE_BODY, _
                          False, False, wdAlignParagraphJustify, _
                          RGB(0, 0, 0), LINE_SPACING_PT, PARA_SPACE_AFTER)

    ' -- Heading 1 --
    Call ConfigurarEstilo(oDoc, "Heading 1", FONT_HEADING, FONT_SIZE_H1, _
                          True, False, wdAlignParagraphLeft, _
                          COLOR_H1, 0, 12)

    ' -- Heading 2 --
    Call ConfigurarEstilo(oDoc, "Heading 2", FONT_HEADING, FONT_SIZE_H2, _
                          True, False, wdAlignParagraphLeft, _
                          COLOR_H2, 0, 8)

    ' -- Heading 3 --
    Call ConfigurarEstilo(oDoc, "Heading 3", FONT_HEADING, FONT_SIZE_H3, _
                          True, True, wdAlignParagraphLeft, _
                          COLOR_H3, 0, 6)
End Sub

' -------------------------------------------------------
' Configura un estilo de parrafo especifico
' -------------------------------------------------------
Private Sub ConfigurarEstilo(oDoc As Document, sNombre As String, _
                              sFuente As String, nTamano As Integer, _
                              bNegrita As Boolean, bCursiva As Boolean, _
                              nAlineacion As WdParagraphAlignment, _
                              lColor As Long, dEspaciadoLinea As Double, _
                              dEspacioDespues As Double)
    Dim oEstilo As Style
    On Error Resume Next
    Set oEstilo = oDoc.Styles(sNombre)
    On Error GoTo 0

    If oEstilo Is Nothing Then Exit Sub

    With oEstilo.Font
        .Name  = sFuente
        .Size  = nTamano
        .Bold  = bNegrita
        .Italic = bCursiva
        .Color = lColor
    End With

    With oEstilo.ParagraphFormat
        .Alignment = nAlineacion
        If dEspaciadoLinea > 0 Then
            .LineSpacingRule = wdLineSpaceExactly
            .LineSpacing     = dEspaciadoLinea
        End If
        .SpaceAfter  = dEspacioDespues
        .SpaceBefore = 0
    End With
End Sub

' -------------------------------------------------------
' Detecta titulos por heuristicas y aplica estilos Heading
' Heuristicas:
'   H1: parrafo corto (<= 80 car), solo mayusculas O negrita, no termina en punto
'   H2: parrafo corto (<= 80 car), negrita, no termina en punto, inicia con numero
'   H3: parrafo corto (<= 80 car), negrita O cursiva, no termina en punto
' -------------------------------------------------------
Private Sub DetectarAplicarTitulos(oDoc As Document)
    Dim oPar     As Paragraph
    Dim sTexto   As String
    Dim nNivel   As Integer

    For Each oPar In oDoc.Paragraphs
        ' No tocar tablas ni objetos flotantes
        If EstaEnTabla(oPar.Range) Then GoTo SiguienteParrafo

        sTexto = Trim(Left(oPar.Range.Text, Len(oPar.Range.Text) - 1))

        ' Ignorar parrafos vacios
        If Len(sTexto) = 0 Then GoTo SiguienteParrafo

        ' No reclasificar si ya tiene estilo de titulo
        If EsEstiloDeTitulo(oPar.Style.NameLocal) Then GoTo SiguienteParrafo

        nNivel = DetectarNivelTitulo(oPar, sTexto)

        Select Case nNivel
            Case 1
                oPar.Style = oDoc.Styles("Heading 1")
            Case 2
                oPar.Style = oDoc.Styles("Heading 2")
            Case 3
                oPar.Style = oDoc.Styles("Heading 3")
            Case Else
                ' Asegurar que el parrafo tenga estilo Normal si no tiene ninguno conocido
                If Not EsEstiloConocido(oPar.Style.NameLocal) Then
                    oPar.Style = oDoc.Styles("Normal")
                End If
        End Select

SiguienteParrafo:
    Next oPar
End Sub

' -------------------------------------------------------
' Determina el nivel de titulo de un parrafo (0 = no es titulo)
' -------------------------------------------------------
Private Function DetectarNivelTitulo(oPar As Paragraph, sTexto As String) As Integer
    Dim bNegrita   As Boolean
    Dim bCursiva   As Boolean
    Dim bMayusculas As Boolean
    Dim nLongitud  As Integer
    Dim bTerminaPunto As Boolean

    nLongitud = Len(sTexto)

    ' Los titulos suelen ser cortos
    If nLongitud = 0 Or nLongitud > 80 Then
        DetectarNivelTitulo = 0
        Exit Function
    End If

    ' Detectar si termina en punto (los titulos normalmente no terminan en punto)
    bTerminaPunto = (Right(sTexto, 1) = ".")

    ' Detectar formato del primer run del parrafo
    If oPar.Range.Runs.Count > 0 Then
        bNegrita  = (oPar.Range.Runs(1).Bold = True)
        bCursiva  = (oPar.Range.Runs(1).Italic = True)
    End If

    ' Detectar si el texto esta completamente en mayusculas
    bMayusculas = (sTexto = UCase(sTexto)) And (sTexto <> LCase(sTexto))

    ' Reglas para H1: completamente en mayusculas Y corto Y sin punto final
    If bMayusculas And nLongitud <= 80 And Not bTerminaPunto Then
        DetectarNivelTitulo = 1
        Exit Function
    End If

    ' Reglas para H2: negrita Y empieza con numero de seccion (ej: "1.2 Titulo")
    If bNegrita And Not bTerminaPunto And nLongitud <= 80 Then
        If EmpiezaConNumeroSeccion(sTexto) Then
            ' Determinar si es H1 o H2 segun profundidad del numero
            If EsNumeroSeccionProfundo(sTexto) Then
                DetectarNivelTitulo = 2
            Else
                DetectarNivelTitulo = 1
            End If
            Exit Function
        End If
    End If

    ' Reglas para H3: negrita o cursiva, corto, sin punto, sin numero de seccion
    If (bNegrita Or bCursiva) And Not bTerminaPunto And nLongitud <= 80 Then
        DetectarNivelTitulo = 3
        Exit Function
    End If

    DetectarNivelTitulo = 0
End Function

' -------------------------------------------------------
' Devuelve True si el texto empieza con patron de seccion (1., 1.2, A.)
' -------------------------------------------------------
Private Function EmpiezaConNumeroSeccion(sTexto As String) As Boolean
    Dim oRegex As Object
    On Error Resume Next
    Set oRegex = CreateObject("VBScript.RegExp")
    If Err.Number <> 0 Then
        ' Fallback sin regex: primer caracter debe ser digito o letra mayuscula
        ' y estar seguido de un punto en posicion 2 o 3
        If Len(sTexto) >= 2 Then
            Dim c1 As String : c1 = Left(sTexto, 1)
            If (c1 >= "0" And c1 <= "9") Or (c1 >= "A" And c1 <= "Z") Then
                EmpiezaConNumeroSeccion = (Mid(sTexto, 2, 1) = "." Or _
                                           (Len(sTexto) >= 3 And Mid(sTexto, 3, 1) = "."))
            Else
                EmpiezaConNumeroSeccion = False
            End If
        Else
            EmpiezaConNumeroSeccion = False
        End If
        Exit Function
    End If
    On Error GoTo 0

    oRegex.Pattern = "^(\d+\.)+\s|^[A-Z]\.\s"
    oRegex.IgnoreCase = False
    EmpiezaConNumeroSeccion = oRegex.Test(sTexto)
    Set oRegex = Nothing
End Function

' -------------------------------------------------------
' Devuelve True si el numero de seccion tiene sub-nivel (ej: "1.2")
' -------------------------------------------------------
Private Function EsNumeroSeccionProfundo(sTexto As String) As Boolean
    Dim oRegex As Object
    On Error Resume Next
    Set oRegex = CreateObject("VBScript.RegExp")
    If Err.Number <> 0 Then
        ' Fallback sin regex: patron "d.d" requiere al menos 4 caracteres
        If Len(sTexto) >= 4 Then
            EsNumeroSeccionProfundo = (Mid(sTexto, 2, 1) = "." And Mid(sTexto, 4, 1) = ".")
        Else
            EsNumeroSeccionProfundo = False
        End If
        Exit Function
    End If
    On Error GoTo 0

    oRegex.Pattern = "^\d+\.\d+"
    EsNumeroSeccionProfundo = oRegex.Test(sTexto)
    Set oRegex = Nothing
End Function

' -------------------------------------------------------
' Devuelve True si el estilo ya es un estilo de titulo Word
' -------------------------------------------------------
Private Function EsEstiloDeTitulo(sNombreEstilo As String) As Boolean
    Dim sLower As String
    sLower = LCase(sNombreEstilo)
    ' Estilos de titulo en ingles y espanol (titulo / título)
    If sLower Like "heading #" Then EsEstiloDeTitulo = True : Exit Function
    If sLower Like "titulo #"  Then EsEstiloDeTitulo = True : Exit Function
    If sLower Like "t?tulo #"  Then EsEstiloDeTitulo = True : Exit Function
    EsEstiloDeTitulo = False
End Function

' -------------------------------------------------------
' Devuelve True si el estilo es un estilo conocido/nativo de Word
' -------------------------------------------------------
Private Function EsEstiloConocido(sNombreEstilo As String) As Boolean
    Select Case LCase(sNombreEstilo)
        Case "normal", "default paragraph font", "no spacing", _
             "body text", "body text 2", "body text 3", _
             "list paragraph", "caption", "quote", "intense quote", _
             "subtitle", "title"
            EsEstiloConocido = True
        Case Else
            EsEstiloConocido = EsEstiloDeTitulo(sNombreEstilo)
    End Select
End Function

' -------------------------------------------------------
' Elimina el indice falso (texto plano que simula un TOC)
' Busca patrones de "Indice", "Contents", "Table of Contents" etc.
' y elimina todos los parrafos del bloque hasta el primer Heading real
' o hasta que se detecte un bloque diferente de TOC.
' -------------------------------------------------------
Private Sub EliminarIndiceFalso(oDoc As Document)
    Dim oPar      As Paragraph
    Dim bEnIndice As Boolean
    Dim sTexto    As String
    Dim i         As Long
    Dim nInicioIndice As Long
    Dim nFinIndice    As Long

    ' --- Paso 1: Eliminar campos TOC existentes (si los hay) ---
    Dim oField As Field
    For Each oField In oDoc.Fields
        If oField.Type = wdFieldTOC Then
            oField.Delete
        End If
    Next oField

    ' --- Paso 2: Buscar y eliminar texto plano que simula un indice ---
    '     Heuristica: bloque que empieza con "indice", "contenido", "tabla de",
    '     y cada linea tiene el patron "Texto....N" o "Texto  N" (numero de pagina)
    bEnIndice      = False
    nInicioIndice  = 0
    nFinIndice     = 0

    Dim nNumPars As Long
    nNumPars = oDoc.Paragraphs.Count

    For i = 1 To nNumPars
        Set oPar = oDoc.Paragraphs(i)
        sTexto = Trim(Left(oPar.Range.Text, Len(oPar.Range.Text) - 1))

        If Not bEnIndice Then
            ' Detectar inicio de indice falso
            If EsEncabezadoIndice(sTexto) Then
                bEnIndice     = True
                nInicioIndice = i
            End If
        Else
            ' Detectar fin de indice falso
            If Not EsLineaDeIndice(sTexto) And Len(sTexto) > 0 Then
                nFinIndice = i - 1
                bEnIndice  = False
                Call EliminarParrafos(oDoc, nInicioIndice, nFinIndice)
                ' Reiniciar contador despues de eliminar
                Exit For
            End If
        End If
    Next i

    ' Si termino el documento dentro del indice
    If bEnIndice And nFinIndice = 0 Then
        Call EliminarParrafos(oDoc, nInicioIndice, nNumPars)
    End If
End Sub

' -------------------------------------------------------
' Elimina los parrafos en el rango [nInicio, nFin] del documento
' -------------------------------------------------------
Private Sub EliminarParrafos(oDoc As Document, nInicio As Long, nFin As Long)
    Dim oRngElim As Range
    If nInicio > nFin Then Exit Sub
    If nInicio < 1 Then nInicio = 1
    If nFin > oDoc.Paragraphs.Count Then nFin = oDoc.Paragraphs.Count

    Set oRngElim = oDoc.Range( _
        oDoc.Paragraphs(nInicio).Range.Start, _
        oDoc.Paragraphs(nFin).Range.End)
    oRngElim.Delete
End Sub

' -------------------------------------------------------
' Devuelve True si el texto es un encabezado de indice
' -------------------------------------------------------
Private Function EsEncabezadoIndice(sTexto As String) As Boolean
    Dim sLower As String
    sLower = LCase(Trim(sTexto))
    Select Case sLower
        Case "indice", "contenido", "tabla de contenido", _
             "tabla de contenidos", "table of contents", "contents", _
             "index", "indice general"
            EsEncabezadoIndice = True
        Case Else
            ' Cubrir variantes con acento: "índice", "índice general"
            If sLower = Chr(237) & "ndice" Or _
               sLower = Chr(237) & "ndice general" Then
                EsEncabezadoIndice = True
            Else
                EsEncabezadoIndice = False
            End If
    End Select
End Function

' -------------------------------------------------------
' Devuelve True si la linea parece ser una entrada de indice falso
' Patron: texto seguido de puntos o espacios y un numero de pagina
' -------------------------------------------------------
Private Function EsLineaDeIndice(sTexto As String) As Boolean
    If Len(sTexto) = 0 Then
        EsLineaDeIndice = True   ' Lineas en blanco dentro del indice
        Exit Function
    End If

    ' Verificar si termina en numero (posiblemente precedido de puntos o espacios)
    Dim sUltimos As String
    sUltimos = Right(Trim(sTexto), 1)
    If sUltimos >= "0" And sUltimos <= "9" Then
        ' Verificar que antes del numero haya puntos suspensivos o espacios
        Dim j As Integer
        For j = Len(sTexto) To 1 Step -1
            Dim c As String
            c = Mid(sTexto, j, 1)
            If c >= "0" And c <= "9" Then
                ' Numero de pagina
            ElseIf c = "." Or c = " " Or c = Chr(9) Then
                ' Separador
            Else
                ' Texto del titulo
                EsLineaDeIndice = True
                Exit Function
            End If
        Next j
    End If
    EsLineaDeIndice = False
End Function

' -------------------------------------------------------
' Inserta la Tabla de Contenido automatica de Word al inicio del documento,
' justo despues del primer titulo H1 si lo hay, o al inicio absoluto.
' -------------------------------------------------------
Private Sub InsertarTOCAutomatico(oDoc As Document)
    Dim oRng      As Range
    Dim oTOC      As TableOfContents
    Dim nInsertar As Long
    Dim oPar      As Paragraph

    ' Buscar posicion: despues del primer Heading 1
    nInsertar = 0
    For Each oPar In oDoc.Paragraphs
        If EsEstiloDeTitulo(oPar.Style.NameLocal) Then
            If InStr(LCase(oPar.Style.NameLocal), "1") > 0 Then
                nInsertar = oPar.Range.End
                Exit For
            End If
        End If
    Next oPar

    ' Crear rango de insercion
    Set oRng = oDoc.Range(nInsertar, nInsertar)

    ' Insertar dos parrafos (uno de separacion, uno para el TOC)
    oRng.InsertAfter vbCr & vbCr
    oRng.Collapse wdCollapseEnd

    ' Retroceder un parrafo para posicionar el TOC
    oRng.MoveStart wdParagraph, -1
    oRng.Collapse wdCollapseStart

    ' Insertar campo de Tabla de Contenido
    Set oTOC = oDoc.TablesOfContents.Add( _
        Range             := oRng, _
        UseHeadingStyles  := True, _
        UpperHeadingLevel := 1, _
        LowerHeadingLevel := 3, _
        UseHyperlinks     := True, _
        HidePageNumbers   := False, _
        IncludePageNumbers := True)

    oTOC.Update
End Sub

' -------------------------------------------------------
' Aplica los margenes corporativos a todas las secciones
' -------------------------------------------------------
Private Sub AplicarMargenesCorporativos(oDoc As Document)
    Dim oSec As Section
    Dim oPgSetup As PageSetup
    Dim dCmPt As Double
    dCmPt = 28.3464566929134   ' 1 cm = 28.3464... pt

    For Each oSec In oDoc.Sections
        Set oPgSetup = oSec.PageSetup
        With oPgSetup
            .TopMargin    = MARGIN_TOP_CM    * dCmPt
            .BottomMargin = MARGIN_BOTTOM_CM * dCmPt
            .LeftMargin   = MARGIN_LEFT_CM   * dCmPt
            .RightMargin  = MARGIN_RIGHT_CM  * dCmPt
        End With
    Next oSec
End Sub

' -------------------------------------------------------
' Aplica encabezado y pie de pagina corporativos
' -------------------------------------------------------
Private Sub AplicarEncabezadoPieDePagina(oDoc As Document)
    Dim oSec  As Section
    Dim oHdr  As HeaderFooter
    Dim oFtr  As HeaderFooter

    For Each oSec In oDoc.Sections
        ' -- Encabezado --
        Set oHdr = oSec.Headers(wdHeaderFooterPrimary)
        oHdr.LinkToPrevious = False
        With oHdr.Range
            .Text = oDoc.BuiltInDocumentProperties(wdPropertyTitle) & ""
            With .ParagraphFormat
                .Alignment = wdAlignParagraphRight
            End With
            With .Font
                .Name  = FONT_BODY
                .Size  = 9
                .Color = RGB(128, 128, 128)
            End With
        End With

        ' -- Pie de pagina con numero de pagina --
        Set oFtr = oSec.Footers(wdHeaderFooterPrimary)
        oFtr.LinkToPrevious = False
        oFtr.Range.Fields.Add oFtr.Range, wdFieldPage
        With oFtr.Range.ParagraphFormat
            .Alignment = wdAlignParagraphCenter
        End With
        With oFtr.Range.Font
            .Name  = FONT_BODY
            .Size  = 9
            .Color = RGB(128, 128, 128)
        End With
    Next oSec
End Sub

' -------------------------------------------------------
' Preparaciones para traduccion con herramientas CAT (Trados, memoQ, etc.)
'   - Asegurar que el documento use UTF-8 (ya garantizado por .docx)
'   - Asegurar que las tablas esten correctamente delimitadas
'   - Asegurar que los estilos de parrafo sean consistentes
'   - Desactivar la correccion automatica de idioma por parrafo
'   - Marcar el idioma del documento
' -------------------------------------------------------
Private Sub PrepararParaTraduccion(oDoc As Document)
    ' Establecer idioma del documento (espanol de Espana como ejemplo)
    ' Cambie wdSpanishModernSort por el idioma origen de su documento
    On Error Resume Next
    oDoc.Content.LanguageID = wdSpanish
    On Error GoTo 0

    ' Desactivar deteccion automatica de idioma por parrafo
    Dim oPar As Paragraph
    For Each oPar In oDoc.Paragraphs
        On Error Resume Next
        oPar.Range.NoProofing = False
        oPar.Range.DetectLanguage = False
        On Error GoTo 0
    Next oPar

    ' Asegurar que las propiedades del documento esten rellenas
    On Error Resume Next
    With oDoc.BuiltInDocumentProperties
        If Len(.Item(wdPropertyTitle).Value) = 0 Then
            .Item(wdPropertyTitle).Value = oDoc.Name
        End If
        If Len(.Item(wdPropertyAuthor).Value) = 0 Then
            .Item(wdPropertyAuthor).Value = Application.UserName
        End If
    End With
    On Error GoTo 0

    ' Asegurar que el documento este en modo de compatibilidad Word moderno
    oDoc.CompatibilityMode = wdWord2013
End Sub
