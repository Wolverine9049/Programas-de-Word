# Programas-de-Word

Macro VBA para dar formato profesional a documentos Word convertidos desde PDF, lista para MindManager y herramientas CAT como Trados.

---

## `FormatWordDocument.bas` — Macro de reformateo profesional

### ¿Qué hace?

Esta macro toma el documento Word **activo** (convertido desde PDF), crea una **copia** y la transforma en un documento Word 100 % nativo y profesional:

| Problema original (PDF → Word)        | Solución aplicada                                    |
|---------------------------------------|------------------------------------------------------|
| Saltos de línea manuales innecesarios | Convertidos a espacios dentro del párrafo            |
| Párrafos vacíos múltiples             | Reducidos a uno solo                                 |
| Espacios y tabulaciones incorrectos   | Normalizados párrafo a párrafo                       |
| Títulos con formato manual            | Detectados automáticamente → Estilos **Heading 1/2/3** |
| Índice falso (texto plano)            | Eliminado y reemplazado por TOC automático de Word   |
| Márgenes y fuentes inconsistentes     | Plantilla corporativa aplicada                       |
| Idioma no marcado                     | Marcado para herramientas CAT (Trados, memoQ)        |

### Lo que **no** modifica

- **Tablas**: se tratan como objetos fijos; el contenido interno no se extrae ni modifica.
- **Imágenes y frames**: permanecen exactamente donde están.
- **Contenido de texto**: no se inventa ni elimina ninguna sección.

---

## Instalación

1. Abra el documento Word a reformatear.
2. Pulse **Alt + F11** para abrir el Editor de Visual Basic.
3. En el menú `Archivo` del editor → **Importar archivo…**
4. Seleccione `FormatWordDocument.bas` y haga clic en **Abrir**.
5. Cierre el editor (Alt + F11 de nuevo).

---

## Uso

1. Asegúrese de que el documento a formatear esté **guardado y activo**.
2. Pulse **Alt + F8**, seleccione `FormatearDocumentoDesdePDF` y haga clic en **Ejecutar**.
3. Confirme la operación en el cuadro de diálogo.
4. La macro creará automáticamente el archivo `<nombre_original>_formateado.docx` en la misma carpeta.
5. Al finalizar, el nuevo documento quedará abierto y listo.

> **Para actualizar el índice:** haga clic derecho sobre la Tabla de Contenidos → **Actualizar campo** → **Actualizar toda la tabla**.

---

## Resultado esperado

✔ Documento 100 % Word nativo (.docx)  
✔ Índice automático actualizable con clic derecho  
✔ Estilos de título aplicados (Heading 1 / 2 / 3)  
✔ Totalmente editable sin romperse  
✔ Apto para traducción con herramientas CAT (Trados, memoQ, Wordfast)  
✔ Nivel profesional, listo para publicación  

---

## Personalización

Las constantes al inicio del módulo permiten ajustar la plantilla corporativa sin tocar la lógica:

```vba
Private Const MARGIN_LEFT_CM   As Double = 3.0      ' Margen izquierdo
Private Const FONT_BODY        As String = "Calibri" ' Fuente de cuerpo
Private Const FONT_SIZE_BODY   As Integer = 11       ' Tamaño de fuente
Private Const COLOR_H1         As Long = RGB(31, 73, 125) ' Color Heading 1
```

Para cambiar el idioma del documento (útil para Trados), modifique la línea en `PrepararParaTraduccion`:

```vba
oDoc.Content.LanguageID = wdSpanish   ' Cambiar por el idioma origen
```

---

## Heurísticas de detección de títulos

La macro identifica títulos automáticamente usando las siguientes reglas:

| Nivel   | Regla de detección                                                    |
|---------|-----------------------------------------------------------------------|
| Heading 1 | Texto ≤ 80 caracteres + **TODO EN MAYÚSCULAS** + sin punto final |
| Heading 1 | Texto ≤ 80 caracteres + **negrita** + número de sección simple (`1.`) |
| Heading 2 | Texto ≤ 80 caracteres + **negrita** + número de subsección (`1.2`) |
| Heading 3 | Texto ≤ 80 caracteres + negrita o cursiva + sin punto final        |

Los párrafos que ya tienen un estilo de título asignado no son modificados.
