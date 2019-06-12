Attribute VB_Name = "LibFunciones"
Option Explicit

Public NombreLibroActivo As String
Public HojaActiva As String
Public CeldaActivaDatos As String

Type Celda
    NombreHoja As String
    Address As Range
    Valor As Variant
End Type

'#########################################################################################################
'#################################################       CELDA       ###################################### (funciones sobre celdas)

Sub Celda_PegarEnCelda(ByVal NombreLibroConExtension As String, ByVal NombreHoja As String, ByVal Fila As Long, ByVal Columna As Long, ByVal Dato As Variant)

Workbooks(NombreLibroConExtension).Sheets(NombreHoja).Cells(Fila, Columna) = Dato

End Sub



'#########################################################################################################
'#################################################       TABLA       ###################################### (funciones sobre tablas)

Function Tabla_ObtenerTabla(ByVal NombreHoja As String, ByVal NombreTabla As String, Optional Libro As Workbook) As ListObject

'dada un libro, hoja y nombre tabla, devuelve el objeto tabla

If Libro Is Nothing Then Set Libro = ThisWorkbook

Set Tabla_ObtenerTabla = Libro.Sheets(NombreHoja).ListObjects(NombreTabla)

End Function

Function Tabla_ObtenerColumnaUno(ByVal Tabla As ListObject) As Range

'devuelve la primer columna menos la fila de titulos de la tabla

Set Tabla_ObtenerColumnaUno = Tabla.ListColumns(1).Range

End Function


Function Tabla_ObtenerFilaUno(ByVal Tabla As ListObject) As Range

'dada una tabla devuelve el rango de los titulos

Set Tabla_ObtenerFilaUno = Tabla.HeaderRowRange

End Function

Function Tabla_PosicionIdEnColumnaUno(ByVal Tabla As ListObject, ByVal NombreIdColumnaUno As String) As Long

'Devuelve dentro de una misma tabla, el numero de fila donde se encuentra el ID buscado

Tabla_PosicionIdEnColumnaUno = WorksheetFunction.Match(NombreIdColumnaUno, Tabla_ObtenerColumnaUno(Tabla), 0)

If IsError(Tabla_PosicionIdEnColumnaUno) Then
    MsgBox "No existe el id " & Tabla_PosicionIdEnColumnaUno
End If

End Function

Function Tabla_PosicionTitulo(ByVal Tabla As ListObject, ByVal NombreTituloBuscado As String) As Long

'Devuelve dentro de una misma tabla, el numero de columna donde se encuentra el titulo buscado

Tabla_PosicionTitulo = WorksheetFunction.Match(NombreTituloBuscado, Tabla_ObtenerFilaUno(Tabla), 0)

If IsError(Tabla_PosicionTitulo) Then
    MsgBox "No existe el titulo " & NombreTituloBuscado
End If

End Function

Function Tabla_BuscarV(ByVal Tabla As ListObject, ByVal DatoBuscado As Variant, ByVal TituloColumnaADevolver As String) As Variant

'devuelve el dato de una tabla, dado el numero de columna, la tabla y el valor a buscar

Tabla_BuscarV = Application.VLookup(DatoBuscado, Tabla.Range, Tabla_PosicionTitulo(Tabla, TituloColumnaADevolver), 0)

If IsError(Tabla_BuscarV) Then
    MsgBox "No existe el dato " & DatoBuscado
End If

End Function

Sub Tabla_EscribirDato(ByVal Tabla As ListObject, ByVal IdFilaUno As Variant, ByVal Titulo As Variant, ByVal DatoAEscribir As Variant)

Tabla.Range.Cells(Tabla_PosicionIdEnColumnaUno(Tabla, IdFilaUno), Tabla_PosicionTitulo(Tabla, Titulo)) = DatoAEscribir

End Sub


'#########################################################################################################
'#################################################       HOJA       ###################################### (funciones sobre hojas)

Sub Hoja_AjustarTamañosFilasColumnas(Optional Hoja As Worksheet)

'ajusta tamaños de filas y columnas de la hoja activa

If Hoja Is Nothing Then Set Hoja = ActiveSheet

Hoja.Cells.Select
Hoja.Cells.EntireColumn.AutoFit
Hoja.Cells.EntireRow.AutoFit
    
End Sub

Function Hoja_SiExisteHoja(NombreHoja As String, Optional Libro As Workbook) As Boolean

'Retorna si la hoja existe en un libro dado o por defecto este libro

Dim sht As Worksheet

If Libro Is Nothing Then Set Libro = ThisWorkbook
On Error Resume Next
Set sht = Libro.Sheets(NombreHoja)
On Error GoTo 0

Hoja_SiExisteHoja = Not sht Is Nothing

End Function

Sub Hoja_CrearHoja(ByVal NombreHoja As String, Optional Libro As Workbook)

'crear hoja a la derecha de las existentes en libro pasado como parametro o en su defecto en este libro

If Libro Is Nothing Then Set Libro = ThisWorkbook

Libro.Activate
Libro.Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = NombreHoja

End Sub

Sub Hoja_BorrarHoja(ByVal NombreHoja As String, Optional Libro As Workbook)

'Borra hoja de libro pasado como parametro o en su defecto este libro

If Libro Is Nothing Then Set Libro = ThisWorkbook

Application.DisplayAlerts = False
Libro.Sheets(NombreHoja).Select
ActiveWindow.SelectedSheets.Delete 'borra hoja sin preguntar
Application.DisplayAlerts = True

End Sub

Sub Hoja_ActivarHoja(ByVal NombreHoja As String, Optional Libro As Workbook)

'Activa hoja de libro pasado como parametro o en su defecto este libro

If Libro Is Nothing Then Set Libro = ThisWorkbook

Libro.Sheets(NombreHoja).Activate
    
End Sub

Sub Hoja_CopiarHoja(ByVal NombreLibroOrigen As String, ByVal NombreHojaACopiar As String, ByVal NombreLibroDestino)

'copia hoja de libro origen al final de libro destino

Workbooks(NombreLibroOrigen).Sheets(NombreHojaACopiar).Copy After:=Workbooks(NombreLibroDestino).Sheets(Workbooks(NombreLibroDestino).Worksheets.Count)
    
End Sub

Sub Hoja_GuardarHojaComoPDF(ByVal Ruta As String, ByVal NombrePdf As String, Optional Hoja As Worksheet)

'Guarda hoja como pdf

If Hoja Is Nothing Then Set Hoja = ActiveSheet

Hoja.ExportAsFixedFormat Type:=xlTypePDF, filename:=Ruta & NombrePdf & ".pdf", Quality:= _
xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

End Sub


'########################################################################################################
'##############################################      LIBRO     ##########################################

Sub Libro_GuardarCopiaLibroExcel(ByVal Ruta As String, ByVal NombreLibro As String)

'Guarda copia libro

ActiveWorkbook.SaveAs filename:=Ruta & NombreLibro & ".xlsx", _
FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

End Sub

Sub Libro_CrearLibro(ByVal Ruta As String, ByVal NombreConExtension As String)

'crea libro

Workbooks.Add
With ActiveWorkbook
    .Title = NombreConExtension
    .SaveAs filename:=Ruta & NombreConExtension
End With

End Sub

Sub Libro_BorrarLibro(ByVal Ruta As String, ByVal NombreConExtension As String)

'borra libro

Kill Ruta & NombreConExtension

End Sub

Function Libro_AbrirLibro(ByVal RutayNombreLibroConExtension As String) As Workbook

'abre libro y lo devuelve

Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(RutayNombreLibroConExtension) Then
    Workbooks.Open filename:=RutayNombreLibroConExtension
    Set Libro_AbrirLibro = ActiveWorkbook
Else
    MsgBox ("El archivo " & RutayNombreLibroConExtension & " no Existe!")
    Libro_AbrirLibro = Nothing
    
End If

End Function

Sub Libro_CerrarLibro(ByVal Libro As Workbook, ByVal Guardar As Boolean)

'Cierra Libro

Application.DisplayAlerts = False

Libro.Close savechanges:=Guardar

Application.DisplayAlerts = True

End Sub

Function Libro_LevantarLibro(ByVal TituloBusqueda As String, ByVal RutaBusqueda As String) As Workbook

'Abre libro a partir de una ruta y lo devuelve, si no elije nada devuelve nothing

Dim RutaArchivoSeleccionado

With Application.FileDialog(msoFileDialogFilePicker)
    .Title = TituloBusqueda
    .Filters.Clear
    .Filters.Add "All Files", "*.*"
    '.Filters.Add ".xls*", "*.xlsx*"
    .FilterIndex = 2
    .AllowMultiSelect = False
    .InitialFileName = RutaBusqueda
    '.Show
    If .Show Then
        RutaArchivoSeleccionado = .SelectedItems.Item(1)
        'Workbooks.Open arch    'abrir archivo
    End If
End With

If RutaArchivoSeleccionado <> "" Then
    Set Libro_LevantarLibro = Libro_AbrirLibro(RutaArchivoSeleccionado)
Else
    Set Libro_LevantarLibro = Nothing

End If
    
End Function


'#########################################    CARPETA   ################################
'requires reference to Microsoft Scripting Runtime
Sub Carpeta_CrearSiNoExiste(NombreCarpeta As String, Ruta As String)

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim path As String

'examples for what are the input arguments
'NombreCarpeta = "Folder"
'Ruta = "C:\"

path = Ruta & NombreCarpeta

If Not fso.FolderExists(path) Then

' doesn't exist, so create the folder
          fso.CreateFolder path

End If

End Sub




Function UltimaFilaConDatosDeHojaActiva() As Double

UltimaFilaConDatosDeHojaActiva = WorksheetFunction.Max(Range("A1048576").End(xlUp).Row, Range("b1048576").End(xlUp).Row, Range("c1048576").End(xlUp).Row)

End Function

Sub EliminarFilaHojaActiva(ByVal Fila As Double)

Rows(Fila & ":" & Fila).Select
Selection.Delete Shift:=xlUp
    
End Sub

Sub CargarDiasEnPersona(ByVal Cuil As String, ByVal DiasHabiles As Integer, ByVal Presentes As Integer, ByVal Ausentes As Integer, ByVal NumHoja As Integer)

Call GuardarEstadoPlanillas

Select Case NumHoja
    Case 1
        ColumnaDiasHabiles = "N"
        ColumnaPresentes = "O"
        ColumnaAusentes = "P"
    Case 2
        ColumnaDiasHabiles = "Q"
        ColumnaPresentes = "R"
        ColumnaAusentes = "S"
    Case 3
        ColumnaDiasHabiles = "T"
        ColumnaPresentes = "U"
        ColumnaAusentes = "V"
End Select

Call ActivarHoja(NombreLibroMacro, NombreHojaBase)

Range("C3").Activate
Seguir = True
Do While ActiveCell <> "" And Seguir
    If ActiveCell = Cuil Then
        Range(ColumnaDiasHabiles & ActiveCell.Row) = DiasHabiles
        Range(ColumnaPresentes & ActiveCell.Row) = Presentes
        Range(ColumnaAusentes & ActiveCell.Row) = Ausentes
        Seguir = False
    Else
        ActiveCell.Offset(1, 0).Activate
    End If
Loop

Call VolverAEstadoPlanillas

End Sub

Function PrimerColumnaSinDatos() As String

Call GuardarEstadoPlanillas

Range("A13").Activate
Do While ActiveCell <> ""
    If ActiveCell = "Dias Habiles" Then
        ActiveCell.Offset(0, 1).Activate
        PrimerColumnaSinDatos = Right(Left(ActiveCell.Offset(0, -1).Address, 3), 2)
    End If
    ActiveCell.Offset(0, 1).Activate
Loop

Call VolverAEstadoPlanillas

End Function

Function UltimaColumna() As String

Call GuardarEstadoPlanillas

Range("A13").Activate
Do While ActiveCell <> ""
    If ActiveCell = "Dias Habiles" Then
        UltimaColumna = Right(Left(ActiveCell.Offset(0, -1).Address, 3), 2)
    End If
    ActiveCell.Offset(0, 1).Activate
Loop

Call VolverAEstadoPlanillas

End Function

Sub abreArchivo()
'Por.Dam
Ruta = ThisWorkbook.path
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Seleccione archivo de excel"
        .Filters.Clear
        .Filters.Add "All Files", "*.*"
        '.Filters.Add "xls.*", "*.xls*"
        .FilterIndex = 2
        .AllowMultiSelect = False
        .InitialFileName = Ruta
        '.Show
        If .Show Then
            RutaArchivoSeleccionado = .SelectedItems.Item(1)
            'Workbooks.Open arch    'abrir archivo
        End If
    End With
    
End Sub

Sub PrimerCeldaVisible()

Seguir = True
Range("a2").Activate
Do While ActiveCell <> "" And Seguir
    If ActiveCell.EntireRow.Hidden Then
        ActiveCell.Offset(1, 0).Activate
    Else
        Seguir = False
    End If
Loop

End Sub

Sub MandarMailOutlook(ByVal MailDestino As String, ByVal Asunto As String, ByVal Mensaje As String, ByVal RutaArchivo As String)

'variable para "manejar" el objeto Outlook
Dim OutApp As Object
'variable para "manejar" el objeto mail
Dim OutMail As Object
  
'creamos el objeto Outlook, para acceder a sus
'propiedades, métodos y eventos:
Set OutApp = CreateObject("Outlook.Application")
'logeamos: ojo acá, debemos tener la cuenta bien configurada
OutApp.session.logon
'creo el mail
Set OutMail = OutApp.CreateItem(0)
  
'y acá comienza el "proceso de envío"
On Error Resume Next
With OutMail
    .To = MailDestino  '"raul_taghon@hotmail.com"
    .CC = "" 'si queremos  agregar alguna copia
    .BCC = "" 'si queremos agregar alguna copia oculta
    .Subject = Asunto 'el asunto
    .Body = Mensaje 'cuerpo del mensaje
    If RutaArchivo <> "" Then
        .Attachments.Add RutaArchivo 'adjunto el archivo actual
    End If
    .Send 'y envío el correo
End With
  
'destruyo los objetos para liberar recursos
Set OutMail = Nothing
Set OutApp = Nothing
    
End Sub

Function CantidadDeCeldas(ByVal OffsetFila As Integer, ByVal OffsetColumna As Integer) As Integer

Call GuardarEstadoPlanillas

Call ActivarHoja(NombreLibroMacro, NombreHojaReclamos)

CantidadDeCeldas = 0
Range("a1").Activate
Do While ActiveCell <> ""

    CantidadDeCeldas = CantidadDeCeldas + 1
    ActiveCell.Offset(OffsetFila, OffsetColumna).Activate
Loop

Call VolverAEstadoPlanillas

End Function

Function MaximoId() As Integer

Call GuardarEstadoPlanillas

MaximoId = 0

Range("a2").Activate
Do While ActiveCell <> ""
    If ActiveCell > MaximoId Then
        MaximoId = ActiveCell
    End If
    
    ActiveCell.Offset(1, 0).Activate
Loop

Call VolverAEstadoPlanillas

End Function

Sub GuardarEstadoPlanillas()

NombreLibroActivo = ActiveWorkbook.Name 'guardo el estado de las planillas
HojaActiva = ActiveSheet.Name
CeldaActivaDatos = ActiveCell.Address

End Sub

Sub VolverAEstadoPlanillas()

Call Hoja_ActivarHoja(HojaActiva, ThisWorkbook) 'vuelvo al lugar donde estaba cuando fue llamada la funcion
Range(CeldaActivaDatos).Activate

End Sub

Function SeleccionoObjeto(ByVal Lista As Object, ByVal Objeto As Integer) As Boolean 'si selecciono de una lista multiple determinado objeto

SeleccionoObjeto = False

For i = 0 To Lista.ListCount - 1
    If Lista.Selected(i) And Lista.List(i) = Objeto Then
        SeleccionoObjeto = True
    End If
Next

End Function

Function Selecciono(ByVal Lista As Object) As Boolean 'si selecciono de una lista multiple

Selecciono = False

For i = 0 To Lista.ListCount - 1
    If Lista.Selected(i) Then
        Selecciono = True
    End If
Next

End Function

Sub PrimerLugarVacio(ByVal NombreLibro As String, ByVal NombreHoja As String, ByVal DesdeHaciaAbajo)

Workbooks(NombreLibro).Sheets(NombreHoja).Activate

Range(DesdeHaciaAbajo).Activate

Do While ActiveCell <> ""
    ActiveCell.Offset(1, 0).Activate
Loop
    
End Sub

Sub SeleccionarTodo()
    
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Copy

End Sub

Function IsFileOpen(filename As String)
    Dim filenum As Integer, errnum As Integer

    On Error Resume Next   ' Turn error checking off.
    filenum = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open filename For Input Lock Read As #filenum
    Close filenum          ' Close the file.
    errnum = Err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case errnum

        ' No error occurred.
        ' File is NOT already open by another user.
        Case 0
         IsFileOpen = False

        ' Error number for "Permission Denied."
        ' File is already opened by another user.
        Case 70
            IsFileOpen = True

        ' Another error occurred.
        Case Else
            IsFileOpen = False
    End Select

End Function

'---------------------------------------------------------------------


Private Sub RecorrerDirectorio(ByVal Ruta As String)

'al abrir acumulo

Set acumulado = ActiveWorkbook

carchivo = Dir(Ruta, 0) 'me devuelve el primer archivo del directorio

Do While carchivo > ""          'recorro el directorio
    
    If Not IsFileOpen(Ruta & carchivo) Then 'si esta abierto no lo intento abrir
        Workbooks.Open (Ruta & carchivo)    'abro el primer archivo temporal
        
        Set Temporal = ActiveWorkbook
        
        Range("A2").Select
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        
        acumulado.Activate
        
        Range("A2").Activate    'primer vacio
        Do While ActiveCell <> ""
            ActiveCell.Offset(1, 0).Activate
        Loop
        ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.DisplayAlerts = True
        
        ActiveWorkbook.Save 'guardo el acumulado
        
        Temporal.Activate
        Application.DisplayAlerts = False
        nombretemporal = ActiveWorkbook.Name
        Temporal.Close
        Application.DisplayAlerts = False
        
        Kill Ruta & nombretemporal  'borro archivo temporal
    End If

    carchivo = Dir  'avanza al siguiente archivo
Loop

End Sub

Sub intro()

'trucos aconsejados para acelerar procesos 07/11/2014
Application.CutCopyMode = False
Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = False

Sheets(NombreModeloListado).Visible = True

End Sub

Sub outro()

'trucos aconsejados para acelerar procesos 07/11/2014
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.CutCopyMode = True

Call Hoja_ActivarHoja(NombreHojaMacro, ThisWorkbook)

Sheets(NombreModeloListado).Visible = False

ThisWorkbook.Save

End Sub

Sub CruzRoja()

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)    'Si sale con la X roja
    If CloseMode = vbFormControlMenu Then
        Ejecutar = False
    End If
End Sub

End Sub

Sub CargaDinamicaListbox()

Sheets("Datos").Activate
Range("K2", Range("K" & Rows.Count).End(xlUp)).Name = "Dynamic"
Me.ListBox1.RowSource = "Dynamic"

End Sub

Sub RecorreListBoxyCargaString(ByVal Lista As Object, ByRef Feriados As String)

Feriados = ","

For i = 0 To Lista.ListCount - 1

    If Lista.Selected(i) = True Then
        Feriados = Feriados & Lista.List(i) & ","
    End If

Next

End Sub

Sub ElementoSeleccionadoDeListaMultiple(ByVal Lista As ListBox, ByRef ElementoSeleccionadoDeListaMultiple() As String) 'devuelve en un arreglo de strings los elementos seleccionados de una lista de seleccion multiple

Dim ElementosSeleccionados(1) As String

'Cuento elementos seleccionados
cantSeleccionados = 0

For i = 0 To Lista.ListCount - 1
    If Lista.Selected(i) Then
        cantSeleccionados = cantSeleccionados + 1
    End If
Next

ReDim ElementosSeleccionados(cantSeleccionados) As String

indice = 0
For i = 0 To Lista.ListCount - 1
    If Lista.Selected(i) Then
        ElementosSeleccionados(indice) = Lista.List(i)
        indice = indice + 1
    End If
Next

'copio los resultados
For i = 0 To Lista.ListCount - 1
    ElementoSeleccionadoDeListaMultiple(i) = ElementosSeleccionados(i)
Next

End Sub

Function CantElementosSeleccionadosDeLista(ByVal Lista As ListBox) As Integer

'Cuento elementos seleccionados
cantSeleccionados = 0

For i = 0 To Lista.ListCount - 1
    If Lista.Selected(i) Then
        cantSeleccionados = cantSeleccionados + 1
    End If
Next

CantElementosSeleccionadosDeLista = cantSeleccionados

End Function

Function BuscaColumnaPorTitulo(ByVal Titulo) As String

'Empieza desde la celda (1,1) y va hacia la derecha, una vez encontrado el titulo devuelve solo las letras o sea, la columna sola
Dim Fila, Columna
BuscaColumnaPorTitulo = ""
Fila = 1
Columna = 1
Do While Columna <> 1000
    If Cells(Fila, Columna) = Titulo Then
        BuscaColumnaPorTitulo = Cells(Fila, Columna).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        BuscaColumnaPorTitulo = Left(BuscaColumnaPorTitulo, Len(BuscaColumnaPorTitulo) - 1)
    End If

    Columna = Columna + 1
Loop

If BuscaColumnaPorTitulo = "" Then
    MsgBox "no encuentra titulo columna " & Titulo
End If

End Function

Function ColumnaTitulo(ByVal Titulo As String) As Long

Workbooks(NombreLibroMacro).Sheets(NombreHojaDatos).Range("A2") = Titulo

ColumnaTitulo = CLng(Workbooks(NombreLibroMacro).Sheets(NombreHojaDatos).Range("B2"))

End Function



