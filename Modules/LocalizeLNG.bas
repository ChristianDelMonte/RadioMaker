Attribute VB_Name = "LocalizeLNG"

'////////////////////////////////////////////////////////
'*
'*  ////////// LOCALIZE module for Vb. 6+ ///////////
'*  *** this module is for lenguaje localization  ***
'*  ********* and is for Radiomaker 1.0 only ********
'*
'*    Copyright (c) 1987-2008 Only development Inc.
'*    Christian A. Del Monte
'///////////////////////////////////////////////////////

Private Const LocalizeCFGFile = "Localize.cfg"

Global LNGDef As String      'dimension global de idioma

Private LNGFilename As String
Private DataId As String
Private Data As String

'//// Funcion para extraer la traduccion de un componente predeterminado del programa
'//// extrayendo los mismos del archivo de lenguaje LNG correspondiente
'//// modo de uso:
'//// WLenguageType = nombre del lenguaje en cuestion: ESPAÑOL, INGLES, etc.
'//// CompId = numero de referencia del componente a buscar la traduccion dentro del archivo LNG
Public Function GetComLng_ByID(WLenguajeType As String, CompId As String) As String

LNGFilename = GetLNGFilename(WLenguajeType)
'si no se encuentra lenguaje correspondiente al seleccionado se usa el ESPAÑOL por defecto
If LNGFilename = "." Then
    LNGFilename = App.path & "\" & App.EXEName & ".ES.lng"     'por defecto
Else
    LNGFilename = App.path & "\" & LNGFilename
End If

On Error GoTo err
Open LNGFilename For Input As #11

Do Until EOF(11)
Line Input #11, Data

    DataId = Left$(Data, 4)
    Data = Mid$(Data, 5, Len(Data))
    Data = Trim(Data)
    If Trim(CompId) = Trim(DataId) Then
        GetComLng_ByID = Data
        Close #11
        Exit Function
    End If
    
Loop
Close #11
Exit Function

err:
Close #11
'DisplayMsg "Error en GetComLng_ByID > Module LocalizeLNG", " Function_data: LNG:" & WLenguajeType & " COMPID:" & CompId & " Filename:" & LNGFilename, err.Number, False
End Function

'//// Funcion para extraer el nombre de archivo correspondiente a un lenguaje determinado
'//// determinado previamente en el archivo localize.cfg dentro de la carpeta del programa
'//// modo de uso:
'//// WLenguageType = nombre del lenguaje en cuestion: ESPAÑOL, INGLES, etc.
'//// devuelve el nombre de archivo LNG correspondiente al lenguaje buscado
'//// en caso de no encontrarse el mismo devuelve "." un punto.
Public Function GetLNGFilename(WLenguajeType As String) As String

On Error GoTo err
Open App.path & "\" & App.EXEName & "." & LocalizeCFGFile For Input As #12
Do Until EOF(12)
Line Input #12, Data

    DataId = Left$(Data, 10)
    DataId = UCase(Trim(DataId))
    Data = Mid$(Data, 11, Len(Data))
    Data = Trim(Data)
    If UCase(Trim(WLenguajeType)) = DataId Then
        If UCase(Trim(Data)) = "." Then
            GetLNGFilename = "."
        Else
            GetLNGFilename = Data
        End If
    End If

Loop
Close #12
Exit Function

err:
Close #12
'DisplayMsg "Error en GetLNGFilename > Module LocalizeLNG", " Function_data: LNG:" & WLenguajeType & " CFG:" & LocalizeCFGFile, err.Number, False
End Function
