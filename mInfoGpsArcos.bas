Attribute VB_Name = "mInfoGpsArcos"
Option Compare Database
Option Explicit

Private Const gbMaxNumber = 1.79769313486231E+308
Private Const gbMaxLongNumber As Long = 2147483647
Private Const maxNumShapes As Long = 900
Private Const maxDifSeg As Long = 60
'
Private nombreArchivoImportado As String
'
Private procesarSoloWaypoints As Boolean
Private juntarTracks As Boolean
Private procesarSoloGeo As Boolean
Private utilizarNodos As Boolean
Private segmentar As Boolean
Private tipoSalida As Integer
'
Private FormatoGrid As String
Private CoordSubfields As Long
Private outputFile As String
Private multipleFiles As Boolean
Private inputFile As String
Private geoId As Long
Private outputFh As Long
'
Private Const fmtGrLatLong_hddd_dddddG As String = "Lat/Lon hddd.ddddd°"
Private Const fmtGrLatLong_hdddG_mm_mmmM As String = "Lat/Long hddd°mm.mmm'"
Private Const fmtGrLatLon_hdddG_mm_mmmM As String = "Lat/Lon hddd°mm.mmm'"

Public gTolerDistAlArco As Double
Public gTolerDistMedidLongArco As Double
Public gVirajeMaximoUnMetro As Integer

Private links As New Collection

Public Sub ProcesarWaypoints()
  Call ProcesarBatch(Nothing, True, False, False, False, False, 0, False, 29, 20, 67, 18, 0)
End Sub

Public Function ProcesarBatch(ByRef frm As Form, ByVal pProcesarSoloWaypoints As Boolean, ByVal pJuntarTracks As Boolean, ByVal pProcesarSoloGeo As Boolean, ByVal pUtilizarNodos As Boolean, ByVal pSegmentar As Boolean, ByVal pTipoSalida As Integer, ByVal pTrackPoint_6_15_11 As Boolean, ByVal pDistMaxBusq As Long, ByVal pPorcMaxDistAlArcoLongArco As Long, ByVal pPorcMaxDistMedidLongArco As Long, ByVal pVirajeMaximoUnMetro As Integer, ByVal pDistMaxSep As Long) As Boolean
  Dim inputFh As Long
  Dim linenumber As Long
  Dim blankLineFg As Boolean
  Dim firstsHeaderFg As Boolean
  Dim firstsHeaderEndedFg As Boolean
  Dim headerFg As Boolean
  Dim trackHeaderReadFg As Boolean
  Dim inTrackFg As Boolean
  Dim trackpointSec As Long
  Dim trackId As Long
  Dim trackName As Variant
  Dim trackPointId As Long
  Dim linea As String
  Dim field As Variant
  Dim firstField As String
  Dim subExitCode As Long
  Dim rec As Long
  Dim db As DAO.Database
  Dim rs As DAO.Recordset
  Dim listaArchivos As New Collection
  Dim archivos As String
  Dim archivo As String
  Dim i As Long
  Dim filtroArchivos As String
  Dim buff As String
  Dim carpetaCbi As String
  Dim outputFileGt As String
  Dim prevMousePointer As Long
  '
  Dim WaypointId As Long
  Dim rsWp As DAO.Recordset
  '
  Dim distanciaMarcaTrackpoint As Long
  '
  distanciaMarcaTrackpoint = pDistMaxBusq
  '
  gTolerDistAlArco = pPorcMaxDistAlArcoLongArco / 100#
  gTolerDistMedidLongArco = pPorcMaxDistMedidLongArco / 100#
  '
  gVirajeMaximoUnMetro = pVirajeMaximoUnMetro
  '
  If gTolerDistAlArco > 2 Or gTolerDistAlArco < 0.01 Then
    MsgBox "Error en el parámetro: Relación máxima entre Distancia al arco y Longitud del arco (porcentaje)"
    Exit Function
  End If
  '
  If gTolerDistMedidLongArco > 0.99 Or gTolerDistMedidLongArco < 0.25 Then
    MsgBox "Error en el parámetro: Relación máxima entre Longitud medida y Longitud del arco (porcentaje)"
    Exit Function
  End If
  '
  prevMousePointer = Screen.MousePointer
  procesarSoloWaypoints = pProcesarSoloWaypoints
  procesarSoloGeo = pProcesarSoloGeo
  juntarTracks = pJuntarTracks
  utilizarNodos = pUtilizarNodos
  segmentar = pSegmentar
  tipoSalida = pTipoSalida
  '
  nombreArchivoImportado = ""
  '
  filtroArchivos = ahtAddFilterItem("", "Archivos GPS (*.txt)", "*.txt")
  inputFile = ahtCommonFileOpenSave(OpenFile:=True, flags:=ahtOFN_ALLOWMULTISELECT, Filter:=filtroArchivos, FileName:=Mid(CurrentDb.Name, 1, InStrRev(CurrentDb.Name, Chr(92))) + "*.txt")
  If Len(inputFile) = 0 Then
    MsgBox "Operation Cancelled - file not loaded", vbCritical, "ProcesarBatch - " & CurrentProject.Name
    ProcesarBatch = False
    Exit Function
  End If
  '
  archivos = inputFile
  If InStr(archivos, vbNullChar) > 0 Then
    archivos = Replace(archivos, vbNullChar, "\", Count:=1)
  End If
  archivos = Replace(archivos, vbNullChar, ", ")
  '
  If procesarSoloWaypoints Then
  
  Else
    If Not procesarSoloGeo Then
      filtroArchivos = ahtAddFilterItem("", "Archivos Csv (*.csv)", "*.csv")
      outputFile = ahtCommonFileOpenSave(OpenFile:=False, Filter:=filtroArchivos, flags:=ahtOFN_OVERWRITEPROMPT Or ahtOFN_READONLY, DefaultExt:="csv", FileName:=Mid(CurrentDb.Name, 1, InStrRev(CurrentDb.Name, Chr(92))) + "*.csv")
      If Len(outputFile) = 0 Then
        MsgBox "Operation Cancelled - file not set", vbCritical, "ProcesarBatch - " & CurrentProject.Name
        ProcesarBatch = False
        Exit Function
      End If
    Else
      outputFile = fullName(nombreCarpeta(archivos), "*.geo,*.csv")
      If MsgBox("¿Desea sobre-escribir los archivos " + outputFile + " ?", vbCritical + vbYesNoCancel) <> vbYes Then
        MsgBox "Operation Cancelled - file not saved", vbCritical, "ProcesarBatch - " & CurrentProject.Name
        ProcesarBatch = False
        Exit Function
      End If
    End If
  End If
  '
  Screen.MousePointer = 11
  '
  If Not frm Is Nothing Then
    frm.SetFocus
    frm.txtCarpetaEntrada.SetFocus
    frm.txtCarpetaEntrada.text = nombreCarpeta(archivos)
    frm.txtArchivoSalida.SetFocus
    frm.txtArchivoSalida.text = outputFile
  End If
  '
  buff = inputFile + vbNullChar
  carpetaCbi = ""
  Set listaArchivos = New Collection
  Do While Len(buff) > 0
    If carpetaCbi = "" Then
      carpetaCbi = StripDelimitedItem(buff, vbNullChar)
      If buff = "" Then
        listaArchivos.Add carpetaCbi
      Else
        carpetaCbi = fullName(carpetaCbi, "")
      End If
    Else
      listaArchivos.Add carpetaCbi + StripDelimitedItem(buff, vbNullChar)
    End If
  Loop
  multipleFiles = (listaArchivos.Count > 1)
  '
  If listaArchivos.Count > 0 Then
    ShowProgressBar "Archivos de entrada...", listaArchivos.Count
  End If
  '
  If Not procesarSoloWaypoints Then
    InicializarInfoSalida
    Rem REVISAR esto requiere MEJORAMIENTO (ene-2011)
    Rem Esto debe ser parametrizable
    Rem InicializarArcosAmbosSentidos
  End If
  '
  If procesarSoloGeo And tipoSalida = 3 Then
    borrarArcosSalida
    outputFile = nombreCarpeta(outputFile)
    outputFile = fullName(outputFile, sinExtens(soloNombreArch(outputFile)))
    outputFh = FreeFile
    outputFileGt = outputFile
    BorrarArchivo outputFileGt + ".geo"
    BorrarArchivo outputFileGt + ".csv"
    Open outputFileGt + ".geo" For Output As #outputFh
    geoId = 0
  End If
  '
  For i = 1 To listaArchivos.Count
    '
    inputFh = 0
    linenumber = 0
    blankLineFg = False
    firstsHeaderFg = False
    firstsHeaderEndedFg = False
    headerFg = False
    trackHeaderReadFg = False
    inTrackFg = False
    trackpointSec = 0
    trackId = 0
    trackName = ""
    trackPointId = 0
    linea = 0
    field = Empty
    firstField = ""
    subExitCode = 0
    rec = 0
    '
    Set db = CurrentDb
    If procesarSoloWaypoints Then
      Set rsWp = db.OpenRecordset("waypoints", dbOpenTable)
    Else
      Set rs = db.OpenRecordset("trackpoints", dbOpenTable)
    End If
    '
    inputFile = listaArchivos(i)
    If procesarSoloGeo Then
      If tipoSalida = 1 Or tipoSalida = 2 Then
        outputFile = sinExtens(inputFile)
        crearCarpeta outputFile
        outputFile = outputFile + "\" + soloNombreArch(outputFile)
      End If
    End If
    '
    nombreArchivoImportado = Mid(inputFile, InStrRev(inputFile, "\") + 1)
    '
    ShowProgressBar "Archivo de entrada " & nombreArchivoImportado & "...", listaArchivos.Count
    UpdateProgressBar i - 1
    '
    If Len(inputFile) = 0 Then
      MsgBox "Operation Cancelled - file not loaded", vbCritical, "Importar() function"
      ProcesarBatch = False
      Exit Function
    End If
    inputFh = FreeFile
    Open inputFile For Input Access Read As #inputFh
    If EOF(inputFh) Then
      MsgBox "El archivo " + inputFile + " no contiene datos", vbCritical, "Importar() function"
      ProcesarBatch = False
      Exit Function
    End If
    If procesarSoloWaypoints Then
      BorrarWaypoints
    Else
      BorrarDatos
    End If
    rec = 0
    Do While Not EOF(inputFh)
      rec = rec + 1
      field = ""
      Line Input #inputFh, linea
      linenumber = linenumber + 1
      blankLineFg = (linea = "")
      If blankLineFg Then
        firstField = ""
        If firstsHeaderFg = False Then
          MsgBox "No se encontraron los primeros encabezados del archivo (Grid, Datum)", vbCritical, "Importar() function"
          ProcesarBatch = False
          Exit Function
        End If
        Corte subExitCode, linenumber, blankLineFg, firstsHeaderFg, firstsHeaderEndedFg, firstField, trackpointSec, inTrackFg, trackHeaderReadFg
        If subExitCode <> 0 Then
          ProcesarBatch = False
          Exit Function
        End If
      Else
        field = getField(linea)
        firstField = field
        If firstsHeaderFg = False Or Not firstsHeaderEndedFg Then
          If firstField = "Grid" Or firstField = "Datum" Then
            If firstField = "Grid" Then
              FormatoGrid = getField(linea)
              Select Case FormatoGrid
                Case fmtGrLatLong_hddd_dddddG
                  CoordSubfields = 1
                Case fmtGrLatLong_hdddG_mm_mmmM, fmtGrLatLon_hdddG_mm_mmmM
                  CoordSubfields = 2
                Case Else
                  MsgBox "Formato de parámetro Grid no reconocido: " + FormatoGrid + ", linea: " + ("" & linenumber), vbCritical, "Importar() function"
                  ProcesarBatch = False
              End Select
            End If
            If Not firstsHeaderFg Then
              firstsHeaderFg = True
            End If
          Else
            MsgBox "No se reconoce el encabezado " + firstField + " como parte de los primeros encabezados, linea " + Trim(CStr(linenumber)), vbCritical, "Importar() function"
            ProcesarBatch = False
            Exit Function
          End If
        Else
          If firstField = "Header" Then
            headerFg = True
            '
            Corte subExitCode, linenumber, blankLineFg, firstsHeaderFg, firstsHeaderEndedFg, firstField, trackpointSec, inTrackFg, trackHeaderReadFg
            If subExitCode <> 0 Then
              ProcesarBatch = False
              Exit Function
            End If
          ElseIf firstField = "Track" Then
            trackHeaderReadFg = True
            trackId = trackId + 1
            field = getField(linea)
            trackName = field
            trackpointSec = 0
            If Not juntarTracks Then
              trackPointId = 0
            End If
            Rem Track subExitCode, linenumber, rsTr, firstField, TrackId, trackName, linea
            '
          ElseIf firstField = "Trackpoint" Then
            If trackHeaderReadFg = False Or trackId = 0 Then
              MsgBox "Trackpoint sin track asociada, linea " + Trim(CStr(linenumber)), vbCritical, "Importar() function"
              ProcesarBatch = False
              Exit Function
            End If
            inTrackFg = True
            trackpointSec = trackpointSec + 1
            trackPointId = trackPointId + 1
            Trackpoint subExitCode, linenumber, rs, firstField, trackId, trackName, trackpointSec, trackPointId, linea, pTrackPoint_6_15_11
            If subExitCode <> 0 Then
              ProcesarBatch = False
              Exit Function
            End If
          ElseIf firstField = "Waypoint" Then
            inTrackFg = False
            WaypointId = WaypointId + 1
            Waypoint subExitCode, linenumber, rsWp, firstField, WaypointId, linea
            If subExitCode <> 0 Then
               ProcesarBatch = False
               Exit Function
            End If
          End If
        End If
      End If
      If rec Mod 100 = 0 Then
        DoEvents
      End If
    Loop
    '
    Close #inputFh
    '
    If Not procesarSoloWaypoints Then
      rs.Close
      Set rs = Nothing
    Else
      rsWp.Close
      Set rsWp = Nothing
    End If
    db.Close
    Set db = Nothing
    '
    ProcesarTracks distanciaMarcaTrackpoint, pDistMaxSep
    '
  Next
  '
  If procesarSoloGeo And outputFh <> 0 And tipoSalida = 3 Then
    Close #outputFh
    outputFh = 0
    outputFileGt = outputFile
    exportarArcosTrackAcsv outputFileGt + ".csv", 0
  End If
  '
  If Not procesarSoloGeo And Not procesarSoloWaypoints Then
    BorrarArchivo outputFile
    DoCmd.TransferText TransferType:=acExportDelim, SpecificationName:="Info_salida Export Specification", TableName:="Info_salida", FileName:=outputFile, HasFieldNames:=True
  End If
  '
  If listaArchivos.Count > 0 Then
    HideProgressBar
  End If
  '
  ProcesarBatch = True
  Screen.MousePointer = prevMousePointer
  '
End Function

Private Sub InicializarArcosAmbosSentidos()
  Dim db As DAO.Database
  Dim sql As String
  '
  sql = "delete from [arcos_ambos_sentidos]"
  Set db = CurrentDb
  db.Execute sql, dbFailOnError
  '
  sql = "insert into [arcos_ambos_sentidos] ([nodo_a], [nodo_b], [longitud])"
  sql = sql + " select [nodo_a], [nodo_b], [longitud] from [arcos_entrada]"
  db.Execute sql, dbFailOnError
  '
  sql = "insert into [arcos_ambos_sentidos] ([nodo_b], [nodo_a], [longitud])"
  sql = sql + " select [f].[nodo_a], [f].[nodo_b], [f].[longitud] from [arcos_entrada] as [f]"
  sql = sql + " where not exists (select [e].* from [arcos_ambos_sentidos] as [e]"
  sql = sql + "   where [e].[nodo_a] = [f].[nodo_b] and [e].[nodo_b] = [f].[nodo_a])"
  db.Execute sql, dbFailOnError
  '
End Sub

Private Sub cargaLinks()
  Dim db As DAO.Database
  Dim rs As DAO.Recordset
  Dim Link As New Link
  Dim sql As String
  '
  Set links = New Collection
  Set db = CurrentDb
  Rem sql = "select [Shapes].* from [Shapes] inner join [arcos_ambos_sentidos] on [Shapes].[nodo_a] = [arcos_ambos_sentidos].[nodo_a] and [Shapes].[nodo_b] = [arcos_ambos_sentidos].[nodo_b] order by [correlativo]"
  sql = "arcos_ambos_sentidos"
  Set rs = db.OpenRecordset(sql, dbOpenTable)
  Do While Not rs.EOF
    Link.IdNodeA = rs.Fields("nodo_a").Value
    Link.IdNodeB = rs.Fields("nodo_b").Value
    Link.loadSegments
    links.Add Link, Key:=Link.Key
    Set Link = New Link
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  
End Sub

Private Sub InicializarInfoSalida()
  Dim db As DAO.Database
  Dim sql As String
  '
  sql = "delete from [info_salida]"
  Set db = CurrentDb
  db.Execute sql, dbFailOnError
  '
End Sub

Private Function ProcesarTracks(ByVal pDistMaxBusq As Long, ByVal pDistMaxSep As Long)
  Dim db As DAO.Database
  Dim rs As DAO.Recordset
  Dim trackpointSec As Long
  Dim lastTrackId As Long
  Dim lastAltitude As Long
  Dim field As String
  Dim lastLegCourseTot As Long
  Dim lastLegCourseSeg As Long
  Dim segmentSec As Long
  Dim sql As String
  Dim angulo As Long
  Dim rec As Long
  Dim progKm As Double
  Dim lastTime As Variant
  '
  Const Minuto = 0.000694444
  '
  Set db = CurrentDb
  sql = "update trackpoints set segmentoID = null, AltitudeMts = null, LengthMts = null"
  sql = sql + ", Segundos = null, KmsHr = null, Angle = null, IndiceGiro = null"
  sql = sql + ", DifAltitudeMts = null, Pendiente = null, IndPendiente = null"
  sql = sql + ", IndPendAbs = null"
  sql = sql + ", longitude = null, latitude = null"
  sql = sql + ", Grilla = null, ProgresivaKm = null"
  db.Execute sql, dbFailOnError
  '
  Set rs = db.OpenRecordset("trackpoints", dbOpenTable)
  rs.Index = "PrimaryKey"
  rs.MoveFirst
  '
  lastTrackId = -1
  rec = 0
  progKm = 0
  lastTime = Null
  Do While Not rs.EOF
    rec = rec + 1
    rs.Edit
    If IsNull(rs.Fields("Altitude").Value) Then
      rs.Fields("AltitudeMts").Value = Null
    Else
      rs.Fields("AltitudeMts").Value = Val(rs.Fields("Altitude").Value)
    End If
    If IsNull(rs.Fields("Leg Length").Value) Then
      rs.Fields("LengthMts").Value = Null
    Else
      rs.Fields("LengthMts").Value = Val(rs.Fields("Leg Length").Value) * IIf(InStr(rs.Fields("Leg Length").Value, "km") > 0, 1000, 1)
    End If
    If IsNull(rs.Fields("Leg Speed").Value) Then
      rs.Fields("KmsHr").Value = Null
    Else
      rs.Fields("KmsHr").Value = Val(rs.Fields("Leg Speed").Value)
    End If
    If IsNull(rs.Fields("Leg Time").Value) Then
      rs.Fields("Segundos").Value = Null
    Else
      rs.Fields("Segundos").Value = Segundos(rs.Fields("Leg Time").Value)
    End If
    If rs.Fields("TrackID").Value <> lastTrackId Then
      trackpointSec = 1
      lastTrackId = rs.Fields("TrackID").Value
      lastLegCourseTot = -361
      angulo = -361
      lastLegCourseSeg = -361
      lastAltitude = -1000000
      segmentSec = 1
      progKm = 0
    Else
      If Not IsNull(rs.Fields("Leg Length").Value) Then
        progKm = progKm + (rs.Fields("LengthMts").Value / 1000)
      End If
      trackpointSec = trackpointSec + 1
      If lastLegCourseTot <> -361 And Not IsNull(rs.Fields("Leg Course").Value) Then
        angulo = diferenciaGrados(lastLegCourseTot, Val(rs.Fields("Leg Course").Value))
      Else
        angulo = -361
      End If
      If lastLegCourseSeg <> -361 Or rs.Fields("LengthMts").Value > 500 Or rs.Fields("Leg Time").Value > 2 * Minuto Then
        If Abs(angulo) >= 135 Then
          segmentSec = segmentSec + 1
          lastLegCourseSeg = -361
        Else
          lastLegCourseSeg = Val(rs.Fields("Leg Course").Value)
        End If
      Else
        If IsNull(rs.Fields("Leg Course").Value) Then
          lastLegCourseSeg = -361
        Else
          lastLegCourseSeg = Val(rs.Fields("Leg Course").Value)
        End If
        If IsNull(rs.Fields("Time").Value) Or ((Not IsNull(lastTime)) And diffSeconds(lastTime, rs.Fields("Time").Value) > maxDifSeg) Then
          segmentSec = segmentSec + 1
        End If
      End If
      If IsNull(rs.Fields("Leg Course").Value) Then
        lastLegCourseTot = -361
      Else
        lastLegCourseTot = Val(rs.Fields("Leg Course").Value)
      End If
    End If
    If angulo = -361 Then
      rs.Fields("Angle").Value = Null
      rs.Fields("IndiceGiro").Value = Null
    Else
      If IsNull(rs.Fields("LengthMts").Value) Then
        rs.Fields("Angle").Value = IIf(Abs(angulo) > gVirajeMaximoUnMetro, gVirajeMaximoUnMetro * Sgn(angulo), angulo)
        rs.Fields("IndiceGiro").Value = Null
      Else
        rs.Fields("Angle").Value = IIf(Abs(angulo) > gVirajeMaximoUnMetro * rs.Fields("LengthMts").Value, gVirajeMaximoUnMetro * Sgn(angulo) * rs.Fields("LengthMts").Value, angulo)
        If rs.Fields("LengthMts").Value = 0 Then
          rs.Fields("IndiceGiro").Value = Null
        Else
          rs.Fields("IndiceGiro").Value = Abs(rs.Fields("Angle").Value) / rs.Fields("LengthMts").Value * 1000
        End If
      End If
    End If
    If lastAltitude = -1000000 Or IsNull(rs.Fields("AltitudeMts").Value) Then
      rs.Fields("DifAltitudeMts").Value = Null
      rs.Fields("AltitudeMtsAvg").Value = Null
    Else
      rs.Fields("DifAltitudeMts").Value = rs.Fields("AltitudeMts").Value - lastAltitude
      rs.Fields("AltitudeMtsAvg").Value = (rs.Fields("AltitudeMts").Value + lastAltitude) / 2
    End If
    If IsNull(rs.Fields("DifAltitudeMts").Value) Or IsNull(rs.Fields("LengthMts").Value) Then
      rs.Fields("Pendiente").Value = Null
    Else
      If rs.Fields("LengthMts").Value <= 0 Then
        rs.Fields("Pendiente").Value = Null
      Else
        rs.Fields("Pendiente").Value = rs.Fields("DifAltitudeMts").Value / rs.Fields("LengthMts").Value
      End If
    End If
    If IsNull(rs.Fields("Pendiente").Value) Then
      rs.Fields("IndPendiente").Value = Null
      rs.Fields("IndPendAbs").Value = Null
    Else
      rs.Fields("IndPendiente").Value = rs.Fields("Pendiente").Value * 100
      rs.Fields("IndPendAbs").Value = Abs(rs.Fields("IndPendiente").Value)
    End If
    If IsNull(rs.Fields("AltitudeMts").Value) Then
      lastAltitude = -1000000
    Else
      lastAltitude = rs.Fields("AltitudeMts").Value
    End If
    rs.Fields("ProgresivaKm").Value = progKm
    rs.Fields("SegmentoID").Value = segmentSec
    field = rs.Fields("Position").Value
    rs.Fields("Latitude").Value = latitude(getCoordField(field))
    rs.Fields("Longitude").Value = longitude(getCoordField(field))
    rs.Update
    lastTime = rs.Fields("Time").Value
    rs.MoveNext
    If rec Mod 100 = 0 Then
      DoEvents
    End If
  Loop
  '
  rs.Close
  Set rs = Nothing
  db.Close
  Set db = Nothing
  '
  If utilizarNodos Then
    BuscarNodosCercanosTps pDistMaxBusq, pDistMaxSep
  End If
  '
  If procesarSoloGeo Then
    generarGeosTracks
  Else
    If Not procesarSoloWaypoints Then
      AgregarArcosInfoSalida pDistMaxSep
    End If
  End If
  '
End Function

Private Sub CorteRecorrido(ByRef rs As DAO.Recordset, ByRef LastCtTrackId As Long, ByRef LastCtTrackpointId As Long, ByVal distPrec As Long, ByVal distMargen As Long, ByVal grColumns As Long, ByVal gSize As Long)
  Dim bm As Variant
  Dim idx As String
  Dim trackId As Long
  Dim trackPointId As Long
  Dim progresivaKM As Double
  Dim cercanos As New Collection
  Dim distProg As Long
  Dim distLineal As Long
  Dim distPuntos As Long
  Dim factor As Double
  Dim i As Long
  Dim minDistPuntos As Double
  Dim retroceder As Long
  Dim f As Long
  Dim maxGiro As Double
  Dim maxGiroBm As Variant
  '
  Const maxFactor As Double = 2#
  '
  bm = rs.Bookmark
  idx = rs.Index
  trackId = rs.Fields("TrackId").Value
  trackPointId = rs.Fields("TrackPointId").Value
  progresivaKM = rs.Fields("ProgresivaKm").Value
  Set cercanos = BuscarCercanos(rs, distPrec, grColumns, gSize, trackId, trackPointId)
  minDistPuntos = gbMaxNumber
  For i = 1 To cercanos.Count
    If LastCtTrackId = -1 Or LastCtTrackId <> trackId Or cercanos(i)(1) > LastCtTrackpointId Then
      distProg = Round((progresivaKM - cercanos(i)(4)) * 1000, 0)
      distLineal = cercanos(i)(2)
      distPuntos = trackPointId - cercanos(i)(1)
      If distLineal = 0 Then
        factor = 999999
      Else
        factor = distProg / distLineal
      End If
      If (distProg - distLineal) > distMargen Then
        If factor > maxFactor Then
          If distPuntos < minDistPuntos Then
            minDistPuntos = distPuntos
          End If
        End If
      End If
    End If
  Next
  If minDistPuntos <> gbMaxNumber Then
    rs.Index = "PrimaryKey"
    rs.Bookmark = bm
    maxGiro = -1
    maxGiro = rs.Fields("IndiceGiro").Value
    maxGiroBm = rs.Bookmark
    For f = 1 To minDistPuntos
      rs.MovePrevious
      If rs.BOF Or rs.EOF Then
        Exit For
      End If
      If rs.Fields("IndiceGiro").Value > maxGiro Then
        maxGiro = rs.Fields("IndiceGiro").Value
        maxGiroBm = rs.Bookmark
      End If
    Next
    '''    retroceder = Round(minDistPuntos / 2, 0)
    '''    rs.Move retroceder * -1
    rs.Bookmark = maxGiroBm
    rs.Edit
    '''    rs.Fields("NuevoArco").Value = True
    rs.Update
    LastCtTrackId = trackId
    LastCtTrackpointId = trackPointId
  End If
  rs.Index = idx
  rs.Bookmark = bm
  '
End Sub

Private Sub AgregarArcosInfoSalida(ByVal pDistMaxSep As Double)
  Dim db As DAO.Database
  Dim rsNodoA As Recordset
  Dim rsNodoB As Recordset
  Dim rsNodo As Recordset
  Dim rsTp As Recordset
  Dim rsTps As Recordset
  Dim sql As String
  Dim rsIs As Recordset
  Dim trackId As Long
  Dim NodoA As Long
  Dim NodoB As Long
  Dim coorA As Variant
  Dim coorB As Variant
  Dim coorMinDist As Variant
  Dim progMinDist As Long
  Dim difDist As Long
  Dim difProg As Long
  Dim distLineal As Long
  Dim coorTpA As Variant
  Dim coorTpB As Variant
  Dim nodosAorig As New Collection
  Dim nodosA As New Collection
  Dim nodosB As New Collection
  Dim idxAelegido As Long
  Dim idxBelegido As Long
  Dim arcosAB As New Collection
  Dim i As Long
  Dim j As Long
  Dim repetido As Boolean
  Dim superpuesto As Boolean
  Dim distLinealAb As Long
  Dim distPerpRel As Double
  Dim distPerpAbs As Double
  Dim Link As New Link
  Dim dsep As Double
  '
  Call cargaLinks
  '
  Set db = CurrentDb()
  '
  Set rsIs = db.OpenRecordset("info_salida")
  '
  Set rsNodo = db.OpenRecordset("nodos")
  rsNodo.Index = "PrimaryKey"
  If Not rsNodo.EOF Then
    rsNodo.MoveFirst
  End If
  '
  Set rsTp = db.OpenRecordset("trackpoints")
  rsTp.Index = "PrimaryKey"
  If Not rsTp.EOF Then
    rsTp.MoveFirst
  End If
  '
  sql = "select distinct [nodos_a].[trackId]"
  sql = sql + ", [arcos].[nodo_a]"
  sql = sql + ", [nodos_a].[trackPointId] as [trackpointId_a]"
  sql = sql + ", [nodos_a].[distancia] as [distancia_nodo_a]"
  sql = sql + ", [arcos].[longitud] as [longitud]"
  sql = sql + " from [arcos_ambos_sentidos] as [arcos]"
  sql = sql + " inner join [nodos_trackpoints] as [nodos_a]"
  sql = sql + " on [arcos].[nodo_a] = [nodos_a].[nodo]"
  sql = sql + " order by [nodos_a].[trackId]"
  sql = sql + ", [arcos].[nodo_a]"
  '
  Set rsNodoA = db.OpenRecordset(sql)
  trackId = -1
  NodoA = -1
  Do While Not rsNodoA.EOF
    trackId = rsNodoA.Fields("trackId").Value
    NodoA = rsNodoA.Fields("nodo_a").Value
    Set nodosAorig = New Collection
    Do While True
      If rsNodoA.EOF() Then
        Exit Do
      End If
      If Not (rsNodoA.Fields("trackId").Value = trackId _
          And rsNodoA.Fields("nodo_a").Value = NodoA) Then
        Exit Do
      End If
      coorTpA = CoordenadaTrackpoint(rsTp, trackId, rsNodoA.Fields("TrackpointId_a").Value)
      nodosAorig.Add Array(rsNodoA.Fields("TrackpointId_a").Value, rsNodoA.Fields("distancia_nodo_a").Value, coorTpA(2), coorTpA, 0, CoordenadaNodo(rsNodo, rsNodoA.Fields("nodo_a").Value))
      rsNodoA.MoveNext
    Loop
    '
    sql = "select [arcos].[nodo_b]"
    sql = sql + ", [nodos_b].[trackPointId] as [trackpointId_b]"
    sql = sql + ", [nodos_b].[distancia] as [distancia_nodo_b]"
    sql = sql + ", [arcos].[longitud]"
    sql = sql + " from [arcos_ambos_sentidos] as [arcos]"
    sql = sql + " inner join [nodos_trackpoints] as [nodos_b]"
    sql = sql + " on [arcos].[nodo_b] = [nodos_b].[nodo]"
    sql = sql + " where [arcos].[nodo_a] = " & NodoA
    sql = sql + " and [nodos_b].[trackId] = " & trackId
    sql = sql + " order by [arcos].[nodo_b]"
    sql = sql + ", [nodos_b].[trackPointId]"
    '
    Rem sql = sql + " and ([nodos_b].[progresiva_km] - x)/[arcos].[longitud] between gTolerDistMedidLongArco and (1 / gTolerDistMedidLongArco)"
    '
    DoEvents
    '
    Set rsNodoB = db.OpenRecordset(sql)
    NodoB = -1
    Do While Not rsNodoB.EOF()
      NodoB = rsNodoB.Fields("nodo_b").Value
      Set nodosA = New Collection
      For i = 1 To nodosAorig.Count
        nodosA.Add nodosAorig(i)
      Next
      Set nodosB = New Collection
      Do While True
        If rsNodoB.EOF() Then
          Exit Do
        End If
        If Not (rsNodoB.Fields("nodo_b").Value = NodoB) Then
          Exit Do
        End If
        coorTpB = CoordenadaTrackpoint(rsTp, trackId, rsNodoB.Fields("TrackpointId_b").Value)
        nodosB.Add Array(rsNodoB.Fields("TrackpointId_b").Value, rsNodoB.Fields("distancia_nodo_b").Value, coorTpB(2), coorTpB, rsNodoB.Fields("longitud").Value, CoordenadaNodo(rsNodo, rsNodoB.Fields("nodo_b").Value))
        rsNodoB.MoveNext
      Loop
      '
      DoEvents
      '
      Set Link = links(NodoA & "-" & NodoB)
      '
      Set arcosAB = New Collection
      '
      Do While nodosA.Count > 0 And nodosB.Count > 0
        '
        idxAelegido = -1
        For i = 1 To nodosA.Count
          If idxAelegido = -1 Then
            dsep = Link.minDistToPoint(nodosA(i)(3)(0), nodosA(i)(3)(1))
            If dsep <= pDistMaxSep Then
              idxAelegido = i
            End If
          Else
            difDist = nodosA(i)(1) - nodosA(idxAelegido)(1)
            difProg = nodosA(i)(2) - nodosA(idxAelegido)(2)
            distLineal = Round(distancia(nodosA(i)(3)(0), nodosA(i)(3)(1), nodosA(idxAelegido)(3)(0), nodosA(idxAelegido)(3)(1)), 0)
            If difDist < (difProg - distLineal) Then
              dsep = Link.minDistToPoint(nodosA(i)(3)(0), nodosA(i)(3)(1))
              If dsep <= pDistMaxSep Then
                idxAelegido = i
              End If
            End If
          End If
        Next
        '
        If idxAelegido = -1 Then
          Exit Do
        End If
        '
        idxBelegido = -1
        For i = 1 To nodosB.Count
          If nodosB(i)(0) > nodosA(idxAelegido)(0) Then
            superpuesto = False
            For j = 1 To arcosAB.Count
              If nodosA(idxAelegido)(0) <= arcosAB(j)(0)(0) And nodosB(i)(0) >= arcosAB(j)(1)(0) Then
                superpuesto = True
                Exit For
              End If
            Next
            If Not superpuesto Then
              distLinealAb = Round(distancia(nodosA(idxAelegido)(3)(0), nodosA(idxAelegido)(3)(1), nodosB(i)(3)(0), nodosB(i)(3)(1)), 0)
              If distLinealAb > 0 Then
                If nodosA(idxAelegido)(1) / distLinealAb <= gTolerDistAlArco And nodosB(i)(1) / distLinealAb <= gTolerDistAlArco Then
                  If (nodosB(i)(2) - nodosA(idxAelegido)(2)) / nodosB(i)(4) >= gTolerDistMedidLongArco And (nodosB(i)(2) - nodosA(idxAelegido)(2)) / nodosB(i)(4) <= (1 / gTolerDistMedidLongArco) Then
                    If idxBelegido = -1 Then
                      dsep = Link.minDistToPoint(nodosB(i)(3)(0), nodosB(i)(3)(1))
                      If dsep <= pDistMaxSep Then
                        idxBelegido = i
                      End If
                    Else
                      difDist = nodosB(i)(1) - nodosB(idxBelegido)(1)
                      difProg = nodosB(i)(2) - nodosB(idxBelegido)(2)
                      distLineal = Round(distancia(nodosB(i)(3)(0), nodosB(i)(3)(1), nodosB(idxBelegido)(3)(0), nodosB(idxBelegido)(3)(1)), 0)
                      If difDist < (difProg - distLineal) * -1 Then
                        dsep = Link.minDistToPoint(nodosB(i)(3)(0), nodosB(i)(3)(1))
                        If dsep <= pDistMaxSep Then
                          idxBelegido = i
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        Next
        '
        If idxBelegido = -1 Then
          nodosA.Remove idxAelegido
        Else
          arcosAB.Add Array(nodosA(idxAelegido), nodosB(idxBelegido))
          nodosA.Remove idxAelegido
          nodosB.Remove idxBelegido
          '
          i = 0
          Do While True
            i = i + 1
            If i > nodosA.Count Then
              Exit Do
            End If
            j = arcosAB.Count
            If nodosA(i)(0) >= arcosAB(j)(0)(0) And nodosA(i)(0) <= arcosAB(j)(1)(0) Then
              nodosA.Remove i
              i = i - 1
            End If
          Loop
          '
          i = 0
          Do While True
            i = i + 1
            If i > nodosB.Count Then
              Exit Do
            End If
            j = arcosAB.Count
            If nodosB(i)(0) >= arcosAB(j)(0)(0) And nodosB(i)(0) <= arcosAB(j)(1)(0) Then
              nodosB.Remove i
              i = i - 1
            End If
          Loop
          '
        End If
        '
      Loop
      '
      For i = 1 To arcosAB.Count
        '
        sql = "SELECT Sum([trackpoints].LengthMts) AS LengthMts"
        sql = sql + ", Sum([trackpoints].Segundos) AS Segundos"
        sql = sql + ", round((Sum([trackpoints].[LengthMts])/1000)"
        sql = sql + " / (Sum([trackpoints].[Segundos])/3600), 0) AS KmsHr"


        sql = sql + ", round(Sum(iif([trackpoints].DifAltitudeMts < 0, [trackpoints].DifAltitudeMts * -1, 0)), 0) AS DifAltMts_Bajada"
        sql = sql + ", round(Sum(iif([trackpoints].DifAltitudeMts > 0, [trackpoints].DifAltitudeMts, 0)), 0) AS DifAltMts_Subida"

        sql = sql + ", round(Sum(iif([trackpoints].DifAltitudeMts < 0, [trackpoints].LengthMts, 0)), 0) AS LongDifAltMts_Bajada"
        sql = sql + ", round(Sum(iif([trackpoints].DifAltitudeMts > 0, [trackpoints].LengthMts, 0)), 0) AS LongDifAltMts_Subida"

        sql = sql + ", Max([trackpoints].Time) AS Hora"


        sql = sql + ", round(Sum(Abs([trackpoints].Angle))/(Sum([trackpoints].[LengthMts])/1000), 0) AS IndiceGiroGrKm"
        sql = sql + ", round(Avg([trackpoints].AltitudeMts),0) AS AltitudeMts"
        sql = sql + ", max([trackpoints].TrackName) AS TrackName"

        sql = sql + " from [trackpoints]"
        sql = sql + " where [trackId] = " & trackId
        sql = sql + " and [trackPointId] between " & (arcosAB(i)(0)(0) + 1)
        sql = sql + " and " & arcosAB(i)(1)(0)
        Set rsTps = db.OpenRecordset(sql)
        If Not (arcosAB(i)(0)(1) / rsTps.Fields("LengthMts").Value > gTolerDistAlArco Or _
            arcosAB(i)(1)(1) / rsTps.Fields("LengthMts").Value > gTolerDistAlArco) Then
            '
            If checkTrackShape(rsTp, trackId, (arcosAB(i)(0)(0) + 1), arcosAB(i)(1)(0), Link, pDistMaxSep) Then
              '
              Call CalcularDistPerpendTpsArco(trackId, arcosAB(i), distPerpRel, distPerpAbs, rsTp)
              '
              rsIs.AddNew
              '
              rsIs.Fields("Archivo").Value = nombreArchivoImportado
              rsIs.Fields("TrackId").Value = trackId
              rsIs.Fields("Nodo_A").Value = NodoA
              rsIs.Fields("Nodo_B").Value = NodoB
              rsIs.Fields("Cod_Pasada").Value = i
              rsIs.Fields("trackpoint_a").Value = arcosAB(i)(0)(0)
              rsIs.Fields("trackpoint_b").Value = arcosAB(i)(1)(0)
              rsIs.Fields("distanodoa").Value = arcosAB(i)(0)(1)
              rsIs.Fields("distanodob").Value = arcosAB(i)(1)(1)
              '
              rsIs.Fields("TrackName").Value = rsTps.Fields("TrackName").Value
              rsIs.Fields("LengthMts").Value = rsTps.Fields("LengthMts").Value
              rsIs.Fields("Segundos").Value = rsTps.Fields("Segundos").Value
              rsIs.Fields("KmsHr").Value = rsTps.Fields("KmsHr").Value
              rsIs.Fields("AltitudeMts").Value = rsTps.Fields("AltitudeMts").Value
              rsIs.Fields("DifAltMts_Bajada").Value = rsTps.Fields("DifAltMts_Bajada").Value
              rsIs.Fields("DifAltMts_Subida").Value = rsTps.Fields("DifAltMts_Subida").Value
              rsIs.Fields("LongDifAltMts_Bajada").Value = rsTps.Fields("LongDifAltMts_Bajada").Value
              rsIs.Fields("LongDifAltMts_Subida").Value = rsTps.Fields("LongDifAltMts_Subida").Value
              rsIs.Fields("Hora").Value = rsTps.Fields("Hora").Value
              rsIs.Fields("IndiceGiroGrKm").Value = rsTps.Fields("IndiceGiroGrKm").Value
              '
              rsIs.Fields("DistPerpRelativa").Value = distPerpRel
              rsIs.Fields("DistPerpAbsoluta").Value = distPerpAbs
              '
              rsIs.Update
              '
          End If
        End If
      Next
    Loop
    rsNodoB.Close
    Set rsNodoB = Nothing
  Loop
  rsNodoA.Close
  Set rsNodoA = Nothing
  db.Close
  Set db = Nothing
  '
End Sub

Private Function checkTrackShape(ByRef rsTp As DAO.Recordset, ByVal trackId As Long, ByVal trackPointAid As Long, ByVal trackPointBid As Long, ByRef Link As Link, ByVal pDistMaxSep As Double) As Boolean
  Dim coor As Variant
  Dim trp As Long
  Dim dsep As Double
  Dim ok As Boolean
  Dim error As Boolean
  
  error = False
  For trp = trackPointAid To trackPointBid
    coor = CoordenadaTrackpoint(rsTp, trackId, trp)
    dsep = Link.minDistToPoint(coor(0), coor(1))
    If dsep <= pDistMaxSep Then
      ok = True
    Else
      error = True
      Exit For
    End If
  Next
  If error Then
    ok = False
  End If
  checkTrackShape = ok
  
End Function

Private Function CoordenadaNodo(ByRef rsNodo As DAO.Recordset, ByVal IdNodo As Long) As Variant
  rsNodo.Seek "=", IdNodo
  CoordenadaNodo = Array(rsNodo.Fields("Longitude").Value, rsNodo.Fields("Latitude").Value)
End Function

Private Function CoordenadaTrackpoint(ByRef rsTp As DAO.Recordset, ByVal trackId As Long, ByVal trackPointId As Long) As Variant
  rsTp.Seek "=", trackId, trackPointId
  CoordenadaTrackpoint = Array(rsTp.Fields("Longitude").Value, rsTp.Fields("Latitude").Value, rsTp.Fields("ProgresivaKm").Value * 1000)
End Function

Private Function Segundos(fecha As Date) As Long
  Segundos = DatePart("s", fecha) + DatePart("n", fecha) * 60 + DatePart("h", fecha) * 3600
End Function

Private Sub BuscarNodosCercanosTps(ByVal pDistMaxBusq As Long, ByVal pDistMaxSep As Long)
  Dim db As DAO.Database
  Dim rs As DAO.Recordset
  Dim rsTp As DAO.Recordset
  Dim x As Long
  Dim y As Long
  Dim posicion As String
  Dim sql As String
  Dim minX As Long
  Dim minY As Long
  Dim maxx As Long
  Dim maxy As Long
  Dim sigx As Long
  Dim sigy As Long
  Const gSize = 5000
  Dim orix As Long
  Dim oriy As Long
  Dim grColumns As Long
  Dim gr As Long
  Dim tpEnc As Boolean
  Dim auxResto As Double
  Dim rec As Long
  Dim LastCtTrackId As Long
  Dim LastCtTrackpointId As Long
  Dim distTp As Long
  Dim rsNodosTp As Recordset
  Dim nodo As Long
  '
  Set db = CurrentDb
  sql = "select min(longitude), min(latitude), max(longitude), min(latitude) from trackpoints"
  Set rs = db.OpenRecordset(sql)
  minX = rs.Fields(0).Value
  minY = rs.Fields(1).Value
  maxx = rs.Fields(2).Value
  maxy = rs.Fields(3).Value
  rs.Close
  Set rs = Nothing
  '
  orix = minX - ModPos(minX, gSize)
  oriy = minY - ModPos(minY, gSize)
  auxResto = ModPos(maxx, gSize)
  maxx = maxx + IIf(auxResto = 0, 0, gSize - auxResto)
  auxResto = ModPos(maxy, gSize)
  maxy = maxy + IIf(auxResto = 0, 0, gSize - auxResto)
  grColumns = (maxx - orix) / gSize
  '
  sql = "update trackpoints"
  sql = sql + " set grilla = fix((latitude - (" & oriy & ")) / " & gSize & ") * " & grColumns
  sql = sql + " + fix((longitude - (" & orix & ")) / " & gSize & ")"
  db.Execute sql, dbFailOnError
  '
  Set rsTp = db.OpenRecordset("trackpoints", dbOpenTable)
  rsTp.Index = "Grilla"
  rsTp.MoveFirst
  '
  Set rsNodosTp = db.OpenRecordset("nodos_trackpoints", dbOpenTable)
  rsNodosTp.Index = "PrimaryKey"
  If Not rsNodosTp.EOF Then
    rsNodosTp.MoveFirst
  End If
  '
  sql = "update [" + "nodos" + "]"
  sql = sql + " set Grilla = null"
  sql = sql + ", TrackpointEncontrado = False"
  db.Execute sql, dbFailOnError
  '
  sql = "delete from [nodos_trackpoints]"
  db.Execute sql, dbFailOnError
  '
  Set rs = db.OpenRecordset("nodos", dbOpenTable)
  rs.Index = "Grilla"
  rs.MoveFirst
  '
  rec = 0
  Do While Not rs.EOF
    rec = rec + 1
    x = rs.Fields("Longitude").Value
    y = rs.Fields("Latitude").Value
    gr = grid(x, y, orix, oriy, gSize, grColumns)
    nodo = rs.Fields("nodo").Value
    rs.Edit
    rs.Fields("Grilla").Value = gr
    If procesarSoloGeo Then
      tpEnc = MarcarTrackpointGeo(x, y, gr, grColumns, gSize, pDistMaxBusq, rsTp, distTp, nodo, rsNodosTp)
    Else
      tpEnc = MarcarTrackpoint(x, y, gr, grColumns, gSize, pDistMaxBusq, rsTp, distTp, nodo, rsNodosTp)
    End If
    If tpEnc Then
      rs.Fields("TrackpointEncontrado").Value = True
      rs.Fields("DistanciaTrackPoint").Value = distTp
    Else
      rs.Fields("TrackpointEncontrado").Value = False
      rs.Fields("DistanciaTrackPoint").Value = Null
    End If
    rs.Update
    If rec Mod 100 = 0 Then
      DoEvents
    End If
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  db.Close
  Set db = Nothing
  '
End Sub

Private Function ModPos(number, divisor) As Double
  Dim resto As Double
  resto = number Mod divisor
  If number < 0 Then
    resto = Abs(divisor) + resto
  End If
  ModPos = resto
End Function

Private Function BuscarCercanos(ByRef rsTp As DAO.Recordset, ByVal maxDistBusq As Long, ByVal columns As Long, ByVal gSize As Long, ByVal trackId As Long, ByVal trackPointId As Long) As Collection
  Dim idx As String
  Dim bm As Variant
  '
  Dim f As Long
  Dim c As Long
  Dim nGr As Long
  Dim minDist As Double
  Dim minBm As Variant
  Dim d As Double
  Dim trackpoints As New Collection
  Dim minBmIdx As Long
  Dim cambio As Boolean
  Dim i As Long
  Dim gr As Long
  Dim x As Long
  Dim y As Long
  Dim TrackIdOrig As Long
  Dim TrackPointIdOrig As Long
  Dim AmplitudGrilla As Long
  '
  If (maxDistBusq * 11) / gSize < 1 Then
    AmplitudGrilla = 1
  Else
    If (maxDistBusq * 11) Mod gSize > 0 Then
      AmplitudGrilla = ((maxDistBusq * 11) \ gSize) + 1
    Else
      AmplitudGrilla = ((maxDistBusq * 11) \ gSize)
    End If
  End If
  '
  bm = rsTp.Bookmark
  idx = rsTp.Index
  rsTp.Index = "Grilla"
  rsTp.Bookmark = bm
  '
  x = rsTp.Fields("Longitude").Value
  y = rsTp.Fields("Latitude").Value
  TrackIdOrig = rsTp.Fields("TrackId").Value
  TrackPointIdOrig = rsTp.Fields("TrackPointId").Value
  gr = rsTp.Fields("Grilla").Value
  '
  minDist = gbMaxNumber
  For f = -AmplitudGrilla To AmplitudGrilla
    For c = -AmplitudGrilla To AmplitudGrilla
      nGr = gr + c + f * columns
      rsTp.Seek "=", nGr
      If Not rsTp.NoMatch Then
        Do While Not rsTp.EOF
          If rsTp.Fields("Grilla").Value <> nGr Then
            Exit Do
          End If
          If rsTp.Fields("TrackId").Value <> trackId Then
            Exit Do
          End If
          If rsTp.Fields("TrackPointId").Value >= trackPointId Then
            Exit Do
          End If
          If rsTp.Fields("TrackId").Value = TrackIdOrig And rsTp.Fields("TrackpointId").Value = TrackPointIdOrig Then
            Exit Do
          End If
          d = Round(distancia(x, y, rsTp.Fields("Longitude").Value, rsTp.Fields("Latitude").Value), 0)
          If d <= maxDistBusq Then
            trackpoints.Add Array(rsTp.Fields("TrackId").Value, rsTp.Fields("TrackPointId").Value, d, rsTp.Bookmark, rsTp.Fields("ProgresivaKm").Value)
            If d < minDist Then
              minDist = d
              minBm = rsTp.Bookmark
              minBmIdx = trackpoints.Count
            End If
          End If
          rsTp.MoveNext
        Loop
      End If
    Next
  Next
  '
  Set BuscarCercanos = trackpoints
  rsTp.Index = idx
  rsTp.Bookmark = bm
  '
End Function

Private Function MarcarTrackpoint(x As Long, y As Long, gr, columns, gSize, maxDistBusq, rsTp As DAO.Recordset, ByRef distTp As Long, ByVal nodo As Long, ByRef rsNodosTp As DAO.Recordset) As Boolean
  Dim f As Long
  Dim c As Long
  Dim nGr As Long
  Dim minDist As Double
  Dim minBm As Variant
  Dim d As Double
  Dim trackpoints As New Collection
  Dim minBmIdx As Long
  Dim cambio As Boolean
  Dim bm As Variant
  Dim i As Long
  Dim pbm As Variant
  Dim AmplitudGrilla As Long
  '
  If (maxDistBusq * 11) / gSize < 1 Then
    AmplitudGrilla = 1
  Else
    If (maxDistBusq * 11) Mod gSize > 0 Then
      AmplitudGrilla = ((maxDistBusq * 11) \ gSize) + 1
    Else
      AmplitudGrilla = ((maxDistBusq * 11) \ gSize)
    End If
  End If
  '
  minDist = gbMaxNumber
  For f = -AmplitudGrilla To AmplitudGrilla
    For c = -AmplitudGrilla To AmplitudGrilla
      nGr = gr + c + f * columns
      rsTp.Seek "=", nGr
      If Not rsTp.NoMatch Then
        Do While Not rsTp.EOF
          If rsTp.Fields("Grilla").Value <> nGr Then
            Exit Do
          End If
          d = Round(distancia(x, y, rsTp.Fields("Longitude").Value, rsTp.Fields("Latitude").Value), 0)
          If d <= maxDistBusq Then
            AgregarNodoTrackpoint rsNodosTp, rsTp.Fields("TrackId").Value, rsTp.Fields("TrackPointId").Value, nodo, d
            If d < minDist Then
              minDist = d
              minBm = rsTp.Bookmark
              minBmIdx = trackpoints.Count
            End If
          End If
          rsTp.MoveNext
        Loop
      End If
    Next
  Next
  '
  If minDist = gbMaxNumber Then
    MarcarTrackpoint = False
  Else
    rsTp.Index = "Grilla"
    MarcarTrackpoint = True
  End If
  '
End Function

Private Function MarcarTrackpointGeo(x As Long, y As Long, gr, columns, gSize, maxDistBusq, rsTp As DAO.Recordset, ByRef distTp As Long, ByVal nodo As Long, ByRef rsNodosTp As DAO.Recordset) As Boolean
  Dim f As Long
  Dim c As Long
  Dim nGr As Long
  Dim minDist As Double
  Dim minBm As Variant
  Dim d As Double
  Dim trackpoints As New Collection
  Dim minBmIdx As Long
  Dim cambio As Boolean
  Dim bm As Variant
  Dim i As Long
  Dim pbm As Variant
  Dim AmplitudGrilla As Long
  '
  If (maxDistBusq * 11) / gSize < 1 Then
    AmplitudGrilla = 1
  Else
    If (maxDistBusq * 11) Mod gSize > 0 Then
      AmplitudGrilla = ((maxDistBusq * 11) \ gSize) + 1
    Else
      AmplitudGrilla = ((maxDistBusq * 11) \ gSize)
    End If
  End If
  '
  minDist = gbMaxNumber
  For f = -AmplitudGrilla To AmplitudGrilla
    For c = -AmplitudGrilla To AmplitudGrilla
      nGr = gr + c + f * columns
      rsTp.Seek "=", nGr
      If Not rsTp.NoMatch Then
        Do While Not rsTp.EOF
          If rsTp.Fields("Grilla").Value <> nGr Then
            Exit Do
          End If
          d = Round(distancia(x, y, rsTp.Fields("Longitude").Value, rsTp.Fields("Latitude").Value), 0)
          If d <= maxDistBusq Then
            trackpoints.Add Array(rsTp.Fields("TrackId").Value, rsTp.Fields("TrackPointId").Value, d, rsTp.Bookmark)
            If d < minDist Then
              minDist = d
              minBm = rsTp.Bookmark
              minBmIdx = trackpoints.Count
            End If
          End If
          rsTp.MoveNext
        Loop
      End If
    Next
  Next
  '
  If minDist = gbMaxNumber Then
    MarcarTrackpointGeo = False
  Else
    distTp = minDist
    rsTp.Index = "PrimaryKey"
    Do While minDist <> gbMaxNumber
      rsTp.Bookmark = minBm
'''      Debug.Print rsTp.Fields("TrackId").Value
'''      Debug.Print rsTp.Fields("TrackPointId").Value
      trackpoints.Remove minBmIdx
      rsTp.MovePrevious
      Do While (Not rsTp.EOF) And (Not rsTp.BOF)
        cambio = False
        For i = 1 To trackpoints.Count
          If rsTp.Fields("TrackId").Value = trackpoints(i)(0) And rsTp.Fields("TrackPointId").Value = trackpoints(i)(1) Then
            cambio = True
            trackpoints.Remove i
            Exit For
          End If
        Next
        If Not cambio Then
          Exit Do
        End If
        rsTp.MovePrevious
      Loop
      rsTp.Bookmark = minBm
      rsTp.MoveNext
      Do While (Not rsTp.EOF) And (Not rsTp.BOF)
        cambio = False
        For i = 1 To trackpoints.Count
          If rsTp.Fields("TrackId").Value = trackpoints(i)(0) And rsTp.Fields("TrackPointId").Value = trackpoints(i)(1) Then
            cambio = True
            trackpoints.Remove i
            Exit For
          End If
        Next
        If Not cambio Then
          Exit Do
        End If
        rsTp.MoveNext
      Loop
      rsTp.Bookmark = minBm
      rsTp.Edit
      rsTp.Fields("NuevoArco").Value = True
      rsTp.Update
      '
      minDist = gbMaxNumber
      For i = 1 To trackpoints.Count
        d = trackpoints(i)(2)
        If d < minDist Then
          minDist = d
          minBm = trackpoints(i)(3)
          minBmIdx = i
        End If
      Next
      '
    Loop
    rsTp.Index = "Grilla"
    MarcarTrackpointGeo = True
  End If
  '
End Function

Private Function AgregarNodoTrackpoint(rsNodosTp As Recordset, trackId, trackPointId, nodo, d)
  rsNodosTp.Seek "=", trackId, trackPointId, nodo
  If rsNodosTp.NoMatch Then
    '
    rsNodosTp.AddNew
    rsNodosTp.Fields("TrackId").Value = trackId
    rsNodosTp.Fields("TrackPointId").Value = trackPointId
    rsNodosTp.Fields("Nodo").Value = nodo
    rsNodosTp.Fields("Distancia").Value = d
    rsNodosTp.Update
  End If
  '
End Function

'''Private Function distancia(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Double
'''  distancia = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
'''End Function

Private Function grid(x, y, orix, oriy, gxSize, columns) As Long
  grid = Fix((y - oriy) / gxSize) * columns + Fix((x - orix) / gxSize)
End Function

Private Function diferenciaGrados(ByVal desde As Long, ByVal hasta As Long) As Long
  If hasta >= desde Then
    If (hasta - desde) > 180 Then
      diferenciaGrados = desde + 360 - hasta
    Else
      diferenciaGrados = hasta - desde
    End If
  Else
    If (desde - hasta) >= 180 Then
      diferenciaGrados = -1 * (hasta + 360 - desde)
    Else
      diferenciaGrados = -1 * (desde - hasta)
    End If
  End If
End Function

Private Function generarGeosTracks()
  Dim db As DAO.Database
  Dim rs As DAO.Recordset
  Dim field As String
  Dim filtroArchivos As String
  Dim outputFileGt As String
  Dim fullFileName As String
  Dim csvPendiente As Boolean
  '
  Dim lastTrackId As Long
  Dim lastSegmentId As Long
  Dim NuevoArco As Boolean
  '
  Dim linea As String
  Dim ar As Arco
  Dim sec As Long
  Dim rsArcos As Recordset
  Dim sql As String
  Dim rec As Long
  Dim TrackFileSeq As Long
  '
  Set db = CurrentDb
  Set rs = db.OpenRecordset("trackpoints", dbOpenTable)
  rs.Index = "PrimaryKey"
  rs.MoveFirst
  Set rsArcos = db.OpenRecordset("arcos_salida", dbOpenTable)
  lastTrackId = -1
  '
  If tipoSalida = 1 Then
    borrarArcosSalida
    TrackFileSeq = 0
    geoId = 0
  ElseIf tipoSalida = 2 Then
    borrarArcosSalida
    outputFh = FreeFile
    outputFileGt = outputFile
    BorrarArchivo outputFileGt + ".geo"
    BorrarArchivo outputFileGt + ".csv"
    Open outputFileGt + ".geo" For Output As #outputFh
    geoId = 0
  End If
  '
  Set ar = Nothing
  sec = 0
  lastSegmentId = -1
  NuevoArco = False
  rec = 0
  Do While Not rs.EOF
    rec = rec + 1
    If rs.Fields("TrackID").Value <> lastTrackId Then
      If tipoSalida = 1 Then
        geoId = 0
      End If
      GuardarArco ar, outputFh, db, rs, sec, geoId, rsArcos, NuevoArco
      If tipoSalida = 1 Then
        outputFileGt = outputFile + "_track" + Trim(CStr(TrackFileSeq))
        If outputFh <> 0 Then
          Close #outputFh
          outputFh = 0
          exportarArcosTrackAcsv outputFileGt + ".csv", lastTrackId
        End If
        TrackFileSeq = TrackFileSeq + 1
        outputFileGt = outputFile + "_track" + Trim(CStr(TrackFileSeq))
        BorrarArchivo outputFileGt + ".geo"
        BorrarArchivo outputFileGt + ".csv"
        outputFh = FreeFile
        Open outputFileGt + ".geo" For Output As #outputFh
      End If
    Else
      If segmentar And rs.Fields("SegmentoID").Value <> lastSegmentId Then
        GuardarArco ar, outputFh, db, rs, sec, geoId, rsArcos, NuevoArco
      Else
        If NuevoArco Or (utilizarNodos = False) Then
          GuardarArco ar, outputFh, db, rs, sec, geoId, rsArcos, NuevoArco
        End If
      End If
    End If
    If sec >= maxNumShapes Then
      GuardarArco ar, outputFh, db, rs, sec, geoId, rsArcos, NuevoArco
    End If
    If ar Is Nothing Then
      Set ar = New Arco
      geoId = geoId + 1
      sec = 0
      ar.ID = geoId
      ar.trackId = rs.Fields("TrackID").Value
      ar.trackName = rs.Fields("TrackName").Value
      ar.HoraInicial = rs.Fields("Time").Value
      ar.SegmentID = rs.Fields("SegmentoID").Value
      ar.NroPuntos = 0
      '
    End If
    sec = sec + 1
    If sec > 1 Then
      If IsNull(rs.Fields("Segundos").Value) Then
        If Not (IsNull(ar.HoraFinal) Or IsNull(rs.Fields("Time").Value)) Then
          If Not diffSeconds(ar.HoraFinal, rs.Fields("Time").Value) >= gbMaxLongNumber Then
            ar.Segundos = ar.Segundos + diffSeconds(ar.HoraFinal, rs.Fields("Time").Value)
          End If
        End If
      Else
        ar.Segundos = ar.Segundos + rs.Fields("Segundos").Value
      End If
      If IsNull(rs.Fields("LengthMts").Value) Then
        ar.LengthMts = Round(distancia(ar.Coordenadas.item(ar.Coordenadas.Count)(0), ar.Coordenadas.item(ar.Coordenadas.Count)(1), rs.Fields("Longitude").Value, rs.Fields("Latitude").Value), 0)
      Else
        ar.LengthMts = ar.LengthMts + rs.Fields("LengthMts").Value
      End If
      '
      If Not IsNull(rs.Fields("AltitudeMts").Value) Then
        ar.DistPesoAltitudeMts = ar.DistPesoAltitudeMts + ar.LengthMts
        ar.AltitudeMts = ar.AltitudeMts + rs.Fields("AltitudeMts").Value * ar.LengthMts
      End If
      If Not IsNull(rs.Fields("IndPendAbs").Value) Then
        ar.DistPesoIndPendAbs = ar.DistPesoIndPendAbs + rs.Fields("LengthMts").Value
        ar.IndPendAbs = ar.IndPendAbs + rs.Fields("IndPendAbs").Value * rs.Fields("LengthMts").Value
      End If
      If Not IsNull(rs.Fields("IndiceGiro").Value) Then
        ar.DistPesoIndiceGiro = ar.DistPesoIndiceGiro + rs.Fields("LengthMts").Value
        ar.IndiceGiro = ar.IndiceGiro + rs.Fields("IndiceGiro").Value * rs.Fields("LengthMts").Value
      End If
      '
    End If
    ar.HoraFinal = rs.Fields("Time").Value
    ar.NroPuntos = ar.NroPuntos + 1
    ar.Coordenadas.Add Array(rs.Fields("Longitude").Value, rs.Fields("Latitude").Value)
    lastTrackId = rs.Fields("TrackID").Value
    lastSegmentId = rs.Fields("SegmentoID").Value
    NuevoArco = rs.Fields("NuevoArco").Value
    If rec Mod 100 = 0 Then
      DoEvents
    End If
    rs.MoveNext
  Loop
  '
  If outputFh <> 0 Then
    GuardarArco ar, outputFh, db, rs, sec, geoId, rsArcos, NuevoArco
    If tipoSalida < 3 Then
      Close #outputFh
      outputFh = 0
    End If
    If tipoSalida = 1 Then
      outputFileGt = outputFile + "_track" + Trim(CStr(TrackFileSeq))
      exportarArcosTrackAcsv outputFileGt + ".csv", lastTrackId
    ElseIf tipoSalida = 2 Then
      outputFileGt = outputFile
      exportarArcosTrackAcsv outputFileGt + ".csv", 0
    End If
  End If
  '
  rs.Close
  Set rs = Nothing
  db.Close
  Set db = Nothing
  '
End Function

Private Sub GuardarArco(ByRef ar As Arco, ByVal outputFh As Long, ByRef db As DAO.Database, ByRef rs As DAO.Recordset, ByRef sec As Long, ByRef geoId As Long, ByRef rsArcos As DAO.Recordset, ByRef NuevoArco As Boolean)
  Dim i As Long
  Dim linea As String
  Dim prevCoordenada As Variant
  Dim prevTrack As Long
  '
  If Not ar Is Nothing Then
    If ar.NroPuntos >= 2 Then
      linea = ar.ID & "," & "1" & "," & ar.NroPuntos
      For i = 1 To ar.Coordenadas.Count
        linea = linea & "," & ar.Coordenadas.item(i)(0)
        linea = linea & "," & ar.Coordenadas.item(i)(1)
      Next
      Print #outputFh, linea
      '
      rsArcos.AddNew
      rsArcos.Fields("ID").Value = ar.ID
      rsArcos.Fields("Shapepoints").Value = ar.NroPuntos
      rsArcos.Fields("TrackID").Value = ar.trackId
      rsArcos.Fields("Archivo").Value = soloNombreArch(inputFile)
      rsArcos.Fields("TrackName").Value = ar.trackName
      rsArcos.Fields("SegmentoID").Value = ar.SegmentID
      rsArcos.Fields("LengthMts").Value = ar.LengthMts
      rsArcos.Fields("Segundos").Value = ar.Segundos
      rsArcos.Fields("KmsHr").Value = ar.KmsHr
      rsArcos.Fields("Hora").Value = ar.Hora
      If ar.DistPesoAltitudeMts <> 0 Then
        rsArcos.Fields("AltitudeMts").Value = ar.AltitudeMts / ar.DistPesoAltitudeMts
      End If
      If ar.DistPesoIndPendAbs <> 0 Then
        rsArcos.Fields("IndPendAbs").Value = ar.IndPendAbs / ar.DistPesoIndPendAbs
      End If
      If ar.DistPesoIndiceGiro <> 0 Then
        If ar.IndiceGiro / ar.DistPesoIndiceGiro > 2 ^ 15 - 1 Then
          rsArcos.Fields("IndiceGiro").Value = 2 ^ 15 - 1
        Else
          rsArcos.Fields("IndiceGiro").Value = ar.IndiceGiro / ar.DistPesoIndiceGiro
        End If
      End If
      rsArcos.Update
      '
    Else
      geoId = geoId - 1
    End If
    '
    prevTrack = ar.trackId
    prevCoordenada = ar.Coordenadas(ar.Coordenadas.Count)
    Set ar = Nothing
    '
    If Not rs.EOF Then
      Set ar = New Arco
      geoId = geoId + 1
      sec = 0
      ar.ID = geoId
      ar.trackId = rs.Fields("TrackID").Value
      ar.trackName = rs.Fields("TrackName").Value
      ar.SegmentID = rs.Fields("SegmentoID").Value
      ar.HoraInicial = rs.Fields("Time").Value
      ar.NroPuntos = 0
      '
      If prevTrack = rs.Fields("TrackId").Value Then
        sec = sec + 1
        ar.NroPuntos = ar.NroPuntos + 1
        ar.Coordenadas.Add prevCoordenada
      End If
    End If
    '
  End If
  NuevoArco = False
End Sub

Private Sub track(ByRef subExitCode As Long, ByRef linenumber As Long, ByRef rsTr As DAO.Recordset, ByRef firstField As String, ByRef trackId As Long, ByRef trackName As Variant, ByRef linea As String)
  subExitCode = 0
  rsTr.AddNew
  rsTr.Fields("ID").Value = trackId
  rsTr.Fields("Name").Value = trackName
  rsTr.Fields("Header").Value = firstField
  rsTr.Fields("Start Time").Value = getField(linea)
  rsTr.Fields("Elapsed Time").Value = getField(linea)
  rsTr.Fields("Length").Value = getField(linea)
  rsTr.Fields("Average Speed").Value = getField(linea)
  rsTr.Fields("Link").Value = getField(linea)
  rsTr.Update
End Sub

Private Sub Trackpoint(ByRef subExitCode As Long, ByRef linenumber As Long, ByRef rs As DAO.Recordset, ByRef firstField As String, ByRef trackId As Long, ByRef trackName As Variant, ByRef trackpointSec As Long, ByRef trackPointId As Long, ByRef linea As String, ByVal trackPoint_6_15_11 As Boolean)
  Dim auxTime As String
  subExitCode = 0
  If Not procesarSoloWaypoints Then
    rs.AddNew
    If juntarTracks Then
      rs.Fields("TrackID").Value = 9999
    Else
      rs.Fields("TrackID").Value = trackId
    End If
    rs.Fields("TrackpointID").Value = trackPointId
    rs.Fields("TrackName").Value = trackName
    rs.Fields("Header").Value = firstField
    rs.Fields("Position").Value = getField(linea)
    auxTime = getField(linea, True)
    If InStr(auxTime, "(") > 0 Then
      auxTime = Mid(auxTime, 1, InStr(auxTime, "(") - 1)
    End If
    rs.Fields("Time").Value = auxTime
    rs.Fields("Altitude").Value = getField(linea)
    rs.Fields("Depth").Value = getField(linea)
    If trackPoint_6_15_11 Then
      Call getField(linea)
    End If
    rs.Fields("Leg Length").Value = getField(linea)
    rs.Fields("Leg Time").Value = getField(linea)
    rs.Fields("Leg Speed").Value = getField(linea)
    rs.Fields("Leg Course").Value = getField(linea)
    rs.Update
  End If
End Sub

Private Sub Waypoint(ByRef subExitCode As Long, ByRef linenumber As Long, ByRef rsWp As DAO.Recordset, ByRef firstField As String, ByRef WaypointId As Long, ByRef linea As String)
  Dim posicion As String
  subExitCode = 0
  If procesarSoloWaypoints Then
    rsWp.AddNew
    rsWp.Fields("ID").Value = WaypointId
    rsWp.Fields("IdPrevio").Value = IIf(WaypointId = 1, Null, WaypointId - 1)
    rsWp.Fields("Header").Value = firstField
    rsWp.Fields("Name").Value = getField(linea)
    rsWp.Fields("Description").Value = getField(linea)
    rsWp.Fields("Type").Value = getField(linea)
    posicion = getField(linea)
    rsWp.Fields("Position").Value = posicion
    rsWp.Fields("Latitude").Value = latitude(getCoordField(posicion))
    rsWp.Fields("Longitude").Value = longitude(getCoordField(posicion))
    rsWp.Fields("Altitude").Value = getField(linea)
    rsWp.Fields("Depth").Value = getField(linea)
    rsWp.Fields("Proximity").Value = getField(linea)
    rsWp.Fields("Temperature").Value = getField(linea)
    rsWp.Fields("Display Mode").Value = getField(linea)
    rsWp.Fields("Color").Value = getField(linea)
    rsWp.Fields("Symbol").Value = getField(linea)
    rsWp.Fields("Facility").Value = getField(linea)
    rsWp.Fields("City").Value = getField(linea)
    rsWp.Fields("State").Value = getField(linea)
    rsWp.Fields("Country").Value = getField(linea)
    rsWp.Fields("Date Modified").Value = getField(linea)
    rsWp.Fields("Link").Value = getField(linea)
    rsWp.Fields("Categories").Value = getField(linea)
    rsWp.Update
  End If
End Sub

Private Sub Corte(ByRef subExitCode As Long, ByRef linenumber As Long, ByRef blankLineFg As Boolean, ByRef firstsHeaderFg As Boolean, ByRef firstsHeaderEndedFg As Boolean, ByRef firstField As String, ByRef trackpointSec As Long, ByRef inTrackFg As Boolean, ByRef trackHeaderReadFg As Boolean)
  subExitCode = 0
  If firstsHeaderFg Then
    firstsHeaderEndedFg = True
  End If
  If inTrackFg Then
    If trackpointSec = 0 Then
      MsgBox "No se encontraron trackpoints, linea " + Trim(CStr(linenumber)), vbCritical, "Importar() function"
      subExitCode = -1
      Exit Sub
    End If
    inTrackFg = False
    If trackHeaderReadFg Then
      trackHeaderReadFg = False
    End If
  End If
  trackpointSec = 0
End Sub

Private Function getField(ByRef pLine As String, Optional ByVal pRepDecSep As Boolean = False) As Variant
  Dim pos As Integer
  If LTrim(pLine) = "" Then
    getField = nullIfEmpty("")
    pLine = ""
  Else
    pos = InStr(pLine, vbTab)
    If pos = 0 Then
      getField = nullIfEmpty(Trim(pLine))
      pLine = ""
    Else
      getField = nullIfEmpty(Trim(Mid(pLine, 1, pos - 1)))
      If pos = Len(pLine) Then
        pLine = ""
      Else
        pLine = Trim(Mid(pLine, pos + 1))
      End If
    End If
  End If
  If pRepDecSep Then
    If IsNull(getField) = False Then
      getField = Replace(getField, ",", ".")
    End If
  End If
End Function

Private Function getCoordField(ByRef pLine As String) As Variant
  Dim pos As Integer
  Dim aux As String
  Dim i As Long
  Const separador As String = " "
  '
  aux = ""
  For i = 1 To CoordSubfields
    If LTrim(pLine) = "" Then
      pLine = ""
      Exit For
    Else
      pos = InStr(pLine, separador)
      If pos = 0 Then
        If i = 1 Then
          aux = Trim(pLine)
        Else
          aux = aux + separador + Trim(pLine)
        End If
        pLine = ""
        Exit For
      Else
        If i = 1 Then
          aux = Trim(Mid(pLine, 1, pos - 1))
        Else
          aux = aux + separador + Trim(Mid(pLine, 1, pos - 1))
        End If
        If pos = Len(pLine) Then
          pLine = ""
          Exit For
        Else
          pLine = Trim(Mid(pLine, pos + 1))
        End If
      End If
    End If
  Next
  If aux = "" Or (Trim(aux) = "" And Len(aux) <= CoordSubfields) Then
    getCoordField = Null
  Else
    getCoordField = aux
  End If
  '
End Function

Private Function getSubField(ByRef pLine As String) As Variant
  Dim pos As Integer
  Const separador As String = " "
  If LTrim(pLine) = "" Then
    getSubField = nullIfEmpty("")
    pLine = ""
  Else
    pos = InStr(pLine, separador)
    If pos = 0 Then
      getSubField = nullIfEmpty(Trim(pLine))
      pLine = ""
    Else
      getSubField = nullIfEmpty(Trim(Mid(pLine, 1, pos - 1)))
      If pos = Len(pLine) Then
        pLine = ""
      Else
        pLine = Trim(Mid(pLine, pos + 1))
      End If
    End If
  End If
End Function

Private Function ImportFile() As Integer
'For testing purposes
' This function will return 1 if a file is imported, 0 if not imported
  Dim FileSpec As String
  FileSpec = OpenTextFile("D:\Download\")

  If Len(FileSpec) = 0 Then
    MsgBox "Operation Cancelled - file not loaded", vbOKOnly, "ImportFile() function"
    ImportFile = 0
  Else
    DoCmd.TransferText A_IMPORTFIXED, "temp_text", "temp_text", FileSpec
    ImportFile = 1
  End If
End Function

Private Function nullIfEmpty(ByVal pTexto As String) As Variant
  nullIfEmpty = IIf(pTexto = "", Null, pTexto)
End Function

Private Function longitude(ByVal pCampoLongitude As Variant) As Variant
  Dim signo As Integer
  Dim aux As String
  Dim grados As String
  Dim minutos As String
  '
  If IsNull(pCampoLongitude) Or Trim(pCampoLongitude) = "" Then
    longitude = Null
  Else
    Select Case FormatoGrid
      Case fmtGrLatLong_hdddG_mm_mmmM, fmtGrLatLon_hdddG_mm_mmmM
        signo = IIf(Mid(pCampoLongitude, 1, 1) = "W", -1, 1)
        aux = Mid(pCampoLongitude, 2)
        grados = getSubField(aux)
        minutos = getSubField(aux)
        longitude = CLng(Round((Val(grados) + Val(minutos) / 60) * signo * 1000000, 0))
      Case Else    'fmtGrLatLong_hddd_dddddG
        signo = IIf(Mid(pCampoLongitude, 1, 1) = "W", -1, 1)
        longitude = CLng(Round(Val(Mid(pCampoLongitude, 2)) * signo * 1000000, 0))
    End Select
  End If
  '
End Function

Private Function latitude(ByVal pCampoLatitude As Variant) As Variant
  Dim signo As Integer
  Dim aux As String
  Dim grados As String
  Dim minutos As String
  '
  If IsNull(pCampoLatitude) Or Trim(pCampoLatitude) = "" Then
    latitude = Null
  Else
    Select Case FormatoGrid
      Case fmtGrLatLong_hdddG_mm_mmmM, fmtGrLatLon_hdddG_mm_mmmM
        signo = IIf(Mid(pCampoLatitude, 1, 1) = "S", -1, 1)
        aux = Mid(pCampoLatitude, 2)
        grados = getSubField(aux)
        minutos = getSubField(aux)
        latitude = CLng(Round((Val(grados) + Val(minutos) / 60) * signo * 1000000, 0))
      Case Else    'fmtGrLatLong_hddd_dddddG
        signo = IIf(Mid(pCampoLatitude, 1, 1) = "S", -1, 1)
        latitude = CLng(Round(Val(Mid(pCampoLatitude, 2)) * signo * 1000000, 0))
    End Select
  End If
  '
End Function

Private Function existeArchivo(ByVal pArchivo As String) As Boolean
  existeArchivo = (Dir(pArchivo) > "")
End Function

Private Function exportarArcosTrackAcsv(ByVal pFileName As String, ByVal pTrack As Long) As Boolean
  Dim configuracionEsquema As String
  Dim sql As String
  Dim db As DAO.Database
  Dim direct As String
  Dim FileName As String
  Dim qd As DAO.QueryDef
  '
  Set db = CurrentDb
  Set qd = db.QueryDefs("ExportacionAcsv")
  sql = "SELECT ID, TrackID, Archivo, TrackName, SegmentoID, LengthMts, Segundos, KmsHr, AltitudeMts, IndPendAbs, IndiceGiro, """" & arcos_salida.Hora as Hora"
  sql = sql + " FROM arcos_salida"
  If pTrack <> 0 Then
    sql = sql + " WHERE TrackID = " + Trim(CStr(pTrack))
  End If
  qd.sql = sql
  BorrarArchivo pFileName
  DoCmd.TransferText acExportDelim, "ExportacionAcsv Export Specification", "ExportacionAcsv", pFileName, HasFieldNames:=True
  '
End Function

Private Sub BorrarArchivo(ByVal pArchivo As String)
  If Dir(pArchivo) > "" Then
    Kill pArchivo
  End If
End Sub

Public Function ShowProgressBar(ByVal mensaje As String, ByVal totalRegistros As Long) As Long
  ShowProgressBar = SysCmd(acSysCmdInitMeter, mensaje, totalRegistros)
  DoEvents
End Function

Public Function UpdateProgressBar(ByVal registro As Long) As Long
  UpdateProgressBar = SysCmd(acSysCmdUpdateMeter, registro)
  DoEvents
End Function

Public Function HideProgressBar() As Long
  HideProgressBar = SysCmd(acSysCmdRemoveMeter)
  DoEvents
End Function

Public Function crearCarpeta(ByVal carpeta As String) As Boolean
  If Not Dir(carpeta, vbDirectory) > "" Then
    MkDir carpeta
    crearCarpeta = True
  End If
End Function

Public Function soloNombreArch(ByVal camino As String) As String
  Dim aux As String
  '
  aux = camino
  If InStrRev(aux, "\") > 0 Then
    aux = Mid(aux, InStrRev(aux, "\") + 1)
  End If
  If InStrRev(aux, ".") > 0 And InStrRev(aux, "\") < InStrRev(aux, ".") Then
    aux = Mid(aux, 1, InStrRev(aux, ".") - 1)
  End If
  soloNombreArch = aux
  '
End Function

Public Function sinExtens(ByVal camino As String) As String
  Dim aux As String
  '
  aux = camino
  If InStrRev(aux, ".") > 0 And InStrRev(aux, "\") < InStrRev(aux, ".") Then
    sinExtens = Mid(aux, 1, InStrRev(aux, ".") - 1)
  Else
    sinExtens = aux
  End If
  '
End Function

Public Function fullName(ByVal prmPath As String, ByVal prmFile As String) As String
  If Right(prmPath, 1) = Chr(92) Then
    fullName = prmPath + prmFile
  Else
    fullName = prmPath + Chr(92) + prmFile
  End If
End Function

Public Function nombreCarpeta(ByVal prmFile As String) As String
  Dim aux As String
  '
  aux = prmFile
  If Right(aux, 1) = Chr(92) Then
    aux = Mid(aux, 1, Len(aux) - 1)
  End If
  If InStrRev(aux, "\") > 0 Then
    aux = Mid(aux, 1, InStrRev(aux, "\") - 1)
    If Right(aux, 1) = ":" Then 'root
      aux = aux + "\"
    End If
  Else
    aux = "" 'padre de root
  End If
  nombreCarpeta = aux
  '
End Function

Public Function IsRoot(ByVal prmFolder As String) As Boolean
  Dim aux As String
  '
  aux = prmFolder
  If Right(aux, 1) = ":" Then
    prmFolder = True
  ElseIf Right(aux, 2) = ":\" Then
    prmFolder = True
  Else
    prmFolder = False
  End If
  '
End Function

Private Function BorrarDatos()
  Dim db As DAO.Database
  Dim sql As String
  '
  sql = "delete from [trackpoints]"
  Set db = CurrentDb
  db.Execute sql, dbFailOnError
  '
End Function

Private Function BorrarWaypoints()
  Dim db As DAO.Database
  Dim sql As String
  '
  sql = "delete from [waypoints]"
  Set db = CurrentDb
  db.Execute sql, dbFailOnError
  '
End Function
Private Function diffSeconds(ByVal pTime1 As Variant, ByVal pTime2 As Variant) As Long
  Dim aux As Double
  '
  aux = gbMaxNumber
  If Not (IsNull(pTime1) Or IsNull(pTime2)) Then
    If Abs(DatePart("yyyy", pTime1) - DatePart("yyyy", pTime2)) > 50 Then
      aux = gbMaxLongNumber
    Else
      aux = DateDiff("s", pTime1, pTime2)
    End If
  End If
  '
  If aux > gbMaxLongNumber Then
    aux = gbMaxLongNumber
  End If
  diffSeconds = aux
  '
End Function

Private Sub borrarArcosSalida()
  Dim db As DAO.Database
  Dim sql As String
  '
  Set db = CurrentDb
  sql = "delete from [arcos_salida]"
  Set db = CurrentDb
  db.Execute sql, dbFailOnError
  '
End Sub

Private Function CalcularDistPerpendTpsArco(ByRef trackId As Long, ByRef arcoElegido, ByRef distPerpRel As Double, ByRef distPerpAbs As Double, ByRef rsTp As DAO.Recordset)
    Dim progInicial As Double
    Dim tp As Variant
    Dim tpAnt As Variant
    Dim i As Long
    Dim ta As Long
    Dim tb As Long
    Dim dPrel As Double
    Dim dPrelAnt As Double
    Dim distSegm As Double
    Dim distAcum As Double
    '
    Dim ax As Double
    Dim ay As Double
    Dim bx As Double
    Dim by As Double
    Rem Dim distAb As Double
    '
    ax = arcoElegido(0)(5)(0)
    ay = arcoElegido(0)(5)(1)
    bx = arcoElegido(1)(5)(0)
    by = arcoElegido(1)(5)(1)
    Rem distAb = arcoElegido(1)(4)
    '
    distPerpRel = 0
    distPerpAbs = 0
    '
    ta = arcoElegido(0)(0)
    tb = arcoElegido(1)(0)
    tpAnt = CoordenadaTrackpoint(rsTp, trackId, ta)
    dPrelAnt = CalcularDistPerpend(ax, ay, bx, by, tpAnt(0), tpAnt(1))
    '
    For i = ta + 1 To tb
        '
        tp = CoordenadaTrackpoint(rsTp, trackId, i)
        distSegm = tp(2) - tpAnt(2) 'diferenca de progresiva respecto a trackpoint anterior
        distAcum = distAcum + distSegm
        dPrel = CalcularDistPerpend(ax, ay, bx, by, tp(0), tp(1))
        distPerpRel = distPerpRel + ((dPrel + dPrelAnt) / 2 * distSegm)
        distPerpAbs = distPerpAbs + ((Abs(dPrel) + Abs(dPrelAnt)) / 2 * distSegm)
        '
        tpAnt = tp
        dPrelAnt = dPrel
        '
    Next
    '
    distPerpRel = distPerpRel / distAcum
    distPerpAbs = distPerpAbs / distAcum
    '
End Function

Private Function CalcularDistPerpend(ByRef ax, ByRef ay, ByRef bx, ByRef by, ByRef px, ByRef py) As Double
    CalcularDistPerpend = ((px - ax) * (by - ay) - (py - ay) * (bx - ax)) / Hipotenusa(ax - bx, ay - by) / 11#
End Function

Private Function Hipotenusa(ByRef x, ByRef y) As Double
    Hipotenusa = Sqr(x ^ 2 + y ^ 2)
End Function
