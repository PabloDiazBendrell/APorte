
Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports SBN_APORTES.AporteActoService


Public Class APORTE_ACTO

    Dim cusBuscar As Integer
    Dim areaBuscar As String

    Private Sub APORTE_ACTO_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim epsgs = New List(Of Combo)

        epsgs.Add(New Combo(24877, "UTM - PSAD56 ZONA 17"))
        epsgs.Add(New Combo(24878, "UTM - PSAD56 ZONA 18"))
        epsgs.Add(New Combo(24879, "UTM - PSAD56 ZONA 19"))
        epsgs.Add(New Combo(32717, "UTM - WGS84 -ZONA 17"))
        epsgs.Add(New Combo(32718, "UTM - WGS84 -ZONA 18"))
        epsgs.Add(New Combo(32719, "UTM - WGS84 -ZONA 19"))

        cmbEpsg.DataSource = epsgs
        cmbEpsg.DisplayMember = "Name"
        cmbEpsg.ValueMember = "ID"


        Dim areas = New List(Of Combo)

        areas.Add(New Combo(0, "TODAS"))
        areas.Add(New Combo(1, "SDDI"))
        areas.Add(New Combo(2, "SDAPE"))

        cmbAreaOrganicaBuscar.DataSource = areas
        cmbAreaOrganicaBuscar.DisplayMember = "Name"
        cmbAreaOrganicaBuscar.ValueMember = "ID"
    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click

        lblTotal.Text = "Total:"
        lblTotal.Refresh()

        Dim clase As Clase = New Clase()
        pnlDatos.Enabled = False
        lblCus.Text = ""
        dgvActos.DataSource = DBNull.Value

        btnPsad17.Enabled = False
        btnPsad17.Image = My.Resources.dResources.Download_gray

        btnPsad17.BaBackColor = System.Drawing.SystemColors.Control

        btnPsad18.Enabled = False
        btnPsad18.Image = My.Resources.Resources.Download_gray
        btnPsad18.BackColor = System.Drawing.SystemColors.Control

        btnPsad19.Enabled = False
        btnPsad19.Image = My.Resources.Resources.Download_gray
        btnPsad19.BackColor = System.Drawing.SystemColors.Control

        btnWgs17.Enabled = False
        btnWgs17.Image = My.Resources.Resources.Download_gray
        btnWgs17.BackColor = System.Drawing.SystemColors.Control

        btnwgs18.Enabled = False
        btnwgs18.Image = My.Resources.Resources.Download_gray
        btnwgs18.BackColor = System.Drawing.SystemColors.Control

        btnWgs19.Enabled = False
        btnWgs19.Image = My.Resources.Resources.Download_gray
        btnWgs19.BackColor = System.DrawingSystemColors.Control


        Dim num As New System.Text.RegularExpressions.Regex("^\d+$")

        Dim cus As String = txtCusBuscar.Text

        If num.IsMatch(cus) = False Then
            MsgBox("Solo se permiten nros.", MsgBoxStyle.Exclamation, "Aviso")
        Else
            Dim client As AporteActoClient = WCF.acto()
            Dim clientPredio As AportePredioService.AportePredioClient = WCF.predio()
            Try
                Dim p_cus As Integer = CType(cus, Integer)

                Dim area As Combo = cmbAreaOrganicaBuscar.SelectedItem

                Dim rpta As RespuestaDataOfActov8beLqCf = client.buscarActos_x_cus(p_cus, If(area.ID = 0, "", area.Name))

                client.Close()

                If rpta.error Then
                    MsgBox(rpta.mensaje, MsgBoxStyle.Exclamation, "Aviso")
                Else
                    lblTotal.Text = "Total:" + rpta.data.Length.ToString()
                    lblTotal.Refresh()

                    cusBuscar = p_cus
                    areaBuscar = If(area.ID = 0, "", area.Name)

                    Dim datos As Acto() = rpta.data
                    dgvActos.DataSource = datos

                    pnlDatos.Enabled = True
                    lblCus.Text = cus 'dato.cap_cus

                    Dim rptaPredio As AportePredioService.RespuestaDataOfCusv8beLqCf = clientPredio.buscarCus(p_cus)

                    clientPredio.Close()
                    If (rptaPredio.error = True) Then
                        MsgBox(rptaPredio.mensaje, MsgBoxStyle.Critical, "Aviso")
                    Else
                        If (rptaPredio.data.Length <> 1) Then
                            MsgBox("No hay aporte gráfico aportado para el CUS buscado.", MsgBoxStyle.Information, "Aviso")
                        Else
                            Dim dato As AportePredioService.Cus = rptaPredio.data(0)

                            btnPsad17.Enabled = If(dato.cap_epsg_24877 > 0, True, False)
                            btnPsad17.Image = If(dato.cap_epsg_24877 > 0, My.Resources.Resources.Download_blue, My.Resources.Resources.Download_gray)
                            btnPsad17.BackColor = If(dato.cap_epsg_24877 > 0, System.Drawing.SystemColors.ActiveCaption, System.Drawing.SystemColors.Control)

                            btnPsad18.Enabled = If(dato.cap_epsg_24878 > 0, True, False)
                            btnPsad18.Image = If(dato.cap_epsg_24878 > 0, My.Resources.Resources.Download_blue, My.Resources.Resources.Download_gray)
                            btnPsad18.BackColor = If(dato.cap_epsg_24878 > 0, System.Drawing.SystemColors.ActiveCaption, System.Drawing.SystemColors.Control)

                            btnPsad19.Enabled = If(dato.cap_epsg_24879 > 0, True, False)
                            btnPsad19.Image = If(dato.cap_epsg_24879 > 0, My.Resources.Resources.Download_blue, My.Resources.Resources.Download_gray)
                            btnPsad19.BackColor = If(dato.cap_epsg_24879 > 0, System.Drawing.SystemColors.ActiveCaption, System.Drawing.SystemColors.Control)

                            btnWgs17.Enabled = If(dato.cap_epsg_32717 > 0, True, False)
                            btnWgs17.Image = If(dato.cap_epsg_32717 > 0, My.Resources.Resources.Download_blue, My.Resources.Resources.Download_gray)
                            btnWgs17.BackColor = If(dato.cap_epsg_32717 > 0, System.Drawing.SystemColors.ActiveCaption, System.Drawing.SystemColors.Control)

                            btnwgs18.Enabled = If(dato.cap_epsg_32718 > 0, True, False)
                            btnwgs18.Image = If(dato.cap_epsg_32718 > 0, My.Resources.Resources.Download_blue, My.Resources.Resources.Download_gray)
                            btnwgs18.BackColor = If(dato.cap_epsg_32718 > 0, System.Drawing.SystemColors.ActiveCaption, System.Drawing.SystemColors.Control)

                            btnWgs19.Enabled = If(dato.cap_epsg_32719 > 0, True, False)
                            btnWgs19.Image = If(dato.cap_epsg_32719 > 0, My.Resources.Resources.Download_blue, My.Resources.Resources.Download_gray)
                            btnWgs19.BackColor = If(dato.cap_epsg_32719 > 0, System.Drawing.SystemColors.ActiveCaption, System.Drawing.SystemColors.Control)
                        End If
                    End If
                End If
            Catch ex As Exception
                client.Abort()
                clientPredio.Abort()
                Dim msg As String = "Error: Aportar Acto - buscar . " + ex.Message
                Dim rptaExc As AportePermisoService.Respuesta = clase.errorMsg(msg, "app", MyClass.Name, autocadVersion.ToString)
                MsgBox(rptaExc.mensaje, MsgBoxStyle.Exclamation, "Aviso")
            End Try
        End If
    End Sub

    Private Sub cmbEpsg_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbEpsg.KeyPress
        e.Handled = True
    End Sub

    Private Sub cmbAreaOrganicaBuscar_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cmbAreaOrganicaBuscar.KeyPress
        e.Handled = True
    End Sub

    Private Sub btnPsad17_Click(sender As Object, e As EventArgs) Handles btnPsad17.Click
        cmbEpsg.SelectedValue = 24877
        cargaDatos("24877")
    End Sub

    Private Sub btnPsad18_Click(sender As Object, e As EventArgs) Handles btnPsad18.Click
        cmbEpsg.SelectedValue = 24878
        cargaDatos("24878")
    End Sub

    Private Sub btnPsad19_Click(sender As Object, e As EventArgs) Handles btnPsad19.Click
        cmbEpsg.SelectedValue = 24879
        cargaDatos("24879")
    End Sub

    Private Sub btnWgs17_Click(sender As Object, e As EventArgs) Handles btnWgs17.Click
        cmbEpsg.SelectedValue = 32717
        cargaDatos("32717")
    End Sub

    Private Sub btnwgs18_Click(sender As Object, e As EventArgs) Handles btnwgs18.Click
        cmbEpsg.SelectedValue = 32718
        cargaDatos("32718")
    End Sub

    Private Sub btnWgs19_Click(sender As Object, e As EventArgs) Handles btnWgs19.Click
        cmbEpsg.SelectedValue = 32719
        cargaDatos("32719")
    End Sub

    Private Sub btnCargar_Click(sender As Object, e As EventArgs) Handles btnCargar.Click
        lblMsg.Text = ""
        lblMsg.Refresh()
        Dim clase As Clase = New Clase()
        Dim num As New System.Text.RegularExpressions.Regex("^\d+$")

        Dim cus As String = lblCus.Text.ToString().Trim()

        If num.IsMatch(cus) = False Then
            MsgBox("Cus Seleccionado no es correcto.", MsgBoxStyle.Exclamation, "Aviso")
        Else
            Dim selectedRowCount As Int32 = dgvActos.Rows.GetRowCount(DataGridViewElementStates.Selected)
            If selectedRowCount <> 1 Then
                MsgBox("Seleccionar un acto..", MsgBoxStyle.Exclamation, "Aviso")
            Else
                Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim acDocEd As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

                Dim selectionOpts As PromptSelectionOptions = New PromptSelectionOptions()
                Dim selectionRes As PromptSelectionResult = acDocEd.SelectImplied()
                selectionRes = acDocEd.GetSelection(selectionOpts)

                If selectionRes.Status = PromptStatus.OK Then
                    Dim acSSet1 As SelectionSet = selectionRes.Value
                    'If acSSet1.Count <> 2 Then
                    'MsgBox("Debe seleccionar mínimo 2 objetos: texto y Polilinea.", MsgBoxStyle.Exclamation, "Aviso")
                    If acSSet1.Count < 1 Then
                        MsgBox("Debe seleccionar mínimo 1 objeto Polilinea.", MsgBoxStyle.Exclamation, "Aviso")
                    Else
                        lblMsg.Text = "Procesando Datos..."
                        lblMsg.Refresh()
                        Dim dbLocal As Database = doc.Database
                        Using doc.LockDocument()
                            Dim trLocal As Transaction = doc.TransactionManager.StartTransaction()
                            Using (trLocal)
                                Dim contaPoligono As Integer = 0
                                Dim finaliza As Boolean = False

                                Dim acObjIdColl = New ObjectIdCollection()

                                Dim ext As New List(Of Extents3d)()
                                Dim PolylineTemp As Polyline

                                For index As Integer = 0 To acSSet1.Count - 1
                                    Dim objeto As Entity = trLocal.GetObject(acSSet1.Item(index).ObjectId, OpenMode.ForRead)
                                    If IsDBNull(objeto) Then
                                        finaliza = True
                                        MsgBox("Objeto seleccionado es nulo", MsgBoxStyle.Exclamation, "Aviso")
                                        lblMsg.Text = ""
                                        lblMsg.Refresh()
                                        Exit For
                                    Else
                                        If objeto.GetType().Name.ToUpper() <> "POLYLINE" And objeto.GetType().Name.ToUpper() <> "DBTEXT" And objeto.GetType().Name.ToUpper() <> "MTEXT" Then
                                            finaliza = True
                                            MsgBox("Debe seleccionar un objeto de tipo Polilínea o Texto.", MsgBoxStyle.Exclamation, "Aviso")
                                            lblMsg.Text = ""
                                            lblMsg.Refresh()
                                            Exit For
                                        Else
                                            If objeto.GetType().Name.ToUpper() = "POLYLINE" Then
                                                contaPoligono = contaPoligono + 1
                                                ext.Add(objeto.GeometricExtents)
                                                PolylineTemp = objeto
                                                If PolylineTemp.Closed = False Then
                                                    MsgBox("Polilinea debe estar cerrado.", MsgBoxStyle.Exclamation, "Aviso")
                                                    finaliza = True
                                                    Exit For
                                                End If
                                                If PolylineTemp.Elevation <> 0 Then
                                                    MsgBox("Elevación de polilinea debe ser 0.", MsgBoxStyle.Exclamation, "Aviso")
                                                    finaliza = True
                                                    Exit For
                                                End If
                                            End If

                                            Dim rb As ResultBuffer = objeto.XData
                                            If IsNothing(rb) = False Then
                                                objeto.UpgradeOpen()
                                                Dim n As Integer = 0
                                                For Each tv As TypedValue In rb
                                                    n = n + 1
                                                    If tv.TypeCode = 1001 Then
                                                        objeto.XData = New ResultBuffer(New TypedValue(tv.TypeCode, tv.Value))
                                                    End If
                                                Next
                                                rb.Dispose()
                                                objeto.DowngradeOpen()
                                            End If
                                            objeto.Dispose()
                                            objeto = trLocal.GetObject(acSSet1.Item(index).ObjectId, OpenMode.ForRead)

                                            acObjIdColl.Add(objeto.ObjectId)

                                        End If
                                    End If
                                Next

                                If finaliza = False Then
                                    If contaPoligono < 1 Then
                                        MsgBox("Debe seleccionar mínimo 1 Polígono.", MsgBoxStyle.Exclamation, "Aviso")
                                        lblMsg.Text = ""
                                        lblMsg.Refresh()
                                    Else
                                        'Dim acObjIdColl = New ObjectIdCollection()

                                        'Dim objeto1 As Entity = trLocal.GetObject(acSSet1.Item(0).ObjectId, OpenMode.ForRead)
                                        'Dim objeto2 As Entity = trLocal.GetObject(acSSet1.Item(1).ObjectId, OpenMode.ForRead)

                                        'If IsDBNull(objeto1) Or IsDBNull(objeto2) Then
                                        '    MsgBox("Objeto seleccionado es nulo", MsgBoxStyle.Exclamation, "Aviso")
                                        '    lblMsg.Text = ""
                                        '    lblMsg.Refresh()
                                        'Else
                                        '    If (objeto1.GetType().Name.ToUpper() <> "POLYLINE" And objeto2.GetType().Name.ToUpper() <> "POLYLINE") Then
                                        '        lblMsg.Text = ""
                                        '        lblMsg.Refresh()
                                        '        MsgBox("Debe seleccionar un polígono.", MsgBoxStyle.Exclamation, "Aviso")
                                        '    Else
                                        '        If (objeto1.GetType().Name.ToUpper() <> "DBTEXT" And
                                        '           objeto1.GetType().Name.ToUpper() <> "MTEXT" And
                                        '           objeto2.GetType().Name.ToUpper() <> "DBTEXT" And
                                        '           objeto2.GetType().Name.ToUpper() <> "MTEXT"
                                        '        ) Then
                                        '            lblMsg.Text = ""
                                        '            lblMsg.Refresh()
                                        '            MsgBox("Debe seleccionar un texto.", MsgBoxStyle.Exclamation, "Aviso")
                                        '        Else
                                        '            Dim rb As ResultBuffer = objeto1.XData
                                        '            If IsNothing(rb) = False Then
                                        '                objeto1.UpgradeOpen()
                                        '                Dim n As Integer = 0
                                        '                For Each tv As TypedValue In rb
                                        '                    n = n + 1
                                        '                    If tv.TypeCode = 1001 Then
                                        '                        objeto1.XData = New ResultBuffer(New TypedValue(tv.TypeCode, tv.Value))
                                        '                    End If
                                        '                Next
                                        '                rb.Dispose()
                                        '                objeto1.DowngradeOpen()
                                        '            End If
                                        '            objeto1.Dispose()
                                        '            objeto1 = trLocal.GetObject(acSSet1.Item(0).ObjectId, OpenMode.ForRead)

                                        '            Dim rb1 As ResultBuffer = objeto2.XData
                                        '            If IsNothing(rb1) = False Then
                                        '                objeto2.UpgradeOpen()
                                        '                Dim n As Integer = 0
                                        '                For Each tv As TypedValue In rb1
                                        '                    n = n + 1
                                        '                    If tv.TypeCode = 1001 Then
                                        '                        objeto2.XData = New ResultBuffer(New TypedValue(tv.TypeCode, tv.Value))
                                        '                    End If
                                        '                Next
                                        '                rb1.Dispose()
                                        '                objeto2.DowngradeOpen()
                                        '            End If
                                        '            objeto2.Dispose()
                                        '            objeto2 = trLocal.GetObject(acSSet1.Item(1).ObjectId, OpenMode.ForRead)

                                        '            acObjIdColl.Add(objeto1.ObjectId)
                                        '            acObjIdColl.Add(objeto2.ObjectId)



                                        Dim row As DataGridViewRow = dgvActos.SelectedRows.Item(0)
                                        Dim positionRow As Integer = row.Index

                                        Dim dato As Acto = row.DataBoundItem

                                        Dim client As AporteActoClient = WCF.acto()

                                        Try
                                            Dim rptaExists As RespuestaDataOfcatastro_actov8beLqCf = client.buscar_acto_x_id_aporte_origen(dato.IdAporteGeografico, dato.origen)
                                            If (rptaExists.error) Then
                                                client.Close()
                                                MsgBox(rptaExists.mensaje, MsgBoxStyle.Exclamation, "Aviso")
                                            Else
                                                If rptaExists.data.Length > 1 Then
                                                    client.Close()
                                                    MsgBox("Se encontraron varios acto con el mismo ID.", MsgBoxStyle.Exclamation, "Aviso")
                                                Else
                                                    Dim msg1 As String = "Seguro de aportar Acto?"
                                                    Dim title As String = "Aportar Acto Nuevo"

                                                    If rptaExists.data.Length = 1 Then
                                                        Dim datoCatastro As catastro_acto = rptaExists.data(0)


                                                        msg1 = "Acto ya fue aportado en los siguientes sistemas:"

                                                        msg1 += If(datoCatastro.cac_epsg_24877 > 0, Chr(13) + " * UTM - PSAD56 ZONA 17", "")
                                                        msg1 += If(datoCatastro.cac_epsg_24878 > 0, Chr(13) + " * UTM - PSAD56 ZONA 18", "")
                                                        msg1 += If(datoCatastro.cac_epsg_24879 > 0, Chr(13) + " * UTM - PSAD56 ZONA 19", "")
                                                        msg1 += If(datoCatastro.cac_epsg_32717 > 0, Chr(13) + " * UTM - WGS84 ZONA 17", "")
                                                        msg1 += If(datoCatastro.cac_epsg_32718 > 0, Chr(13) + " * UTM - WGS84 ZONA 18", "")
                                                        msg1 += If(datoCatastro.cac_epsg_32719 > 0, Chr(13) + " * UTM - WGS84 ZONA 19", "")

                                                        msg1 += Chr(13) + "Desea aportar Acto?"   ' Define message.

                                                        title = "Acto ya aportado"   ' Define title.
                                                    End If
                                                    Dim style As MsgBoxStyle
                                                    Dim response As MsgBoxResult

                                                    style = MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Information Or MsgBoxStyle.YesNo

                                                    response = MsgBox(msg1, style, title)

                                                    If response = MsgBoxResult.Yes Then   ' User chose Yes.
                                                        lblMsg.Text = "Subiendo Datos..."
                                                        lblMsg.Refresh()
                                                        Dim dbDestino As Database = New Database()
                                                        Dim mapping As IdMapping = New IdMapping()
                                                        Dim trDestino As Transaction = dbDestino.TransactionManager.StartTransaction()
                                                        Using (trDestino)
                                                            If ext.Count > 0 Then
                                                                Dim vpt As ViewportTable = trDestino.GetObject(dbDestino.ViewportTableId, OpenMode.ForRead)
                                                                Dim vptr As ViewportTableRecord = trDestino.GetObject(vpt("*Active"), OpenMode.ForWrite)

                                                                'Point2d pmin = Point2d.Origin;
                                                                'Point2d pmax = New Point2d(3000, 1500);
                                                                'Dim ed As Editor = doc.Editor
                                                                Dim xPointMin As Double
                                                                Dim yPointMin As Double
                                                                Dim xPointMax As Double
                                                                Dim yPointMax As Double
                                                                Dim first As Boolean = True
                                                                For Each obj As Extents3d In ext
                                                                    obj.TransformBy(acDocEd.CurrentUserCoordinateSystem.Inverse())
                                                                    If first Then
                                                                        xPointMin = obj.MinPoint.X
                                                                        yPointMin = obj.MinPoint.Y
                                                                        xPointMax = obj.MaxPoint.X
                                                                        yPointMax = obj.MaxPoint.Y
                                                                    Else
                                                                        xPointMin = If(obj.MinPoint.X < xPointMin, obj.MinPoint.X, xPointMin)
                                                                        yPointMin = If(obj.MinPoint.Y < yPointMin, obj.MinPoint.Y, yPointMin)
                                                                        xPointMax = If(obj.MaxPoint.X > xPointMax, obj.MaxPoint.X, xPointMax)
                                                                        yPointMax = If(obj.MaxPoint.Y > yPointMax, obj.MaxPoint.Y, yPointMax)
                                                                    End If
                                                                    first = False
                                                                Next


                                                                Dim pmin As Point2d = New Point2d(xPointMin, yPointMin)
                                                                Dim pmax As Point2d = New Point2d(xPointMax, yPointMax)
                                                                Dim Height As Double = pmax.Y - pmin.Y
                                                                Dim Width As Double = pmax.X - pmin.X

                                                                vptr.CenterPoint = pmin + ((pmax - pmin) / 2.0)
                                                                vptr.Height = Height
                                                                vptr.Width = Width
                                                            End If
                                                            '' Open the Block table for read 
                                                            Dim btDestino As BlockTable = trDestino.GetObject(dbDestino.BlockTableId, OpenMode.ForRead)
                                                            '' Open the Block table record Model space for write 
                                                            Dim btrDestino As BlockTableRecord = trDestino.GetObject(btDestino(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                                                            '' Clone the objects to the new database           
                                                            Dim acIdMap As IdMapping = New IdMapping()
                                                            dbLocal.WblockCloneObjects(acObjIdColl, btrDestino.ObjectId, acIdMap, DuplicateRecordCloning.Replace, False)
                                                            trDestino.Commit()
                                                        End Using

                                                        Try
                                                            'lblProceso.Text = "Subiendo archivo..."
                                                            'lblProceso.Refresh()


                                                            'Dim v_id_archivo As String = clase.GetRandomString(6)
                                                            'Dim v_nombre_archivo As String = path_procesamiento_acto + "\" + v_id_archivo + ".dxf"

                                                            'If (File.Exists(v_nombre_archivo)) Then
                                                            'v_id_archivo = clase.GetRandomString(11)
                                                            'v_nombre_archivo = path_procesamiento_acto + "\" + v_id_archivo + ".dxf"
                                                            'End If
                                                            'Dim v_nombre_archivo As String = dato.origen + "_" + If(dato.cus = -1, "CUS", dato.cus.ToString()) + "_" + dato.codigo_resolucion.ToString() + "_" + dato.IdAporteGeografico.ToString()
                                                            Dim v_nombre_archivo As String = dato.IdAporteGeografico.ToString()

                                                            dbDestino.DxfOut(path_procesamiento_acto + "\" + v_nombre_archivo + ".dxf", 16, DwgVersion.AC1015)

                                                            'Try
                                                            'lblProceso.Text = "Procesando información..."
                                                            'lblProceso.Refresh()

                                                            Dim rpta As Respuesta = client.aporteIndividual(dato.cac_id, v_nombre_archivo, dato.origen, dato.cus, dato.IdAporteGeografico, cmbEpsg.SelectedValue, usuario, nombre_pc, autocadVersion.ToString())

                                                            lblMsg.Text = ""
                                                            lblMsg.Refresh()
                                                            MsgBox(rpta.mensaje, MsgBoxStyle.Information, "Aviso")

                                                            Dim rptaBus As RespuestaDataOfActov8beLqCf = client.buscarActos_x_cus(cusBuscar, areaBuscar)
                                                            If (rptaBus.error = False) Then
                                                                Dim datos As Acto() = rptaBus.data
                                                                dgvActos.DataSource = datos
                                                                dgvActos.Rows(positionRow).Selected = True
                                                            End If
                                                            client.Close()
                                                        Catch ex As IOException
                                                            client.Abort()
                                                            lblMsg.Text = ""
                                                            lblMsg.Refresh()
                                                            Dim msg As String = "Error: Copiar dxf al servidor al aportar Acto. " + ex.Message
                                                            Dim rptaExc As AportePermisoService.Respuesta = clase.errorMsg(msg, "app", MyClass.Name, autocadVersion.ToString)
                                                            MsgBox(rptaExc.mensaje, MsgBoxStyle.Critical, "Aviso")
                                                        Catch ex As Exception
                                                            client.Abort()
                                                            lblMsg.Text = ""
                                                            lblMsg.Refresh()
                                                            Dim msg As String = "Error: Aportar Acto. " + ex.Message
                                                            Dim rptaExc As AportePermisoService.Respuesta = clase.errorMsg(msg, "app", MyClass.Name, autocadVersion.ToString)
                                                            MsgBox(rptaExc.mensaje, MsgBoxStyle.Critical, "Aviso")
                                                        End Try
                                                    Else
                                                        client.Close()
                                                        lblMsg.Text = ""
                                                        lblMsg.Refresh()
                                                        '            End If
                                                    End If
                                                End If
                                            End If
                                        Catch ex As Exception
                                            client.Abort()
                                            lblMsg.Text = ""
                                            lblMsg.Refresh()
                                            Dim msg As String = "Error: Aportar Acto. " + ex.Message
                                            Dim rptaExc As AportePermisoService.Respuesta = clase.errorMsg(msg, "app", MyClass.Name, autocadVersion.ToString)
                                            MsgBox(rptaExc.mensaje, MsgBoxStyle.Critical, "Aviso")
                                        End Try
                                    End If
                                End If
                                trLocal.Dispose()
                            End Using
                        End Using
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub cargaDatos(ByVal p_epsg As String)
        Dim clase As Clase = New Clase()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        'doc.LockDocument()
        '' Get the current document editor
        Dim acDocEd As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
        Dim dbLocal As Database = doc.Database

        Dim ext
        Dim eye
        Using doc.LockDocument()
            Using dbLocal
                Using trLocal As Transaction = dbLocal.TransactionManager.StartTransaction()

                    Dim ltLocal As LayerTable = trLocal.GetObject(dbLocal.LayerTableId, OpenMode.ForRead)

                    Dim filenm As String = path_historico_predio + "\" + p_epsg + "\ACTUAL\" + lblCus.Text.ToString() + ".dxf"

                    Dim acObjIdColl = New ObjectIdCollection()


                    Dim dbCargo As Database = New Database(False, True)
                    Using (dbCargo)
                        ''Cargamos dxf
                        Try
                            dbCargo.DxfIn(filenm, path_log_predio)
                            Dim laysToClone As ObjectIdCollection = New ObjectIdCollection()
                            Using trCargo As Transaction = dbCargo.TransactionManager.StartTransaction()
                                Dim btCargo As BlockTable = trCargo.GetObject(dbCargo.BlockTableId, OpenMode.ForRead)
                                Dim ltCargo As LayerTable = trCargo.GetObject(dbCargo.LayerTableId, OpenMode.ForRead)

                                For Each btrIdCargo As ObjectId In btCargo
                                    Dim btrCargo As BlockTableRecord = trCargo.GetObject(btrIdCargo, OpenMode.ForRead)

                                    Dim vptCargo As ViewportTable = trCargo.GetObject(dbCargo.ViewportTableId, OpenMode.ForRead)
                                    Dim vptrCargo As ViewportTableRecord = trCargo.GetObject(vptCargo("*Active"), OpenMode.ForRead)

                                    eye = Matrix3d.Rotation(-vptrCargo.ViewTwist, vptrCargo.ViewDirection, vptrCargo.Target) *
                                     Matrix3d.Displacement(vptrCargo.Target - Point3d.Origin) *
                                     Matrix3d.PlaneToWorld(vptrCargo.ViewDirection)

                                    For Each objIdCargo As ObjectId In btrCargo
                                        Dim ent2 As Entity = trCargo.GetObject(objIdCargo, OpenMode.ForRead)
                                        'If ent2.Layer.ToUpper = "BG_CUS" Or ent2.Layer.ToUpper = "BG_CUS_TXT" Or
                                        'ent2.Layer.ToUpper = "BG_CUS_CANCELADO" Or ent2.Layer.ToUpper = "BG_CUS_CANCELADO_TXT" Or
                                        'ent2.Layer.ToUpper = "BG_CUS_REFERENCIAL" Or ent2.Layer.ToUpper = "BG_CUS_REFERENCIAL_TXT" Then
                                        Dim Ent As Entity = ent2.Clone()
                                        acObjIdColl.Add(ent2.ObjectId)
                                        If ltLocal.Has(ent2.Layer.ToUpper) = False Then
                                            laysToClone.Add(ltCargo(ent2.Layer))
                                        End If
                                        If ent2.GetType().Name.ToUpper() = "POLYLINE" Then
                                            ext = ent2.GeometricExtents
                                        End If
                                        'End If
                                    Next
                                    btrCargo.Dispose()
                                Next
                                trCargo.Commit()
                            End Using

                            If acObjIdColl.Count > 0 Then
                                '' Open the Block table for read 
                                Dim btLocal As BlockTable = trLocal.GetObject(dbLocal.BlockTableId, OpenMode.ForRead)
                                '' Open the Block table record Model space for write 
                                Dim btrLocal As BlockTableRecord = trLocal.GetObject(btLocal(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                                '' Clone the objects to the new database           
                                Dim acIdMap As IdMapping = New IdMapping()
                                dbCargo.WblockCloneObjects(acObjIdColl, btrLocal.ObjectId, acIdMap, DuplicateRecordCloning.Replace, False)

                                MsgBox("Se copiaron " + acObjIdColl.Count.ToString() + " objeto(s).", MsgBoxStyle.Information, "Aviso")
                            Else
                                MsgBox("No se copio ningun objeto.", MsgBoxStyle.Exclamation, "Aviso")
                            End If
                        Catch ex As IOException
                            Dim msg As String = "Error: Aportar Acto - descarga Cus." + filenm + " " + ex.Message
                            Dim rptaExc As AportePermisoService.Respuesta = clase.errorMsg(msg, "app", MyClass.Name, autocadVersion.ToString)
                            MsgBox(rptaExc.mensaje, MsgBoxStyle.Critical, "Aviso")
                        Catch ex As Exception
                            Dim msg As String = "Error: Aportar Acto - descarga Cus." + filenm + " " + ex.Message
                            Dim rptaExc As AportePermisoService.Respuesta = clase.errorMsg(msg, "app", MyClass.Name, autocadVersion.ToString)
                            MsgBox(rptaExc.mensaje, MsgBoxStyle.Critical, "Aviso")
                        End Try
                    End Using

                    dbCargo.Dispose()

                    If IsNothing(ext) = False Then
                        Using View As ViewTableRecord = acDocEd.GetCurrentView()
                            'Dim vpt As ViewportTable = trLocal.GetObject(dbLocal.ViewportTableId, OpenMode.ForRead)
                            'Dim vptr As ViewportTableRecord = trLocal.GetObject(vpt("*Active"), OpenMode.ForWrite)

                            'Dim vptr As ViewTableRecord = New ViewTableRecord()
                            'Dim WCS2DCS As Matrix3d = Matrix3d.PlaneToWorld(View.ViewDirection)
                            'WCS2DCS = Matrix3d.Displacement(View.Target - Point3d.Origin) * WCS2DCS
                            'WCS2DCS = Matrix3d.Rotation(-View.ViewTwist, View.ViewDirection, View.Target) * WCS2DCS
                            'WCS2DCS = WCS2DCS.Inverse()

                            'MsgBox(ext.ToString())
                            'ext.TransformBy(WCS2DCS)
                            ext.TransformBy(acDocEd.CurrentUserCoordinateSystem.Inverse())

                            Dim pmin As Point2d = New Point2d(ext.MinPoint.X, ext.MinPoint.Y)
                            Dim pmax As Point2d = New Point2d(ext.MaxPoint.X, ext.MaxPoint.Y)
                            Dim Height As Double = pmax.Y - pmin.Y
                            Dim Width As Double = pmax.X - pmin.X

                            View.CenterPoint = pmin + ((pmax - pmin) / 2.0)
                            View.Height = Height
                            View.Width = Width

                            'Dim lower = New Double(3) {pmin.X, pmin.Y, pmin.Z};

                            'Dim upper As Double[]= New Double[3] { max.X, max.Y, max.Z };

                            ' MsgBox(Width)
                            acDocEd.SetCurrentView(View)

                            'Dim acadApp As Object = Application.AcadApplication

                            'AcadApp.ZoomExtents()
                            'app.zo
                            'AcadApp.ZoomWindow(pmin, pmax)
                        End Using
                    End If
                    trLocal.Commit()
                End Using
            End Using
        End Using
    End Sub

    Private Sub dgvActos_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvActos.CellContentClick
        Dim senderGrid = DirectCast(sender, DataGridView)
        'if (e.ColumnIndex == senderGrid.Columns["Opn"].Index && e.RowIndex >= 0)
        If TypeOf senderGrid.Columns(e.ColumnIndex) Is DataGridViewButtonColumn AndAlso e.RowIndex >= 0 Then
            Dim dato As Acto = dgvActos.Rows(e.RowIndex).DataBoundItem
            Dim clase As Clase = New Clase()
            Dim client = WCF.acto()
            Dim ruta As String = If(e.ColumnIndex = 1, dato.ruta_resolucion, dato.ruta_partida)
            Try
                If (e.ColumnIndex = 0) Then
                    Dim planos = client.buscar_planos(dato.codigo_informe, dato.anho_resolucion)
                    If (planos.error) Then
                        MsgBox(planos.mensaje, MsgBoxStyle.Exclamation, "Aviso")
                    Else
                        If planos.data.Length > 0 Then
                            For Each plano As String In planos.data
                                If File.Exists(plano) = True Then
                                    Process.Start(plano)
                                Else
                                    MsgBox("Archivo no existe:" + plano)
                                End If
                            Next
                        Else
                            MsgBox("No se encontraron planos..")
                        End If
                    End If
                Else
                    If String.IsNullOrEmpty(ruta) Then
                        MsgBox("Fila selecciona no contiene dato para abrir.")
                    Else
                        If File.Exists(ruta) = True Then
                            Process.Start(ruta)
                        Else
                            MsgBox("Archivo no existe:" + ruta)
                        End If
                    End If
                End If

            Catch ex As Exception
                client.Abort()
                Dim msg As String = "Error: " + ex.Message
                Dim rptaExc As AportePermisoService.Respuesta = clase.errorMsg(msg, "app", MyClass.Name, autocadVersion.ToString)
                MsgBox(rptaExc.mensaje, MsgBoxStyle.Critical, "Aviso")
            End Try
            'TODO - Button Clicked - Execute Code Here
        End If
    End Sub

    Private Sub dgvActos_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles dgvActos.CellPainting

        If e.RowIndex < 0 Then
            Return
        End If
        Dim dato As Acto = dgvActos.Rows(e.RowIndex).DataBoundItem
        'dgvSupervision.Rows(e.RowIndex).Cells("cus_propiedad_text").Value = If(dato.cus_propiedad < 0, Nothing, dato.cus_propiedad.ToString())
        'dgvSupervision.Rows(e.RowIndex).Cells("area_text").Value = If(dato.area < 0, Nothing, dato.area.ToString())



        If Not (String.IsNullOrEmpty(dato.cac_usuario)) Then
            dgvActos.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightGreen
        End If
        'I supposed your button column Is at index 0
        If (e.ColumnIndex = 0 Or e.ColumnIndex = 1 Or e.ColumnIndex = 2) Then
            e.Paint(e.CellBounds, DataGridViewPaintParts.All)
            Dim w = My.Resources.Resources.doc.Width
            Dim h = My.Resources.Resources.doc.Height
            Dim x = e.CellBounds.Left + (e.CellBounds.Width - w) / 2
            Dim y = e.CellBounds.Top + (e.CellBounds.Height - h) / 2

            e.Graphics.DrawImage(If(e.ColumnIndex = 0, My.Resources.Resources.autocad, (If(e.ColumnIndex = 1, My.Resources.Resources.doc, My.Resources.Resources.pdf))), New Rectangle(x, y, w, h))
            e.Handled = True
            dgvActos.Rows(e.RowIndex).Cells(e.ColumnIndex).ToolTipText = If(e.ColumnIndex = 0, "Plano", (If(e.ColumnIndex = 1, "Resolución", "Partida")))
        End If


    End Sub

    Private Sub dgvActos_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dgvActos.RowPostPaint
        Using b As SolidBrush = New SolidBrush(dgvActos.RowHeadersDefaultCellStyle.ForeColor)
            Dim pos = e.RowIndex + 1
            e.Graphics.DrawString(pos.ToString(System.Globalization.CultureInfo.CurrentUICulture), dgvActos.DefaultCellStyle.Font, b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 4)
        End Using
    End Sub

End Class