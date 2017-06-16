Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports System
Imports System.Timers
Imports System.IO
Imports CarteraPSE.clFunciones
Imports System.Xml
Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form1
    Public Nombre As String = ""
    Public Nombre2 As String = ""
    Public Carpeta As String = ""

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        DateTimePicker3.Value = DateTime.Now.ToString("dd/MM/yyyy")
        ComboBanco.Items.Add("")
        ComboBanco.Items.Add("PSE-Avisor")
        ComboBanco.Items.Add("Bancolombia")
        ComboBanco.Items.Add("Caja Social")
        ComboBanco.Items.Add("CDT")

        ComboTipo.Items.Add("")
        ComboTipo.Items.Add("Pension")
        ComboTipo.Items.Add("Matricula")

        Label1.Text = SociedadActual
    End Sub

    Private Sub ComboBanco_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBanco.SelectedIndexChanged
        TextBox1.Enabled = True
        If ComboBanco.SelectedIndex = 1 Then
            TextBox1.Enabled = True
            ComboTipo.Enabled = True
        Else
            TextBox1.Text = ""
            TextBox1.Enabled = False
            ComboTipo.Text = ""
            ComboTipo.Enabled = False
        End If
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        If ComboBanco.SelectedIndex <= 0 Then
            MsgBox("Debe seleccionar el archivo a generar", MsgBoxStyle.Exclamation, "Aviso")
            ComboBanco.Focus()
            Exit Sub
        End If

        'Dim FechaIni As String = FormatDateTime(DateTimePicker2.Value, DateFormat.ShortDate).Substring(0, 10)
        'Dim FechaFin As Date = FormatDateTime(DateTimePicker1.Value, DateFormat.ShortDate).Substring(0, 10)
        Dim FechaArchivo As Date = FormatDateTime(DateTimePicker3.Value, DateFormat.ShortDate).Substring(0, 10)
        Dim Consec As String = TextBox1.Text
        Try
            Dim Variables As String = "Declare @FechaArchivo As DateTime " & _
                                        "Declare @Valor As Numeric(19,6) " & _
                                        "Declare @Conteo As Int " & _
                                        "Declare @Consec As Nvarchar(2) " & _
                                        "Set Dateformat DMY " & _
                                        "Set @FechaArchivo = '" & FechaArchivo & "' " & _
                                        "Set @Consec = '" & Consec & "' "

            Dim NomArchi As String = ""
            Dim NomArchi2 As String = ""
            Dim Fecha = DateTime.Now.ToString("yyyyMMdd")
            Dim CadenaSQL As String = ""
            Dim CadenaSQL2 As String = ""
            Dim Extension As String = ""

            If ComboBanco.SelectedIndex = 1 Then
                NomArchi = "DET114770897"
                NomArchi2 = "USUARIOS114770897"
                Carpeta = "PSE\"
                If ComboTipo.SelectedIndex = 1 Then
                    CadenaSQL = Variables & My.Resources.CarteraPSE
                    CadenaSQL2 = Variables & My.Resources.UsuariosPSE
                ElseIf ComboTipo.SelectedIndex = 2 Then
                    CadenaSQL = Variables & My.Resources.CarteraPSE1
                    CadenaSQL2 = Variables & My.Resources.UsuariosPSE1
                End If
                Extension = ".txt"
            ElseIf ComboBanco.SelectedIndex = 2 Then
                NomArchi = "PROVEEDORES"
                Carpeta = "Pago Electronico\"
                CadenaSQL = Variables & My.Resources.Bancolombia
                Extension = ".FIL"
            ElseIf ComboBanco.SelectedIndex = 3 Then
                NomArchi = "BCS"
                Carpeta = "Pago Electronico\"
                CadenaSQL = Variables & My.Resources.CajaSocial
                Extension = ".txt"
            ElseIf ComboBanco.SelectedIndex = 4 Then
                NomArchi = "PAGO"
                Carpeta = "Pago Electronico\"
                CadenaSQL = Variables & My.Resources.CDT
                CadenaSQL2 = Variables & My.Resources.CDT2
                Extension = ".xlsx"
            End If

            Dim Filas1, Columnas1, Fila_Actual1, Columna_Actual1 As Integer
            Dim Filas2, Columnas2, Fila_Actual2, Columna_Actual2 As Integer
            Dim Texto1 As String
            Dim Texto2 As String
            Dim cnn As New SqlConnection
            Dim cnn2 As New SqlConnection

            If ComboBanco.SelectedIndex = 4 Then

                cnn = SetConectionSQL(SociedadActual)
                Dim da As New SqlDataAdapter(CadenaSQL, cnn)

                Dim ds As New DataSet
                da.Fill(ds)
                DataGridView1.DataSource = ds.Tables(0)
                
                Nombre = NomArchi & Fecha.ToString & Extension
                Nombre2 = NomArchi2 & Fecha.ToString & Extension
                Filas1 = DataGridView1.RowCount - 1
                Columnas1 = DataGridView1.ColumnCount - 1
                Fila_Actual1 = 0

                cnn2 = SetConectionSQL(SociedadActual)
                Dim da2 As New SqlDataAdapter(CadenaSQL2, cnn2)

                Dim ds2 As New DataSet
                da2.Fill(ds2)
                DataGridView2.DataSource = ds2.Tables(0)
                Nombre = NomArchi & Fecha.ToString & Extension
                Filas2 = DataGridView2.RowCount - 1
                Columnas2 = DataGridView2.ColumnCount - 1
                Fila_Actual2 = 0
            ElseIf ComboBanco.SelectedIndex = 1 Then

                cnn = SetConectionSQL(SociedadActual)
                Dim da As New SqlDataAdapter(CadenaSQL, cnn)

                Dim ds As New DataSet
                da.Fill(ds)
                DataGridView1.DataSource = ds.Tables(0)

                Nombre = NomArchi & Fecha.ToString & Extension
                Nombre2 = NomArchi2 & Fecha.ToString & Extension
                Filas1 = DataGridView1.RowCount - 1
                Columnas1 = DataGridView1.ColumnCount - 1
                Fila_Actual1 = 0

                cnn2 = SetConectionSQL(SociedadActual)
                Dim da2 As New SqlDataAdapter(CadenaSQL2, cnn2)

                Dim ds2 As New DataSet
                da2.Fill(ds2)
                DataGridView2.DataSource = ds2.Tables(0)

                Nombre2 = NomArchi2 & Fecha.ToString & Extension
                Filas2 = DataGridView2.RowCount - 1
                Columnas2 = DataGridView2.ColumnCount - 1
                Fila_Actual2 = 0

            Else
                cnn = SetConectionSQL(SociedadActual)
                Dim da As New SqlDataAdapter(CadenaSQL, cnn)

                Dim ds As New DataSet
                da.Fill(ds)
                DataGridView1.DataSource = ds.Tables(0)
                Nombre = NomArchi & Fecha.ToString & Extension
                Filas1 = DataGridView1.RowCount - 1
                Columnas1 = DataGridView1.ColumnCount - 1
                Fila_Actual1 = 0
            End If

            If DataGridView1.RowCount > 0 Then

                If ComboBanco.SelectedIndex = 4 Then

                    'Creamos las variables
                    Dim exApp As New Microsoft.Office.Interop.Excel.Application
                    Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
                    Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet
                    Dim exHoja2 As Microsoft.Office.Interop.Excel.Worksheet

                    'Añadimos el Libro al programa, y la hoja al libro
                    exLibro = exApp.Workbooks.Add
                    exHoja = exLibro.Sheets.Add()
                    exHoja = CType(exLibro.Sheets(1), Excel.Worksheet)
                    exHoja2 = CType(exLibro.Sheets(2), Excel.Worksheet)

                    'Añadimos la hoja al libro
                    GridAExcel(DataGridView1, DataGridView2, exApp, exLibro, exHoja, exHoja2)

                    Me.Cursor = Cursors.Default

                    'Aplicación visible
                    exApp.Application.Visible = False
                    exLibro.SaveAs("C:\SAP\" & Carpeta & Nombre)

                    exApp.Quit()
                    exLibro = Nothing
                    exApp = Nothing
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(exHoja)

                ElseIf ComboBanco.SelectedIndex = 1 Then
                    Dim strm1 As New StreamWriter("C:\SAP\" & Carpeta & Nombre)
                    While Fila_Actual1 < Filas1
                        Texto1 = ""
                        Columna_Actual1 = 0
                        Texto1 = Texto1 + (DataGridView1(0, Fila_Actual1).Value)
                        strm1.WriteLine(Texto1)
                        Fila_Actual1 = Fila_Actual1 + 1
                    End While
                    strm1.Close()
                    Dim strm2 As New StreamWriter("C:\SAP\" & Carpeta & Nombre2)
                    While Fila_Actual2 < Filas2
                        Texto2 = ""
                        Columna_Actual2 = 0
                        Texto2 = Texto2 + (DataGridView2(0, Fila_Actual2).Value)
                        strm2.WriteLine(Texto2)
                        Fila_Actual2 = Fila_Actual2 + 1
                    End While
                    strm2.Close()
                    'ElseIf ComboBanco.SelectedIndex = 2 Then
                    '    Dim strm1 As New StreamWriter("C:\SAP\" & Carpeta & Nombre)
                    '    Texto1 = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"
                    '    While Fila_Actual1 < Filas1
                    '        Texto1 = Texto1 + (DataGridView1(0, Fila_Actual1).Value)
                    '        Fila_Actual1 = Fila_Actual1 + 1
                    '    End While
                    '    Texto1 = Replace(Texto1, (">"), (">" & vbCrLf))
                    '    strm1.WriteLine(Texto1)
                    '    strm1.Close()
                Else
                    Dim strm1 As New StreamWriter("C:\SAP\" & Carpeta & Nombre)
                    While Fila_Actual1 < Filas1
                        Texto1 = ""
                        Columna_Actual1 = 0
                        Texto1 = Texto1 + (DataGridView1(0, Fila_Actual1).Value)
                        strm1.WriteLine(Texto1)
                        Fila_Actual1 = Fila_Actual1 + 1
                    End While
                    strm1.Close()
                End If

                MsgBox("Archivo generado con éxito")
            Else
                MessageBox.Show("No existen datos en el rango de fechas seleccionado")
            End If

            Me.Cursor = Cursors.Default
            DataGridView1.DataSource.Clear()
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
    End Sub
    Function GridAExcel(ByVal ElGrid As DataGridView, ByVal ElGrid2 As DataGridView, exApp As Microsoft.Office.Interop.Excel.Application _
                            , exLiibro As Microsoft.Office.Interop.Excel.Workbook _
                            , exHoja2 As Microsoft.Office.Interop.Excel.Worksheet, exHoja As Microsoft.Office.Interop.Excel.Worksheet
                        ) As Boolean

        Try
            ' ¿Cuantas columnas y cuantas filas?
            Dim NCol As Integer = ElGrid.ColumnCount
            Dim NRow As Integer = ElGrid.RowCount
            Dim NCol2 As Integer = ElGrid2.ColumnCount
            Dim NRow2 As Integer = ElGrid2.RowCount

            'Aqui recorremos todas las filas, y por cada fila todas las columnas y vamos escribiendo.
            For i As Integer = 1 To NCol
                exHoja.Cells.Item(1, i) = ElGrid.Columns(i - 1).Name.ToString
            Next

            For Fila As Integer = 0 To NRow - 1
                For Col As Integer = 0 To NCol - 1
                    exHoja.Cells.Item(Fila + 2, Col + 1) = ElGrid.Rows(Fila).Cells(Col).Value
                Next
            Next
            exHoja.Columns.AutoFit()

            For i As Integer = 1 To NCol2
                exHoja2.Cells.Item(1, i) = ElGrid2.Columns(i - 1).Name.ToString
            Next

            For Fila As Integer = 0 To NRow2 - 1
                For Col As Integer = 0 To NCol2 - 1
                    exHoja2.Cells.Item(Fila + 2, Col + 1) = ElGrid2.Rows(Fila).Cells(Col).Value
                Next
            Next
            exHoja2.Columns.AutoFit()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al exportar a Excel")
            Return False
        End Try
    End Function
End Class