Imports System.Data
Imports System.Data.SqlClient
Imports System
Imports System.Timers
Imports System.IO
Imports CarteraPSE.clFunciones
Imports System.Text
Imports System.Data.DataRow
Imports System.Linq
Imports System.Linq.Expressions
Imports System.Collections.Generic
Imports System.Globalization
Imports CarteraPSE.MDIParentPin
Public Class Pagos
    Public Shared Cuenta As String
    Public Shared CuentaAnti As String
    Public Shared Serie As String
    Public Shared TipoFac As String
    Public Shared FechaLog As String
    Public connection As SqlConnection

    'Recupera las cuentas de nivel 5
    Public Function TraerCuentas()
        Dim sQuery As String
        Dim BaseSelec As String = SociedadActual
        sQuery = "Select Cast(Acctcode As Nvarchar (40)) + ' - ' + Cast(AcctName As Nvarchar (500)) As Cuenta From OACT Where Levels = '5' AND U_IncluirAppCartera = 'SI' "
        connection = SetConectionSQL(BaseSelec)
        Dim command As New SqlCommand(sQuery, connection)
        command.Connection.Open()
        Dim reader As SqlDataReader = command.ExecuteReader()
        Try
            If reader.HasRows = True Then
                ComboCuenta.Items.Clear()
                ComboCuentaAnti.Items.Clear()
                While reader.Read()
                    ComboCuenta.Items.Add(reader("Cuenta"))
                    ComboCuentaAnti.Items.Add(reader("Cuenta"))
                End While
            End If
        Finally
            ' Always call Close when done reading.
            reader.Close()
        End Try
    End Function
    'Recupera las series para los documentos de pagos recibidos
    Public Function TraerSeries()
        Dim sQuery As String
        Dim BaseSelec As String = SociedadActual
        sQuery = "Select Cast(Series As Nvarchar(20)) + ' - ' + SeriesName As Series From NNM1 Where ObjectCode = '24'"
        connection = SetConectionSQL(BaseSelec)
        Dim command As New SqlCommand(sQuery, connection)
        command.Connection.Open()
        Dim reader As SqlDataReader = command.ExecuteReader()
        Try
            If reader.HasRows = True Then
                ComboSerie.Items.Clear()
                While reader.Read()
                    ComboSerie.Items.Add(reader("Series"))
                End While
            End If
        Finally
            ' Always call Close when done reading.
            reader.Close()
        End Try
    End Function

    Private Sub Pagos_Click(sender As Object, e As EventArgs) Handles Me.Click

    End Sub
    'Carga los campos
    Private Sub Pagos_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        ComboArchivo.Items.Add("")
        ComboArchivo.Items.Add("Colpatria")
        ComboArchivo.Items.Add("Avisor")
        ComboTipo.Items.Add("")
        ComboTipo.Items.Add("Pension")
        ComboTipo.Items.Add("Matricula")
        ComboTipo.Items.Add("Arrirndo/Hipoteca")
        Cuenta = ""
        CuentaAnti = ""
        Using connection
            TraerCuentas()
            TraerSeries()
        End Using
    End Sub
    'Habilita campo de seleccion se archivos
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim folderDlg As New OpenFileDialog()
        folderDlg.InitialDirectory = "c:\"
        folderDlg.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
        folderDlg.FilterIndex = 2
        folderDlg.RestoreDirectory = True
        If folderDlg.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            TextBox1.Text = folderDlg.FileName()
        End If
    End Sub
    Private Sub Button2_Click_1(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If ComboCuenta.Text = "" Then
            MsgBox("Debe seleccionar la cuenta de recaudo", MsgBoxStyle.Exclamation, "Aviso")
            ComboCuenta.Focus()
            Exit Sub
        ElseIf ComboCuentaAnti.Text = "" Then
            MsgBox("Debe seleccionar la cuenta de anticipos", MsgBoxStyle.Exclamation, "Aviso")
            ComboCuentaAnti.Focus()
            Exit Sub
        ElseIf ComboSerie.Text = "" Then
            MsgBox("Debe seleccionar la serie de numeracion", MsgBoxStyle.Exclamation, "Aviso")
            ComboCuentaAnti.Focus()
            Exit Sub
        ElseIf ComboTipo.Text = "" Then
            MsgBox("Debe seleccionar el tipo de pago", MsgBoxStyle.Exclamation, "Aviso")
            ComboTipo.Focus()
            Exit Sub
        ElseIf TextBox1.Text = "" Then
            MsgBox("Debe seleccionar el archivo", MsgBoxStyle.Exclamation, "Aviso")
            TextBox1.Focus()
            Exit Sub
        End If

        Cursor = Cursors.WaitCursor

        Serie = ComboSerie.Text
        dtPagos.Rows.Clear()
        dtPagos.Clear()
        Dim Count As Integer = 0
        Dim CountFail As Integer = 0
        Dim Directorio1 As String = TextBox1.Text
        Try
            ' Leer el contenido del archivo seleccionado mientras no se llegue al final
            Dim objFSO1 = CreateObject("Scripting.FileSystemObject")
            Dim objFile1 = objFSO1.OpenTextFile(Directorio1.ToString)
            Dim strText As String = objFile1.ReadAll 'o ReadLine

            Dim lector As New IO.StreamReader(Directorio1)

            Dim ArrLineaArchivo As String()
            Dim relation As DataRow

            While lector.Peek() <> -1

                Dim linea As String = lector.ReadLine()
                Dim CodAlum As Long = 0
                Dim ValorPago As Double
                Dim FechaPago As String
                Dim NumFac As Integer = 0
                Dim TipoEmpresa As Integer = 0


                If ComboArchivo.SelectedIndex = 1 Then
                    'Estructura archivo recaudo Colpatria
                    If Mid(linea, 1, 2) = "01" Then
                        FechaPago = Mid(linea, 13, 8)
                        TipoFac = Mid(linea, 30, 4)
                    End If
                End If

                If ComboArchivo.SelectedIndex = 1 Then
                    If Mid(linea, 1, 2) = "06" Then
                        CodAlum = QuitarCeros(Mid(linea, 38, 13))
                        ValorPago = CDbl(Val(PonerDecimal(QuitarCeros(Mid(linea, 51, 14)))))
                        NumFac = 0

                        'Almacena los datos recuperados en la Tabla
                        If ComboTipo.SelectedIndex = 3 Then
                            TipoEmpresa = QuitarCeros(Mid(linea, 36, 2))
                            ArrLineaArchivo = {FechaPago, CodAlum, ValorPago, TipoFac, NumFac, TipoEmpresa}
                        Else
                            ArrLineaArchivo = {FechaPago, CodAlum, ValorPago, TipoFac, NumFac, TipoEmpresa}
                        End If
                        relation = dtPagos.NewRow()
                        relation.ItemArray = ArrLineaArchivo
                        dtPagos.Rows.Add(relation)
                    End If
                ElseIf ComboArchivo.SelectedIndex = 2 Then
                    'Estructura archivo recaudo Avisor
                    FechaPago = ""
                    If Mid(linea, 1, 1) = "1" Then
                        GoTo NextLine
                    ElseIf Mid(linea, 1, 1) = "2" Then
                        CodAlum = Replace(QuitarCeros(Mid(linea, 303, 80)), " ", "")
                        NumFac = 0 'QuitarCeros(Mid(linea, 233, 80))
                        ValorPago = CDbl(Val(PonerDecimal(QuitarCeros(Mid(linea, 107, 18)))))
                        FechaPago = Mid(linea, 433, 8)

                        'Almacena los datos recuperados en la Tabla
                        ArrLineaArchivo = {FechaPago, CodAlum, ValorPago, TipoFac, NumFac}
                        relation = dtPagos.NewRow()
                        relation.ItemArray = ArrLineaArchivo
                        dtPagos.Rows.Add(relation)

                    End If
                End If
NextLine:
            End While

            'For Each DataRow para Validacion Archivo
            Dim Fecha = DateTime.Now.ToString("yyyyMMdd hh mm ss")
            Dim strm1 As New StreamWriter("C:\SAP\Pagos Banco\Log Pagos\Log" & Fecha.ToString & ".txt")
            Dim Texto1 As String = ""
            objFile1.Close()
            lector.Close()

            Texto1 = SociedadActual & vbCrLf
            Texto1 = Texto1 & "------------------VALIDACION ARCHIVO------------------" & vbCrLf & vbCrLf
            Texto1 = Texto1 & "Codigo" & vbTab & "Valor" & vbTab & "Fecha" & vbTab & "Comentarios" & vbCrLf

            For Each row As DataRow In dtPagos.Rows

                Dim VarValor As Double
                Dim VarAlumn As Long
                Dim FechaPago As String
                Dim Comentario As String
                Dim rs As SAPbobsCOM.Recordset

                FechaPago = row("FechaPago")
                VarValor = row("ValorPago")
                VarAlumn = row("CodAlum")

                Dim CadenaSQL As String
                CadenaSQL = ""
                CadenaSQL = "Select Case When Count(Code) >= 1 Then 'OK' Else 'No existe el alumno' End Comentario From [@SEI_ALUMNOS] Where Code = '" & VarAlumn & "'"
                If ComboTipo.SelectedIndex = 3 Then
                    CadenaSQL = "Select Case When Count(AddID) >= 1 Then 'OK' Else 'No existe el SN' End Comentario From OCRD Where AddID = '" & VarAlumn & "'"
                End If


                rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rs.DoQuery(CadenaSQL)
                If Not (rs.EoF) Then
                    Comentario = rs.Fields.Item(0).Value
                    rs.MoveNext()
                End If

                Texto1 = Texto1 & VarAlumn & vbTab & VarValor & vbTab & FechaPago & vbTab & Comentario & vbCrLf

            Next
            Texto1 = Texto1 & vbCrLf & vbCrLf & "------------------VALIDACION CREACION PAGOS----------------" & vbCrLf

            'For Each DataRow realizar pagos
            For Each row As DataRow In dtPagos.Rows

                Dim objPay As SAPbobsCOM.Payments
                Dim VarValor As Double
                Dim VarAlumn As Long
                Dim VarTipoEmpresa As Integer
                Dim FechaPago As String
                Dim NumFac As Integer
                Dim TipoFac As String
                Dim dValor As Double
                Dim status As Object
                Dim errorcode As Integer
                Dim errordesc As String
                Dim ValorApli As String

                status = Nothing
                errorcode = Nothing
                errordesc = Nothing
                ValorApli = Nothing
                dValor = Nothing

                Dim cnn As New SqlConnection
                Dim ds As New DataSet
                Dim CadenaSQL As String
                cnn = SetConectionSQL(SociedadActual)

                FechaPago = row("FechaPago")
                VarValor = row("ValorPago") ''Val(Resdata02.valor)
                VarAlumn = row("CodAlum") ''Resdata02.idpago
                If ComboTipo.SelectedIndex = 3 Then
                    VarTipoEmpresa = row("TipoEmpresa")
                End If
                'TipoFac = row("TipoFac")
                'NumFac = row("NumFac")

                ValorApli = VarValor
                dValor = CDbl(Val(ValorApli))
                Cuenta = ComboCuenta.Text
                CuentaAnti = ComboCuentaAnti.Text
                Serie = ComboSerie.Text

                'Traer CardCode segun alumno
                Dim ValidaCliente As String
                Dim Cliente As String = ""
                If ComboTipo.SelectedIndex = 3 Then
                    ValidaCliente = "SELECT CardCode FROM OCRD WHERE AddID = '" & VarAlumn & "'"
                Else
                    ValidaCliente = "Select Top 1 IsNull(U_SEITitul,''), 0,'',0 , 0 As DocNum From [@SEI_CONC_ALUMN] Where Code = '" & VarAlumn & "'"
                End If

                Dim Cli As SAPbobsCOM.Recordset
                Cli = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Cli.DoQuery(ValidaCliente)
                If Not (Cli.EoF) Then
                    Cliente = Cli.Fields.Item(0).Value
                    Cli.MoveNext()
                End If

                'Validacion Tipo de pago
                If ComboTipo.SelectedIndex = 1 Then
                    TipoFac = "1"
                ElseIf ComboTipo.SelectedIndex = 2 Then
                    TipoFac = "2"
                ElseIf ComboTipo.SelectedIndex = 3 Then
                    TipoFac = "3"
                Else
                    TipoFac = "2"
                End If

                'Consulta para recuperar los documentos pendientes por pagar
                cnn = SetConectionSQL(SociedadActual)
                CadenaSQL = ""
                If ComboArchivo.SelectedIndex = 1 Then
                    CadenaSQL = "Select CardCode, TransId, ObjType, Saldo From (Select * From (Select A2.CardCode, A0.TransId, A0.ObjType, Sum(IsNull(A1.Debit,0))- Case When A4.Canceled = 'Y' Then 0 Else IsNull(Sum(A3.SumApplied),0) End 'Saldo', A2.DocDate, 1 'Orden' From OJDT A0 Inner Join JDT1 A1 On A0.TransId = A1.TransId Left Join OINV A2 On A0.Ref1 = A2.DocNum Left Join RCT2 A3 On A0.TransId = A3.DocEntry And A0.ObjType = A3.InvType Left Join ORCT A4 On A3.DocNum = A4.DocEntry Where A0.TransCode = 'INT' And A0.Ref2 = '" & VarAlumn & "' And A2.U_SEI_FACTU = '" & TipoFac & "' And IsNull(A0.StornoToTr,'') = '' And A1.BalDueDeb > 0 Group By A2.CardCode, A0.TransId, A0.ObjType, A2.DocDate, A4.Canceled) C0 Where Saldo > 0 And TransId Not In (Select StornoToTr From OJDT Where IsNull(StornoToTr,'') <> '') Union All Select A0.CardCode, A0.DocEntry, A0.ObjType, Cast((A0.DocTotal-IsNull(A0.PaidToDate,0)) As Numeric(19,2))'Saldo', A0.DocDate, 2 'Orden' From OINV A0 Inner Join INV1 A1 On A0.DocEntry = A1.DocEntry Where A0.DocStatus = 'O' And (A0.DocTotal-IsNull(A0.PaidToDate,0)) > 0 And A0.NumAtCard = '" & VarAlumn & "' And A0.U_SEI_FACTU = '" & TipoFac & "' And A1.TargetType != '14' Group By CardCode, A0.DocEntry, A0.ObjType, A0.DocTotal, A0.PaidToDate, A0.DocDate Union All Select * From (Select A0.CardCode, A0.DocEntry, A0.ObjType, Cast((A0.DocTotal-IsNull(Sum(A1.SumApplied),0)) As Numeric(19,2))*-1 'Saldo', A0.DocDate , 3 'Orden' From ORIN A0 Left Join RCT2 A1 On A0.DocEntry = A1.DocEntry And A0.ObjType = A1.InvType Where A0.DocStatus = 'O' And A0.NumAtCard = '" & VarAlumn & "' Group By A0.CardCode, A0.DocEntry, A0.ObjType, A0.DocTotal, A0.DocDate) B0 Where Saldo < 0 Union All Select Top 1 IsNull(U_SEITitul,''), 0,'',0 ,GETDATE() , 4 'Orden'  From [@SEI_CONC_ALUMN] Where Code = '" & VarAlumn & "')P0 WHERE YEAR(DocDate) > 2015 Order By Orden, DocDate, Saldo desc"
                ElseIf ComboArchivo.SelectedIndex = 2 Then
                    CadenaSQL = "Select CardCode, TransId, ObjType, Saldo From (Select A0.CardCode, A0.DocEntry'TransId', A0.ObjType, Cast((A0.DocTotal-IsNull(A0.PaidToDate,0)) As Numeric(19,2))'Saldo', A0.DocDate, 1 'Orden' From OINV A0 Where A0.DocStatus = 'O' And (A0.DocTotal-IsNull(A0.PaidToDate,0)) > 0 And A0.DocNum = '" & NumFac & "' And A0.NumAtCard = '" & VarAlumn & "' Union All Select * From (Select A2.CardCode, A0.TransId, A0.ObjType, Sum(IsNull(A1.Debit,0))- Case When A4.Canceled = 'Y' Then 0 Else IsNull(Sum(A3.SumApplied),0) End 'Saldo', A2.DocDate, 1 'Orden' From OJDT A0 Inner Join JDT1 A1 On A0.TransId = A1.TransId Left Join OINV A2 On A0.Ref1 = A2.DocNum Left Join RCT2 A3 On A0.TransId = A3.DocEntry And A0.ObjType = A3.InvType Left Join ORCT A4 On A3.DocNum = A4.DocEntry Where A0.TransCode = 'INT' And A0.Ref2 = '" & VarAlumn & "' And A2.U_SEI_FACTU = '" & TipoFac & "' And IsNull(A0.StornoToTr,'') = ''And A1.BalDueDeb > 0 Group By A2.CardCode, A0.TransId, A0.ObjType, A2.DocDate, A4.Canceled) C0 Where Saldo > 0 And TransId Not In (Select StornoToTr From OJDT Where IsNull(StornoToTr,'') <> '') Union All Select A0.CardCode, A0.DocEntry, A0.ObjType, Cast((A0.DocTotal-IsNull(A0.PaidToDate,0)) As Numeric(19,2))'Saldo' , A0.DocDate, 3 'Orden' From OINV A0 Inner Join INV1 A1 On A0.DocEntry = A1.DocEntry Where A0.DocStatus = 'O' And (A0.DocTotal-IsNull(A0.PaidToDate,0)) > 0 And A0.NumAtCard = '" & VarAlumn & "' And A0.U_SEI_FACTU = '" & TipoFac & "'  And A1.TargetType != '14' Group By CardCode, A0.DocEntry, A0.ObjType, A0.DocTotal, A0.PaidToDate, A0.DocDate Union All Select * From (Select A0.CardCode, A0.DocEntry, A0.ObjType, Cast((A0.DocTotal-IsNull(Sum(A1.SumApplied),0)) As Numeric(19,2))*-1 'Saldo', A0.DocDate, 4 'Orden' From ORIN A0 Left Join RCT2 A1 On A0.DocEntry = A1.DocEntry And A0.ObjType = A1.InvType Where A0.DocStatus = 'O' And A0.NumAtCard = '" & VarAlumn & "' Group By A0.CardCode, A0.DocEntry, A0.ObjType, A0.DocTotal, A0.DocDate) B0 Where Saldo > 0 Union All Select Top 1 IsNull(U_SEITitul,''), 0,'',0, GetDate(), 5 'Orden'  From [@SEI_CONC_ALUMN] Where Code = '" & VarAlumn & "')P0 WHERE YEAR(DocDate) > 2015 Order By Orden, DocDate, Saldo desc"
                End If

                If TipoFac = "3" Then

                    If VarTipoEmpresa = 8 Then
                        CadenaSQL = "SELECT * FROM (" & _
                               "SELECT * FROM (" & _
                               "SELECT CardCode, DocEntry, ObjType, (DocTotal - PaidToDate) AS Saldo, 0 Orden, 0 InstlmntID, DocDueDate, DocDate  " & _
                               "FROM OINV WHERE CardCode = '" & Cliente & "' AND DocType = 'S' AND DocStatus = 'O' " & _
                               "UNION " & _
                               "SELECT T0.CardCode, T0.DocEntry, T0.ObjType, (T1.InsTotal - T1.PaidToDate) AS Saldo, 3 Orden, T1.InstlmntID, T0.DocDueDate, T0.DocDate  FROM OINV T0 " & _
                               "INNER JOIN INV6 T1 ON T1.DocEntry = T0.DocEntry  WHERE T0.CardCode = '" & Cliente & "' AND T0.DocType = 'I' AND T1.Status = 'O'  " & _
                               ") P0 " & _
                               "WHERE YEAR(P0.DocDueDate) < 2016 " & _
                               "UNION " & _
                               "SELECT * FROM (" & _
                               "SELECT CardCode, DocEntry, ObjType, (DocTotal - PaidToDate) AS Saldo, 1 Orden, 0 InstlmntID, DocDueDate, DocDate  " & _
                               "FROM OINV WHERE CardCode = '" & Cliente & "' AND DocType = 'S' AND DocStatus = 'O' AND U_SaldoInicial = 'Y' " & _
                               "UNION " & _
                               "SELECT T0.CardCode, T0.DocEntry, T0.ObjType, (T1.InsTotal - T1.PaidToDate) AS Saldo, 4 Orden, T1.InstlmntID, T0.DocDueDate, T0.DocDate  FROM OINV T0 " & _
                               "INNER JOIN INV6 T1 ON T1.DocEntry = T0.DocEntry  WHERE T0.CardCode = '" & Cliente & "' AND T0.DocType = 'I' AND T1.Status = 'O' AND T0.U_SaldoInicial = 'Y' " & _
                               ") P0 " & _
                               "WHERE P0.DocDueDate = '20160229' " & _
                               "UNION " & _
                               "SELECT * FROM (" & _
                               "SELECT CardCode, DocEntry, ObjType, (DocTotal - PaidToDate) AS Saldo, 2 Orden, 0 InstlmntID, DocDueDate, DocDate  " & _
                               "FROM OINV WHERE CardCode = '" & Cliente & "' AND DocType = 'S' AND DocStatus = 'O' AND (U_SaldoInicial IS NULL  OR U_SaldoInicial = 'N')  " & _
                               "UNION " & _
                               "SELECT T0.CardCode, T0.DocEntry, T0.ObjType, (T1.InsTotal - T1.PaidToDate) AS Saldo, 5 Orden, T1.InstlmntID, T0.DocDueDate, T0.DocDate  FROM OINV T0 " & _
                               "INNER JOIN INV6 T1 ON T1.DocEntry = T0.DocEntry  WHERE T0.CardCode = '" & Cliente & "' AND T0.DocType = 'I' AND T1.Status = 'O' AND (T0.U_SaldoInicial IS NULL OR T0.U_SaldoInicial = 'N')   " & _
                               ") P0 " & _
                               "WHERE YEAR(P0.DocDueDate) > 2015  " & _
                               ") P1 WHERE P1.DocDate <= '" & FechaPago.Substring(0, 4) & FechaPago.Substring(4, 2) & FechaPago.Substring(6, 2) & "' ORDER BY P1.Orden , P1.DocDueDate  "

                    End If

                    If VarTipoEmpresa = 9 Then
                        CadenaSQL = "SELECT CardCode, DocEntry, ObjType, (DocTotal - PaidToDate) AS Saldo, 1 Orden, 0 InstlmntID, DocDueDate  " & _
                                  "FROM OINV WHERE CardCode = '" & Cliente & "' AND DocStatus = 'O' AND DocDate <= '" & FechaPago.Substring(0, 4) & FechaPago.Substring(4, 2) & FechaPago.Substring(6, 2) & "' ORDER BY DocDueDate "
                    End If

                   
                End If

                Dim da As New SqlDataAdapter(CadenaSQL, cnn)

                da.Fill(ds)
                DataGridView1.DataSource = ds.Tables(0)
                Dim ClienteFac As String
                Dim DocEntry As Integer
                Dim Cuota As Integer
                Dim TipoDeuda As Integer
                Dim ObjType As String
                Dim Saldo As Double
                ClienteFac = ""
                DocEntry = 0
                Cuota = 0
                TipoDeuda = 0
                ObjType = ""
                Saldo = 0
                Dim Filas1, Columnas1, Fila_Actual1, Columna_Actual1 As Integer
                Filas1 = DataGridView1.RowCount
                Columnas1 = DataGridView1.ColumnCount

                Fila_Actual1 = 0

                Cuenta = ComboCuenta.Text
                Serie = ComboSerie.Text

                '------------------------------------------------------------------Cabecera Pago Doc y abono cuenta------------------------------------------------------------------
                objPay = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                objPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
                objPay.DocDate = FechaPago.Substring(0, 4) & "/" & FechaPago.Substring(4, 2) & "/" & FechaPago.Substring(6, 2)
                objPay.Remarks = VarAlumn
                objPay.Series = Replace(Mid(Serie, 1, InStr(Serie, "-") - 1), " ", "")
                objPay.CardCode = Cliente
                objPay.TransferDate = FechaPago.Substring(0, 4) & "/" & FechaPago.Substring(4, 2) & "/" & FechaPago.Substring(6, 2)
                objPay.TransferSum = dValor
                objPay.TransferAccount = Mid(Cuenta, 1, 8)
                objPay.TransferReference = VarAlumn
                If TipoFac = "3" Then

                Else
                    objPay.CounterReference = VarAlumn
                    objPay.Reference2 = VarAlumn
                End If
                objPay.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
                objPay.JournalRemarks = "Pagos recibidos-" & Cliente

                'Lineas para el documento 
                If DataGridView1.RowCount > 0 Then
                    While Fila_Actual1 < Filas1
                        Columna_Actual1 = 0
                        ClienteFac = (DataGridView1(0, Fila_Actual1).Value)
                        DocEntry = (DataGridView1(1, Fila_Actual1).Value)
                        ObjType = (DataGridView1(2, Fila_Actual1).Value)
                        Saldo = (DataGridView1(3, Fila_Actual1).Value)
                        'DocNum = (DataGridView1(4, Fila_Actual1).Value)
                        If TipoFac = "3" Then
                            If VarTipoEmpresa = 8 Then
                                TipoDeuda = (DataGridView1(4, Fila_Actual1).Value)
                                Cuota = (DataGridView1(5, Fila_Actual1).Value)
                            End If
                        End If

                        If DocEntry <> "0" And CDbl(Val(VarValor)) > 0 Then

                            If VarValor = 0 Then
                                GoTo NextPago
                            ElseIf VarValor >= Saldo Then
                                ValorApli = Saldo
                            Else
                                ValorApli = VarValor
                            End If

                            'dValor = CDbl(Val(ValorApli))
                            dValor = ValorApli
                            objPay.Invoices.InvoiceType = ObjType
                            objPay.Invoices.DocEntry = DocEntry
                            objPay.Invoices.SumApplied = dValor


                            If TipoFac = "3" Then
                                If VarTipoEmpresa = 8 And TipoDeuda = 2 Then
                                    objPay.Invoices.InstallmentId = Cuota
                                End If
                            End If

                            If ObjType = "13" Then
                                objPay.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
                            ElseIf ObjType = "14" Then
                                objPay.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_CredItnote
                            Else
                                objPay.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_JournalEntry
                            End If
                            VarValor = CDbl(Val(VarValor)) - CDbl(Val(dValor))

                            objPay.Invoices.Add()
                        Else

                        End If
                        Fila_Actual1 = Fila_Actual1 + 1
                    End While

                    'Validacion de abono a cuenta
                    If CDbl(Val(VarValor)) > 0 Then
                        objPay.ControlAccount = Mid(CuentaAnti, 1, 8)
                    End If

                    status = objPay.Add()

                    'Escribir log Pagos
                    If (status <> 0) Then
                        CountFail = CountFail + 1
                        oCompany.GetLastError(errorcode, errordesc)
                        'MsgBox("Error: " & errorcode & " - " & errordesc)
                        FechaLog = ""
                        FechaLog = DateTime.Now().ToString("dd/MM/yyyy hh:mm:ss tt")
                        Texto1 = Texto1 & vbCrLf & (FechaLog & vbTab & "Error Pago: " & errorcode & " - " & errordesc & " Alumno: " & VarAlumn)
                    Else
                        Count = Count + 1
                        Dim NewDoc As String = ""
                        oCompany.GetNewObjectCode(NewDoc)
                        FechaLog = ""
                        FechaLog = DateTime.Now().ToString("dd/MM/yyyy hh:mm:ss tt")
                        If ActualizarAsientoPago(NewDoc) Then
                            If TipoFac = "3" Then
                                Texto1 = Texto1 & vbCrLf & (FechaLog & vbTab & "Numero Interno del Pago: " & NewDoc & " Codigo Adicional SN: " & VarAlumn)
                            Else
                                Texto1 = Texto1 & vbCrLf & (FechaLog & vbTab & "Numero Interno del Pago: " & NewDoc & " Valor Pagado: " & ValorApli & " Alumno: " & VarAlumn)
                            End If

                        Else
                            Texto1 = Texto1 & vbCrLf & (FechaLog & vbTab & "Numero Interno del Pago: " & NewDoc & " Valor Pagado: " & ValorApli & " Alumno: " & VarAlumn & " --- ERROR ACTUALIZACION DE TERCERO EN ASIENTO CONTABLE DEL PAGO")
                        End If
                    End If
                Else

                    'Validacion de abono a cuenta
                    If CDbl(Val(VarValor)) > 0 Then
                        objPay.ControlAccount = Mid(CuentaAnti, 1, 8)
                    End If

                    status = objPay.Add()

                    'Escribir log Pagos
                    If (status <> 0) Then
                        CountFail = CountFail + 1
                        oCompany.GetLastError(errorcode, errordesc)
                        'MsgBox("Error: " & errorcode & " - " & errordesc)
                        FechaLog = ""
                        FechaLog = DateTime.Now().ToString("dd/MM/yyyy hh:mm:ss tt")
                        Texto1 = Texto1 & vbCrLf & (FechaLog & vbTab & "Error Pago: " & errorcode & " - " & errordesc & " Alumno: " & VarAlumn)
                    Else
                        Count = Count + 1
                        Dim NewDoc As String = ""
                        oCompany.GetNewObjectCode(NewDoc)
                        FechaLog = ""
                        FechaLog = DateTime.Now().ToString("dd/MM/yyyy hh:mm:ss tt")
                        If ActualizarAsientoPago(NewDoc) Then
                            If TipoFac = "3" Then
                                Texto1 = Texto1 & vbCrLf & (FechaLog & vbTab & "Numero Interno del Pago: " & NewDoc & " Codigo Adicional SN: " & VarAlumn)
                            Else
                                Texto1 = Texto1 & vbCrLf & (FechaLog & vbTab & "Numero Interno del Pago: " & NewDoc & " Valor Pagado: " & ValorApli & " Alumno: " & VarAlumn)
                            End If

                        Else
                            Texto1 = Texto1 & vbCrLf & (FechaLog & vbTab & "Numero Interno del Pago: " & NewDoc & " Valor Pagado: " & ValorApli & " Alumno: " & VarAlumn & " --- ERROR ACTUALIZACION DE TERCERO EN ASIENTO CONTABLE DEL PAGO")
                        End If
                    End If


                    GoTo NextPago
                End If
NextPago:
            Next
            Texto1 = Texto1 & vbCrLf & "Numero de pagos realizados: " & Count & " Numero de pagos fallidos: " & CountFail
            strm1.WriteLine(Texto1)
            strm1.Close()

            Cursor = Cursors.Default

            MsgBox("Se crearon " & Count & " pagos")
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
    End Sub

    Public Function ActualizarAsientoPago(ByRef docEntry As String) As Boolean

        Dim oRs As SAPbobsCOM.Recordset
        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sql As String = ""
        sql = "SELECT DocNum, CounterRef, CardCode FROM ORCT WHERE DocEntry  = '" & docEntry & "'"
        Dim DocNum As String = ""
        Dim Referencia As String = ""
        Dim Tercero As String = ""
        oRs.DoQuery(sql)
        While Not oRs.EoF
            DocNum = oRs.Fields.Item(0).Value
            Referencia = oRs.Fields.Item(1).Value
            Tercero = oRs.Fields.Item(2).Value
            oRs.MoveNext()
        End While


        Dim Cli As SAPbobsCOM.Recordset
        Cli = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sql1 As String = ""
        sql1 = "UPDATE JDT1 SET U_HBT_Tercero = '" & Tercero & "' WHERE Ref2 = '" & Referencia & "' AND Ref1 = '" & DocNum & "'"
        Try
            Cli.DoQuery(sql1)
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    'Funcion para quitar ceros a la izquierda de una cadena
    Public Shared Function QuitarCeros(cadena As String) As String
        Dim resultado As String = ""
        Dim Largo As Integer = Len(cadena)
        Dim i As Integer = 1
        Dim Caracter As String

        While i <= Largo
            Caracter = Mid(cadena, i, 1)
            If Caracter <> "0" Then
                resultado = Mid(cadena, i, (Len(cadena) + 1) - i)
                Exit While
            End If

            i = i + 1
        End While

        If Replace(resultado, " ", "") <> "" Then
            Return resultado
        Else
            Return "0"
        End If

    End Function
    'Funcion para poner separador de decimal
    Public Shared Function PonerDecimal(Cadena2 As String) As String
        Dim resultado As String = ""
        Dim Largo As Integer = Len(Cadena2)
        Dim i As Integer = 1
        'Dim Caracter As String

        resultado = Mid(Cadena2, 1, Len(Cadena2) - 2) & "." & Mid(Cadena2, Len(Cadena2) - 1)

        Return resultado

    End Function
End Class