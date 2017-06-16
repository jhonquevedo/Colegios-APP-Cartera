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
Public Class Novedades


    'Public SBO_Company = New SAPbobsCOM.Company
    'Public servidor As String = ""
    'Public BaseDatos As String = ""
    'Public usuario As String = ""
    'Public clave As String = ""
    'Public usuarioSAP As String = ""
    'Public claveSAP As String = ""
    'Public FamiliaID As String = ""
    'Public Responsable As String = ""
    'Public seleccion As Boolean = False
    'Public TotalCancelado As String = ""
    'Public CodigoBanco As String = ""
    'Public NombreBanco As String = ""
    'Public ListaSeleccion As List(Of String) = New List(Of String)
    'Public PagoAutomatico As Boolean = False
    'Public FamiliaIDConsulta As String = ""

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim folderDlg As New OpenFileDialog()
        folderDlg.InitialDirectory = "c:\"
        folderDlg.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
        folderDlg.FilterIndex = 2
        folderDlg.RestoreDirectory = True
        If folderDlg.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            TextBox1.Text = folderDlg.FileName()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click




        'Dim lRetCode, lErrCode As Long
        'Dim sErrMsg As String

        'LeerDatosCompania("C:\SAP\Novedades\Credenciales.txt")
        ''Instantiate a Company object
        'SBO_Company = New SAPbobsCOM.Company
        'SBO_Company.Server = servidor '"ASUSX555LD\JHONMSSQL"
        'SBO_Company.CompanyDB = BaseDatos '"PRUEBA_ANGLO_0807"
        'SBO_Company.UserName = usuarioSAP  '"manager"
        'SBO_Company.Password = claveSAP  '"S31_COL"
        'SBO_Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
        'SBO_Company.DbUserName = usuario '"sa"
        'SBO_Company.DbPassword = clave '"1234"
        'SBO_Company.UseTrusted = False


        'lRetCode = SBO_Company.Connect()

        ''Check Return Code
        'If lRetCode <> 0 Then

        '    SBO_Company.GetLastError(lErrCode, sErrMsg)
        '    MsgBox("Error al conectar: " & sErrMsg)
        'Else
        '    ActualizarUDO()
        '    'MsgBox("Conexion exitosa a: " & SBO_Company.CompanyDB)
        'End If
        ActualizarUDO()

    End Sub

    'Public Sub LeerDatosCompania(ByRef ruta As String)
    '    Dim objReader As New StreamReader(ruta)
    '    Dim sLine As String = ""
    '    Dim arrText As New ArrayList()
    '    Do
    '        sLine = objReader.ReadLine()
    '        If Not sLine Is Nothing Then
    '            If sLine.Split(":")(0) = "Servidor" Then
    '                servidor = sLine.Split(":")(1)
    '            End If
    '            If sLine.Split(":")(0) = "BD" Then
    '                BaseDatos = sLine.Split(":")(1)
    '            End If
    '            If sLine.Split(":")(0) = "Usuario" Then
    '                usuario = sLine.Split(":")(1)
    '            End If
    '            If sLine.Split(":")(0) = "Clave" Then
    '                clave = sLine.Split(":")(1)
    '            End If
    '            If sLine.Split(":")(0) = "UsuarioSAP" Then
    '                usuarioSAP = sLine.Split(":")(1)
    '            End If
    '            If sLine.Split(":")(0) = "ClaveSAP" Then
    '                claveSAP = sLine.Split(":")(1)
    '            End If

    '        End If
    '    Loop Until sLine Is Nothing
    '    objReader.Close()
    '    Console.ReadLine()
    'End Sub

    Public Sub ActualizarUDO()
        Cursor = Cursors.WaitCursor

        Dim Count As Integer = 0
        Dim CountFail As Integer = 0
        Dim Directorio1 As String = TextBox1.Text
        Dim aux As String = ""

        Dim Fecha = DateTime.Now.ToString("yyyyMMdd hh mm ss")
        Dim strm1 As New StreamWriter("C:\SAP\Pagos Banco\Log Pagos\Log" & Fecha.ToString & ".txt")
        Dim Texto1 As String = ""
        Dim FechaLog As String

        Dim codigoAlumno As String = ""
        Dim codigoConcepto As String = ""
        Dim nombreConcepto As String = ""
        Dim cardCode As String = ""
        Dim cardName As String = ""
        Dim tarifaF As String = ""
        Dim tarifaE As String = ""
        Dim descuentoSN As String = ""
        Dim descuentoProcen As String = ""
        Dim porcentajePago As String = ""
        Dim precio As String = ""
        Dim proyecto As String = ""
        Dim viaDePago As String = ""
        Dim periodo As String = ""
        Dim centroCosto1 As String = ""
        Dim centroCosto2 As String = ""
        Dim ErrorLog As Boolean = False
        Dim contador As Integer = 0
        Dim numErrores As Integer = 0

        Try
            Dim objFSO1 = CreateObject("Scripting.FileSystemObject")
            Dim objFile1 = objFSO1.OpenTextFile(Directorio1.ToString)
            Dim strText As String = objFile1.ReadAll 'o ReadLine

            Dim lector As New IO.StreamReader(Directorio1)
            ' Leer el contenido mientras no se llegue al final         

            
            Dim numLinea As Integer = 0

         

            Texto1 = SociedadActual & vbCrLf
            Texto1 = Texto1 & vbCrLf & "------------------VALIDACION ARCHIVO------------------" & vbCrLf & vbCrLf
            Texto1 = Texto1 & "Codigo" & vbTab & "Valor" & vbTab & "Fecha" & vbTab & "Comentarios" & vbCrLf

            While lector.Peek() <> -1
                Dim linea As String = ""
                If numLinea < 2 Then
                    numLinea = numLinea + 1
                    linea = lector.ReadLine()
                Else
                    linea = lector.ReadLine()
                    codigoAlumno = linea.Split("|")(0)
                    codigoConcepto = linea.Split("|")(1)
                    nombreConcepto = linea.Split("|")(2)
                    cardCode = linea.Split("|")(3)
                    cardName = linea.Split("|")(4)
                    tarifaF = linea.Split("|")(5)
                    tarifaE = linea.Split("|")(6)
                    descuentoSN = linea.Split("|")(7)
                    descuentoProcen = linea.Split("|")(8)
                    porcentajePago = linea.Split("|")(9)
                    precio = linea.Split("|")(10)
                    proyecto = linea.Split("|")(11)
                    viaDePago = linea.Split("|")(12)
                    periodo = linea.Split("|")(13)
                    centroCosto1 = linea.Split("|")(14)
                    centroCosto2 = linea.Split("|")(15)

                 

                    If Not ConsultarAlumno(codigoAlumno) Then
                        FechaLog = ""
                        FechaLog = DateTime.Now().ToString("dd/MM/yyyy hh:mm:ss tt")
                        Texto1 = Texto1 & vbCrLf & (FechaLog & vbTab & "Error: No existe el Alumno - Alumno: " & codigoAlumno & " Concepto: " & codigoConcepto)
                        ErrorLog = True
                        numErrores = numErrores + 1
                    Else
                        If Not ColsultarConcepto(codigoAlumno, codigoConcepto) Then
                            Dim oGeneralService As SAPbobsCOM.GeneralService
                            Dim oGeneralData As SAPbobsCOM.GeneralData
                            Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
                            Dim oChild As SAPbobsCOM.GeneralData
                            Dim oChildren As SAPbobsCOM.GeneralDataCollection
                            Dim sCmp As SAPbobsCOM.CompanyService
                            Try
                                sCmp = oCompany.GetCompanyService
                                oGeneralService = sCmp.GetGeneralService("SEI_ALUMNOS")
                                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)


                                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                                oGeneralParams.SetProperty("Code", codigoAlumno)
                                oGeneralData = oGeneralService.GetByParams(oGeneralParams)


                                oChildren = oGeneralData.Child("SEI_CONC_ALUMN")
                                oChild = oChildren.Add

                                oChild.SetProperty("U_SEIConc", codigoConcepto) '-> Concepto
                                oChild.SetProperty("U_SEINConc", nombreConcepto)
                                oChild.SetProperty("U_SEITitul", cardCode)
                                oChild.SetProperty("U_SEINTitu", cardName)
                                oChild.SetProperty("U_SEITariF", tarifaF)
                                oChild.SetProperty("U_SEITariE", tarifaE)
                                oChild.SetProperty("U_SEIDEMan", descuentoSN)
                                oChild.SetProperty("U_SEIDesE", descuentoProcen)
                                oChild.SetProperty("U_SEIPPCT", porcentajePago.Trim)
                                oChild.SetProperty("U_SEIPrCoT", precio)
                                oChild.SetProperty("U_SEIProj", proyecto)
                                oChild.SetProperty("U_SEIViaPa", viaDePago)
                                oChild.SetProperty("U_SEI_Peri", periodo)
                                oChild.SetProperty("U_SEI_CogsOcrCod", centroCosto1)
                                oChild.SetProperty("U_SEI_CogsOcrCod2", centroCosto2)
                                oGeneralService.Update(oGeneralData)
                                contador = contador + 1
                            Catch ex As Exception
                                ErrorLog = True
                                FechaLog = ""
                                FechaLog = DateTime.Now().ToString("dd/MM/yyyy hh:mm:ss tt")
                                Cursor = Cursors.Default
                                Texto1 = Texto1 & vbCrLf & (FechaLog & vbTab & "Error: " & ex.Message & " - Alumno: " & codigoAlumno & " Concepto: " & codigoConcepto)
                                numErrores = numErrores + 1
                            End Try

                        End If
                    End If


                   
                End If

NextLine:
            End While
            FechaLog = ""
            FechaLog = DateTime.Now().ToString("dd/MM/yyyy hh:mm:ss tt")
            Texto1 = Texto1 & vbCrLf & (FechaLog & vbTab & "Fin del proceso: ")
            Texto1 = Texto1 & vbCrLf & (FechaLog & vbTab & "Novedades cargadas: " & contador & " Novedades no cargadas: " & numErrores)
            strm1.WriteLine(Texto1)
            strm1.Close()
            If ErrorLog Then
                MsgBox("Proceso terminado con errores")
            Else
                MsgBox("Proceso terminado con exito")
            End If

            objFile1.Close()
            lector.Close()
        Catch ex As Exception
            'FechaLog = ""
            'FechaLog = DateTime.Now().ToString("dd/MM/yyyy hh:mm:ss tt")
            'Cursor = Cursors.Default
            'Texto1 = Texto1 & vbCrLf & (FechaLog & vbTab & "Error: " & ex.Message & " - Alumno: " & codigoAlumno & " Concepto: " & codigoConcepto)
            'strm1.WriteLine(Texto1)
            'strm1.Close()
            MsgBox("Proceso terminado con errores")
            Exit Sub
        End Try
    End Sub

    Public Function ColsultarConcepto(ByRef codAlumno As String, ByRef codConcepto As String) As Boolean
        Dim oRs As SAPbobsCOM.Recordset
        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sql As String = ""
        Try
            sql = "SELECT * FROM [@SEI_CONC_ALUMN] WHERE Code = '" & codAlumno & "' AND U_SEIConc = '" & codConcepto & "' AND (U_SEIFeUlF IS NULL OR U_SEIFeUlF = '')"

            oRs.DoQuery(sql)
            While Not oRs.EoF
                Return True
            End While
            Return False
        Catch ex As Exception
            MsgBox(ex.Message & " " & sql)
            Return False
        End Try

    End Function

    Public Function ConsultarAlumno(ByRef codAlumno As String) As Boolean
        Dim oRs As SAPbobsCOM.Recordset
        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sql As String = ""
        Try
            sql = "SELECT * FROM [@SEI_ALUMNOS]  WHERE Code = '" & codAlumno & "' "

            oRs.DoQuery(sql)
            While Not oRs.EoF
                Return True
            End While
            Return False
        Catch ex As Exception
            MsgBox(ex.Message & " " & sql)
            Return False
        End Try

    End Function
End Class