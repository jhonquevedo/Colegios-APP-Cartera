Imports System.Data
Imports System.Data.SqlClient
Imports System
Imports System.Timers
Imports System.IO
Module Connection

    Public oRecordSet As SAPbobsCOM.Recordset ' A recordset object
    Public oCompany As SAPbobsCOM.Company ' The company object
    Public sErrMsg As String
    Public lErrCode As Integer
    Public lRetCode As Integer
    Public Item1 As String
End Module

Public Class From1
    Private Sub BrowseButton_Click(ByVal sender As System.Object, _
                                   ByVal e As System.EventArgs) Handles Button2.Click

        Dim folderDlg As New FolderBrowserDialog
        folderDlg.ShowNewFolderButton = True
        If (folderDlg.ShowDialog() = DialogResult.OK) Then
            TextBox1.Text = folderDlg.SelectedPath
            ''Dim root As Environment.SpecialFolder = folderDlg.RootFolder
        End If

    End Sub
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim Directorio1 As String = TextBox1.Text & "\"
        Try
            Dim FileStr As String() = Directory.GetFiles(Directorio1, "*.txt")
            For Each Files In FileStr
                Dim objFSO = CreateObject("Scripting.FileSystemObject")
                Dim objFile = objFSO.OpenTextFile(Files.ToString)
                Dim FileName As String = My.Computer.FileSystem.GetName(Files)
                Dim SizeFile = objFSO.GetFile(Directorio1 & FileName)
                Dim strSize As Integer = SizeFile.Size
                objFile.Close()
                If strSize = 0 Then
                    objFile.Close()
                    My.Computer.FileSystem.DeleteFile(Directorio1 & FileName)
                Else
                    Dim objFSO1 = CreateObject("Scripting.FileSystemObject")
                    Dim objFile1 = objFSO1.OpenTextFile(Files.ToString)
                    Dim strText As String = objFile1.ReadAll 'o ReadLine

                    Dim lector As New IO.StreamReader(Directorio1 & FileName)
                    ' Leer el contenido mientras no se llegue al final

                    Dim i As Integer
                    oCompany = New SAPbobsCOM.Company()
                    oCompany.Server = "Localhost"
                    oCompany.CompanyDB = ComboComp.Text
                    oCompany.DbUserName = "sap"
                    oCompany.DbPassword = "AtomCol3*"
                    oCompany.UserName = "shernand"
                    oCompany.Password = "T&QGroup1"
                    oCompany.LicenseServer = "LocalHost" & ":30000"
                    oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
                    Dim TipoRegistro As String = ""
                    Dim NITEmpresa As String = ""
                    Dim FechaPago As String = ""

                    While lector.Peek() <> -1
                        Dim linea As String = lector.ReadLine()


                        If Mid(linea, 1, 2) = "01" Then
                            TipoRegistro = Mid(linea, 1, 2)
                            NITEmpresa = Mid(linea, 3, 10)
                            FechaPago = Mid(linea, 13, 8)
                        End If

                        Dim FacPagar As String = ""
                        Dim Valor As String = ""
                        Dim CondPago As String = ""

                        If Mid(linea, 1, 2) = "06" Then

                            FacPagar = QuitarCeros(Mid(linea, 3, 48))
                            Valor = QuitarCeros(Mid(linea, 51, 14))
                            CondPago = QuitarCeros(Mid(linea, 65, 2))

                            i = oCompany.Connect()
                            If i = 0 Then
                                Try
                                    Dim status As Object
                                    Dim errorcode As Integer
                                    Dim errordesc As String

                                    Dim objPay As SAPbobsCOM.Payments
                                    objPay = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                                    objPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments

                                    objPay.DocDate = Date.Now.Date
                                    objPay.CardCode = "C900074272"
                                    objPay.CashSum = Valor
                                    objPay.UserFields.Fields.Item("U_asesor").Value = CondPago
                                    objPay.Invoices.DocEntry = FacPagar
                                    objPay.Invoices.SumApplied = Valor
                                    objPay.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
                                    objPay.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
                                    status = objPay.Add()
                                    'MsgBox(oCompany.GetLastErrorDescription)
                                    If (status <> 0) Then
                                        oCompany.GetLastError(errorcode, errordesc)
                                        MsgBox("Error: " & errorcode & " - " & errordesc)
                                        'MsgBox("Failed to add a payment")
                                    End If
                                Catch ex As Exception
                                    MsgBox(ex.Message)
                                End Try

                            End If

                        End If
                    End While

                    objFile1.Close()
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try
    End Sub

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

        Return resultado

    End Function
End Class
