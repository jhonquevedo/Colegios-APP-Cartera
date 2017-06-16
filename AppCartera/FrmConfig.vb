Imports System.Data.SqlClient
Imports CarteraPSE.clFunciones

Public Class FrmConfig
    Public ServidorIni As String = My.Settings.Servidor
    Public UserdbIni As String = My.Settings.UserDB
    Public PassdbIni As String = My.Settings.PassDB
    'Public dsnMySQLIni As String = My.Settings.dsnMySql
    Public SQLVersionIni As String = My.Settings.SQLVersion

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub BtnSave_Click(sender As System.Object, e As System.EventArgs) Handles BtnSave.Click

        If TxtServidor.Text = "" Then
            MsgBox("Los datos ingresados no son correctos", MsgBoxStyle.Information)
            TxtServidor.Focus()
            Exit Sub
        End If

        Try
            My.Settings.Servidor = TxtServidor.Text
            My.Settings.UserDB = TxtUserDB.Text
            My.Settings.PassDB = TxtPassDB.Text
            'My.Settings.dsnMySql = TxtdsnMySQL.Text
            My.Settings.SQLVersion = CmbSQLVersion.Text

            Dim cnn As SqlConnection
            cnn = SetConectionSQL("master")
            Using cnn
                cnn.Open()
                cnn.Close()
            End Using

            'datos correctos
            MsgBox("Datos registrados con exito", MsgBoxStyle.Information)
            LoginForm1.GroupBox1.Enabled = True
            LoginForm1.GroupBox3.Enabled = True
            LoginForm1.TraerCompanias()
            Me.Close()

        Catch ex As Exception
            My.Settings.Servidor = ServidorIni
            My.Settings.UserDB = UserdbIni
            My.Settings.PassDB = PassdbIni
            'My.Settings.dsnMySql = dsnMySQLIni
            My.Settings.SQLVersion = SQLVersionIni

            MsgBox("Los datos ingresados no son correctos. Error: " & ex.Message, MsgBoxStyle.Information)
            DatosEnCampos()
            Exit Sub
        End Try
        

    End Sub

    Public Function DatosEnCampos()

        TxtServidor.Text = My.Settings.Servidor
        TxtUserDB.Text = My.Settings.UserDB
        TxtPassDB.Text = My.Settings.PassDB
        'TxtdsnMySQL.Text = My.Settings.dsnMySql
        CmbSQLVersion.Text = My.Settings.SQLVersion

        Return 0
    End Function

    Private Sub FrmConfig_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load


        Dim Versiones() As String
        Versiones = {"2005", "2008", "2012", "2014"}
        For Each VSQL In Versiones
            CmbSQLVersion.Items.Add(VSQL)
        Next

        DatosEnCampos()

    End Sub
End Class