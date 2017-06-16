Imports System.Data.SqlClient
Imports ActDatosSDP.clFunciones

Public Class LoginForm1

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    Public connection As SqlConnection
    Public oCompany As SAPbobsCOM.Company

    Private Sub LoginForm1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Dim BaseSelec As String
        Dim Servidor As String
        TxtServidor.Text = My.Settings.Servidor
        Servidor = TxtServidor.Text '"LocalHost"

        Try
            If My.Settings.Servidor = "" Or My.Settings.UserDB = "" Or My.Settings.PassDB = "" Then
                GroupBox1.Enabled = False
                GroupBox3.Enabled = False
            End If

            'IW-00163E0058b0
            BaseSelec = "SBO-COMMON"
            'Cadena = "Data Source=" & Servidor & ",1433;Network Library=DBMSSOCN;Initial Catalog=" & BaseSelec & ";User ID=SAP;Password=AtomCol3*;"

            connection = SetConectionSQL(BaseSelec)

            Using connection

                TraerCompanias()

            End Using
            'la conexion esta ok

        Catch ex As Exception
            MsgBox("Se ha presentado el siguiente error: " & ex.Message, MsgBoxStyle.Critical, "Mensaje")
        End Try

        
    End Sub

    Public Function TraerCompanias()
        Dim sQuery As String
        Dim BaseSelec As String = "SBO-COMMON"
        sQuery = "Select dbName, cmpName From SRGC "
        connection = SetConectionSQL(BaseSelec)
        Dim command As New SqlCommand(sQuery, connection)
        command.Connection.Open()
        Dim reader As SqlDataReader = command.ExecuteReader()
        Try
            If reader.HasRows = True Then
                ComboSociedad.Items.Clear()
                While reader.Read()
                    ComboSociedad.Items.Add(reader("dbName"))

                End While
            End If

        Finally
            ' Always call Close when done reading.
            reader.Close()
        End Try
    End Function

    Private Sub BtnChange_Click(sender As System.Object, e As System.EventArgs) Handles BtnChange.Click

        If BtnChange.Text.ToUpper = "EDIT" Then
            BtnChange.Text = "Save"
            TxtServidor.ReadOnly = False
            Exit Sub
        End If

        If BtnChange.Text.ToUpper = "SAVE" Then

            If TxtServidor.Text = "" Then
                TxtServidor.Text = "LocalHost"
            End If
            My.Settings.Servidor = TxtServidor.Text
            BtnChange.Text = "Edit"
            TxtServidor.ReadOnly = True

            Exit Sub
        End If

    End Sub

    Private Sub OK_Click_1(sender As System.Object, e As System.EventArgs) Handles OK.Click

        If ComboSociedad.Text = "" Then
            MsgBox("No ha seleccionado ninguna Base de Datos", MsgBoxStyle.Information, "Aviso")
            ComboSociedad.Focus()
            Exit Sub
        End If

        If Trim(TxtUsername.Text.ToString) = "" Then
            MsgBox("Ingrese un Usuario", MsgBoxStyle.Information, "Aviso")
            TxtUsername.Focus()
            Exit Sub
        End If

        If Trim(TxtPassword.Text.ToString) = "" Then
            MsgBox("Ingrese una Contraseña", MsgBoxStyle.Information, "Aviso")
            TxtPassword.Focus()
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        GroupBox1.Enabled = False
        GroupBox2.Enabled = False
        GroupBox3.Enabled = False

        Application.DoEvents()
        If ConecCompany(TxtUsername.Text, TxtPassword.Text, ComboSociedad.Text) = False Then
            'no se logro conectar a la sociedad
            MsgBox("Los datos ingresados no son validos", MsgBoxStyle.Information)

            GroupBox1.Enabled = True
            GroupBox2.Enabled = True
            GroupBox3.Enabled = True
            Me.Cursor = Cursors.Default

        Else
            'Acceso correcto
            'MsgBox("datos ingresados validos", MsgBoxStyle.Exclamation)
            GroupBox1.Enabled = True
            GroupBox2.Enabled = True
            GroupBox3.Enabled = True
            Me.Cursor = Cursors.Default
            SociedadActual = ComboSociedad.Text
            UserActual = TxtUsername.Text.ToUpper
            PassActual = TxtPassword.Text
            FrmPrincipal.Show()
            Me.Close()

        End If
        


    End Sub

    Public Function ConecCompany(User As String, Pass As String, Company As String) As Boolean

        Dim i As Integer

        oCompany = New SAPbobsCOM.Company()
        oCompany.Server = My.Settings.Servidor
        oCompany.CompanyDB = Company
        oCompany.DbUserName = My.Settings.UserDB
        oCompany.DbPassword = My.Settings.PassDB
        oCompany.UserName = User
        oCompany.Password = Pass
        oCompany.LicenseServer = My.Settings.Servidor & ":30000"

        Select Case My.Settings.SQLVersion
            Case 2005
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
            Case 2008
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            Case 2012
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
            Case 2014
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014
            Case Else
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL
        End Select

        
        i = oCompany.Connect()
        
        If i = 0 Then
            Return True
        Else
            Return False
        End If


    End Function

    Private Sub BtnConfig_Click(sender As System.Object, e As System.EventArgs) Handles BtnConfig.Click
        FrmConfig.Show()

    End Sub
End Class
