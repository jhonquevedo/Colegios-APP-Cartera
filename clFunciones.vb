Imports System.Data.SqlClient

Public Class clFunciones

    Public oCompany As SAPbobsCOM.Company
    Public Shared SociedadActual As String
    Public Shared UserActual As String
    Public Shared PassActual As String

    Public Shared Function SetConectionSQL(Base As String) As SqlConnection
        Dim Cadena As String
        Dim Servidor As String = My.Settings.Servidor
        Dim UserDB As String = My.Settings.UserDB
        Dim PassDb As String = My.Settings.PassDB

        Cadena = "Data Source=" & Servidor & ",1433;Network Library=DBMSSOCN;Initial Catalog=" & Base & ";User ID=" & UserDB & ";Password=" & PassDb & ";"
        SetConectionSQL = New SqlConnection(Cadena)
        Return SetConectionSQL
    End Function


End Class
