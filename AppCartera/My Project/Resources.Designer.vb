﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System

Namespace My.Resources
    
    'This class was auto-generated by the StronglyTypedResourceBuilder
    'class via a tool like ResGen or Visual Studio.
    'To add or remove a member, edit your .ResX file then rerun ResGen
    'with the /str option, or rebuild your VS project.
    '''<summary>
    '''  A strongly-typed resource class, for looking up localized strings, etc.
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0"),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.Microsoft.VisualBasic.HideModuleNameAttribute()>  _
    Friend Module Resources
        
        Private resourceMan As Global.System.Resources.ResourceManager
        
        Private resourceCulture As Global.System.Globalization.CultureInfo
        
        '''<summary>
        '''  Returns the cached ResourceManager instance used by this class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend ReadOnly Property ResourceManager() As Global.System.Resources.ResourceManager
            Get
                If Object.ReferenceEquals(resourceMan, Nothing) Then
                    Dim temp As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager("CarteraPSE.Resources", GetType(Resources).Assembly)
                    resourceMan = temp
                End If
                Return resourceMan
            End Get
        End Property
        
        '''<summary>
        '''  Overrides the current thread's CurrentUICulture property for all
        '''  resource lookups using this strongly typed resource class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend Property Culture() As Global.System.Globalization.CultureInfo
            Get
                Return resourceCulture
            End Get
            Set
                resourceCulture = value
            End Set
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Set @Conteo = 
        '''(
        '''	Select C0.Filas From 
        '''	(
        '''		Select Count(*) As Filas from 
        '''		(
        '''			Select
        '''			&apos;1&apos; +
        '''			REPLICATE(&apos;0&apos;, 6 - Len(LTRIM(A2.BankCode))) + LTRIM(A2.BankCode) +
        '''			REPLICATE(&apos;0&apos;, 4 - Len(IsNull(A5.Street,&apos;&apos;))) + IsNull(A5.Street,&apos;&apos;) +
        '''			REPLICATE(&apos;0&apos;, 17 - Len(IsNull(A2.DflAccount,&apos;0&apos;))) + IsNull(A2.DflAccount,&apos;0&apos;) +
        '''			Case when A2.DflIBAN = &apos;01&apos; Then &apos;1&apos; When A2.DflIBAN = &apos;02&apos; Then &apos;2&apos; Else &apos;&apos; End +
        '''			REPLICATE(&apos;0&apos;, 17 - DATALENGTH(Replace(Cast(IsNull(A1.SumApplied,&apos;0&apos;) As Numeric(19 [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property Bancolombia() As String
            Get
                Return ResourceManager.GetString("Bancolombia", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Select
        '''&apos;6&apos; + 
        '''Case when A2.DflIBAN = &apos;01&apos; Then &apos;22&apos; When A2.DflIBAN = &apos;02&apos; Then &apos;32&apos; When A2.DflIBAN = &apos;03&apos; Then &apos;23&apos; When A2.DflIBAN = &apos;04&apos; Then &apos;33&apos; Else &apos;00&apos; End +
        '''REPLICATE(&apos;0&apos;, 12 - DATALENGTH(Replace(Cast(sum(A1.SumApplied) As Numeric(19,2)),&apos;.&apos;,&apos;&apos;))) + Replace(Cast(sum(A1.SumApplied) As Numeric(19,2)),&apos;.&apos;,&apos;&apos;) +
        '''Cast(IsNull(A2.DflAccount,&apos;&apos;) As Char(17)) +
        '''&apos;000010320&apos; +
        '''Cast(Replace(IsNull(A2.LicTradNum,&apos;&apos;), &apos;-&apos;, &apos;&apos;) As Char(15)) +
        '''Cast(A2.CardName As Char (22))+
        '''Cast(&apos;V&apos; As Char (2)) +
        '''Space [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property CajaSocial() As String
            Get
                Return ResourceManager.GetString("CajaSocial", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Set @Valor =
        '''(
        '''	Select SUM(CAST(SUBSTRING(P0.Cadena,449,8) as numeric(16,2)))  As Total from 
        '''		(			
        '''			
        '''			Select C0.* from (
        '''			/*Cabecera Factura 00*/
        '''			Select P0.Cadena1 + REPLICATE(&apos;0&apos;, 18 - DATALENGTH(Replace(Cast(Sum(P0.Valor) as numeric(16,2)),&apos;.&apos;,&apos;&apos;))) + Replace(Cast(Sum(P0.Valor) as numeric(16,2)),&apos;.&apos;,&apos;&apos;) + P0.Cadena As Cadena,
        '''			CASE WHEN CHARINDEX(&apos;-&apos;, P0.NumAtCard) &gt; 0 THEN LEFT(P0.NumAtCard, CHARINDEX(&apos;-&apos;, P0.NumAtCard)-1) ELSE P0.NumAtCard  END as NumAtCard ,0 DocEntry, 2 As Orden
        ''' [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property CarteraPSE() As String
            Get
                Return ResourceManager.GetString("CarteraPSE", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Set @Valor =
        '''(
        '''	Select SUM(CAST(SUBSTRING(P0.Cadena,449,8) as numeric(16,2)))  As Total from 
        '''		(			
        '''			
        '''			Select C0.* From (
        '''			Select
        '''			&apos;1&apos; +
        '''			Convert(VARCHAR(8), GetDate(), 112) +
        '''			REPLICATE(&apos;0&apos;, 6 - DATALENGTH(LTRIM(@Conteo))) + LTRIM(@Conteo) +
        '''			REPLICATE(&apos;0&apos;, 18 - DATALENGTH(Replace(Cast(@Valor as numeric(16,2)),&apos;.&apos;,&apos;&apos;))) + Replace(Cast(@Valor as numeric(16,2)),&apos;.&apos;,&apos;&apos;) +
        '''			@Consec +
        '''			Space(527) As Cadena, 0 As NumAtCard, 0 As DocEntry, 1 As Orden
        '''			)C0
        '''
        '''			Union All
        '''
        '''			Se [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property CarteraPSE1() As String
            Get
                Return ResourceManager.GetString("CarteraPSE1", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Select &apos;CDU026NUMO&apos;As &apos;NO TOCAR&apos;,&apos;CDU026NIDE&apos; As &apos;NO TOCAR&apos;,&apos;CDU026NCLI&apos;As &apos;NO TOCAR&apos;,&apos;CDU026TICC&apos; As &apos;NO TOCAR&apos;,&apos;CDU026NRCC&apos; As &apos;NO TOCAR&apos;,
        '''&apos;CDU026NITB&apos;As &apos;NIT DEL CLIENTE RECEPTOR&apos;,&apos;CDU026NOMB&apos; As &apos;NOMBRE DEL CLIENTE RECEPTOR&apos;,&apos;CDU026BANB&apos; As &apos;CODIGO BANCO&apos;,
        '''&apos;CDU026TIPC&apos; As &apos;1= CTA CTE / 2= CTA AHOR&apos;,&apos;CDU026NRCB&apos; As &apos;NUMERO DE CUENTA&apos;,&apos;CDU026CLAM&apos; As &apos;NO TOCAR&apos;,&apos;CDU026VLRB&apos; As &apos;VALOR SIN DECIMALES&apos;
        '''Union All
        '''Select
        '''Cast(Row_number() over (ORDER BY A2.LicTradNum) As nvarchar(40)) As &apos;NO TOCAR&apos;,
        '''&apos;&apos; As [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property CDT() As String
            Get
                Return ResourceManager.GetString("CDT", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Select &apos;CDU026NUMO&apos;As &apos;NO TOCAR&apos;,&apos;CDU026NIDE&apos; As &apos;NO TOCAR&apos;,&apos;CDU026NCLI&apos;As &apos;NO TOCAR&apos;,&apos;CDU026TICC&apos; As &apos;NO TOCAR&apos;,&apos;CDU026NRCC&apos; As &apos;NO TOCAR&apos;,
        '''&apos;CDU026NITB&apos;As &apos;NIT DEL CLIENTE RECEPTOR&apos;,&apos;CDU026NOMB&apos; As &apos;NOMBRE DEL CLIENTE RECEPTOR&apos;,&apos;CDU026BANB&apos; As &apos;CODIGO BANCO&apos;,
        '''&apos;CDU026TIPC&apos; As &apos;1= CTA CTE / 2= CTA AHOR&apos;,&apos;CDU026NRCB&apos; As &apos;NUMERO DE CUENTA&apos;,&apos;CDU026CLAM&apos; As &apos;NO TOCAR&apos;,&apos;CDU026VLRB&apos; As &apos;VALOR SIN DECIMALES&apos;
        '''Union All
        '''Select
        '''Cast(Row_number() over (ORDER BY A2.LicTradNum) As nvarchar(40)) As &apos;NO TOCAR&apos;,
        '''&apos;&apos; As [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property CDT2() As String
            Get
                Return ResourceManager.GetString("CDT2", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Set @Conteo = 
        '''(
        '''	Select C0.Filas From 
        '''	(
        '''		Select Count(*) As Filas from 
        '''		(
        '''			Select
        '''			REPLICATE(&apos;0&apos;, 5 - Len(Cast(Row_number() over (ORDER BY A1.DocEntry) + 1 As nvarchar(40)))) + Cast(Row_number() over (ORDER BY A1.DocEntry) + 1 As nvarchar(40)) As REGISTRO,
        '''			&apos;02&apos; As TIPO_REGISTRO,
        '''			REPLICATE(&apos;0&apos;, 12 - Len(IsNull(Cast(A2.DflAccount As Nvarchar(500)),&apos;0&apos;))) + IsNull(Cast(A2.DflAccount As Nvarchar(500)),&apos;0&apos;) As NUMERO_CUENTA,
        '''			REPLICATE(&apos;0&apos;, 11 - Len(Replace(IsNull(A2.LicTradNum,&apos;&apos;), &apos; [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property Colpatria() As String
            Get
                Return ResourceManager.GetString("Colpatria", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Select * from (
        '''Select
        '''CASE WHEN CHARINDEX(&apos;-&apos;, A2.LicTradNum) &gt; 0 THEN Cast(LEFT(A2.LicTradNum, CHARINDEX(&apos;-&apos;, A2.LicTradNum)-1) As char(80)) ELSE Cast(A2.LicTradNum As char(80)) END +
        '''&apos;01&apos; +
        '''--Cast(IsNull(A0.NumAtCard,&apos;&apos;) As Char(35)) +
        '''CASE WHEN CHARINDEX(&apos;-&apos;, A2.LicTradNum) &gt; 0 THEN Cast(LEFT(A2.LicTradNum, CHARINDEX(&apos;-&apos;, A2.LicTradNum)-1) As char(35)) ELSE Cast(A2.LicTradNum As char(35)) END +
        '''&apos;00&apos; +
        '''Cast(UPPER(Replace(Replace(Replace(Replace(Replace(Replace(IsNull(A0.CardName,&apos;&apos;),&apos;ñ&apos;,&apos;n&apos;), &apos;á&apos;, [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property UsuariosPSE() As String
            Get
                Return ResourceManager.GetString("UsuariosPSE", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Select * from (
        '''Select
        '''CASE WHEN CHARINDEX(&apos;-&apos;, A2.LicTradNum) &gt; 0 THEN Cast(LEFT(A2.LicTradNum, CHARINDEX(&apos;-&apos;, A2.LicTradNum)-1) As char(80)) ELSE Cast(A2.LicTradNum As char(80)) END +
        '''&apos;01&apos; +
        '''Cast(IsNull(A0.NumAtCard,&apos;&apos;) As Char(35)) +
        '''&apos;00&apos; +
        '''Cast(UPPER(Replace(Replace(Replace(Replace(Replace(Replace(IsNull(A0.CardName,&apos;&apos;),&apos;ñ&apos;,&apos;n&apos;), &apos;á&apos;, &apos;a&apos;), &apos;é&apos;,&apos;e&apos;), &apos;í&apos;, &apos;i&apos;), &apos;ó&apos;, &apos;o&apos;), &apos;ú&apos;,&apos;u &apos;) ) as Char(50)) +
        '''Space(50) As Cadena
        '''From ORDR A0
        '''Inner Join OCRD A2 ON A0.CardCode = A2.CardCode
        '''Inner Join OADM [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property UsuariosPSE1() As String
            Get
                Return ResourceManager.GetString("UsuariosPSE1", resourceCulture)
            End Get
        End Property
    End Module
End Namespace
