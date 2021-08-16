<%On Error Resume Next 

'conecta a base de dados, usando a string de conexão e a conexão passada no parâmetro
Function AbreOra(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.CursorLocation = 3
'rs_nome.open string,"driver={ibm db2 odbc driver}; dsn=PROD; uid=gbadmin; pwd=gbadmin"
rs_nome.open string,"Provider=msdaora;Data Source=sisttst;User Id=asprea;Password=asprea$;"
End Function

Function AbreOraRS(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.CursorLocation = 3
rs_nome.open string,"Provider=OraOLEDB.Oracle.1;Password=rea$;Persist Security Info=False;User ID=asprea;Data Source=SISTTST;"
End Function

'conecta a base de dados, usando a string de conexão e a conexão passada no parâmetro
Function Abre(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.CursorLocation = 3
'rs_nome.open string,"driver={ibm db2 odbc driver}; dsn=PROD; uid=gbadmin; pwd=gbadmin"
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=gbadmin; pwd=gbadmin; Data Source=PROD"
End Function

'conecta a base de dados de SuporteCPD
Function AbreSU(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.CursorLocation = 3
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=db2su; pwd=sup$usu; Data Source=PROD"
End Function

'conecta a base de dados de SuporteCPD
Function AbreSUTeste(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.CursorLocation = 3
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=db2su; pwd=sup$usu; Data Source=DB2TESTE"
End Function

'conecta a base de dados de teste, usando a string de conexão e a conexão passada no parâmetro
Function AbreTeste(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.CursorLocation = 3
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=gbadmin; pwd=gbadmin; Data Source=DB2TESTE"
End Function

'conecta a base SANBX (DB2/UDB)     * 22/setembro/2015
Function AbreBX(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.open string,"Provider=MSDASQL.1;Password=UEIB7F5H;Persist Security Info=True;User ID=TASPBX;Data Source=DB2TESTE"
End Function

Function AbreD(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
'rs_nome.open string,"driver={ibm db2 odbc driver}; dsn=PROD; uid=gbadmin; pwd=gbadmin"
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=gbadmin; pwd=gbadmin; Data Source=PRODD"
End Function

'conecta a base de dados da plataforma alta (Main Frame), usando a string de conexão e a conexão passada no parâmetro
Function AbreMF(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
'rs_nome.open string,"driver={ibm db2 odbc driver}; dsn=sqldba; uid=drdaprod; pwd=prod"
rs_nome.open string,"Provider=MSDASQL.1;Password=access;Persist Security Info=True;User ID=vagenp;Data Source=prdmdb"
End Function

Function AbreMFTeste(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
'rs_nome.open string,"driver={ibm db2 odbc driver}; dsn=tstmdb; uid=drdatest; pwd=teste"
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=drdatest;Data Source=TSTMDB; pwd=teste"
End Function

'conecta a base de dados da automação
Function Abre_Automacao(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.open string,"Provider=MSDASQL.1;Password=gbadmin;Persist Security Info=False;User ID=gbadmin;Data Source=AUTOMA"
End Function

'conecta a base de dados da automação
Function Abre_AutomacaoTeste(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.open string,"Provider=MSDASQL.1;Password=gbadmin;Persist Security Info=False;User ID=gbadmin;Data Source=AUTESTE"
'rs_nome.open string,"driver={ibm db2 odbc driver}; dsn=AUTESTE; uid=gbadmin; pwd=gbadmin"
End Function 

'conecta a base de dados, usando a string de conexão e a conexão passada no parâmetro
Function AbreGQ(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.CursorLocation = 3
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=db2gq; pwd=qualidd$; Data Source=PROD"
End Function

Function AbreGQTeste(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.CursorLocation = 3
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=db2gq; pwd=qualidd$; Data Source=DB2TESTE"
End Function

'conecta a base de dados, usando a string de conexão e a conexão passada no parâmetro
Function AbreTS(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.CursorLocation = 3
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=db2ts; pwd=inspeee$; Data Source=PROD"
End Function

Function AbreTSTeste(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.CursorLocation = 3
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=db2ts; pwd=inspeee$; Data Source=DB2TESTE"
End Function


'conecta a base de dados, usando a string de conexão e a conexão passada no parâmetro
Function AbreTF(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.CursorLocation = 3
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=db2tf; pwd=cperdas$; Data Source=db2teste"
End Function

'conecta a base de dados, usando a string de conexão e a conexão passada no parâmetro
Function Abrebcv(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.CursorLocation = 3
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=db2bcv;pwd=bcurric$; Data Source=PROD"
End Function

'conecta a base de dados, usando a string de conexão e a conexão passada no parâmetro
Function AbreCO(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.CursorLocation = 3
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=db2co1;pwd=comunic$; Data Source=PROD"
End Function

'conecta a base DT (udb)
Function AbreDt(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.open string,"Provider=MSDASQL.1;Password=desliga$;Persist Security Info=True;User ID=db2dt;Data Source=prod"
End Function

Function AbreDtTeste(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.open string,"Provider=MSDASQL.1;Password=desliga$;Persist Security Info=True;User ID=db2dt;Data Source=db2teste"
End Function

'conecta a base de dados da plataforma alta (Main Frame), usando a string de conexão e a conexão passada no parâmetro - Cotação Eletrônica CE
Function AbreMF_CE_Prod(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=prodce;Data Source=PRDMDB; pwd=cesan"
End Function

Function AbreMF_CE_Teste(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=testce;Data Source=TSTMDB; pwd=cesan"
End Function

'conecta a base de dados de teste, das tabelas BE (dispositivos da brigada)
Function AbreBE(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.CursorLocation = 3
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=db2be;pwd=brigadem$; Data Source=PROD"
End Function

'conecta a base de dados da operação PIO
Function Abre_PIOTeste(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.CursorLocation = 3
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=db2pi; pwd=servpio$; Data Source=PROD"
End Function 

'conecta a base de dados para tabela site - desabilita/habilita sistema 
Function AbreSite(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
'rs_nome.open string,"driver={ibm db2 odbc driver}; dsn=PROD; uid=gbadmin; pwd=gbadmin"
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=db2web; pwd=manuweb$; Data Source=PROD"
End Function

'conecta a base de dados para tabelas de indicadores
Function AbreIndicadorT(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
'rs_nome.open string,"driver={ibm db2 odbc driver}; dsn=PROD; uid=gbadmin; pwd=gbadmin"
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=db2in; pwd=cindica$; Data Source=db2teste"
End Function
Function AbreIndicador(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
'rs_nome.open string,"driver={ibm db2 odbc driver}; dsn=PROD; uid=gbadmin; pwd=gbadmin"
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=db2in; pwd=cindica$; Data Source=prod"
End Function

'conecta a base de dados para tabelas de indicadores
Function AbreLoginWebT(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
'rs_nome.open string,"driver={ibm db2 odbc driver}; dsn=PROD; uid=gbadmin; pwd=gbadmin"
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=db2co3; pwd=med$remot; Data Source=db2teste"
End Function
Function AbreLoginWeb(string,rs_nome)
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
'rs_nome.open string,"driver={ibm db2 odbc driver}; dsn=PROD; uid=gbadmin; pwd=gbadmin"
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=db2co3; pwd=med$remot; Data Source=prod"
End Function

'******************************* Functions para Conexão *************************************

'---------- faz 1 conexão para vários recordset ----------
Function AbreCONEXAO(conn) 'conn = variável para conexao
Set conn = Server.CreateObject("ADODB.connection")
'conn.open "PROD", "gbadmin", "gbadmin" 'dsn(BD), uid(user), pwd(password)
conn.open "Provider=MSDASQL.1;Persist Security Info=False;User ID=gbadmin; pwd=gbadmin; Data Source=PROD"
End Function
'---------- faz 1 conexão para vários recordset (o DB2 não suporta) ------------
Function AbreCONEXAO_MF(conn) '- conn = variável para conexao NA PLATAFORMA ALTA
Set conn = Server.CreateObject("ADODB.connection")
'conn.open "sqldba", "drdaprod", "prod" ' CONEXAO PLATAFORMA ALTA - dsn(BD), uid(user), pwd(password)
conn.open "Provider=MSDASQL.1;Persist Security Info=False;User ID=drdaprod; pwd=prod; Data Source=sqldba"
End Function
'--- AUTOMACAO ------- faz 1 conexão para vários recordset ----------
Function AbreCONEXAO_AU(conn) 'conn = variável para conexao
Set conn = Server.CreateObject("ADODB.connection")
'conn.open "PROD", "gbadmin", "gbadmin" 'dsn(BD), uid(user), pwd(password)
conn.open "Provider=MSDASQL.1;Persist Security Info=False;User ID=gbadmin; pwd=gbadmin; Data Source=AUTOMA"
End Function
'--------- executa recordset ----------
Function AbreRS(conn,String,rs_nome) '- conexao, select, recordset
set rs_nome = conn.execute (string) 
End Function

'******************************* Functions Historian *************************************

Function AbreHistorianC(string,rs_nome)
	Set rs_nome = conecta.Execute(string)
End Function

Function AbreHistorian(string,rs_nome)'para teste
Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
rs_nome.CursorLocation = 3
'rs_nome.open string,"driver={ibm db2 odbc driver}; dsn=PROD; uid=gbadmin; pwd=gbadmin"
rs_nome.open string,"Provider=MSDASQL.1;Persist Security Info=False;User ID=gbadmin; pwd=gbadmin; Data Source=PROD"
End Function


'******************************* Functions Autorização *************************************
Function Autoriza'logar no sitema novo
	dim user_atual
	user_atual = trim(ucase(Request.ServerVariables("logon_user")))
	user_atual = Right(user_atual,Len(user_atual)-InstrRev(user_atual,"\"))
'	if user_atual="RCPD2040" then user_atual="RFMT2010"'valeria
'	if user_atual="RCPD2040" then user_atual="RFMT2020"'tania
'	if user_atual<>"RCPD2040" then
		if len(user_atual)="0" then user_atual = right(user_atual,InstrRev(user_atual,"/")+1)
		'if user_atual="RCPD2040" then user_atual = "RTF2002"
		str = "Select * from prod.co14, prod.co57 where co14codf=co57codf and co14user='"& user_atual &"'" 
		call abre(str,rs) 
		if not rs.eof then
		session("codf")=trim(rs("co14codf"))
		session("codd")=trim(rs("co14codd"))
			while not rs.eof
				session(rs("co57codi"))="1"
				session("D"&rs("co57codi"))=trim(rs("co57sigl"))
				rs.movenext
			wend
		else
			call fecha(rs)
			str = "Select * from prod.co14 where co14user='"& user_atual &"'" : call abre(str,rs) 
			if not rs.eof then
				session("codf")=trim(rs("co14codf"))
				session("codd")=trim(rs("co14codd"))
			end if	
		end if
		call fecha(rs)
'	end if
	session("User")=user_atual
end function

Function Autorizacao(var_codi, var_indi, var_tpac)
dim rs57, str57, var_usuario
'if trim(var_usuario)="" or isnull(trim(var_usuario)) then: var_usuario = "SANASA\RCPD2040": end if
str57 ="select * from prod.co57 where co57codi=" & var_codi & " and co57codf=" & codf
if var_indi<>"" then str57 = str57 & " and co57indi=" & var_indi 
if var_tpac<>"" then str57 = str57 & " and co57tpac=" & var_tpac 
Call abre(str57,rs57)
if not rs57.eof then
	Autorizacao=true
else
	Autorizacao=false
end if
Call fecha(rs57)
End Function

Function Fecha(rs_nome)'fecha conexão
rs_nome.close
set rs_nome = nothing
End Function


Function VerificaErros%>
	<Center><font class="boldText">Status: 
	<%str="select * from prod.gg03, prod.gg23 where gg03anop=gg23anop and gg03nrop=gg23nrop and gg03anop= 2002 and gg03nrop=10587 and gg03cdas=310115" 
	Call AbreCONEXAO_MF(conn) 
	Call AbreRS(conn,str,rs)

	If err.number>0 then%>
		Ocorreram Erros no Script:<P>
		Número do erro=<%=err.number%><P>
		Descrição do erro=<%=err.description%><P>
		Help Context=<%=err.helpcontext%><P>" 
		Help Path=<%=err.helppath%><P>
		Native Error=<%=err.nativeerror%><P>
		Source=<%=err.source%><P>
		SQLState=<%=err.sqlstate%><P>
	<%else%>
		<!--Nenhum problema aconteceu!<p>-->
	<%end if
	IF conn.errors.count> 0 then%>
		Ocorreram erros com o Database<P><%'=str%><P>
		<%for counter= 0 to conn.errors.count%>
		Erro #<%=conn.errors(counter).number%><P>
		Descrição -><%=conn.errors(counter).description%><p>
	<%next
	else%>
		<font face="Arial" size=2 color=red>Acesso ao DRDA executado com sucesso.
	<%end if
	call fecha(rs)
End Function

function Visitas(var_pagi)
	dim vis2
	if trim(session(var_pagi))<> "intranet" then 
		str="select * from prod.co37 where co37data=current date and co37pagi='" & var_pagi & "'": Call abre(str,rs)
		if rs.eof then
			Call fecha(rs)
			str = "insert into prod.co37 (co37pagi, co37data, co37cont) values ('" & var_pagi & "', current date, 1)"
		else
			Call fecha(rs)
			str = "update prod.co37 set co37cont=(co37cont+1) where co37data=current date and co37pagi='" & var_pagi & "'"
		end if
		Call abre(str,rs)
'	response.write "<br><br><br><br><br><br><br><br><br>"&var_pagi
'	response.write "<br><br><br><br><br><br><br><br><br>"&session(var_pagi)
	end if
	session(var_pagi)="intranet"
end function

'ACCESS
'Set rs_nome = Server.CreateObject("ADODB.RECORDSET")
'DBQ = Server.mappath(Request.ServerVariables("PATH_INFO"))'endereço atual
'set Object = Server.CreateObject("Scripting.FileSystemObject")'Cria Object
'DBQ = Object.GetParentFolderName(Object.GetParentFolderName(Object.GetParentFolderName(DBQ)))'tira o último diretório
'DBQ = DBQ & "../dados/xxx.mdb"
%>