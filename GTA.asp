<%@language="vbscript"%>

<!--#include file="../asp/_conexao.asp"-->

<%
dim rFSO, rArquivo, wFSO, wArquivo
dim sReadLine, sLineAux, sLinha, sArqTxt, sBtxt
dim dAuxData
dim valida1, valida2, valida3
dim xData
dim wFaz

dim dUltima, dArrec, sCodigo, wPriLote

dim str3, rs3, sLinh


sLineAux=""
sBtxt=""

dAuxData=""

valida1=0
valida2=0
valida3=0

wFaz="N"
wPriLote="S"

dUltima=""
dArrec=""
sCodigo=""


sArqTxt = "c:\_asp\benarrec-"
'o Rogerio Scaion adaptou o caminho abaixo:'
'sArqTxt = "\\SERVAPLERPP\Benner\Integrator\ERP_ARRECADACAO\benarrec-" '


'Gravação de txt c/Totais de Arrecadação'

CALL checa_se_ha_dados()
dAuxData = dUltima

if dAuxData<>"" then
   sBtxt=sArqTxt & dAuxData & ".txt"
end if


valida3=1
dim fs
set fs=Server.CreateObject("Scripting.FileSystemObject")
if fs.FileExists( sBtxt ) then
   valida3=0
   response.write("ja existe "& sBtxt)
else
   response.write("novo " & sBtxt)
end if


if (dAuxData<>"" and sBtxt<>"" and valida1>0 and valida2>0 and valida3>0 ) then

   wFaz = "N"

   Set wFSO=Server.CreateObject("Scripting.FileSystemObject")
   Set wArquivo=wFSO.OpenTextFile( sBtxt,2,true)

   ' fazer leituras e gravações '


   str3="select tp02codi, tp02linh, tp02tmst from prod.tp02 where tp02codi="&sCodigo&" order by tp02codi, tp02tmst"
   CALL AbreMF(str3,rs3)

   if rs3.eof = false then
      do while (not rs3.eof)

         sLinh = rs3("tp02linh")

         if mid(sLinh, 2, 20) = "K9_TS_LOTESBANCARIOS"  then

            if wPriLote = "N"  then
	           wArquivo.writeline "#NEWPACK"
               wArquivo.writeline " "
	        end if

            wArquivo.writeline "[" & mid(sLinh,  2, 20) & "]"

            wArquivo.writeline mid(sLinh, 23, 10)
            wArquivo.writeline mid(sLinh, 33,  9)
            wArquivo.writeline mid(sLinh, 42, 18)
            wArquivo.writeline mid(sLinh, 60, 31)
            wArquivo.writeline mid(sLinh, 91, 81)
            wArquivo.writeline mid(sLinh,172, 11)
            wArquivo.writeline mid(sLinh,183, 17)
            wArquivo.writeline mid(sLinh,200, 58)
            wArquivo.writeline mid(sLinh,258, 29)
            wArquivo.writeline mid(sLinh,287, 27)
            wArquivo.writeline mid(sLinh,314, 11)
            wArquivo.writeline mid(sLinh,325, 22)

         end if


         if mid(sLinh, 2, 16) = "K9_TS_TRANSACOES"  then
            wArquivo.writeline " "

            wArquivo.writeline "[" & mid(sLinh,  2, 16) & "]"

            wArquivo.writeline mid(sLinh, 19, 10)
            wArquivo.writeline mid(sLinh, 29, 35)
            wArquivo.writeline mid(sLinh, 64, 11)
            wArquivo.writeline mid(sLinh, 75, 11)
            wArquivo.writeline mid(sLinh, 86, 22)
            wArquivo.writeline mid(sLinh,108, 22)
            wArquivo.writeline mid(sLinh,130, 58)
            wArquivo.writeline mid(sLinh,188, 38)
            if mid(sLinh,226, 10) <> "          " then
               wArquivo.writeline mid(sLinh,226, 31)
            end if
            wArquivo.writeline mid(sLinh,257, 24)
            wArquivo.writeline mid(sLinh,281, 15)
            wArquivo.writeline mid(sLinh,296, 30)

         end if
         wPriLote = "N"

         rs3.movenext
      loop
      wArquivo.writeline "#NEWPACK"

   end if


' FIM'


end if
%>


</html>


<%
'Verifica se existem registros a serem selecionados na data'
'------------------------------- TP01/TP02'

sub checa_se_ha_dados()
    dim str1, rs1, str2, rs2

    str1="select * from prod.tp01 where tp01lote ='BTARREC' and tp01flag='P' order by tp01data desc"
    CALL AbreMF(str1,rs1)

    if rs1.eof = false then

       dArrec = rs1("tp01data")
       hArrec = rs1("tp01hora")
       sCodigo = rs1("tp01codi")
       valida1 = 1

       str2="select tp02codi, tp02linh  from prod.tp02 where tp02codi="&sCodigo
       CALL AbreMF(str2,rs2)

       if rs2.eof = false then
          dArrec = rs2("tp02linh")

          dUltima = mid(dArrec, 192, 4)&"-"&mid(dArrec, 196, 2)&"-"&mid(dArrec, 198, 2)
          valida2 = 1
       else
          valida2 = 0
       end if
	else
	   valida1 = 0
	end if

	CALL Fecha(rs1)

end sub
%>
