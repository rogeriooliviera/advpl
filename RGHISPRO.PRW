#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOPCONN.CH"
#INCLUDE "RWMAKE.CH"
#INCLUDE "tbiconn.ch
#Include "Fileio.ch"
#Define ENTER Chr(13)+Chr(10)

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  � RGHISPRO													  ���
�������������������������������������������������������������������������͹��
���Desc.     � Retorna o historico do Produto Selicionado.          	  ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � 					                                          ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

User Function RGHISPRO( )

Local aSays		:=	{}
Local aButtons 	:= 	{}
Local nOpca 	:= 	0
Local cCadastro	:= 	"HITORICO DOS PRODUTOS SELECIONADOS"
Local nDataVld	:=  0
Private cPerg := padr("RGHISPRO",Len(SX1->X1_GRUPO))
Private _aTabelas	:= {}

//CRIAR PERGUNTAS CASO NAO EXISTIR
ValidPerg()

AADD(aSays,"Prezado(a) Sr.(a) "+UPPER(AllTrim(PswRet()[1][4])))
AADD(aSays,"Esta rotina tem como objetivo permitir a sele��o de 15 produtos ")
AADD(aSays,"que voc� deseja visualizar o seu hist�rico de compras...")
AADD(aSays,"")
AADD(aSays,"")
AADD(aSays,"Espec�fico Compras")
AADD(aButtons, { 5,.T.,{|| pergunte(cPerg,.T.) } } )
AADD(aButtons, { 1,.T.,{|o| nOpca:= 1,o:oWnd:End() } } )
AADD(aButtons, { 2,.T.,{|o| o:oWnd:End() }} )

FormBatch( cCadastro, aSays, aButtons )

//CLICOU EM CANCELAR, RETURN (NAO PROCESSA)
If nOpca == 1
	
	Processa({|| _aTabelas := PROCPRO(),cCadastro,"",.F.})
Else
	Return Nil
EndIf

If len(_aTabelas) > 0
	MsgRun("Gerando arquivo Excel, aguarde...","Processando ",{||CursorWait(),QuerySQL(),CursorArrow()})
	//msaguarde({||QuerySQL()},"Gerando arquivo")
Endif


Return NIL

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �PROCPRO   �Autor  �Rogerio Oliveira	 � Data �  29/06/16   ���
�������������������������������������������������������������������������͹��
���Desc.     � Selecao dos Produtos que ira compor o relatorio            ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
**********************
Static Function PROCPRO()
**********************

Local cGrupo	:= space(12)
Local cMarca	:= GetMark()
local nOpca 	:= 1
local lInverte	:= .F.
local aEmpTroca	:= '00'
local _nRecSM0 	:= ZC4->(recno())
Local oDlg
local _aRet		:= {}
Local aCampos	:= {{"ZC4_OK","C",02,0 },{"ZC4_CODGEN","C",16,0 },{"ZC4_DESC","C",20,0 },{"ZC4_GRUPO","C",15,0 }}
Local aCampos2 	:= {{"ZC4_OK",,"  ",""},{"ZC4_CODGEN",,"Cod. Produto","@X"},{"ZC4_DESC",,"Descri��o","@X"},{"ZC4_GRUPO",,"Grupo","@X"}}

if select("PROCPRO") > 0
	PROCPRO->(dbclosearea())
endif
dbUseArea(.T.,,CriaTrab(aCampos,.T.),"PROCPRO",.F.,.F.)
IndRegua("PROCPRO",CriaTrab(NIL,.F.),"ZC4_CODGEN",,,"Indexando Produtos")

// Carga do Alias EMPRESA com os dados do SM0
ZC4->(dbgotop())
while ZC4->(!eof())
	
	If aEmpTroca <> ZC4->ZC4_CODGEN
		
		RecLock("PROCPRO",.F.)
			cGrupo := ZC4->ZC4_GRUPO  
		MsUnlock()	
		
		If mv_par03 == 1 .AND. ALLTRIM(cGrupo) == 'MERC' 
		    RecLock("PROCPRO",.T.)
			PROCPRO->ZC4_CODGEN	:= ZC4->ZC4_CODGEN
			PROCPRO->ZC4_DESC 	:= ZC4->ZC4_DESC
			PROCPRO->ZC4_GRUPO 	:= ZC4->ZC4_GRUPO
			MsUnlock()	
		ElseIf mv_par04 == 1 .AND. ALLTRIM(cGrupo) == 'ORTE'  
			RecLock("PROCPRO",.T.)
			PROCPRO->ZC4_CODGEN	:= ZC4->ZC4_CODGEN
			PROCPRO->ZC4_DESC	:= ZC4->ZC4_DESC
			PROCPRO->ZC4_GRUPO 	:= ZC4->ZC4_GRUPO
			MsUnlock()
		ElseIf mv_par05 == 1 .AND. ALLTRIM(cGrupo) == 'ACOU'
			RecLock("PROCPRO",.T.)
			PROCPRO->ZC4_CODGEN	:= ZC4->ZC4_CODGEN
			PROCPRO->ZC4_DESC	:= ZC4->ZC4_DESC
			PROCPRO->ZC4_GRUPO 	:= ZC4->ZC4_GRUPO 
			MsUnlock()
		ElseIf mv_par06 == 1 .AND. ALLTRIM(cGrupo) == 'UTENS'  
			RecLock("PROCPRO",.T.)
			PROCPRO->ZC4_CODGEN	:= ZC4->ZC4_CODGEN
			PROCPRO->ZC4_DESC	:= ZC4->ZC4_DESC
			PROCPRO->ZC4_GRUPO 	:= ZC4->ZC4_GRUPO
			MsUnlock() 
		ElseIf mv_par07 == 1 //DEMAIS GRUPOS
		    IF ALLTRIM(cGrupo) <> 'MERC' .AND. ALLTRIM(cGrupo) <> 'ORTE' .AND. ALLTRIM(cGrupo) <> 'ACOU' .AND. ALLTRIM(cGrupo) <> 'UTENS'
		    RecLock("PROCPRO",.T.)
			PROCPRO->ZC4_CODGEN	:= ZC4->ZC4_CODGEN
			PROCPRO->ZC4_DESC	:= ZC4->ZC4_DESC
			PROCPRO->ZC4_GRUPO 	:= ZC4->ZC4_GRUPO
			MsUnlock()
			EndIF 
		EndIf
		
		aEmpTroca := ZC4->ZC4_CODGEN
	endif
	
	ZC4->(dbskip())
enddo
ZC4->(dbgoto(_nRecSM0))

PROCPRO->( dbGotop() )
DEFINE MSDIALOG oDlg TITLE "Selecione os Produtos que dever�o ser consideradas" From 009,000 To 030,063 OF oMainWnd
oMark := MsSelect():New("PROCPRO","ZC4_OK","",aCampos2,@lInverte,@cMarca,{20,2,140,248})
oMark:oBrowse:bAllMark := {|| PROCPRO->(DBEVAL({||RecLock("PROCPRO",.F.),PROCPRO->ZC4_OK := iif(empty(PROCPRO->ZC4_OK),cMarca,""),MsUnlock()})), PROCPRO->(dbgotop())}
ACTIVATE MSDIALOG oDlg ON INIT EnchoiceBar(oDlg,{|| nOpca := 1,oDlg:End()},{|| nOpca := 2,oDlg:End()}) CENTERED

if nOpca == 1 // Confirmou o processamento
	PROCPRO->(dbgotop())
	while PROCPRO->(!eof())
		dbselectarea("PROCPRO")
		if IsMark("ZC4_OK",cMarca,lInverte)
			
			//Carrega o Array com os dados necessarios
			AADD( _aRet, {PROCPRO->ZC4_CODGEN, PROCPRO->ZC4_DESC, alltrim(PROCPRO->ZC4_GRUPO)})
			
		endif
		PROCPRO->(dbskip())
	enddo
endif

PROCPRO->(dbclosearea())
Return(_aRet)



/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �QuerySQL  �Autor  �Rogerio Oliveira    � Data �  30/06/16   ���
�������������������������������������������������������������������������͹��
���Desc.     � Extracao dos dados via SQL                                 ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function QuerySQL()

Local _aItem	:= {} //Itens
Local _aHeader	:= {} //Cabecalho
Local _cQuery 	:= '' //Query
Local _ni       := 0  //Contador
Local CR 		:= chr(13)+chr(10)

//MONTA A QUERY
_cQuery := " SELECT ZC3_NUM, ZC3_ITEM, ZC3_EMISSA, ZC3_COD, ZC3_DESC,(SELECT ZC4_GRUPO FROM ZC4990 WHERE ZC4_CODGEN = ZC3_CODGEN AND D_E_L_E_T_ <> '*') GRUPO, "
_cQuery += " ZC3_QTD, ZC3_PRECO, ZC3_TOTAL, ZC3_NOME, ZC3_END  " 									   +CR
_cQuery += " FROM "+RetSqlName("ZC3")+ " ZC3 WHERE D_E_L_E_T_ <> '*' "							   	   +CR
If !Empty(MV_PAR01) .and. !Empty(MV_PAR02)
	_cQuery += " AND ZC3_EMISSA BETWEEN  '" +DTOS(MV_PAR01)+ "' AND '" +DTOS(MV_PAR02) + "' "		   +CR
EndIf
_cQuery += " AND ZC3.ZC3_UNIDAD <> 'DVERSUS' "													       +CR
_cQuery += " AND ZC3_CODGEN IN ("																 	   +CR

For _ni := 1 to len(_aTabelas)
	_cQuery += "'" + _aTabelas[_ni][1]+"'"
	If  _ni < len(_aTabelas)
		_cQuery += ","
	End If
Next _ni
_cQuery += ") ORDER BY ZC3_COD, ZC3_PRECO

If Select("TSD1") <> 0 // Fechar caso aberto
	TSD1->(dbCloseArea())
Endif

ChangeQuery(_cQuery)
TCQuery _cQuery Alias "TSD1" New
TCSETFIELD("TSD1","ZC3_EMISSA","D")
TCSETFIELD("TSD1","ZC3_PRECO", "N", 17, 2)
TCSETFIELD("TSD1","ZC3_TOTAL", "N", 17, 2)
TSD1->(dbGoTop())

dbSelectArea("TSD1")  // Posicionar no topo
dbGotop()
IF Eof()
	MsGBox("N�o h� dados !!","Aten��o")
	Return
Endif

//MONTA O ARRAY DE ITENS aItem
While !Eof()
	aAdd(_aItem,{;
	chr(160)+ALLTRIM(TSD1->ZC3_NUM), chr(160)+TSD1->ZC3_ITEM, TSD1->ZC3_EMISSA, chr(160)+TSD1->ZC3_COD, TSD1->ZC3_DESC, TSD1->GRUPO, ;
	TSD1->ZC3_QTD, TSD1->ZC3_PRECO, TSD1->ZC3_TOTAL, TSD1->ZC3_NOME, TSD1->ZC3_END ;
	,""})
	dbSkip()
Enddo

//MONTA O CABECALHO aHeader
_aHeader := {}

AADD(_aHeader, {"ZC3_NUM"			,"C",10, 0})
AADD(_aHeader, {"ZC3_ITEM"	  		,"C",10, 0})
AADD(_aHeader, {"ZC3_EMISSA"		,"C",10, 0})
AADD(_aHeader, {"ZC3_COD"			,"C",20, 0})
AADD(_aHeader, {"ZC3_DESC"			,"C",40, 0})
AADD(_aHeader, {"ZC4_GRUPO"			,"C",40, 0})
AADD(_aHeader, {"ZC3_QTD" 			,"C",10, 0})
AADD(_aHeader, {"ZC3_PRECO" 	    ,"C",10, 0})
AADD(_aHeader, {"ZC3_TOTAL"		    ,"C",10, 0})
AADD(_aHeader, {"ZC3_NOME"		    ,"C",40, 0})
AADD(_aHeader, {"ZC3_END"		    ,"C",20, 0})

DlgToExcel({{"GETDADOS","RELAT�RIO DE PRODUTOS, DE ACORDO COM OS PAR�TROS DEFINIDOS...",_aHeader,_aItem}})

Return( Nil )



/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �VALIDPERG �Autor  �Rogerio Oliveira	 � Data �  23/05/07   ���
�������������������������������������������������������������������������͹��
���Desc.     � Cria as perguntas do SX1                                   ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function VALIDPERG()

// ZC4->ZC4_GRUPO,cOrdem,cPergunt,cPerSpa,cPerEng,cVar,cTipo ,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3, cGrpSxg,cPyme,;
// cVar01,cDef01,cDefSpa1,cDefEng1,cCnt01,cDef02,cDefSpa2,cDefEng2,cDef03,cDefSpa3,cDefEng3,cDef04,cDefSpa4,cDefEng4, cDef05,cDefSpa5,cDefEng5
PutSX1(cPerg,"01","DT Emiss�o Inicial?"			,"DT Emiss�o Inicial?"	    	,"DT Emiss�o Inicial?"		   ,"mv_ch1","D",08,0,0,"G","",""		,"",,"mv_par01","","","","","","","","","","","","","","","","")
PutSX1(cPerg,"02","DT Emiss�o Final?"			,"DT Emiss�o Final?"		    ,"DT Emiss�o Final?"		   ,"mv_ch2","D",08,0,0,"G","",""		,"",,"mv_par02","","","","","","","","","","","","","","","","")
PutSX1(cPerg,"03","Retorna Grupo MERC? *"	   	,"Retorna Grupo MERC? *"	    ,"Retorna Grupo MERC? *"  	   ,"mv_ch3","N",01,0,0,"C","",""		,"",,"mv_par03","Sim","","","","Nao","","","","","","","","","","","")
PutSX1(cPerg,"04","Retorna Grupo ORTE? *"	   	,"Retorna Grupo ORTE? *"	    ,"Retorna Grupo ORTE? *"  	   ,"mv_ch4","N",01,0,0,"C","",""		,"",,"mv_par04","Sim","","","","Nao","","","","","","","","","","","")
PutSX1(cPerg,"05","Retorna Grupo ACOU? *"	   	,"Retorna Grupo ACOU? *"	    ,"Retorna Grupo ACOU? *"  	   ,"mv_ch5","N",01,0,0,"C","",""		,"",,"mv_par05","Sim","","","","Nao","","","","","","","","","","","")
PutSX1(cPerg,"06","Retorna Grupo UTENS? *"	   	,"Retorna Grupo UTENS? *"	    ,"Retorna Grupo UTENS? *"  	   ,"mv_ch6","N",01,0,0,"C","",""		,"",,"mv_par06","Sim","","","","Nao","","","","","","","","","","","")
PutSX1(cPerg,"07","Retorna os Demais Grupos? *"	,"Retorna os Demais Grupos?  *" ,"Retorna os Demais Grupos? *" ,"mv_ch7","N",01,0,0,"C","",""		,"",,"mv_par07","Sim","","","","Nao","","","","","","","","","","","")

Return( Nil )



