
#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOTVS.CH"
#INCLUDE "TOPCONN.CH"  //INCLUDE do TCQuery
#INCLUDE "RWMAKE.CH"   //INCLUDE do MsGBox
#INCLUDE "tbiconn.ch
#INCLUDE "Fileio.ch"

#define PAD_LEFT	0
#define PAD_RIGHT	1
#define PAD_CENTER  2
#define ENTER Chr(13)+Chr(10)

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณ RGHISGRAF												  บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณ Retorna o historico do Produto Selicionado.          	  บฑฑ
ฑฑบ          ณ                                                            บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ 					                                          บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
User Function RGHISGRAF( )

Local aSays		:=	{}
Local aButtons 	:= 	{}
Local nOpca 	:= 	0
Local cCadastro	:= "Rel. Analํtico de Compras."
Local _nY		:=  0  //Contador for
Local nAscan    :=  0

Private cPerg := padr("RGHISGRAF",Len(SX1->X1_GRUPO))
Private _aTabelas	:= {}

Private cDataIN := ''

//RELATำRIO
Private cTitulo		:= "RELATำRIO DE ANALISE DOS PEDIDOS DE COMPRA..."
Private oPrint		:= TMSPrinter():New(cTitulo)
Private CR 			:= chr(13) + chr(10)
Private cQueryA		:= ""			// Query Auxiliar do relatorio
Private dDataRel	:= dDataBase	// Pega a DATA do dia para processamento
Private aDatas		:= {}			// Array para buscar datas
Private aCalGer		:= {}			// Array para calculo Mensal
Private nTotal	    := 0            // Total  
Private nMedia		:= 0            // Media

//FONTES
Private oFont07a		:= TFont():New('Arial',,07,,.F.,,,,.F.,.F.)	//Negrito/Italico/Sublinhado
Private oFont07b		:= TFont():New('Arial',,07,,.T.,,,,.F.,.F.)	//Negrito/Italico/Sublinhado
Private oFont08b		:= TFont():New('Arial',,08,,.T.,,,,.F.,.F.)	//Negrito/Italico/Sublinhado
Private oFont09b		:= TFont():New('Arial',,09,,.T.,,,,.F.,.F.)	//Negrito/Italico/Sublinhado
Private oFont10a		:= TFont():New('Arial',,10,,.F.,,,,.F.,.F.)	//Negrito/Italico/Sublinhado
Private oFont10b		:= TFont():New('Arial',,10,,.T.,,,,.F.,.F.)	//Negrito/Italico/Sublinhado
Private oFont11a		:= TFont():New('Arial',,11,,.F.,,,,.F.,.F.)	//Negrito/Italico/Sublinhado
Private oFont11b		:= TFont():New('Arial',,11,,.T.,,,,.F.,.F.)	//Negrito/Italico/Sublinhado
Private oFont12a		:= TFont():New('Arial',,12,,.F.,,,,.F.,.F.)	//Negrito/Italico/Sublinhado
Private oFont12b		:= TFont():New('Arial',,12,,.T.,,,,.F.,.F.)	//Negrito/Italico/Sublinhado

//ESTRUTURA DE PAGINAS
Private nLin		:= 10
Private nLinLimite	:= 2350			// Limite da folha A4
Private nLinHalf	:= nLinLimite/2	// Limite da folha A4 //Metade de linhas
Private nColInicio 	:= 100			// Coluna de Inicio
Private nColLimite	:= 3400			// Tamanho limite de coluna. Tamanho da folha.
Private nColHalf	:= nColLimite/2	// Tamanho limite de coluna. Tamanho da folha. //Metade de colunas
Private nPag 	   	:= 1			// Numero de pagina
Private nTotPag     := 1            // Total de paginas
Private nRegs		:= 0


//CRIAR PERGUNTAS CASO NAO EXISTIR
ValidPerg()

AADD(aSays,"Prezado(a) Sr.(a) "+UPPER(AllTrim(PswRet()[1][4])))
AADD(aSays,"Esta programa tem como objetivo permitir a voc๊ visualizar")
AADD(aSays,"o calculo geral das compras pelo periodo selecionado...")
AADD(aSays,"")
AADD(aSays,"")
AADD(aSays,"Especํfico PC")
AADD(aButtons, { 5,.T.,{|| pergunte(cPerg,.T.) } } )
AADD(aButtons, { 1,.T.,{|o| nOpca:= 1,o:oWnd:End() } } )
AADD(aButtons, { 2,.T.,{|o| o:oWnd:End() }} )

FormBatch( cCadastro, aSays, aButtons )

//CLICOU EM CANCELAR, RETURN (NAO PROCESSA)
If !nOpca == 1
	Return Nil
EndIf

//RETORNA O CONTEUDO DAS PERGUNTAS
pergunte(cPerg,.F.)

If nOpca == 1
	pergunte(cPerg,.F.)
	
	cQueryA := " SELECT ZC3_EMISSA FROM ZC3990 WHERE D_E_L_E_T_ <> '*' AND "         +CR
	cQueryA += " ZC3_EMISSA  >= '"+DTOS(MV_PAR01)+"' AND "                           +CR
	cQueryA += " ZC3_EMISSA  <= '"+DTOS(MV_PAR02)+"' "								 +CR
	cQueryA += " GROUP BY ZC3_EMISSA ORDER BY 1 "				  					 +CR
	
	If Select("GRAFZC3") <> 0 // Fechar caso aberto
		GRAFZC3->(dbCloseArea())
	Endif
	
	ChangeQuery(cQueryA)
	TCQuery cQueryA Alias "GRAFZC3" New
	GRAFZC3->(dbGoTop())
	
	dbSelectArea("GRAFZC3")  // Posicionar no topo
	dbGotop()
	IF Eof()
		MsGBox("Nใo hแ dados !!","Aten็ใo")
		Return
	Endif
	
	//MONTA O ARRAY DATAS
	While GRAFZC3->(!Eof())
		nAscan := Ascan(aDatas, { |x| x[1] == SUBSTR(GRAFZC3->ZC3_EMISSA,1, 4)+SUBSTR(GRAFZC3->ZC3_EMISSA,5, 2)})
		If nAscan == 0
			AADD( aDatas, {SUBSTR(GRAFZC3->ZC3_EMISSA,1, 4)+SUBSTR(GRAFZC3->ZC3_EMISSA,5, 2)})
		EndIf
		
		dbSkip()
	Enddo
	
	
	For _nY := 1 to len(aDatas)
		
		//Chama Fun็ใo para calculo mensal de gastos
		 AADD( aCalGer,{aDatas[_nY][1], CALCMENSAL(aDatas[_nY][1]) })
		
	Next _nY
	
	Imprime(aCalGer)
	
endif

Return (Nil)




/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณ CALCMENSALบ	Autor  ณ Rogerio Oliveira	 Data 16/10/2016  บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณ Processa a impressao do Relatorio						  บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Departamento Tributario				  					  บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function CALCMENSAL(cParam)
 
Local cQueryP := ""
Local aArea   := GetArea()
Local nCount  := 0
Local cAlias  := GetNextAlias()

//MONTAGEM DA QUERY PRINCIPAL
cQueryP  := " SELECT ISNULL(SUM(ZC3_TOTAL), 0) AS nCout from ZC3990 Z WHERE D_E_L_E_T_ <> '*' AND "        +CR
cQueryP  += " SUBSTRING(Z.ZC3_EMISSA,1,4)+SUBSTRING(Z.ZC3_EMISSA,5,2) = '"+cParam+"'"  +CR

TCQuery cQueryP Alias &cAlias New
dbSelectArea(cAlias)
dbGoTop()

nCount := (cAlias)->nCout
(cAlias)->(dbCloseArea())
RestArea(aArea)

Return nCount


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณ IMPRIME บ	Autor  ณ Rogerio Oliveira	 บ Data 06/06/14  บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณ Processa a impressao do Relatorio						  บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Departamento Tributario				  					  บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function Imprime(aCalGer)

Local nAux     :=0
Local _nY      :=0
Local cValor01 :=space(20) 
Local cValor02 :=space(20)
Local cValor03 :=space(20)
Local cValor04 :=space(20)
Local cValor05 :=space(20)
Local cValor06 :=space(20)
Local cValor07 :=space(20)
Local cValor08 :=space(20)
Local cValor09 :=space(20)
Local cValor10 :=space(20)
Local cValor11 :=space(20)
Local cValor12 :=space(20)    

cValor01 := "0"
cValor02 := "0"
cValor03 := "0"
cValor04 := "0"
cValor05 := "0"
cValor06 := "0"
cValor07 := "0"
cValor08 := "0"
cValor09 := "0"
cValor10 := "0"
cValor11 := "0"
cValor12 := "0"

IF ( Len(aCalGer)==0 )
	MsGBox("Nใo hแ dados !!","Aten็ใo")
	Return
Endif

oPrint:Setup()			// Abre a tela para selecionar a Impressora
oPrint:SetPaperSize(9)  // Define o tamanho da pagina, A4
//oPrint:SetLandscape()	// Paisagem
oPrint:SetPortrait()	// Retrato
oPrint:StartPage()
ImpCabec()
 

For _nY := 1 to len(aCalGer)  
    	
	If SUBSTR(aCalGer[_nY][1],5,2) ="01"
		cValor01 := cValToChar(aCalGer[_nY][2])
	ElseIf SUBSTR(aCalGer[_nY][1],5,2) ="02"
	    cValor02 := cValToChar(aCalGer[_nY][2])
	ElseIf SUBSTR(aCalGer[_nY][1],5,2) ="03"
	    cValor03 := cValToChar(aCalGer[_nY][2])
	ElseIf SUBSTR(aCalGer[_nY][1],5,2) ="04"   
	    cValor04 := cValToChar(aCalGer[_nY][2])
	ElseIf SUBSTR(aCalGer[_nY][1],5,2) ="05"
	    cValor05 := cValToChar(aCalGer[_nY][2])
	ElseIf SUBSTR(aCalGer[_nY][1],5,2) ="06"   
	    cValor06 := cValToChar(aCalGer[_nY][2])
	ElseIf SUBSTR(aCalGer[_nY][1],5,2) ="07"
	    cValor07 := cValToChar(aCalGer[_nY][2])
	ElseIf SUBSTR(aCalGer[_nY][1],5,2) ="08"
	    cValor08 := cValToChar(aCalGer[_nY][2])
	ElseIf SUBSTR(aCalGer[_nY][1],5,2) ="09"
	    cValor09 := cValToChar(aCalGer[_nY][2])
	ElseIf SUBSTR(aCalGer[_nY][1],5,2) ="10"
	    cValor10 := cValToChar(aCalGer[_nY][2])
	ElseIf SUBSTR(aCalGer[_nY][1],5,2) ="11"
	    cValor11 := cValToChar(aCalGer[_nY][2])
	Else
	    cValor12 := cValToChar(aCalGer[_nY][2])
	EndIf   
	
	nTotal := nTotal+aCalGer[_nY][2]  // Soma todos os valores na variแvel                                              
			
	oPrint:Say(nLin+300,0100,SUBSTR(aCalGer[_nY][1],1,4)		 			,oFont11a,,,,PAD_LEFT) //-- ANO
	oPrint:Say(nLin+300,0300,cValor01	 									,oFont11a,,,,PAD_LEFT) //-- JANEIRO
	oPrint:Say(nLin+300,0500,cValor02								   		,oFont11a,,,,PAD_LEFT) //-- FEVEREIRO
	oPrint:Say(nLin+300,0750,cValor03								   		,oFont11a,,,,PAD_LEFT) //-- MARวO
	oPrint:Say(nLin+300,0950,cValor04								   		,oFont11a,,,,PAD_LEFT) //-- ABRIL
	oPrint:Say(nLin+300,1150,cValor05							 			,oFont11a,,,,PAD_LEFT) //-- MAIO
	oPrint:Say(nLin+300,1350,cValor06 										,oFont11a,,,,PAD_LEFT) //-- JUNHO
	oPrint:Say(nLin+300,1550,cValor07								   		,oFont11a,,,,PAD_LEFT) //-- JULHO
	oPrint:Say(nLin+300,1750,cValor08								   		,oFont11a,,,,PAD_LEFT) //-- AGOSTO
	oPrint:Say(nLin+300,2000,cValor09								   		,oFont11a,,,,PAD_LEFT) //-- SETEMBRO
	oPrint:Say(nLin+300,2250,cValor10							 			,oFont11a,,,,PAD_LEFT) //-- OUTUBRO
	oPrint:Say(nLin+300,2500,cValor11	 									,oFont11a,,,,PAD_LEFT) //-- NOVEMBRO
	oPrint:Say(nLin+300,2750,cValor12								   		,oFont11a,,,,PAD_LEFT) //-- DEZEMBRO
	
	nLin +=50
	If nLin >= 3100
		oPrint:EndPage()
		nPag += 1
		nTotPag := nPag
		nLin := 10
		oPrint:StartPage()
		Impcabec()
	Endif     
	
	cValor01 := "0"
	cValor02 := "0"
	cValor03 := "0"
	cValor04 := "0"
	cValor05 := "0"
	cValor06 := "0"
	cValor07 := "0"
	cValor08 := "0"
	cValor09 := "0"
	cValor10 := "0"
	cValor11 := "0"
	cValor12 := "0"
		
Next _nY

oPrint:Say(nLin+400,0100,"Total =  "					 			    ,oFont11a,,,,PAD_LEFT)
oPrint:Say(nLin+400,0100,space(12)+cValtoChar(nTotal)	     			,oFont11a,,,,PAD_LEFT) 

oPrint:Say(nLin+400,0500,"Media =  "    					 		    ,oFont11a,,,,PAD_LEFT) 
oPrint:Say(nLin+400,0500,space(13)+cValtoChar(nTotal/len(aCalGer))  	,oFont11a,,,,PAD_LEFT) 

oPrint:Say(nLin+500,nColLimite -1500,"Total de Pแginas: "				,oFont11a,,,,PAD_CENTER)
oPrint:Say(nLin+500,nColLimite -1500,space(41)+cValtoChar(nTotPag)	 	,oFont11b,,,,PAD_CENTER)
oPrint:EndPage()

oPrint:Preview()
Return ( Nil )


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณ IMPCABEC บ	Autor  ณ Rogerio Oliveira	 บ Data 06/06/14  บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณ Monta o Cabecalho do Relatorio							  บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Departamento Tributario				  					  บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function ImpCabec()

//oPrint:SayBitmap(nLin+15,0070,cFig,350,100)     // Linha/Coluna/Figura/Comprimento/Largura
oPrint:Say(0100,nColHalf-500,cTitulo						 	,oFont11a,,,,PAD_CENTER)
oPrint:Say(0100,nColLimite-1500,"Pแgina:  "+cValtoChar(nPag) 	,oFont11a,,,,PAD_CENTER)

oPrint:LINE(290,100,290,nColLimite)				// oPrint:LINE(nLinha1,nColuna1,nLinha2,nColuna2)

oPrint:Say(235,0100,"ANOS" 						 			    ,oFont11a,,,,PAD_LEFT)
oPrint:Say(235,0300,"JANEIRO"						 	 		,oFont11a,,,,PAD_LEFT)
oPrint:Say(235,0500,"FEVEREIRO"    					 		    ,oFont11a,,,,PAD_LEFT)
oPrint:Say(235,0750,"MARวO"						 	 	   		,oFont11a,,,,PAD_LEFT)
oPrint:Say(235,0950,"ABRIL"     						 	    ,oFont11a,,,,PAD_LEFT)
oPrint:Say(235,1150,"MAIO" 				 	 		            ,oFont11a,,,,PAD_LEFT)
oPrint:Say(235,1350,"JUNHO" 			 	 		            ,oFont11a,,,,PAD_LEFT) 
oPrint:Say(235,1550,"JULHO"				 	 		            ,oFont11a,,,,PAD_LEFT)   
oPrint:Say(235,1750,"AGOSTO"			 	 		            ,oFont11a,,,,PAD_LEFT)   
oPrint:Say(235,2000,"SETEMBRO"			 	 		            ,oFont11a,,,,PAD_LEFT)  
oPrint:Say(235,2250,"OUTUBRO" 			 	 		            ,oFont11a,,,,PAD_LEFT)     
oPrint:Say(235,2500,"NOVEMBRO"			 	 		            ,oFont11a,,,,PAD_LEFT)   
oPrint:Say(235,2750,"DEZEMBRO" 			 	 		            ,oFont11a,,,,PAD_LEFT)

Return ( Nil )


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณVALIDPERG บAutor  ณRogerio Oliveira	 บ Data ณ  01/10/2016 บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณ Cria as perguntas do SX1                                   บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ        
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function VALIDPERG()

// ZC4->ZC4_GRUPO,cOrdem,cPergunt,cPerSpa,cPerEng,cVar,cTipo ,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3, cGrpSxg,cPyme,;
// cVar01,cDef01,cDefSpa1,cDefEng1,cCnt01,cDef02,cDefSpa2,cDefEng2,cDef03,cDefSpa3,cDefEng3,cDef04,cDefSpa4,cDefEng4, cDef05,cDefSpa5,cDefEng5
PutSX1(cPerg,"01","DT Emissใo Inicial?"			,"DT Emissใo Inicial?"	    	,"DT Emissใo Inicial?"		   ,"mv_ch1","D",08,0,0,"G","",""		,"",,"mv_par01","","","","","","","","","","","","","","","","")
PutSX1(cPerg,"02","DT Emissใo Final?"			,"DT Emissใo Final?"		    ,"DT Emissใo Final?"		   ,"mv_ch2","D",08,0,0,"G","",""		,"",,"mv_par02","","","","","","","","","","","","","","","","")

Return( Nil )


