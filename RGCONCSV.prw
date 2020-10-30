#INCLUDE "rwmake.ch"
#INCLUDE "common.ch"
#INCLUDE "topconn.ch"
#INCLUDE "PROTHEUS.CH"

#define DIR_SERVER_TMP "SERVIDOR\arq_temp\"
//#define DIR_SERVER_TMP "C:\"
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณ RGCONCSV บ	Autor  ณ Rogerio Oliveira	 บ Data 17/02/14  บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณ Rotina para atualizar Cabecalho Tabela de Precos, via CSV  บฑฑ
ฑฑบ          ณ                                                            บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Usada Compras						  					  บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
User Function RGCONCSV()

Private aPergs 		:= {}
Private aRet 		:= {}
Private cDelimit 	:= ';'
Private cArquivo	:= ""
Private cLog		:= ""
Private cUserName	:= ""

// Filiais
Private cFilialZC1	:=  xFilial('ZC1')
Private cFilialZC2	:=  xFilial('ZC2')
Private cFilialZC3	:=  xFilial('ZC3')
Private cFilialZC4	:=  xFilial('ZC4')

// Tabelas
DbSelectArea("ZC1")
DbSelectArea("ZC2")
DbSelectArea("ZC3")
DbSelectArea("ZC4")

//Usuario que esta acessando
cUserName	:= UsrFullName ( __cUserId )

aAdd( aPergs ,{6,"Arquivo : ",cArquivo,"@!",'.T.','.T.',85,.T.,"Arquivos CSV (*.CSV) | *.CSV",DIR_SERVER_TMP})
aAdd( aPergs ,{1,"Caracter Demitador : ",cDelimit,"@!",'.T.',,'.F.',10,.T.})

If !ParamBox(aPergs ," Selecione o Arquivo .CSV",aRet)
	Return Nil
EndIf

cArquivo	:= aRet[1]
cDelimit	:= aRet[2]

//Validando se existe o arquivo informado
If !File(cArquivo)
	MsgAlert("Invแlido! Favor Informar um Arquivo *.CSV vแlido!")
	Return(Nil)
EndIf

IMPCSV()

Return (Nil)


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  IMPCSV  บAutor  ณ Rogerio Oliveira   บ Data ณ   12/03/14     บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณ Importa o cabecalho tabela de Precos .csv para Protheus    บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Compras		                                              บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function IMPCSV()

Private _cFile := ""
Private _xFile := ""
Private _cType := ""
Private _nTot  := 0

nHdl := fopen(cArquivo,0)
_nTot:= fSeek(nHdl,0,2)
fClose(nHdl)

If _nTot > 0
	Processa({||EXEIMP(),,"Importando registros. Aguarde..."})
EndIf

Return(Nil)



/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณEXEIMP  บAutor  ณ Rogerio Oliveira   บ Data ณ  12/03/14     บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณ Tratamento e validacao e gravacao dos Dados do Arquivo .CSVบฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Compras		                                              บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function EXEIMP()

Local nX                := 0
Local cBuffer			:= ""
Local aLinha 			:= {}
Local _cTexto			:= ""
Private aRegistro 		:= {}
Private nTotLinha		:= 0

Begin Transaction

FT_FUSE(cArquivo)

nTotLinha := FT_FLASTREC()

If nTotLinha < 2
	MsgAlert("Arquivo Incompleto!!! Favor verificar...")
	DisarmTransaction()
	Return (Nil)
Endif

FT_FGOTOP()
cBuffer := FT_FREADLN()    // Tratamento para nao adcionar o nome das colunas do arquivo *.CSV
FT_FSKIP()
cBuffer := FT_FREADLN()
ProcRegua(nTotLinha)
Do While !FT_FEOF()
	IncProc()
	
	aLinha 		:= {}
	
	// Executa leitura da linha em cBuffer e identifica os campos pelo delimitador ";"
	If Len(cBuffer) >0
		_cTexto	:= ""
		For nX 	:= 1 to Len(cBuffer)
			If Substr(cBuffer,nX,1) == cDelimit //cDelimit Determina o carater delimitador.
				aAdd(aLinha, _cTexto )
				_cTexto := ""
			Else
				_cTexto += Substr(cBuffer,nX,1)
			Endif
		Next nX
		If !Empty(_cTexto)
			aAdd(aLinha, _cTexto )
		Else
			aAdd(aLinha, " " )
		Endif
	Endif
	                                                       
	//If Len(aLinha) == 19 //Padrใo definido para 19 Colunas.
	If Len(aLinha) == 21 //Padrใo definido para 21 Colunas.
		aAdd( aRegistro, { ;
		aLinha[01], ;	// PEDIDO
		aLinha[02], ; 	// ITEM
		aLinha[03], ; 	// NF
		aLinha[04], ; 	// EMISSAO_NF
		aLinha[05], ; 	// COD_PROD_GEN
		aLinha[06], ; 	// DESC_PROD_GENERICA
		aLinha[07], ; 	// COD_PROD
		aLinha[08], ; 	// DESC_PRODUTO
		aLinha[09], ; 	// MARCA
		aLinha[10], ; 	// GRUPO
		aLinha[11], ; 	// QUANTIDADE
		aLinha[12], ; 	// UNIDADE
		aLinha[13], ; 	// PREวO_UNI
		aLinha[14], ; 	// TOTAL
		aLinha[15], ; 	// COD_FOR
		aLinha[16], ; 	// LOJA_FOR
		aLinha[17], ; 	// CNPJ
		aLinha[18], ; 	// FORNECEDOR  
		aLinha[19], ; 	// ENDERECO DO FORNECEDOR  
		aLinha[20], ; 	// CIDADE
		aLinha[21] } )	// UF
	Else
		MsgAlert("O arquivo nใo segue o Layout definido!!!")
		DisarmTransaction()
		Return (Nil)
	Endif
	
	FT_FSKIP()
	cBuffer := FT_FREADLN()
Enddo
FT_FUSE()

End Transaction

If(VLDREGIS(aRegistro)) 	// VALIDA OS REGISTROS ANTES DA IMPORTAวรO
	
	UPFOR(aRegistro)        // UPDATE DADOS DE FORNECEDOR/DEMAIS
EndIf

//If !Empty(cArquivo) //DELETAR ARQUIVO .CSV
//	FErase(cArquivo)
//EndIF

Return (Nil)



/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณ VLDREGIS   บAutor  ณ Rogerio Oliveira   บ Data ณ  05/08/15 บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณ VALIDA DADOS - "VERIFICA SE A COLUNAS EM BRANCO"           บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Compras   		                                          บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function VLDREGIS(aRegistro)

Local nY   := 0
Local lRet := .T.
ProcRegua(Len(aRegistro))

For nY := 1 to Len(aRegistro)
	
	IncProc()
	
	If Empty(aRegistro[nY,01]) 		 // PEDIDO
		MsgAlert("O N๚mero do Pedido nใo Informado na linha..: "+cValToChar(nY))
		lRet := .F.
		Return (Nil)
	ElseIf Empty(aRegistro[nY,02])   // ITEM
		MsgAlert("O Item do Pedido nใo Informado na linha..: "+cValToChar(nY))
		lRet := .F.
		Return (Nil)
	ElseIf Empty(aRegistro[nY,03])   // NF
		MsgAlert("O da NF nใo Informado na linha..: "+cValToChar(nY))
		lRet := .F.
		Return (Nil)
	ElseIf Empty(aRegistro[nY,04])   // EMISSAO_NF
		MsgAlert("O Data de Emissใo da NF nใo Informado na linha..: "+cValToChar(nY))
		lRet := .F.
		Return (Nil)
		//ElseIf Empty(aRegistro[nY,05]) // COD_PROD_GEN
		
	ElseIf Empty(aRegistro[nY,06])   // DESC_PROD_GENERICA
		MsgAlert("A Descri็ใo Generica do Produto nใo Informado na linha..: "+cValToChar(nY))
		lRet := .F.
		Return (Nil)
		//ElseIf Empty(aRegistro[nY,07]) // COD_PROD
		
	ElseIf Empty(aRegistro[nY,08])   // DESC_PRODUTO
		MsgAlert("A Descri็ใo do Produto nใo Informado na linha..: "+cValToChar(nY))
		lRet := .F.
		Return (Nil)
	ElseIf Empty(aRegistro[nY,09])   // MARCA
		MsgAlert("A Marca do Produto nใo Informado na linha..: "+cValToChar(nY))
		lRet := .F.
		Return (Nil)
	ElseIf Empty(aRegistro[nY,10])   // GRUPO
		MsgAlert("O Grupo do Produto nใo Informado na linha..: "+cValToChar(nY))
		lRet := .F.
		Return (Nil)
	ElseIf Empty(aRegistro[nY,11])   // QUANTIDADE
		MsgAlert("A Quantidade nใo Informado na linha..: "+cValToChar(nY))
		lRet := .F.
		Return (Nil)
	ElseIf Empty(aRegistro[nY,12])   // UNIDADE
		MsgAlert("A Unidade de Medidas nใo Informado na linha..: "+cValToChar(nY))
		lRet := .F.
		Return (Nil)
	ElseIf Empty(aRegistro[nY,13])   // PREวO_UNI
		MsgAlert("O Pre็o nใo Informado na linha..: "+cValToChar(nY))
		lRet := .F.
		Return (Nil)
	ElseIf Empty(aRegistro[nY,14])   // TOTAL
		MsgAlert("O Valor Total nใo Informado na linha..: "+cValToChar(nY))
		lRet := .F.
		Return (Nil)
		//ElseIf Empty(aRegistro[nY,15]) // COD_FOR
		
		//ElseIf Empty(aRegistro[nY,16]) // LOJA_FOR
		
	ElseIf Empty(aRegistro[nY,17])   // CNPJ
		MsgAlert("O CNPJ do Fornec. nใo Informado na linha..: "+cValToChar(nY))
		lRet := .F.
		Return (Nil)
	ElseIf Empty(aRegistro[nY,18])	 // FORNECEDOR
		MsgAlert("O Nome do Fornec. nใo Informado na linha..: "+cValToChar(nY))
		lRet := .F.
		Return (Nil) 
   ElseIf Empty(aRegistro[nY,19])	 // ENDERECO DO FORNECEDOR
		MsgAlert("O Endere็o do Fornec. nใo Informado na linha..: "+cValToChar(nY))
		lRet := .F.
		Return (Nil)  
  	ElseIf Empty(aRegistro[nY,20])	 // CIDADE DO FORNECEDOR
		MsgAlert("A Cidade do Fornec. nใo Informado na linha..: "+cValToChar(nY))
		lRet := .F.
		Return (Nil)				
	ElseIf Empty(aRegistro[nY,21])   // ESTADO DO FORNECEDOR
		MsgAlert("O Estado(UF) do Fornec. nใo Informado na linha..: "+cValToChar(nY))
		lRet := .F.
		Return (Nil)
	EndIf
	
Next nY

Return(lRet)


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  UPFOR      บAutor  ณ Rogerio Oliveira   บ Data ณ  12/03/14   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณ Atualisando Base de Dados			        			  บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Compras		                                              บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function UPFOR(aRegistro)

Local nY        := 0
Local cSeq		:= "1"
Local cAux		:= space(2)
Private cEndFor := space(40)
Private cNomFor := space(40)
Private cCGCFor := space(14)
Private cCodFor := space(6)
Private cLojFor := space(2)  
Private cCidade := space(40)
Private cUfFor  := space(2)

/*
Obs. A geracao do LOG esta seguindo a sequencia de linhas da planinha,
ou seja, acrescenta uma linha a mais... por cusa do cabecalho da planilha...
*/

ProcRegua(Len(aRegistro))


ZC1->(DbGotop())
ZC1->(DbSetOrder(2))
//Executando a Gravacao dos Dados de Fornecedor
For nY := 1 to Len(aRegistro)
	
	IncProc()
	
	//cCodFor  := StrZero ( Val(aRegistro[nY,12]), 6 ) //FORNECEDOR
	//cLojFor  := StrZero ( Val(aRegistro[nY,13]), 2 ) //LOJA
	cLojFor  := "01"								   //LOJA
	cCGCFor  := aRegistro[nY,17]  					   //CNPJ DO FORNECEDOR
	cNomFor	 := aRegistro[nY,18]                       //NOME DO FORNECEDOR
	cEndFor  := aRegistro[nY,19]                       //ENDERECO DO FORNECEDOR 
	cCidade  := aRegistro[nY,20]                       //CIDADE DO FORNECEDOR 
	cUfFor   := aRegistro[nY,21]                       //ESTADO(UF) DO FORNECEDOR 
	
	If ZC1->( !MsSeek(cFilialZC1+cCGCFor) )
		
		If RecLock("ZC1", .T.) // Incluir
			      
			//STRZERO(VAL(nY+1),9)
			cLog := "LINHA -"+STRZERO(nY+1,9)+ " -Fornecedor/CNPJ  "+ALLTRIM(cNomFor)+space(10)+"       -"+ALLTRIM(cCGCFor)+"  Processado em " + DToC( MSDate() ) + " เs " + Time() + CRLF
			GrvLog( cLog, cSeq )
			
			cAux	        := VldForLoj(cNomFor)
			ZC1->ZC1_FILIAL := cFilialZC1
			ZC1->ZC1_CODFOR := GetSXENum("ZC1","ZC1_CODFOR")
			ZC1->ZC1_LOJFOR := IIf(Empty(cAux),cLojFor,cAux)
			ZC1->ZC1_CNPJ   := cCGCFor
			ZC1->ZC1_NOME   := cNomFor
			ZC1->ZC1_ENDFOR := cEndFor
			ZC1->ZC1_CIDFOR := cCidade 
			ZC1->ZC1_UFFOR  := cUfFor
			
			
			aRegistro[nY,15] := ZC1_CODFOR
			aRegistro[nY,16] := ZC1_LOJFOR
			
			ZC1->( MsUnlock() )  // Destravar o registro
			ZC1->(ConfirmSX8())	 // para executar a grava็ใo dos dados
			
		EndIf
		
	Else // Carregar o Arry com os dados de Fornecedor
		
		If RecLock("ZC1", .F.)
			
			aRegistro[nY,15] := ZC1_CODFOR 	// codigo
			aRegistro[nY,16] := ZC1_LOJFOR 	// loja
			aRegistro[nY,17] := ZC1_CNPJ   	// CNPJ
			aRegistro[nY,18] := ZC1_NOME   	// nome
			aRegistro[nY,19] := ZC1_ENDFOR 	// endereco
			aRegistro[nY,20] := ZC1_CIDFOR 	// cidade
			aRegistro[nY,21] := ZC1_UFFOR 	// uf
			
			MsUnlock()       // destravar o registro
		Endif
		
	EndIf
	
Next nY

//CHAMA ROTINA QUE GERA O COD. E DES. GENษRICA DOS PRODUTOS
UPRODGE(aRegistro)

//CHAMA ROTINA QUE IMPORTA OS PRODUTOS
UPPROD(aRegistro)

//CHAMA ROTINA QUE IMPORTA OS DADOS DO PEDIDO
UPPECOM(aRegistro)

Return(Nil)




/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณ UPPROD  ณ Rogerio Oliveira   บ Data ณ  27/07/16            บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.      Funcao que realiza o Cadastrado de Produto				  บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Compras   		                                          บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function UPRODGE(aRegistro)

Local nY        := 0
Local cSeq		:= "2"
Local cDesc		:= space(50)
Local cGrupo    := space(12)

ProcRegua(Len(aRegistro))

ZC4->(DbGotop())
ZC4->(DbSetOrder(2))
//Executando a Gravacao dos Dados de Fornecedor
For nY := 1 to Len(aRegistro)
	
	IncProc()
	
	cDesc    := aRegistro[nY,6]    //DESC_PROD_GENERICA
	cGrupo   := aRegistro[nY,10]   //GRUPO
	
	If ZC4->( !MsSeek(cFilialZC4+cDesc) )
		
		If RecLock("ZC4", .T.) // Incluir 
		   
		   //cLog := "LINHA -"+cValToChar(nY+1)+
		    cLog := "LINHA -"+STRZERO(nY+1,9)+ " - Descricao Generica..: "+ALLTRIM(cDesc)+space(4)+"       -Processado em " + DToC( MSDate() ) + " เs " + Time() + CRLF
			GrvLog( cLog, cSeq )
			
			ZC4->ZC4_FILIAL := cFilialZC4
			ZC4->ZC4_CODGEN := GetSXENum("ZC4","ZC4_CODGEN")
			ZC4->ZC4_DESC 	:= cDesc
			ZC4->ZC4_GRUPO  := cGrupo
			
			aRegistro[nY,5] := ZC4_CODGEN
			
			ZC4->( MsUnlock() )  // Destravar o registro
			ZC4->(ConfirmSX8())	 // para executar a grava็ใo dos dados
			
		EndIf
		
	Else // Carregar o Arry com os dados de Fornecedor
		
		If RecLock("ZC4", .F.)
			
			aRegistro[nY,5] := ZC4_CODGEN 	// COD_PROD_GEN
			ZC4->ZC4_GRUPO  := cGrupo
			
			MsUnlock()       // destravar o registro
		Endif
		
	EndIf
	
	
Next nY

Return(Nil)


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณ UPPROD  ณ Rogerio Oliveira   บ Data ณ  13/03/14            บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.      Funcao que realiza o Cadastrado de Produto				  บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Compras   		                                          บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function UPPROD(aRegistro)

Local nY        := 0
Local cSeq		:= "3"
Local nAuxPr	:= 0.00      //Variavel axuliar para Formatar Preco
Local cDescPr	:= space(45) //------------- DESC_PROD
Local cMarcaPr	:= space(45) //------------- MARCA_PROD
Local cGrupo    := space(7)  //------------- GRUPO
Local cUnid		:= space(15) //------------- UNIDADE
Local cPreco	:= space(17) //------------- PRECO_UNI
Local cCod		:= space(15) //------------- COD_PROD
Local cCodGen   := space(12) //------------- COD_PROD_GEN

ProcRegua(Len(aRegistro))

ZC2->(DbGotop())
ZC2->(DbSetOrder(1))
//Executando a Gravacao dos Dados de Produto
For nY := 1 to Len(aRegistro)
	
	IncProc()
	
	cDescPr 	:= aRegistro[nY,8]
	cMarcaPr	:= aRegistro[nY,9]
	cGrupo      := aRegistro[nY,10]
	cUnid		:= aRegistro[nY,12]
	cCodGen     := aRegistro[nY,5]
	
	cPreco   := ALLTRIM(StrTran(aRegistro[nY,13],'R$','' )) // AJUSTA PRECO UNITARIO
	nAuxPr   := Val(StrTran(cPreco,',','.'))
	
	// BUSCA SE EXISTE O PRODUTO
	If ( VldProd(cDescPr, cMarcaPr) == 0)
		
		If RecLock("ZC2", .T.)	  // Incluir
			
			ZC2->ZC2_COD := GetSXENum("ZC2","ZC2_COD")
			
			cLog := "LINHA -"+STRZERO(nY+1,9)+" -Produto/Descricao "+ALLTRIM(cDescPr)+space(4)+"       -"+ALLTRIM(ZC2_COD)+space(4)+"Processado  em " + DToC( MSDate() ) + " เs " + Time() + CRLF
			GrvLog( cLog, cSeq )
			
			ZC2->ZC2_DESC    := cDescPr
			ZC2->ZC2_MARCA   := cMarcaPr
			ZC2->ZC2_GRUPO   := cGrupo
			ZC2->ZC2_UNIDAD  := cUnid
			ZC2->ZC2_ULTPRE  := nAuxPr
			ZC2->ZC2_CODGEN  := cCodGen
			
			ZC2->( MsUnlock() )    		   // Destravar o registro
			ZC2->(ConfirmSX8())   		   // Executar a grava็ใo dos dados
		EndIf
		
	Else
		
		cCod  := RetProd(cDescPr, cMarcaPr) // Funcao que retorna o codigo do produto
		
		If ZC2->( MsSeek(cFilialZC2+cCod) )
			
			If RecLock("ZC2", .F.)	         // Alterar
				
				ZC2->ZC2_MARCA   := cMarcaPr
				ZC2->ZC2_GRUPO   := cGrupo
				ZC2->ZC2_UNIDAD  := cUnid
				ZC2->ZC2_ULTPRE  := nAuxPr
				
				ZC2->( MsUnlock() )    		 // Destravar o registro
				ZC2->(ConfirmSX8())   		 // Executar a grava็ใo dos dados
			EndIf
			
		EndIf
		
	EndIf
	
Next nY

Return(Nil)


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณ UPPECOM  ณ Rogerio Oliveira   บ Data ณ  13/03/14   		  บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.      Funcao que valida se exite o Fornecedor/Loja Cadastrados    บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Compras   		                                          บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function UPPECOM(aRegistro)

Local nY        := 0
Local cSeq		:= "4"
Local nAuxPr	:= 0.00    //Variavel axuliar para Formatar Preco
Local nAxTot	:= 0.00	   //Variavel axuliar para Formatar Total
Local nAxQtd 	:= 0.000   //Variavel axuliar para Formatar Quantidade

Local cNum		:= space(6) //-------------- PEDIDO
Local cItem		:= space(4) //-------------- ITEM
Local cDoc		:= space(9)	//-------------- NF
Local dEmissa	:= CTOD("  /  /    ") //---- EMISSAO_NF
Local cDesc		:= space(45) //------------- DESC_PROD
Local cMarc		:= space(45) //------------  MARCA_PROD
Local cQTD		:= space(15) //------------- QUANTIDADE
Local cUnid		:= space(15)//-------------- UNIDADE
Local cPreco	:= space(17) //------------- PRECO_UNI
Local cTotal	:= space(17) //------------  TOTAL
Local cCodGen	:= space(12) //------------  COD_PROD_GEN
Local cFornece	:= space(6)//--------------- COD_FOR
Local cLoja		:= space(2)//--------------- LOJA_FOR
Local cCnpj		:= space(14)//-------------- CNPJ
Local cNome		:= space(40)//-------------- FORNECEDOR
Local cEnd		:= space(30)//-------------- ENDERECO DO FORNECEDOR
Local cCidade	:= space(30)//-------------- CIDADE DO FORNECEDOR
Local cUfFor	:= space(2) //-------------- UF DO FORNECEDOR

ProcRegua(Len(aRegistro))

cLog	:= "Processo iniciado em " + DToC( MSDate() ) + " เs " + Time() + CRLF
GrvLog( cLog, cSeq )

//Executando a Gravacao dos Dados de Fornecedor
For nY := 1 to Len(aRegistro)
	IncProc()
	
	cNum		:= aRegistro[nY,01] //----------> PEDIDO
	cItem		:= aRegistro[nY,02] //----------> ITEM
	cDoc		:= aRegistro[nY,03] //----------> NF
	dEmissa		:= CTOD(aRegistro[nY,04]) //----> EMISSAO_NF
	cDesc		:= aRegistro[nY,08] //----------> DESC_PRODUTO
	cMarc		:= aRegistro[nY,09] //----------> MARCA_PROD
	cQTD		:= aRegistro[nY,11] //----------> QUANTIDADE
	cUnid		:= aRegistro[nY,12] //----------> UNIDADE
	cPreco		:= aRegistro[nY,13] //----------> PRECO_UNI
	cTotal		:= aRegistro[nY,14] //----------> TOTAL
	cFornece	:= aRegistro[nY,15] //----------> COD_FORN
	cLoja		:= aRegistro[nY,16] //----------> LOJA_FORN
	cCnpj		:= aRegistro[nY,17] //----------> CNPJ
	cNome		:= aRegistro[nY,18] //----------> FORNECEDOR
	cEnd		:= aRegistro[nY,19] //----------> ENDERECO
	cCodGen     := aRegistro[nY,5]  //----------> COD_PROD_GEN
	cCidade     := aRegistro[nY,20]  //----------> ZC3_CIDFOR
	cUfFor      := aRegistro[nY,21]  //----------> ZC3_UFFOR
	
	cPreco   := ALLTRIM(StrTran(aRegistro[nY,13],'R$','' )) // AJUSTA PRECO UNITARIO
	nAuxPr   := Val(StrTran(cPreco,',','.'))
	
	cTotal   := ALLTRIM(StrTran(aRegistro[nY,14],'R$','' )) // AJUSTA PRECO TOTAL
	nAxTot   := Val(StrTran(cTotal,',','.'))
	
	nAxQtd	 := Val(StrTran(cQTD,',','.')) 				   // AJUSTA QUANTIDADE
	
	ZC3->(DbGotop())
	ZC3->(DbSetOrder(1))
	If ZC3->( !MsSeek(cFilialZC3+cNum+cItem) )
		
		If RecLock("ZC3", .T.)	  // Incluir
			
			cLog := "LINHA -"+STRZERO(nY+1,9)+" -Pedido de Compra / ITEM "+space(4)+ALLTRIM(cNum)+"       -"+cItem
			GrvLog( cLog, cSeq )
			
			ZC3->ZC3_FILIAL	:= cFilialZC3
			ZC3->ZC3_NUM    := cNum
			ZC3->ZC3_ITEM   := cItem
			ZC3->ZC3_DOC    := cDoc
			ZC3->ZC3_EMISSA := dEmissa
			ZC3->ZC3_COD    := RetProd(cDesc, cMarc) //Funcao que retorna o codigo do produto
			ZC3->ZC3_DESC   := cDesc
			ZC3->ZC3_MARCA  := cMarc
			ZC3->ZC3_QTD    := nAxQtd    //Auxiliar quantidade
			ZC3->ZC3_UNIDAD := cUnid
			ZC3->ZC3_PRECO  := nAuxPr    //Auxiliar preco
			ZC3->ZC3_TOTAL	:= nAxTot    //Auxiliar total
			ZC3->ZC3_FORNEC	:= cFornece
			ZC3->ZC3_LOJA	:= cLoja
			ZC3->ZC3_CNPJ	:= cCnpj
			ZC3->ZC3_NOME	:= cNome
			ZC3->ZC3_END	:= cEnd
			ZC3->ZC3_CODGEN	:= cCodGen 
			ZC3->ZC3_CIDFOR := cCidade
			ZC3->ZC3_UFFOR  := cUfFor
			
			ZC3->( MsUnlock() )    // Destravar o registro
			ZC3->(ConfirmSX8())    // para executar a grava็ใo dos dados
		Else
			If RecLock("ZC3", .F.)	// Alterar
				
				ZC3->ZC3_DOC    := cDoc
				ZC3->ZC3_MARCA  := cMarc
				ZC3->ZC3_UNIDAD := cUnid
				
				ZC3->( MsUnlock() )    // Destravar o registro
				ZC3->(ConfirmSX8())    // para executar a grava็ใo dos dados
			EndIf
		EndIf
		
	EndIf
	
Next nY

cLog	:= "Processo finalizado em " + DToC( MSDate() ) + " เs " + Time() + CRLF
GrvLog( cLog, cSeq )
MsgAlert("Dados Importados com Sucesso!!!")
Return(Nil)


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  RetProd   		      Rogerio Oliveira   บ Data ณ  13/03/14   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.  Funcao que valida se existe o Produto Cadastrado, caso exista   บฑฑ
ฑฑบela retorna o codigo do produto.                                       บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Compras   		                                          บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function RetProd(cDescPr, cMarcaPr)

Local cQuery := ""
Local aArea  := GetArea()
Local cCod   := space(15)
Local cAlias := GetNextAlias()

cQuery := " SELECT ZC2_COD"
cQuery += " FROM "+RetSqlName("ZC2")
cQuery += " WHERE ZC2_FILIAL = '"+cFilialZC2+"' "
cQuery += " AND ZC2_DESC = '"+cDescPr+"'"+" AND ZC2_MARCA = '"+cMarcaPr+"'"
cQuery += " AND D_E_L_E_T_ <> '*'"

TCQuery cQuery Alias &cAlias New
dbSelectArea(cAlias)
dbGoTop()

cCod := (cAlias)->ZC2_COD
(cAlias)->(dbCloseArea())
RestArea(aArea)

Return cCod


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  VldForLoj   		  Rogerio Oliveira   บ Data ณ  13/03/14   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.      Funcao que valida se exite o Fornecedor/Loja Cadastrados    บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Compras   		                                          บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function VldForLoj(cNomFor)

Local cQuery := ""
Local aArea  := GetArea()
Local cLoja  := space(2)
Local cAlias := GetNextAlias()
Local cAux	 := ""

cQuery := " SELECT max(ZC1_LOJFOR) AS LOJA"
cQuery += " FROM "+RetSqlName("ZC1")
cQuery += " WHERE ZC1_FILIAL = '"+cFilialZC1+"' "
cQuery += " AND ZC1_NOME LIKE '%"+cNomFor+"%' "
cQuery += " AND D_E_L_E_T_ <> '*'"

TCQuery cQuery Alias &cAlias New
dbSelectArea(cAlias)
dbGoTop()

cLoja := (cAlias)->LOJA
(cAlias)->(dbCloseArea())
RestArea(aArea)

cAux := IIf(Empty(cLoja), cLoja, SOMA1(cLoja))

Return cAux


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  VldProd   		      Rogerio Oliveira   บ Data ณ  13/03/14   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.      Funcao que valida se existe o Produto Cadastrado  		  บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Compras   		                                          บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function VldProd(cDescPr, cMarcaPr)

Local cQuery := ""
Local aArea  := GetArea()
Local nCount := 0
Local cAlias := GetNextAlias()

cQuery := " SELECT ISNULL(SUM(1), 0) AS nCout
cQuery += " FROM "+RetSqlName("ZC2")
cQuery += " WHERE ZC2_FILIAL = '"+cFilialZC2+"' "
cQuery += " AND ZC2_DESC = '"+cDescPr+"'"+" AND ZC2_MARCA = '"+cMarcaPr+"'"
cQuery += " AND D_E_L_E_T_ <> '*'"

TCQuery cQuery Alias &cAlias New
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
ฑฑบPrograma  ณ GrvLog   บAutor  ณ Rogerio Oliveira   บ Data ณ  13/03/14   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณGeracao log 			                                      บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Compras   		                                          บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function GrvLog( cError, cSeq )

Local cAux1		:= "\logForn_log\"+__cUserId+" -LOGS.log"
Local cAux2		:= "\logProdGEN_log\"+__cUserId+" -LOGS.log"
Local cAux3		:= "\logProd_log\"+__cUserId+" -LOGS.log"
Local cAux4		:= "\logPedComp_log\"+__cUserId+" -LOGS.log"

Local cArq		:= " "
Local nHdl		:= 0

cLog := cError

If cSeq == "1"
	cArq := cAux1
ElseIf cSeq == "2"
	cArq := cAux2
ElseIf cSeq == "3"
	cArq := cAux3
ElseIf cSeq == "4"
	cArq := cAux4
Endif

If File(cArq)
	If (nHdl := FOpen(cArq,2)) == -1
		Return
	EndIf
Else
	If (nHdl := MSFCreate(cArq,0)) == -1
		Return
	EndIf
Endif

FSeek( nHdl,0,2)
fWrite(nHdl, cLog+Chr(13)+Chr(10))
fClose(nHdl)

Return(Nil)

