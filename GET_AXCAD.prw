#include 'protheus.ch'
#include 'parmtype.ch'

User Function XCadSA2()

Local cAlias := "ZC4"
Local cTitulo := "Cadastro Teste"                    
Local cVldExc := ".T."
Local cVldAlt := ".T."	
	
	
dbSelectArea(cAlias)
dbSetOrder(1)	
AxCadastro(cAlias,cTitulo,cVldExc,cVldAlt)
	
	
Return NIL