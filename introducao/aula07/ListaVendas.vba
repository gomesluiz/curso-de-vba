Sub connect()
    Dim strSQL 		As String
    Dim lngColunas 	AS Long 
    Dim lngLinhas  	AS Long
    Dim vendas() 
		

    Set conexao = getConexao()

    strSQL = "SELECT * FROM vendas"
    rs.Open strSQL, conexao, adOpenStatic
    
    vendas 	= rs.GetRows()	      ,
    colunas 	= UBound(myArray, 1)  'Quantidade de colunas
    linhas 	= UBound(myArray, 2)  'Quantidade de linhas

    For c = 0 To colunas 
        Range("a5").Offset(0, c).Value = rs.Fields(c).Name
        For l = 0 To linhas
           Range("a5").Offset(l + 1, c).Value = vendas(c, l)
        Next
    Next

    rs.Close
    Set rs = Nothing
    conexao.Close
    Set conexao = Nothing
End Sub
