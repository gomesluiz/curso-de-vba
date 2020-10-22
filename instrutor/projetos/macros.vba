' Projeto 3
Option Explicit
Public Sub TrataNome()
    Dim nome As String
    Dim primeiro As String
    Dim ultimo As String
    Dim tamanho As Integer
    Dim espaco As Integer
    
    nome = InputBox("Entre com o seu primeiro e último nome.", "Nome:")
    espaco = InStr(1, nome, " ")
    
    primeiro = Left(nome, espaco - 1)
    
    Range("C3").Value = primeiro
    tamanho = Len(primeiro)
    Range("C4").Value = tamanho
    tamanho = Len(nome)
    
    ultimo = Mid(nome, espaco + 1, tamanho - espaco)
    Range("C5").Value = ultimo
    tamanho = Len(ultimo)
    Range("C6").Value = tamanho
    Range("C7").Value = UCase(nome)
    Range("C8").Value = LCase(nome)
    Range("C9").Value = StrConv(nome, vbProperCase)
    Range("C10").Value = StrReverse(nome)
    Range("C11").Value = ultimo & ", " & primeiro
End Sub

' Projeto 4
Public Sub CalculaEstatistica()

    With ActiveSheet
        'These formulas are entered into the new worksheet.
        .Range("D2").Formula = "=COUNT(" & ActiveWindow.Selection.Address & ")"
        .Range("D3").Formula = "=MIN(" & ActiveWindow.Selection.Address & ")"
        .Range("D4").Formula = "=MAX(" & ActiveWindow.Selection.Address & ")"
        .Range("D5").Formula = "=SUM(" & ActiveWindow.Selection.Address & ")"
        .Range("D6").Formula = "=AVERAGE(" & ActiveWindow.Selection.Address & ")"
        .Range("D7").Formula = "=STDEV(" & ActiveWindow.Selection.Address & ")"
    
        .Range("C2").Value = "Count: "
        .Range("C3").Value = "Min: "
        .Range("C4").Value = "Max: "
        .Range("C5").Value = "Sum: "
        .Range("C6").Value = "Average: "
        .Range("C7").Value = "Stan Dev:"
        .Range("C2:D7").Select
    End With
    
    With Selection
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = vbBlue
        .Font.Name = “Arial”
        .Columns.AutoFit
        .Interior.Color = vbGreen
        .Borders.Weight = xlThick
        .Borders.Color = vbRed
    End With
    Range("D2:D7").Select
    Selection.NumberFormat = "0.00"
    Range("A1").Select
    
End Sub


' Projeto 5
Option Explicit
Public Sub JogaDados()
    Dim dados1 As Integer
    Dim dados2 As Integer
    
    dados1 = (Rnd * 5) + 1
    dados2 = (Rnd * 5) + 1
    
    Range("F7").Value = dados1
    Range("G7").Value = dados2
    
    Range("F9:G9").Merge
    If dados1 = dados2 Then
        Range("F9").Value = "VENCEU!"
        Range("F9").Interior.Color = vbBlue
    Else
        Range("F9").Value = "PERDEU!"
        Range("F9").Interior.Color = vbRed
    End If
    
End Sub

' Projeto 6
Public Sub ChecaSituacao()
    Dim r As Range
    Dim n, aumento, reducao, estavel As Integer
    Dim ultimaLinha As Long
    
    aumento = 0: reducao = 0: estavel = 0
    
    Worksheets("Vendas").Activate
    ultimaLinha = Cells(Rows.Count, 1).End(xlUp).Row
    
    MontaBorda ("a4:e" & ultimaLinha)
    Set r = Range("b4:e" & ultimaLinha)
    For n = 1 To r.Rows.Count
        If (Int(r.Cells(n, 3).Value) <= Int(r.Cells(n, 2).Value)) And _
           (Int(r.Cells(n, 2).Value) <= Int(r.Cells(n, 1).Value)) Then
            r.Cells(n, 4).Interior.Color = RGB(255, 0, 0)
            reducao = reducao + 1
        ElseIf (Int(r.Cells(n, 3).Value) >= Int(r.Cells(n, 2).Value)) And _
               (Int(r.Cells(n, 2).Value) >= Int(r.Cells(n, 1).Value)) Then
            r.Cells(n, 4).Interior.Color = RGB(0, 255, 0)
            aumento = aumento + 1
        Else
            r.Cells(n, 4).Interior.Color = RGB(255, 255, 0)
            estavel = estavel + 1
        End If
    Next n
    
    ' monta tabela resumo
    MontaBorda ("g6:h6")
    Range("G4").Value = "Aumento"
    Range("G4").Interior.Color = RGB(0, 0, 255)
    Range("H4").Value = aumento
    
    Range("G5").Value = "Redução"
    Range("G5").Interior.Color = RGB(255, 0, 0)
    Range("H5").Value = reducao
    
    Range("G6").Value = "Estável"
    Range("G6").Interior.Color = RGB(255, 255, 0)
    Range("H6").Value = estavel
    
End Sub

