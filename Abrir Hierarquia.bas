Attribute VB_Name = "Módulo1"
Sub abrir()

    ' Declarações
    Dim inp, out As Worksheet
    Dim last_row As Long
    Dim rng As Range, cell As Range
    Dim negativeFound As Boolean
    
    ' Atribuições
    Set inp = ThisWorkbook.Sheets("Input")
    Set out = ThisWorkbook.Sheets("Output")

    ' Ultima Linha preenchida
    last_row = inp.Cells(Rows.Count, "B").End(xlUp).Row
    last_rowOut = out.Cells(Rows.Count, "B").End(xlUp).Row
    
    If last_rowOut >= 2 Then
        out.Range("A2:F" & last_rowOut).Clear
    End If
    
    ' Define o intervalo a ser verificado (alterar conforme necessário)
    Set rng = inp.Range("L3:Q" & last_row)

    ' Verifica se há números negativos no intervalo
    negativeFound = False
    For Each cell In rng
        If cell.Value < 0 Then
            negativeFound = True
            Exit For
        End If
    Next cell

    ' Bloqueia a execução se números negativos forem encontrados
    If negativeFound Then
        MsgBox "Erro: Há números negativos no intervalo especificado de Cluster.", vbExclamation
        Exit Sub
    End If
    
    For i = 3 To last_row
            ' Cluster A
        If inp.Range("L" & i).Value <> 0 Then
            j = out.Cells(Rows.Count, "A").End(xlUp).Row
            m = inp.Range("L" & i).Value
            For k = 1 To m
                out.Range("A" & j + k & ":E" & j + k).Value = inp.Range("A" & i & ":E" & i).Value
                out.Range("F" & j + k).Value = inp.Range("L" & 2).Value
            Next k
        End If
            ' Cluster B
        If inp.Range("M" & i).Value <> 0 Then
            j = out.Cells(Rows.Count, "A").End(xlUp).Row
            m = inp.Range("M" & i).Value
            For k = 1 To m
                out.Range("A" & j + k & ":E" & j + k).Value = inp.Range("A" & i & ":E" & i).Value
                out.Range("F" & j + k).Value = inp.Range("M" & 2).Value
            Next k
        End If
            ' Cluster C
        If inp.Range("N" & i).Value <> 0 Then
            j = out.Cells(Rows.Count, "A").End(xlUp).Row
            m = inp.Range("N" & i).Value
            For k = 1 To m
                out.Range("A" & j + k & ":E" & j + k).Value = inp.Range("A" & i & ":E" & i).Value
                out.Range("F" & j + k).Value = inp.Range("N" & 2).Value
            Next k
        End If
            ' Cluster D
        If inp.Range("O" & i).Value <> 0 Then
            j = out.Cells(Rows.Count, "A").End(xlUp).Row
            m = inp.Range("O" & i).Value
            For k = 1 To m
                out.Range("A" & j + k & ":E" & j + k).Value = inp.Range("A" & i & ":E" & i).Value
                out.Range("F" & j + k).Value = inp.Range("O" & 2).Value
            Next k
        End If
            ' Cluster E
        If inp.Range("P" & i).Value <> 0 Then
            j = out.Cells(Rows.Count, "A").End(xlUp).Row
            m = inp.Range("P" & i).Value
            For k = 1 To m
                out.Range("A" & j + k & ":E" & j + k).Value = inp.Range("A" & i & ":E" & i).Value
                out.Range("F" & j + k).Value = inp.Range("P" & 2).Value
            Next k
        End If
            ' Cluster F
        If inp.Range("Q" & i).Value <> 0 Then
            j = out.Cells(Rows.Count, "A").End(xlUp).Row
            m = inp.Range("Q" & i).Value
            For k = 1 To m
                out.Range("A" & j + k & ":E" & j + k).Value = inp.Range("A" & i & ":E" & i).Value
                out.Range("F" & j + k).Value = inp.Range("Q" & 2).Value
            Next k
        End If

    Next i
    
    MsgBox "Pronto!"
        
End Sub
