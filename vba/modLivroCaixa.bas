Attribute VB_Name = "modLivroCaixa"
Option Explicit

' =====================================================================
' Demonstracao (dados ficticios). Duas macros para as abas "Controle Financeiro NN":
'   Saldo            refaz a coluna L (saldo corrido) ate a ultima linha com data
'   AdicionarLinhas  insere N linhas acima da faixa azul do fim, sem quebrar os intervalos
' As duas se ancoram no bloco "Saldo ... Final" da coluna F (nunca em End(xlUp)).
' =====================================================================

Private Function LinhaFinal(ws As Worksheet) As Long
    Dim r As Long
    For r = 7 To ws.UsedRange.Row + ws.UsedRange.Rows.Count
        If InStr(1, CStr(ws.Cells(r, 6).Value), "Final", vbTextCompare) > 0 Then
            LinhaFinal = r
            Exit Function
        End If
    Next r
End Function

Public Sub Saldo(Optional ByVal silencioso As Boolean = False)
    Dim ws As Worksheet, linFinal As Long, ultimaDados As Long, r As Long
    Dim sSaida As String, corpo As String, fPrimeira As String, fDemais As String

    Set ws = ActiveSheet
    linFinal = LinhaFinal(ws)
    If linFinal = 0 Then
        MsgBox "Nao achei o bloco 'Saldo ... Final' na coluna F desta aba.", vbExclamation
        Exit Sub
    End If

    ' ultima linha de dados = primeira com data em A, subindo da faixa azul (linFinal - 1)
    For r = linFinal - 2 To 7 Step -1
        If Len(CStr(ws.Cells(r, 1).Value)) > 0 Then
            ultimaDados = r
            Exit For
        End If
    Next r
    If ultimaDados < 7 Then
        MsgBox "Nao achei lancamentos na coluna A desta aba.", vbExclamation
        Exit Sub
    End If

    sSaida = "Sa" & ChrW(237) & "da"   ' montada com ChrW de proposito: se o acento corromper, o saldo fecha errado sem erro
    corpo = "SUMIF(RC[-5],""Receita"",RC[-1])" & _
            "-SUMIF(RC[-5],""Gasto"",RC[-1])" & _
            "+SUMIF(RC[-4],""Aporte"",RC[-1])" & _
            "-SUMIF(RC[-4],""Retirada"",RC[-1])" & _
            "+SUMIF(RC[-4],""Entrada"",RC[-1])" & _
            "-SUMIF(RC[-4],""" & sSaida & """,RC[-1])" & _
            "-SUMIFS(RC[-1],RC[-4],""Investimento"",RC[-6],""<>VENDA DE ATIVO"")" & _
            "+SUMIFS(RC[-1],RC[-4],""Investimento"",RC[-6],""VENDA DE ATIVO"")"
    fPrimeira = "=SUM(R[-5]C[-5],R[-4]C[-5],R[-3]C[-5],R[-5]C[-1],R[-4]C[-1],R[-3]C[-1])+" & corpo
    fDemais = "=ROUND(R[-1]C+" & corpo & ",2)"

    Application.ScreenUpdating = False
    ws.Cells(7, 12).FormulaR1C1 = fPrimeira
    If ultimaDados >= 8 Then ws.Range(ws.Cells(8, 12), ws.Cells(ultimaDados, 12)).FormulaR1C1 = fDemais
    If linFinal - 2 > ultimaDados Then ws.Range(ws.Cells(ultimaDados + 1, 12), ws.Cells(linFinal - 2, 12)).ClearContents
    ws.Range(ws.Cells(7, 12), ws.Cells(ultimaDados, 12)).NumberFormat = "R$ #,##0.00"
    Application.ScreenUpdating = True

    If Not silencioso Then
        MsgBox "Saldo refeito em " & ws.Name & "." & vbCrLf & "Linhas 7 a " & ultimaDados & "." & vbCrLf & vbCrLf & _
               "Ultimo saldo: " & Format$(ws.Cells(ultimaDados, 12).Value, "#,##0.00"), vbInformation, "Pronto"
    End If
End Sub

Public Sub AdicionarLinhas(Optional ByVal numLinhas As Long = 0, Optional ByVal silencioso As Boolean = False)
    Dim ws As Worksheet, linFinal As Long, faixa As Long, v As Variant

    Set ws = ActiveSheet
    linFinal = LinhaFinal(ws)
    If linFinal = 0 Then
        MsgBox "Nao achei o bloco 'Saldo ... Final' na coluna F desta aba.", vbExclamation
        Exit Sub
    End If
    faixa = linFinal - 1   ' faixa azul com OK/ERRO

    If numLinhas = 0 Then
        v = Application.InputBox("Quantas linhas deseja adicionar?" & vbCrLf & vbCrLf & "Aba: " & ws.Name & vbCrLf & _
                                 "Faixa de validacao na linha: " & faixa, "Adicionar linhas", 100, Type:=1)
        If VarType(v) = vbBoolean Then Exit Sub
        numLinhas = CLng(v)
    End If
    If numLinhas <= 0 Then Exit Sub

    Application.ScreenUpdating = False
    ' insere na PENULTIMA linha util: so assim o Excel estica sozinho os intervalos dos SOMASES do bloco Final
    ws.Rows((faixa - 2) & ":" & (faixa - 2 + numLinhas - 1)).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Application.ScreenUpdating = True

    If Not silencioso Then MsgBox numLinhas & " linhas inseridas na aba " & ws.Name & ".", vbInformation, "Pronto"
End Sub
