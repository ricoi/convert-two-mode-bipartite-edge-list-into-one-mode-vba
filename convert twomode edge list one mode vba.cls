VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plan2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' necessário copiar nomes de titulares para coluna "b"

Sub convert_twomode_edgelist_onemode()

Dim numer_pat1, nome1, nome2 As String
Dim data, IPC As String
Dim num_lin1, num_lin2 As Long

Worksheets("base").Activate
Range("A2").Activate
'roda até acabar dados
While ActiveCell.Value <> ""
    'armazena linha inicial e numero da patente correspondente
    num_lin1 = ActiveCell.Row
    numer_pat1 = ActiveCell.Value
    data = ActiveCell.Offset(0, 2).Value
    'class = ActiveCell.Offset(0, 3).Value
    'class1 = ActiveCell.Offset(0, 4).Value
    'class2 = ActiveCell.Offset(0, 5).Value
    'class3 = ActiveCell.Offset(0, 6).Value
    'class4 = ActiveCell.Offset(0, 7).Value
    'class5 = ActiveCell.Offset(0, 8).Value
    'class6 = ActiveCell.Offset(0, 9).Value
    'class7 = ActiveCell.Offset(0, 10).Value
    'anda uma pra baixo e continua andando enquanto o numero da patente estiver repetindo
    ActiveCell.Offset(1, 0).Activate
    While ActiveCell.Offset(-1, 0).Value = numer_pat1
        ActiveCell.Offset(1, 0).Activate
    Wend
    'armazena numero de linha da ultima relação da mesma patente
    num_lin2 = (ActiveCell.Row - 2)
    
    'ate aqui tem num_lin1(inicio), num_lin2(fim) e patente
    
    'pra cada linha entre a primeira e a ultima excluindo a ultima
    For i = num_lin1 To (num_lin2 - 1) Step 1
    'armazena nome 1
        nome1 = Cells(i, 2)
    'pra cada linha entre a linha atual e a ultima
        For j = i To num_lin2 Step 1
    'armazena o nome 2
            nome2 = Cells(j, 2)
    'se nome1 e nome2 sao diferentes entao
            If nome1 <> nome2 Then
                'copia na outra aba
                Worksheets("relacao").Activate
                ActiveCell.Value = numer_pat1
                ActiveCell.Offset(0, 1).Value = nome1
                ActiveCell.Offset(0, 2).Value = nome2
                ActiveCell.Offset(0, 3).Value = data
                'ActiveCell.Offset(0, 6).Value = class
                'ActiveCell.Offset(0, 7).Value = class1
                'ActiveCell.Offset(0, 8).Value = class2
                'ActiveCell.Offset(0, 9).Value = class3
                'ActiveCell.Offset(0, 10).Value = class4
                'ActiveCell.Offset(0, 11).Value = class5
                'ActiveCell.Offset(0, 12).Value = class6
                'ActiveCell.Offset(0, 13).Value = class7
                'ativa linha de baixo para proximo dado e volta pra aba base
                ActiveCell.Offset(1, 0).Activate
                Worksheets("base").Activate
                
            End If
    
        Next

    Next
'ativa a primeira linha da proxima patente
Cells(num_lin2 + 1, 1).Activate
'volta o primeiro loop (while)

' Escrito por Felício Visnardi
Wend

End Sub

