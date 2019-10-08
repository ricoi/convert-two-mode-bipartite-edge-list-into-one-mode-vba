# Convert two-mode networks to one-mode networks - vba macro

This vba macro was developed to support researchers, students and data analysts to perform Social Network Analysis (SNA). 
The algorithm has application to transform affiliations networks into one mode networks. 
It can be used to establish relations between actors linked to the same thing, i.e. to creating links between:
-	Co-assignees or coinventors of the same patent; 
-	co-authors of the same article; 
-	students of the same classes; 
-	other kinds of affiliation, or common characteristics;


### Requirements

Requires software Microsoft Excel and Visual Basic, any version.
The data pattern in this model is suitable to [Gephi](https://gephi.org/) input, recommended for SNA.

## How to use
To exemplify, this manual shows the construction of an inter-organizational colaboration network for innovation development, which connections between organizations are established by co-ownership of the same patent:

If some patent (N) belongs to three different companies (X, Y, Z), the relations between then are represented by (X-Y, X-Z, Y-Z).

### Preparing data base:
- Into an excel file (.xlsx), rename two different worksheets as: “base” and “relacao”
-	“base” worksheet should have the edge list, in each line with a unique relationship, and columns with actors data:
-	First column (A), header “source”, should have the actors data (organizations names; 
-	Second column (B), header “target”, should have the affiliation data (publication number of patent owned by the organization);
-	The next columns could have complementary data from each relationship, like patent application data, and others (necessary adapt vba code); 
-	Data must be sort, hierarchically, in this order: column “target” (B) followed by column “source” (A);
-	Duplicated relationships should be cleaned, because it represents errors, i.e.: the name of the organization appear two times to the same patent number;
   
 

When run this macro, the new one-mode edge list will be generated at “relacao” worksheet:
- Each line will have a single relantionship;
-	First column (A), will have the patent number, or the common characteristic which originated the relationship;
-	Second column (B), and also third column ( C), will have the actors data (organizations names); 
-	The next columns could have complementary data from each relantionship, like patent application data, and others (necessary adapt vba code);
-	Second column (B) should header “source”, and third column ( C) should header "target", for Gephi input;
 
### Running vba macro
-	Before perform vba macro, select the first line with the first column (A2) where results will be generated, at “relacao” worksheet;

-	[See how to run a macro](https://support.office.com/en-us/article/run-a-macro-5e855fd2-02d1-45f5-90a3-50e645fe3155);

- If you creating your own excel file, paste the code presented below. 
[See how to copy a macro module to another workbook at Office Support](https://support.office.com/en-us/article/copy-a-macro-module-to-another-workbook-13c0938b-8432-4259-9177-a71f7e626de0)

```vb
Sub convert_twomode_edgelist_onemode()
$

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


Wend

End Sub
```






## Authors:

[Ricardo Cruz Gomes](https://orcid.org/0000-0003-4414-4738)

Felício Visnard

## License

This project is licensed under the terms of the MIT license


