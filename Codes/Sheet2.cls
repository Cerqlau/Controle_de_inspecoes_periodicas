VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

'verifica��o para chamada do formul�rio em caso de data

On Error GoTo saida 'inibe os demais itens da planilha ignorando o erro

Cancel = True



If IsDate(Target.Value) = True Then 'verifica se � uma data
    If Target.Interior.Color = 5296274 Then ' se data verifica se verde
        UserForm1.Show
    End If
    If Target.Interior.Color = 10498160 Then ' se data verifica se roxo
        UserForm1.Show
        UserForm2.Show
    End If
    If Target.Interior.Color = 15123099 Then  ' se data verifica se roxo
        UserForm2.Show
    End If
    If Target.Interior.Color = 8696052 Then  ' se data verifica se roxo
        UserForm3.Show
    End If
    If Target.Interior.Color = 65535 Then  ' se data verifica se roxo
        UserForm1.Show
        UserForm3.Show
    End If
    If Target.Interior.Color = 13734634 Then  ' se data verifica se roxo
        UserForm1.Show
        UserForm2.Show
        UserForm3.Show
    End If
    
    
End If


saida:

End Sub
