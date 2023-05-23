VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Inspe��es Mensais "
   ClientHeight    =   9435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15105
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

'Imprimir papel

On Error GoTo Erro
 
    Sheets("Mensal").Visible = xlSheetVisible ' torna a planilha vis�vel
    
    ActiveWorkbook.Sheets("Mensal").Select
    
    Range("C8:C20").ClearContents
    Range("D8:D20").ClearContents
    Range("A5").Value = "Elaborador: "
    Range("D5:D6").Value = "Data: "
    

    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    Application.Dialogs(xlDialogPrinterSetup).Show 'mostra caixa de sele��o de impressora
    ActiveWindow.SelectedSheets.PrintOut , Copies:=1
    
    ActiveWorkbook.Sheets("Calendario de inspe��o 2023").Select
    
    Sheets("Mensal").Visible = xlSheetVeryHidden 'esconde a planilha ap�s a opera��o
    
'    ActiveWindow.SelectedSheets.PrintOut , ActivePrinter:="Microsoft Print to PDF", PrintToFile:=True, Copies:=1, Collate:=True, _
'        IgnorePrintAreas:=False
'    Application.WindowState = xlNormal
    
Erro:

End Sub

Private Sub CommandButton2_Click()

Sheets("Mensal ").Visible = xlSheetVisible

Call desprotege_relatorio

Dim data_limite As Date

data_limite = CDate(Me.TextBox3.Value)


'verifica se o relat�rio j� foi elaborado anteriormente e informa o respons�vel e a data

ActiveWorkbook.Sheets("Relatorio ROP").Select
Range("K4").Select
Range(Selection, Selection.End(xlDown)).Select

For Each sel In Selection
contador = contador + 1
Next

For i = 3 To contador + 3

    If Cells(i + 1, 11).Value = CDate(Me.TextBox3.Value) And Cells(i + 1, 13).Value <> "" Then
    
        MsgBox "Relat�rio j� elaborado por " & Cells(i + 1, 6).Value & " em " & Cells(i + 1, 8).Value, vbInformation
        ActiveWorkbook.Sheets("Calendario de inspe��o 2023").Select
        GoTo saida
    End If
    
Next i



'verifica se est� dentro do prazo X de atraso para entrega

data_ = DateDiff("d", Date, data_limite)

If data_ < -1 Then
   
   MsgBox "Inspe��o fora do prazo, n�o � possivel atualizar", vbExclamation
   ActiveWorkbook.Sheets("Calendario de inspe��o 2023").Select
   GoTo saida
   
End If

'verifica se est� com mais de X de entecipa��o para entrega

If data_ > 5 Then
   
   MsgBox "Inspe��o com 5 dias ou mais de antecipa��o, n�o � possivel atualizar", vbExclamation
   ActiveWorkbook.Sheets("Calendario de inspe��o 2023").Select
   GoTo saida
   
End If

'verifica se o o relat�rio possui elaborador

If Me.ComboBox1.Value = "" Then

   MsgBox "Por favor selecione o nome do elaborador", vbExclamation
   ActiveWorkbook.Sheets("Calendario de inspe��o 2023").Select
   GoTo saida

End If


'verifica��o se todas as checkbox est�o marcadas

Dim cont As Integer

cont = 0
contador = 0

If Me.CheckBox1.Value = False Then 'contador de checkbox que n�o foram marcados
   cont = cont + 1
End If
If Me.CheckBox2.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox3.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox4.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox5.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox6.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox7.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox8.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox9.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox10.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox11.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox12.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox13.Value = False Then
   cont = cont + 1
End If


   
If cont >= 1 And Me.TextBox2.Text = "" Then

    MsgBox "Existem " & Str(cont) & " itens que n�o foram verificados ou n�o foram justificados. Por favor refa�a a verifica��o e insira a justificativa no campo de observa��es"
    ActiveWorkbook.Sheets("Calendario de inspe��o 2023").Select
Else

    'Verifica��o para salvamento na data correta
    
    ActiveWorkbook.Sheets("Relatorio ROP").Select
    Range("K4").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    For Each sel In Selection
    contador = contador + 1
    Next
    
    'loop de identifica��o das c�lulas para inserir as informa��es atrav�s da data limite
    For i = 3 To contador + 3
    
        If Cells(i + 1, 11).Value = CDate(Me.TextBox3.Value) And IsDate(Cells(i + 1, 12).Value) = False Then
        Cells(i + 1, 12).Value = Date 'data de realiza��o
        Cells(i + 1, 13).Value = Me.ComboBox1.Value 'elaboardor
        Cells(i + 1, 14).Value = Me.TextBox2.Value
        Cells(i + 1, 14).Select
        
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        Call imprimir_relatorio_completo
        
        MsgBox "Relat�rio salvo com sucesso", vbInformation
        ActiveWorkbook.Sheets("Calendario de inspe��o 2023").Select
        End If
        
    Next i


End If


saida:

Call protege_relatorio

Sheets("Mensal").Visible = xlSheetVeryHidden

End Sub



Private Sub UserForm_Initialize()

Dim data_limite As Date



With ComboBox1

.AddItem "Luciano"
.AddItem "Pablo"
.AddItem "Valmir"
.AddItem "Anderson"

End With

'insere a data limite no formul�rio a partir do duplo click

data_limite = CDate(ActiveCell.Value)

With TextBox3
.Text = Str(data_limite)
End With

'Insere data ataul de execu��o do formul�rio

With TextBox1

.Text = Str(Date)

End With

' Cria e insere os itens para inspe��o a partir do listbox

With Me.ListBox1

.ColumnCount = 3
.ColumnWidths = "30;200;30"

.AddItem "1"
.AddItem "-------------------------------------------------------------------------------------------------"
.AddItem "2"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "3"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "4"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "5"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "6"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "7"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "8"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "9"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "10"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "11"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "12"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "13"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"




.List(0, 1) = "Insertos e placas Hangoff"
.List(1, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(2, 1) = "Guincho #11 (guincho pneum�tico Popeye"
.List(3, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(4, 1) = "Guincho #12 (guincho pneum�tico Popeye"
.List(5, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(6, 1) = "Roscas de Manilhas"
.List(7, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(8, 1) = "Manilhas Hidr�ulicas"
.List(9, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(10, 1) = "Cabo de a�o"
.List(11, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(12, 1) = "Bandeja Coletora de res�duos"
.List(13, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(14, 1) = "Bomba Sapo"
.List(15, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(16, 1) = "Refresh de utiliza��o dos guinchos"
.List(17, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(18, 1) = "Safety Hook's"
.List(19, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(20, 1) = "Insertos e placas Hangoff parafusos"
.List(21, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(22, 1) = "Lan�a Retinida"
.List(23, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(24, 1) = "Conjunto completo do Hangoff (Tulipa)"
.List(25, 1) = "---------------------------------------------------------------------------------------------------------------------------"


.List(0, 2) = "Lubrifica��o"
.List(1, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(2, 2) = "Inspe��o visual e lubrifica��o"
.List(3, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(4, 2) = "Inspe��o visual e lubrifica��o"
.List(5, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(6, 2) = "Inspe��o visual e teste funcional"
.List(7, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(8, 2) = "Inspe��o visual e teste funcional"
.List(9, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(10, 2) = "Inspe��o visual e lubrifica��o"
.List(11, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(12, 2) = "Inspe��o visual e reparo (caso aplic�vel)"
.List(13, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(14, 2) = "Inspe��o visual e teste funcional"
.List(15, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(16, 2) = "Treinamento da equipe"
.List(17, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(18, 2) = "Checar par�metros de inspe��o de seguran�a"
.List(19, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(20, 2) = "Conferir quantidade e integridade das roscas"
.List(21, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(22, 2) = "Teste funcional"
.List(23, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(24, 2) = "Inspe��o visual nas roscas dos parafusos e limpeza dos furos"
.List(25, 2) = "---------------------------------------------------------------------------------------------------------------------------"

End With


End Sub

Private Sub protege_relatorio()

'Sub para proteger com senha o relat�rio ROP

On Error GoTo Erro

Application.ScreenUpdating = False
ActiveWorkbook.Sheets("Relatorio ROP").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
ActiveWindow.SmallScroll Down:=-21
Selection.Locked = False
Selection.FormulaHidden = False
Range("A1:X37").Select
Selection.Locked = True
Selection.FormulaHidden = False
Range("Y1:AA15").Select
Selection.Locked = True
Selection.FormulaHidden = False
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:="Subsea77"
Range("B5").Select
Application.ScreenUpdating = True

Exit Sub

Erro:

End Sub


Private Sub desprotege_relatorio()

'Sub para desproteger com senha o relat�rio ROP

On Error GoTo Erro

Application.ScreenUpdating = False
ActiveWorkbook.Sheets("Relatorio ROP").Select
ActiveSheet.Unprotect Password:="Subsea77"
Application.ScreenUpdating = True

Exit Sub

Erro:

End Sub


Private Sub imprimir_relatorio_completo()


Sheets("Mensal").Visible = xlSheetVisible 'torna a planilha vis�vel

ActiveWorkbook.Sheets("Mensal").Select 'seleciona a planilha

'Bloco de preenchimento do relat�rio com os dados do relat�rio

If Me.CheckBox1.Value = False Then
   Range("C8").Value = "N�o"
Else
   Range("C8").Value = "Sim"
End If

If Me.CheckBox2.Value = False Then
    Range("C9").Value = "N�o"
Else
   Range("C9").Value = "Sim"
End If
If Me.CheckBox3.Value = False Then
   Range("C10").Value = "N�o"
Else
   Range("C10").Value = "Sim"
End If
If Me.CheckBox4.Value = False Then
   Range("C11").Value = "N�o"
Else
   Range("C11").Value = "Sim"
End If
If Me.CheckBox5.Value = False Then
   Range("C12").Value = "N�o"
Else
   Range("C12").Value = "Sim"
End If
If Me.CheckBox6.Value = False Then
   Range("C13").Value = "N�o"
Else
   Range("C13").Value = "Sim"
End If
If Me.CheckBox7.Value = False Then
   Range("C14").Value = "N�o"
Else
   Range("C14").Value = "Sim"
End If
If Me.CheckBox8.Value = False Then
      Range("C15").Value = "N�o"
Else
   Range("C15").Value = "Sim"
End If
If Me.CheckBox9.Value = False Then
    Range("C16").Value = "N�o"
Else
   Range("C16").Value = "Sim"
End If
If Me.CheckBox10.Value = False Then
      Range("C17").Value = "N�o"
Else
   Range("C17").Value = "Sim"
End If
If Me.CheckBox11.Value = False Then
   Range("C18").Value = "N�o"
Else
   Range("C18").Value = "Sim"
End If
If Me.CheckBox12.Value = False Then
      Range("C19").Value = "N�o"
Else
   Range("C19").Value = "Sim"
End If
If Me.CheckBox13.Value = False Then
      Range("C20").Value = "N�o"
Else
   Range("C20").Value = "Sim"
End If



Range("D8:D20").Value = Me.TextBox2.Value 'copia o valor do textbox para o campo de observa��es da planilha
Range("A5").Value = "Elaborador: " & Me.ComboBox1.Text ' Copia valor do combobox
Range("D5:D6").Value = "Data: " & Me.TextBox1.Value 'copia valor data de realiza��o do relat�rio



Application.PrintCommunication = False
With ActiveSheet.PageSetup
    .PrintTitleRows = ""
    .PrintTitleColumns = ""
End With
Application.PrintCommunication = True
ActiveSheet.PageSetup.PrintArea = ""
Application.PrintCommunication = False
With ActiveSheet.PageSetup
    .LeftHeader = ""
    .CenterHeader = ""
    .RightHeader = ""
    .LeftFooter = ""
    .CenterFooter = ""
    .RightFooter = ""
    .LeftMargin = Application.InchesToPoints(0.7)
    .RightMargin = Application.InchesToPoints(0.7)
    .TopMargin = Application.InchesToPoints(0.75)
    .BottomMargin = Application.InchesToPoints(0.75)
    .HeaderMargin = Application.InchesToPoints(0.3)
    .FooterMargin = Application.InchesToPoints(0.3)
    .PrintHeadings = False
    .PrintGridlines = False
    .PrintComments = xlPrintNoComments
    .PrintQuality = 600
    .CenterHorizontally = False
    .CenterVertically = False
    .Orientation = xlPortrait
    .Draft = False
    .PaperSize = xlPaperA4
    .FirstPageNumber = xlAutomatic
    .Order = xlDownThenOver
    .BlackAndWhite = False
    .Zoom = False
    .FitToPagesWide = 1
    .FitToPagesTall = 1
    .PrintErrors = xlPrintErrorsDisplayed
    .OddAndEvenPagesHeaderFooter = False
    .DifferentFirstPageHeaderFooter = False
    .ScaleWithDocHeaderFooter = True
    .AlignMarginsHeaderFooter = True
    .EvenPage.LeftHeader.Text = ""
    .EvenPage.CenterHeader.Text = ""
    .EvenPage.RightHeader.Text = ""
    .EvenPage.LeftFooter.Text = ""
    .EvenPage.CenterFooter.Text = ""
    .EvenPage.RightFooter.Text = ""
    .FirstPage.LeftHeader.Text = ""
    .FirstPage.CenterHeader.Text = ""
    .FirstPage.RightHeader.Text = ""
    .FirstPage.LeftFooter.Text = ""
    .FirstPage.CenterFooter.Text = ""
    .FirstPage.RightFooter.Text = ""
End With


Data = CStr(Date)

Data = Format(Me.TextBox3.Value, "dd.mm.yyyy")

hora = CStr(Time)

hora = Format(Time, "hh.mm.ss")

Planilha = "ROP_Mensal"

Pasta = VBA.Format(Date, "mmmm")

nome = ThisWorkbook.Path & "\Mensal\" & Pasta & "\" & Planilha & "_" & Data & ".pdf"

nome2 = Planilha & "_" & Data

Application.PrintCommunication = True
ActiveWindow.SelectedSheets.PrintOut , ActivePrinter:="Microsoft Print to PDF", PrintToFile:=True, Copies:=1, Collate:=True, _
IgnorePrintAreas:=False, PrToFileName:=nome
Application.WindowState = xlNormal

Call enviar_email(nome, nome2)

ActiveWorkbook.Sheets("Calendario de inspe��o 2023").Select

Sheets("Mensal").Visible = xlSheetVeryHidden 'esconde a planilha ap�s a opera��o

End Sub


