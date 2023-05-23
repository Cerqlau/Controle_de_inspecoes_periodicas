VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Inspe��es a cada 15 dias "
   ClientHeight    =   9435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15105
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

'Imprimir papel

On Error GoTo Erro
 
    Sheets("Quinzenal").Visible = xlSheetVisible ' torna a planilha vis�vel
    
    ActiveWorkbook.Sheets("Quinzenal").Select
    
    Range("C8:C33").ClearContents
    Range("D8:D33").ClearContents
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
    
    Sheets("5 dias").Visible = xlSheetVeryHidden 'esconde a planilha ap�s a opera��o
    
'    ActiveWindow.SelectedSheets.PrintOut , ActivePrinter:="Microsoft Print to PDF", PrintToFile:=True, Copies:=1, Collate:=True, _
'        IgnorePrintAreas:=False
'    Application.WindowState = xlNormal
    
Erro:

End Sub

Private Sub CommandButton2_Click()

Call desprotege_relatorio

Dim data_limite As Date

Sheets("Quinzenal").Visible = xlSheetVisible

data_limite = CDate(Me.TextBox3.Value)


'verifica se o relat�rio j� foi elaborado anteriormente e informa o respons�vel e a data

ActiveWorkbook.Sheets("Relatorio ROP").Select
Range("F4").Select
Range(Selection, Selection.End(xlDown)).Select

For Each sel In Selection
contador = contador + 1
Next

For i = 3 To contador + 3

    If Cells(i + 1, 6).Value = CDate(Me.TextBox3.Value) And Cells(i + 1, 8).Value <> "" Then
    
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
If Me.CheckBox14.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox15.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox16.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox17.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox18.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox19.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox20.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox21.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox22.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox23.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox24.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox25.Value = False Then
   cont = cont + 1
End If
If Me.CheckBox26.Value = False Then
   cont = cont + 1
End If


   
If cont >= 1 And Me.TextBox2.Text = "" Then

    MsgBox "Existem " & Str(cont) & " itens que n�o foram verificados ou n�o foram justificados. Por favor refa�a a verifica��o e insira a justificativa no campo de observa��es"
   
Else

    'Verifica��o para salvamento na data correta
    
    ActiveWorkbook.Sheets("Relatorio ROP").Select
    Range("F4").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    For Each sel In Selection
    contador = contador + 1
    Next
    
    'loop de identifica��o das c�lulas para inserir as informa��es atrav�s da data limite
    For i = 3 To contador + 3
    
        If Cells(i + 1, 6).Value = CDate(Me.TextBox3.Value) And IsDate(Cells(i + 1, 7).Value) = False Then
        'And Cells(i + 1, 6).Value <= Date
        Cells(i + 1, 7).Value = Date 'data de realiza��o
        Cells(i + 1, 8).Value = Me.ComboBox1.Value 'elaboardor
        Cells(i + 1, 9).Value = Me.TextBox2.Value
        Cells(i + 1, 9).Select
        
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

Sheets("Quinzenal").Visible = xlSheetVeryHidden
Call protege_relatorio

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
.ColumnWidths = "30;250;30" 'largura

'coluna numera��o

.AddItem "1"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
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
.AddItem "14"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "15"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "16"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "17"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "18"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "19"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "20"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "21"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "22"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "23"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "24"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "25"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"
.AddItem "26"
.AddItem "---------------------------------------------------------------------------------------------------------------------------"


'coluna objeto de inspe��o

.List(0, 1) = "Olhal rotativo mesa popa BB"
.List(1, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(2, 1) = "olhal rotativo mesa popa BE"
.List(3, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(4, 1) = "Olhal rotativo mesa meio popa"
.List(5, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(6, 1) = "olhal rotativo mesa proa BB"
.List(7, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(8, 1) = "olhal rotativo mesa proa BE"
.List(9, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(10, 1) = "Olhal rotativo guincho #2"
.List(11, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(12, 1) = "Patesca guincho #7"
.List(13, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(14, 1) = "Patesca guincho #5"
.List(15, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(16, 1) = "Patesca mesa de trabalho"
.List(17, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(18, 1) = "Patesca guincho #8"
.List(19, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(20, 1) = "Patesca calha de popa BB"
.List(21, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(22, 1) = "Patesca guincho popa BE"
.List(23, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(24, 1) = "Parafusos e fura��o de Hangoff"
.List(25, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(26, 1) = "Talha do A&R de 600t"
.List(27, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(28, 1) = "Talha do A&R de 200t"
.List(29, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(30, 1) = "Ferramenta Pneum�tica para fita Bandit"
.List(31, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(32, 1) = "Parafusadeira Maclaren (grande)"
.List(33, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(34, 1) = "Parafusadeira Maclaren (m�dia)"
.List(35, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(36, 1) = "Parafusadeira (pequena)"
.List(37, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(38, 1) = "Bomba pneum�tica 'sapo'"
.List(39, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(40, 1) = "Lixadeira pneum�tica"
.List(41, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(42, 1) = "Lixadeira el�trica"
.List(43, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(44, 1) = "Agulheiro Pneum�tico"
.List(45, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(46, 1) = "Mesa de inspe��o MCV"
.List(47, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(48, 1) = "Calhas de carregamento e Roletes"
.List(49, 1) = "---------------------------------------------------------------------------------------------------------------------------"
.List(50, 1) = "Lubrifica��o das Catracas (caixas de sapatas da torre)"
.List(51, 1) = "---------------------------------------------------------------------------------------------------------------------------"

'coluna a��o

.List(0, 2) = "Inspe��o visual e lubrifica��o"
.List(1, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(2, 2) = "Inspe��o visual e lubrifica��o"
.List(3, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(4, 2) = "Inspe��o visual e lubrifica��o"
.List(5, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(6, 2) = "Inspe��o visual e lubrifica��o"
.List(7, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(8, 2) = "Inspe��o visual e lubrifica��o"
.List(9, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(10, 2) = "Inspe��o visual e lubrifica��o"
.List(11, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(12, 2) = "Inspe��o visual e lubrifica��o"
.List(13, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(14, 2) = "Inspe��o visual e lubrifica��o"
.List(15, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(16, 2) = "Inspe��o visual e lubrifica��o"
.List(17, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(18, 2) = "Inspe��o visual e lubrifica��o"
.List(19, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(20, 2) = "Inspe��o visual e lubrifica��o"
.List(21, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(22, 2) = "Inspe��o visual e lubrifica��o"
.List(23, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(24, 2) = "Inspe��o visual"
.List(25, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(26, 2) = "Inspe��o visual e teste funcional"
.List(27, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(28, 2) = "Inspe��o visual e teste funcional"
.List(29, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(30, 2) = "Inspe��o visual e teste funcional"
.List(31, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(32, 2) = "Inspe��o visual e teste funcional"
.List(33, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(34, 2) = "Inspe��o visual e teste funcional"
.List(35, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(36, 2) = "Inspe��o visual e teste funcional"
.List(37, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(38, 2) = "Inspe��o visual e teste funcional"
.List(39, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(40, 2) = "Inspe��o visual e teste funcional"
.List(41, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(42, 2) = "Inspe��o visual e teste funcional"
.List(43, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(44, 2) = "Inspe��o visual e teste funcional"
.List(45, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(46, 2) = "Inspe��o visual"
.List(47, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(48, 2) = "Inspe��o visual"
.List(49, 2) = "---------------------------------------------------------------------------------------------------------------------------"
.List(50, 2) = "Inspe��o visual"
.List(51, 2) = "---------------------------------------------------------------------------------------------------------------------------"

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


Sheets("5 dias").Visible = xlSheetVisible 'torna a planilha vis�vel

ActiveWorkbook.Sheets("5 dias").Select 'seleciona a planilha

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
If Me.CheckBox14.Value = False Then
      Range("C21").Value = "N�o"
Else
   Range("C21").Value = "Sim"
End If
If Me.CheckBox15.Value = False Then
      Range("C22").Value = "N�o"
Else
   Range("C22").Value = "Sim"
End If
If Me.CheckBox16.Value = False Then
      Range("C23").Value = "N�o"
Else
   Range("C23").Value = "Sim"
End If
If Me.CheckBox17.Value = False Then
      Range("C24").Value = "N�o"
Else
   Range("C24").Value = "Sim"
End If
If Me.CheckBox18.Value = False Then
      Range("C25").Value = "N�o"
Else
   Range("C25").Value = "Sim"
End If
If Me.CheckBox19.Value = False Then
      Range("C26").Value = "N�o"
Else
   Range("C26").Value = "Sim"
End If
If Me.CheckBox20.Value = False Then
      Range("C27").Value = "N�o"
Else
   Range("C27").Value = "Sim"
End If
If Me.CheckBox21.Value = False Then
      Range("C28").Value = "N�o"
Else
   Range("C28").Value = "Sim"
End If
If Me.CheckBox22.Value = False Then
      Range("C29").Value = "N�o"
Else
   Range("C29").Value = "Sim"
End If
If Me.CheckBox23.Value = False Then
    Range("C30").Value = "N�o"
Else
   Range("C30").Value = "Sim"
End If
If Me.CheckBox24.Value = False Then
      Range("C31").Value = "N�o"
Else
   Range("C31").Value = "Sim"
End If
If Me.CheckBox25.Value = False Then
      Range("C32").Value = "N�o"
Else
   Range("C32").Value = "Sim"
End If
If Me.CheckBox26.Value = False Then
      Range("C33").Value = "N�o"
Else
   Range("C33").Value = "Sim"
End If


Range("D8:D33").Value = Me.TextBox2.Value 'copia o valor do textbox para o campo de observa��es da planilha
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

Planilha = "ROP_15_dias"

Pasta = VBA.Format(Date, "mmmm")

nome = ThisWorkbook.Path & "\15 dias\" & Pasta & "\" & Planilha & "_" & Data & ".pdf"

nome2 = Planilha & "_" & Data

Application.PrintCommunication = True
ActiveWindow.SelectedSheets.PrintOut , ActivePrinter:="Microsoft Print to PDF", PrintToFile:=True, Copies:=1, Collate:=True, _
IgnorePrintAreas:=False, PrToFileName:=nome
Application.WindowState = xlNormal

Call enviar_email(nome, nome2)

ActiveWorkbook.Sheets("Calendario de inspe��o 2023").Select

Sheets("Quinzenal").Visible = xlSheetVeryHidden 'esconde a planilha ap�s a opera��o

End Sub


