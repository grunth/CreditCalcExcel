Option Explicit

Private Sub UserForm_Initialize()
'Инициализация
'Заполняем комбы
TextBox1.SetFocus
ComboBox1.AddItem "Год"
ComboBox1.Value = "Год"
ComboBox1.AddItem "Месяц"

ComboBox2.AddItem "Аннуитентные платежи"
ComboBox2.Value = "Аннуитентные платежи"
ComboBox2.AddItem "Дифференцированные платежи"
End Sub

Private Sub CalcButton_Click()
If fCheckValues = True Then
    Call Clear
    If ComboBox2 = "Аннуитентные платежи" Then
            If ComboBox1.Value = "Год" Then
            Call CalcAnnuitent("Year") 'Вычисляем и выводим если выбран год
            Else: Call CalcAnnuitent("Month")
            End If
        Else:
        If ComboBox1.Value = "Год" Then
        Call CalcDifferent("Year") 'Вычисляем и выводим если выбран год
        Else: Call CalcDifferent("Month")
        End If
    End If
    Else: Exit Sub
End If
End Sub

Private Sub ClearButton_Click()
Call Clear
End Sub

Private Sub TextBox1_Change()
'Этот код не даст ввести буквы или какую-нибудь хрень
If Right(TextBox1.Value, 1) = "-" Then TextBox1.Value = ""
If Len(TextBox1.Value) > 0 Then
If Len(TextBox1.Value) = 1 And TextBox1.Value = "-" Then Exit Sub
If IsNumeric(TextBox1.Value) = False Or Right(TextBox1.Value, 1) = "+" Or Right(TextBox1.Value, 1) = "-" = True Then TextBox1.Value = Left(TextBox1.Value, Len(TextBox1.Value) - 1)
End If
End Sub

Private Sub TextBox2_Change()
'Этот код не даст ввести буквы или какую-нибудь хрень
If Right(TextBox2.Value, 1) = "-" Then TextBox2.Value = ""
If Len(TextBox2.Value) > 0 Then
If Len(TextBox2.Value) = 1 And TextBox2.Value = "-" Then Exit Sub
If IsNumeric(TextBox2.Value) = False Or Right(TextBox2.Value, 1) = "+" Or Right(TextBox2.Value, 1) = "-" = True Then TextBox2.Value = Left(TextBox2.Value, Len(TextBox2.Value) - 1)
End If
End Sub

Private Sub TextBox3_Change()
'Этот код не даст ввести буквы или какую-нибудь хрень
If Right(TextBox3.Value, 1) = "-" Then TextBox3.Value = ""
If Len(TextBox3.Value) > 0 Then
If Len(TextBox3.Value) = 1 And TextBox3.Value = "-" Then Exit Sub
If IsNumeric(TextBox3.Value) = False Or Right(TextBox3.Value, 1) = "+" Or Right(TextBox3.Value, 1) = "-" = True Then TextBox3.Value = Left(TextBox3.Value, Len(TextBox3.Value) - 1)
End If
End Sub

Private Sub CalcAnnuitent(flag As String)
'Аннуитентные платежи (равные ежемесячные платежи). Вычисляет и выводит их на лист.
'Dim  As Currency 'D - Общая сумма кредита, Y - Ежемесячный платеж
Dim i, D, Y, Procents, SumProcents, SumDolg, Dolg As Single 'i - процентная ставка как Ставка*0,01
Dim n, month, temp, NumRow  As Integer 'n - срок погашения в годах
D = CSng(TextBox1.Text)
i = CSng(TextBox2.Text) / 100
If flag = "Year" Then
n = CInt(TextBox3.Text)
month = n * 12
Else:
n = CInt(TextBox3.Text) / 12 'Срок кредита в годах
month = CInt(TextBox3.Text) 'Кол-во месяцев
End If
Y = D * i / 12 / (1 - 1 / (1 + i / 12) ^ (n * 12)) 'Ежемесячный платеж
Worksheets("Расчет_кредита").Cells(1, 2).Value = D
Worksheets("Расчет_кредита").Cells(1, 4).Value = i
Worksheets("Расчет_кредита").Cells(2, 2).Value = n
Worksheets("Расчет_кредита").Cells(2, 4).Value = ComboBox2.Value
'Заполняем поля
NumRow = 4
Procents = 0
SumProcents = 0
SumDolg = 0
Dolg = 0
For temp = 0 To month Step 1
Cells(NumRow, 1).Value = temp
Cells(NumRow, 2).Value = D
If temp > 0 Then
Cells(NumRow, 3).Value = Y
Cells(NumRow, 4).Value = Procents
Cells(NumRow, 5).Value = Dolg
Else
Cells(NumRow, 3).Value = "-"
Cells(NumRow, 4).Value = "-"
Cells(NumRow, 5).Value = "-"
End If
NumRow = NumRow + 1
Procents = D * i / 12 'Процентные платежи
Dolg = Y - Procents
D = D - (Y - Procents)
SumProcents = SumProcents + Procents
SumDolg = SumDolg + Dolg
Next temp
Cells(NumRow, 1).Value = "ИТОГО"
Cells(NumRow, 3).Value = Y * month
Cells(NumRow, 4).Value = SumProcents
Cells(NumRow, 5).Value = SumDolg
End Sub

Private Sub CalcDifferent(flag As String)
'Дифференцированные платежи. Вычисляет и выводит их на лист.
'Dim  As Currency 'D - Общая сумма кредита, Y - Ежемесячный платеж
Dim i, D, Y, Procents, YSum, Dolg, SumProcents As Single 'i - процентная ставка как Ставка*0,01
Dim n, month, temp, NumRow  As Integer 'n - срок погашения в годах
D = CSng(TextBox1.Text)
i = CSng(TextBox2.Text) / 100
If flag = "Year" Then
n = CInt(TextBox3.Text) 'Срок погашения в годах
month = n * 12 'Число месяцев
Dolg = D / (n * 12) 'Ежемесячная сумма погашения основного долга
Else:
n = CInt(TextBox3.Text) / 12
month = CInt(TextBox3.Text)
Dolg = D / (n * 12) 'Ежемесячная сумма погашения основного долга
End If
'Заполняем поля
NumRow = 4
YSum = 0
SumProcents = 0
Worksheets("Расчет_кредита").Cells(1, 2).Value = D
Worksheets("Расчет_кредита").Cells(1, 4).Value = i
Worksheets("Расчет_кредита").Cells(2, 2).Value = n
Worksheets("Расчет_кредита").Cells(2, 4).Value = ComboBox2.Value
Cells(NumRow, 1).Value = 0
Cells(NumRow, 2).Value = D
Cells(NumRow, 3).Value = "-"
Cells(NumRow, 4).Value = "-"
Cells(NumRow, 5).Value = "-"
For temp = 1 To month Step 1
NumRow = NumRow + 1
Procents = (D - (temp - 1) * D / (n * 12)) * i / 12
Y = Dolg + Procents
D = D - Dolg
Cells(NumRow, 1).Value = temp
Cells(NumRow, 2).Value = D
Cells(NumRow, 3).Value = Y
Cells(NumRow, 4).Value = Procents
Cells(NumRow, 5).Value = Dolg
YSum = YSum + Y
SumProcents = SumProcents + Procents
Next temp
Cells(NumRow + 1, 1).Value = "ИТОГО"
Cells(NumRow + 1, 3).Value = YSum
Cells(NumRow + 1, 4).Value = SumProcents
Cells(NumRow + 1, 5).Value = Dolg * month
End Sub

Private Sub Clear()
'Очищаем старые данные
Dim aRange As Range
Dim lastRow As Integer
Dim lastColumn As Integer
Set aRange = Worksheets("Расчет_кредита").Range("A4").SpecialCells(xlCellTypeLastCell)
lastRow = aRange.Row + 1
lastColumn = aRange.Column
Worksheets("Расчет_кредита").Range(Cells(4, 1), Cells(lastRow, lastColumn)).Delete
ThisWorkbook.Save
End Sub

Private Function fCheckValues() As Boolean
'Проверяет, что в текстовых полях есть значения
'Если чего-то не хватает то выдает сообщение и
'ставит курсор в первое поле где нет значения
If (Len(TextBox1.Value) = 0 Or Len(TextBox2.Value) = 0 Or Len(TextBox3.Value) = 0) Then
MsgBox ("Заполните все параметры!")
    If Len(TextBox1.Value) = 0 Then
    TextBox1.SetFocus
    ElseIf Len(TextBox2.Value) = 0 Then
    TextBox2.SetFocus
    Else: TextBox3.SetFocus
    End If
fCheckValues = False
Exit Function
Else: fCheckValues = True
End If
End Function
