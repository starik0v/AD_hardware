# Импорт модуля Active Directory в сессию PowerShell
import-module activedirectory

# Сначала откройте книгу Excel:
$ExcelObj = New-Object -comobject Excel.Application
$ExcelWorkBook = $ExcelObj.Workbooks.Open("C:\Users\istarikov\Desktop\Документы\Скрипты\AD_PCs.xlsx")
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("PC_Win")

# Получаем количество заполненных строк в xlsx файле
$rowcount=$ExcelWorkSheet.UsedRange.Rows.Count

#Заполняем первый столбец именами ПК
for($i=2;$i -eq $usercount;$i++)
{
#Получаем имя ПК в AD
$usercount=2
$ADpcProp = Get-ADComputer -properties name |select-object name
$usercount++
#Заполняем ячейку данными из АД
$ExcelWorkSheet.Columns.Item(1).Rows.Item($i) = $ADpcProp.Name
}
$ExcelWorkBook.Save()
$ExcelWorkBook.close($true)