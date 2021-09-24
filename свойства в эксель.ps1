# Импорт модуля Active Directory в сессию PowerShell
import-module activedirectory

# Сначала откройте книгу Excel:
$ExcelObj = New-Object -comobject Excel.Application
$ExcelWorkBook = $ExcelObj.Workbooks.Open("C:\Users\istarikov\Desktop\Документы\Скрипты\AD_PCs.xlsx")
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("PC_Win")

#Получение списка ПК в AD 
#Get-ADComputer -filter * -properties Name | select Name | Export-csv -path C:\Users\istarikov\Documents\ADcomputers.csv -Append -Encoding UTF8 

# Получаем количество заполненных строк в xlsx файле
$rowcount=$ExcelWorkSheet.UsedRange.Rows.Count

# Перебираем все строки в столбце 1, начиная со второй строки (в этих ячейках указано доменное имя пользователя)
for($i=2;$i -le $rowcount;$i++)
{
$ADpcname=$ExcelWorkSheet.Columns.Item(1).Rows.Item($i).Text

# Получаем значение атрибутов пользователя в AD
$ADpcProp = Get-ADComputer $ADpcname -properties OperatingSystem, IPv4Address |select-object OperatingSystem, IPv4Address #name добавить потом
#Заполняем ячейки данными из AD
$ExcelWorkSheet.Columns.Item(2).Rows.Item($i) = $ADpcProp.OperatingSystem
$ExcelWorkSheet.Columns.Item(3).Rows.Item($i) = $ADpcProp.IPv4Address
}
#Сохраните xls файл и закройте Excel
$ExcelWorkBook.Save()
$ExcelWorkBook.close($true)