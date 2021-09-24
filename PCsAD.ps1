# ������ ������ Active Directory � ������ PowerShell
import-module activedirectory

# ������� �������� ����� Excel:
$ExcelObj = New-Object -comobject Excel.Application
$ExcelWorkBook = $ExcelObj.Workbooks.Open("C:\Users\istarikov\Desktop\���������\�������\AD_PCs.xlsx")
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("PC_Win")

# �������� ���������� ����������� ����� � xlsx �����
$rowcount=$ExcelWorkSheet.UsedRange.Rows.Count

#��������� ������ ������� ������� ��
for($i=2;$i -eq $usercount;$i++)
{
#�������� ��� �� � AD
$usercount=2
$ADpcProp = Get-ADComputer -properties name |select-object name
$usercount++
#��������� ������ ������� �� ��
$ExcelWorkSheet.Columns.Item(1).Rows.Item($i) = $ADpcProp.Name
}
$ExcelWorkBook.Save()
$ExcelWorkBook.close($true)