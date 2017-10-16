# Функция получения корректного номера последней используемой строки на вкладке.
function Get-Last-Row ($worksheet)  {

$xlCellTypeLastCell = 11
$used = $worksheet.usedRange 
$lastCell = $used.SpecialCells($xlCellTypeLastCell) 
$row = $lastCell.row 
return $row
}


$file = Read-Host -Prompt "Введите путь к файлу в котором необходимо произвести нумерацию правил (обязательный параметр)"

# Открываем нужный файл и вкладку Policy
$excel = New-Object -comobject Excel.Application

# Блокируем запросы на подтверждение выполнения операции, скрываем окно Excel, оключаем обновление окна Excel (повышаем быстродействие).
$Excel.Visible = $true


$workbook = $excel.workbooks.Open($file)
$worksheet = $workbook.worksheets.item("Policy")

# Получаем адрес последне используемой на листе ячейки, формируем диапазон поиска и производим поиск.
$LastRowAddr = Get-Last-Row($worksheet)
Write-Host (" Максимиальный номер используемой на вкладке Policy строки: " +$LastRowAddr)

# Счётчик текущего номера правила
$CurrRuleNum = 1;
for ($row = 1; $row -le $LastRowAddr; $row++) {
    
    # Текущая ячейка - шапка или заголовок
    if ($worksheet.Cells.Item($row,1).MergeCells -eq $true) {
        Write-Host ("Ячейка участвует в объединении. Следовательно это либо заголовок либо Шапка таблицы. Пропускаем строку: " +$row )
        continue
    }

    # Текущая первая чейка в строке пуста и вторая ячейка в строке не пуста - это правило, значит нужно нумеровать.
    if (([string]::IsNullOrEmpty($Worksheet.Cells.Item($row,1).value2)) -and (!([string]::IsNullOrEmpty($Worksheet.Cells.Item($row,2).value2)))) {
        Write-Host ("Ячейка пуста, нумерую. Текущая строка: " +$row )
        $worksheet.Cells.Item($row,1) = "$CurrRuleNum"
        $CurrRuleNum++
        continue
    }

}

