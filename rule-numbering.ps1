# ������� ��������� ����������� ������ ��������� ������������ ������ �� �������.
function Get-Last-Row ($worksheet)  {

$xlCellTypeLastCell = 11
$used = $worksheet.usedRange 
$lastCell = $used.SpecialCells($xlCellTypeLastCell) 
$row = $lastCell.row 
return $row
}


$file = Read-Host -Prompt "������� ���� � ����� � ������� ���������� ���������� ��������� ������ (������������ ��������)"

# ��������� ������ ���� � ������� Policy
$excel = New-Object -comobject Excel.Application

# ��������� ������� �� ������������� ���������� ��������, �������� ���� Excel, �������� ���������� ���� Excel (�������� ��������������).
$Excel.Visible = $true


$workbook = $excel.workbooks.Open($file)
$worksheet = $workbook.worksheets.item("Policy")

# �������� ����� �������� ������������ �� ����� ������, ��������� �������� ������ � ���������� �����.
$LastRowAddr = Get-Last-Row($worksheet)
Write-Host (" ������������� ����� ������������ �� ������� Policy ������: " +$LastRowAddr)

# ������� �������� ������ �������
$CurrRuleNum = 1;
for ($row = 1; $row -le $LastRowAddr; $row++) {
    
    # ������� ������ - ����� ��� ���������
    if ($worksheet.Cells.Item($row,1).MergeCells -eq $true) {
        Write-Host ("������ ��������� � �����������. ������������� ��� ���� ��������� ���� ����� �������. ���������� ������: " +$row )
        continue
    }

    # ������� ������ ����� � ������ ����� � ������ ������ � ������ �� ����� - ��� �������, ������ ����� ����������.
    if (([string]::IsNullOrEmpty($Worksheet.Cells.Item($row,1).value2)) -and (!([string]::IsNullOrEmpty($Worksheet.Cells.Item($row,2).value2)))) {
        Write-Host ("������ �����, �������. ������� ������: " +$row )
        $worksheet.Cells.Item($row,1) = "$CurrRuleNum"
        $CurrRuleNum++
        continue
    }

}

