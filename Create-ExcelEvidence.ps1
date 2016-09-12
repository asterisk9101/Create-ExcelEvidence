param($Path)

function logger($msg) {
    $msg = (Get-Date -Format "u") + "`t" + $msg
    Write-Host $msg
}

logger "Processing..."
Add-Type -AssemblyName System.Windows.Forms

$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $true
logger "Connected to Excel."
if ($path) {
    try {
        $Excel.Workbooks.Open($Path) > $null
        logger "Open $Path"
    } catch {
        logger "File Open Failed. => $Path"
        $Excel.Quit()
        return;
    }
} else {
    $Excel.Workbooks.Add() > $null
}

[System.Windows.Forms.Clipboard]::SetDataObject("")
logger "Clear Clipboard."
while($true){
    if ([system.windows.forms.Clipboard]::ContainsImage()) {
        $Excel.ActiveCell.Value2 = Get-Date -Format "\'yyyy/MM/dd HH:mm:ss"
        $Excel.ActiveCell.Offset(1, 0).Select() > $null
        $Excel.ActiveSheet.Paste()
        $Graphic = $Excel.Selection
        $Offset = $Excel.ActiveCell.Top + $Graphic.Height
        measure-command {
            $Excel.Caption = "Pasting..."
            $cell = $Excel.ActiveCell
            $cell = $cell.Offset([Math]::Ceiling($Graphic.Height / $cell.Height), 0)
            while ($cell.Top -lt $Offset) { $cell = $cell.Offset( 5,0)}
            while ($cell.Top -gt $Offset) { $cell = $cell.Offset(-5,0)}
            while ($cell.Top -lt $Offset) { $cell = $cell.Offset( 1,0)}
            while ($cell.Top -gt $Offset) { $cell = $cell.Offset(-1,0)}
            $cell.Offset(2,0).Select() > $null
            $Excel.Caption = $null
        } | % { logger $_.TotalMilliseconds.toString("#####.000ms").PadLeft(10) }
        [System.Windows.Forms.Clipboard]::SetDataObject("")
    }
    if ($Excel.Visible -eq $False) {
        logger "Excel Closed."
        break
    }
    Start-Sleep -Milliseconds 100
}
logger "Terminating..."
