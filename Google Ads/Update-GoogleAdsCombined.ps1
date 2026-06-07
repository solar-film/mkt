param(
    [string]$SourceDir = $PSScriptRoot,
    [string]$OutputFileName = "Google_Ads_Combined.csv"
)

$ErrorActionPreference = "Stop"

function Read-GoogleAdsCsv {
    param([System.IO.FileInfo]$File)

    $lines = [System.IO.File]::ReadAllLines($File.FullName, [System.Text.Encoding]::UTF8)
    if ($lines.Count -lt 3) {
        return $null
    }

    $sourceAccount = $lines[0].Trim()
    $reportRange = $lines[1].Trim()
    $csvText = ($lines[2..($lines.Count - 1)] -join [Environment]::NewLine)
    $rows = $csvText | ConvertFrom-Csv

    [pscustomobject]@{
        File          = $File
        SourceAccount = $sourceAccount
        ReportRange   = $reportRange
        Headers       = @($rows | Select-Object -First 1 | ForEach-Object { $_.PSObject.Properties.Name })
        Rows          = @($rows)
    }
}

$sourcePath = (Resolve-Path -LiteralPath $SourceDir).Path
$outputPath = Join-Path $sourcePath $OutputFileName

$inputFiles = Get-ChildItem -LiteralPath $sourcePath -Filter "*.csv" |
    Where-Object {
        $_.FullName -ne $outputPath -and
        $_.Name -notlike "*Combined*" -and
        $_.Name -notlike "*รวม*"
    } |
    Sort-Object Name

$dataSets = @($inputFiles | ForEach-Object { Read-GoogleAdsCsv -File $_ } | Where-Object { $null -ne $_ })

$headers = New-Object System.Collections.Generic.List[string]
foreach ($name in @("SourceAccount", "SourceFile", "ReportRange")) {
    [void]$headers.Add($name)
}

foreach ($dataSet in $dataSets) {
    foreach ($header in $dataSet.Headers) {
        if (-not $headers.Contains($header)) {
            [void]$headers.Add($header)
        }
    }
}

$combinedRows = foreach ($dataSet in $dataSets) {
    foreach ($row in $dataSet.Rows) {
        $combined = [ordered]@{}
        foreach ($header in $headers) {
            $combined[$header] = ""
        }

        $combined["SourceAccount"] = $dataSet.SourceAccount
        $combined["SourceFile"] = $dataSet.File.Name
        $combined["ReportRange"] = $dataSet.ReportRange

        foreach ($property in $row.PSObject.Properties) {
            $combined[$property.Name] = $property.Value
        }

        [pscustomobject]$combined
    }
}

$combinedRows | Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding UTF8

[pscustomobject]@{
    OutputFile = $outputPath
    InputFiles = $inputFiles.Count
    Rows       = @($combinedRows).Count
}
