#requires -Version 5.1
<#
.SYNOPSIS
Erstellt lokal Jahresarchiv.csv aus genau einer Jahresquelle je Kalenderjahr.
.DESCRIPTION
CSV: Windows-1252, Semikolon. XLSX/XLSM: Excel Desktop erforderlich.
Jahresdateien werden nur lesend geöffnet; Makros und externe Links bleiben aus.
Nur die erste ganze Zeile je Fallnummer wird berücksichtigt, danach wird nach
dem Aufnahme-Kalenderjahr gefiltert. Keine Zusammenführung späterer Zeilen.
Das Archiv enthält keine Fallnummern, Patienten- oder Mitarbeiternamen.
Aufnahmedatum, Fachabteilung und Dauern bleiben fallbezogene Statistikdaten.
Mit -DashboardPath wird eine neue Kopie erstellt und deren Archivabfrage
aktiviert/aktualisiert. Das ursprüngliche Dashboard bleibt unverändert.
.EXAMPLE
.\Jahresarchiv_erstellen.ps1 -ManifestPath .\Jahresarchiv_Quellen.json
.EXAMPLE
.\Jahresarchiv_erstellen.ps1 -ManifestPath .\Jahresarchiv_Quellen.json -DashboardPath 'S:\...\StatBoardReakt_2026_Dashboard.xlsm'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ManifestPath,
    [string]$OutputCsv,
    [string]$DashboardPath,
    [string]$UpdatedDashboardPath
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$invariant = [System.Globalization.CultureInfo]::InvariantCulture
$rawHeaders = @('Fallnummer','FA','Aufnahmedatum','Pat. Nachname','Pat Vorname',
    'Anf Dat','Anf Uhrzeit','Befund Anlage Dat','Befund Anlage Uhrzeit',
    'Durchf_Timestamp','Datenerfassung Anlagedatum','Datenerfassung Anlageuhrzeit',
    'Vidiert am','Vidiert um','Nachname','Vorname')
$usedIndexes = @(0,1,2,5,6,7,8,9,12,13)
$archiveHeaders = @('AufnahmeDatum','Jahr','Monat','FA','Dauer1_s','Dauer2_s',
    'Dauer3_s','Dauer4_s','Dauer5_s','QuelleJahr','QuelleDatei','QuellstandDate')
$defaultCsv = 'S:\Belegungsmanagement\Abteilungslaufwerk\Fachbereiche\Sozialdienst\Statistik\Jahresarchiv.csv'
$script:excelApplication = $null

function Get-Option($sourceObject, [string]$propertyName, $defaultValue) {
    $property = $sourceObject.PSObject.Properties[$propertyName]
    if ($null -ne $property) { return $property.Value }
    return $defaultValue
}

function Get-LocalPath([string]$pathText, [string]$baseDirectory) {
    if ([string]::IsNullOrEmpty($pathText)) { throw 'Ein Dateipfad im Manifest fehlt.' }
    if ([IO.Path]::IsPathRooted($pathText)) { return [IO.Path]::GetFullPath($pathText) }
    return [IO.Path]::GetFullPath([IO.Path]::Combine($baseDirectory, $pathText))
}

function Get-ColumnNumber($columnValue) {
    if ($columnValue -is [int] -or $columnValue -is [long]) {
        if ($columnValue -ge 1 -and $columnValue -le 16384) { return [int]$columnValue }
    }
    $columnText = ([string]$columnValue).ToUpperInvariant()
    if ($columnText -notmatch '^[A-Z]{1,3}$') { throw 'Spaltenzuordnung muss Excel-Buchstaben oder Nummern 1 bis 16384 verwenden.' }
    $columnNumber = 0
    foreach ($character in $columnText.ToCharArray()) {
        $columnNumber = $columnNumber * 26 + [int]$character - 64
    }
    if ($columnNumber -gt 16384) { throw 'Spaltenzuordnung liegt hinter XFD.' }
    return $columnNumber
}

function Get-ColumnLetter([int]$columnNumber) {
    $columnLetter = ''
    while ($columnNumber -gt 0) {
        $columnNumber--
        $columnLetter = [string][char](65 + ($columnNumber % 26)) + $columnLetter
        $columnNumber = [int][Math]::Floor($columnNumber / 26)
    }
    return $columnLetter
}

function Normalize-Header($headerValue) {
    return ([regex]::Replace(([string]$headerValue).Replace([char]160, ' ').Trim(), '\s+', ' ')).ToLowerInvariant()
}

function Convert-ToDate($dateValue, [bool]$date1904) {
    if ($null -eq $dateValue -or [string]$dateValue -eq '') { return $null }
    if ($dateValue -is [double] -or $dateValue -is [int] -or $dateValue -is [long] -or $dateValue -is [decimal]) {
        $serial = [double]$dateValue
        if ([double]::IsNaN($serial) -or [double]::IsInfinity($serial) -or $serial -lt 0) { return $null }
        if (-not $date1904 -and $serial -ge 60 -and $serial -lt 61) { return $null }
        if ($date1904) { $epoch = [datetime]::new(1904,1,1) }
        else {
            $epoch = [datetime]::new(1899,12,31)
            if ($serial -ge 60) { $serial-- }
        }
        try { return $epoch.AddSeconds([Math]::Round($serial * 86400)) }
        catch { return $null }
    }
    $formats = [string[]]@('dd.MM.yyyy HH:mm:ss','dd.MM.yyyy HH:mm','dd.MM.yyyy',
        'yyyy-MM-dd HH:mm:ss','yyyy-MM-dd HH:mm','yyyy-MM-dd',
        "yyyy-MM-dd'T'HH:mm:ss","yyyy-MM-dd'T'HH:mm")
    $parsedDate = [datetime]::MinValue
    if ([datetime]::TryParseExact(([string]$dateValue).Trim(), $formats, $invariant,
        [Globalization.DateTimeStyles]::None, [ref]$parsedDate)) { return $parsedDate }
    return $null
}

function Convert-ToTimeSeconds($timeValue) {
    if ($null -eq $timeValue -or [string]$timeValue -eq '') { return $null }
    if ($timeValue -is [double] -or $timeValue -is [int] -or $timeValue -is [long] -or $timeValue -is [decimal]) {
        $fraction = [double]$timeValue
        if ([double]::IsNaN($fraction) -or [double]::IsInfinity($fraction) -or $fraction -lt 0 -or $fraction -ge 1) { return $null }
        $seconds = [long][Math]::Round($fraction * 86400)
        if ($seconds -ge 86400) { return $null }
        return $seconds
    }
    $parsedTime = [datetime]::MinValue
    if ([datetime]::TryParseExact(([string]$timeValue).Trim(), [string[]]@('HH:mm:ss','HH:mm'),
        $invariant, [Globalization.DateTimeStyles]::None, [ref]$parsedTime)) {
        return [long]$parsedTime.TimeOfDay.TotalSeconds
    }
    return $null
}

function Get-PairedTimestamp($sourceRow, [int]$dateIndex, [int]$timeIndex, [bool]$date1904) {
    $pairedDate = Convert-ToDate $sourceRow[$dateIndex] $date1904
    $pairedTimeSeconds = Convert-ToTimeSeconds $sourceRow[$timeIndex]
    if ($null -eq $pairedDate -or $null -eq $pairedTimeSeconds) { return $null }
    return $pairedDate.Date.AddSeconds($pairedTimeSeconds)
}

function Write-ArchiveRow($csvWriter, [object[]]$values) {
    $escaped = foreach ($value in $values) {
        '"' + ([string]$value).Replace('"','""') + '"'
    }
    $csvWriter.WriteLine(($escaped -join ';'))
}

function Get-ExcelApplication {
    if ($null -eq $script:excelApplication) {
        $script:excelApplication = New-Object -ComObject Excel.Application
        $script:excelApplication.Visible = $false
        $script:excelApplication.DisplayAlerts = $false
        $script:excelApplication.AskToUpdateLinks = $false
        $script:excelApplication.EnableEvents = $false
        $script:excelApplication.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
    }
    return $script:excelApplication
}

function Test-EncryptedOfficeFile([string]$filePath) {
    $stream = [IO.File]::OpenRead($filePath)
    try {
        $prefix = New-Object byte[] 8
        $readCount = $stream.Read($prefix, 0, 8)
        return ($readCount -eq 8 -and ([BitConverter]::ToString($prefix)) -eq 'D0-CF-11-E0-A1-B1-1A-E1')
    }
    finally { $stream.Dispose() }
}

function Open-SafeWorkbook([string]$workbookPath, [bool]$readOnly) {
    $application = Get-ExcelApplication
    $passwordText = ''
    $securePassword = $null
    $passwordPointer = [IntPtr]::Zero
    try {
        if (Test-EncryptedOfficeFile $workbookPath) {
            $securePassword = Read-Host -AsSecureString ('Excel-Kennwort für ' + [IO.Path]::GetFileName($workbookPath))
            $passwordPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
            $passwordText = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($passwordPointer)
        }
        # UpdateLinks=0; no macros/events; no recent-file list; no repair mode.
        return $application.Workbooks.Open($workbookPath, 0, $readOnly,
            [Type]::Missing, $passwordText, [Type]::Missing, $true,
            [Type]::Missing, [Type]::Missing, $false, $false,
            [Type]::Missing, $false, $true, 0)
    }
    finally {
        if ($passwordPointer -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordPointer) }
        $passwordText = $null
        if ($null -ne $securePassword) { $securePassword.Dispose() }
    }
}

function Add-CaseToArchive($sourceRow, [bool]$date1904, $sourceConfig,
        $seenCases, $statistics, $csvWriter) {
    $statistics.source_rows++
    $caseValue = $sourceRow[0]
    if ($null -eq $caseValue -or [string]$caseValue -eq '') { $statistics.missing_case_number_rows++; return }
    if ($caseValue -is [double]) {
        if ($caseValue -eq [Math]::Truncate($caseValue)) { $caseKey = $caseValue.ToString('0', $invariant) }
        else { $caseKey = $caseValue.ToString('G17', $invariant) }
    }
    else { $caseKey = [string]$caseValue }
    if (-not $seenCases.Add($caseKey)) { $statistics.later_duplicate_rows++; return }
    $admission = Convert-ToDate $sourceRow[2] $date1904
    if ($null -eq $admission) { $statistics.invalid_admission_rows++; return }
    if ($admission.Year -ne $sourceConfig.year) { $statistics.other_admission_year_rows++; return }
    $request = Get-PairedTimestamp $sourceRow 5 6 $date1904
    $sighting = Get-PairedTimestamp $sourceRow 7 8 $date1904
    $performance = Convert-ToDate $sourceRow[9] $date1904
    $signoff = Get-PairedTimestamp $sourceRow 12 13 $date1904
    $starts = @($admission,$request,$request,$performance,$request)
    $ends = @($request,$sighting,$performance,$signoff,$signoff)
    $durations = New-Object object[] 5
    for ($metricIndex=0; $metricIndex -lt 5; $metricIndex++) {
        if ($null -eq $starts[$metricIndex] -or $null -eq $ends[$metricIndex]) {
            $durations[$metricIndex] = ''; $statistics.missing_duration_counts[$metricIndex]++
        }
        else {
            $seconds = [long]($ends[$metricIndex] - $starts[$metricIndex]).TotalSeconds
            if ($seconds -lt 0) { $durations[$metricIndex] = ''; $statistics.negative_duration_counts[$metricIndex]++ }
            else {
                $durations[$metricIndex] = $seconds
                $statistics.valid_duration_counts[$metricIndex]++
                $statistics.duration_sum_seconds[$metricIndex] += $seconds
            }
        }
    }
    $department = [regex]::Replace(([string]$sourceRow[1]).Replace([char]160, ' ').Trim(), '\s+', ' ')
    $values = @($admission.ToString('yyyy-MM-dd', $invariant), $admission.Year,
        $admission.Month, $department) + $durations + @($sourceConfig.year,
        [IO.Path]::GetFileName($sourceConfig.path), $sourceConfig.stand_date)
    Write-ArchiveRow $csvWriter $values
    $statistics.archive_rows++
    $monthKey = [string]$admission.Month
    if (-not $statistics.months.Contains($monthKey)) { $statistics.months[$monthKey] = 0 }
    $statistics.months[$monthKey]++
}

function Export-CsvSource($sourceConfig, $statistics, $csvWriter) {
    Add-Type -AssemblyName Microsoft.VisualBasic
    $encodingName = Get-Option $sourceConfig 'encoding' 'cp1252'
    if ($encodingName -eq 'cp1252') { $sourceEncoding = [Text.Encoding]::GetEncoding(1252) }
    else { $sourceEncoding = [Text.Encoding]::GetEncoding([string]$encodingName) }
    $reader = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($sourceConfig.path, $sourceEncoding, $true)
    try {
        $reader.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $reader.SetDelimiters([string](Get-Option $sourceConfig 'delimiter' ';'))
        $reader.HasFieldsEnclosedInQuotes = $true
        $reader.TrimWhiteSpace = $false
        if ($reader.EndOfData) { throw 'CSV-Quelle enthält keine Kopfzeile.' }
        $headers = $reader.ReadFields()
        $mapping = Get-Option $sourceConfig 'columns' ([pscustomobject]@{})
        $columnIndexes = @{}
        foreach ($rawIndex in $usedIndexes) {
            $expectedHeader = Normalize-Header (Get-Option $mapping $rawHeaders[$rawIndex] $rawHeaders[$rawIndex])
            $matches = @()
            for ($headerIndex=0; $headerIndex -lt $headers.Count; $headerIndex++) {
                if ((Normalize-Header $headers[$headerIndex]) -eq $expectedHeader) { $matches += $headerIndex }
            }
            if ($matches.Count -ne 1) { throw ('CSV-Kopf fehlt oder ist mehrdeutig: Quellspalte ' + ($rawIndex + 1)) }
            $columnIndexes[$rawIndex] = $matches[0]
        }
        $seenCases = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        while (-not $reader.EndOfData) {
            $fields = $reader.ReadFields()
            if ($fields.Count -ne $headers.Count) { throw 'CSV-Zeile hat eine andere Spaltenzahl als der Dateikopf.' }
            $sourceRow = New-Object object[] 16
            foreach ($rawIndex in $usedIndexes) { $sourceRow[$rawIndex] = $fields[$columnIndexes[$rawIndex]] }
            Add-CaseToArchive $sourceRow $false $sourceConfig $seenCases $statistics $csvWriter
        }
    }
    finally { $reader.Dispose() }
}

function Export-WorkbookSource($sourceConfig, $statistics, $csvWriter) {
    $workbook = $null
    $worksheet = $null
    $range = $null
    try {
        $workbook = Open-SafeWorkbook $sourceConfig.path $true
        $worksheet = $workbook.Worksheets.Item([string](Get-Option $sourceConfig 'sheet' 'Rohdatenimport'))
        $headerRow = [int](Get-Option $sourceConfig 'header_row' 12)
        $firstDataRow = [int](Get-Option $sourceConfig 'first_data_row' ($headerRow + 1))
        if ($headerRow -lt 1 -or $firstDataRow -le $headerRow -or $firstDataRow -gt 1048576) { throw 'Ungültige Kopf-/Datenzeile im Manifest.' }
        $mapping = Get-Option $sourceConfig 'columns' ([pscustomobject]@{})
        $columnNumbers = @{}
        $highestColumn = 1
        $mappedColumns = [Collections.Generic.HashSet[int]]::new()
        foreach ($rawIndex in $usedIndexes) {
            $columnNumber = Get-ColumnNumber (Get-Option $mapping $rawHeaders[$rawIndex] ($rawIndex + 1))
            if (-not $mappedColumns.Add($columnNumber)) { throw 'Mehrere Quellfelder sind derselben Excel-Spalte zugeordnet.' }
            $columnNumbers[$rawIndex] = $columnNumber
            if ($columnNumber -gt $highestColumn) { $highestColumn = $columnNumber }
            if ((Get-Option $sourceConfig 'validate_headers' $true) -and
                (Normalize-Header $worksheet.Cells.Item($headerRow,$columnNumber).Value2) -ne (Normalize-Header $rawHeaders[$rawIndex])) {
                throw ('Unerwarteter Excel-Kopf: Quellspalte ' + ($rawIndex + 1) + '. Zuordnung im Manifest prüfen.')
            }
        }
        $lastDataRow = [int]$worksheet.Cells.Item($worksheet.Rows.Count,$columnNumbers[0]).End(-4162).Row
        if ($lastDataRow -lt $firstDataRow) { return }
        $range = $worksheet.Range(('A' + $firstDataRow), ((Get-ColumnLetter $highestColumn) + $lastDataRow))
        $data = $range.Value2
        $date1904 = [bool]$workbook.Date1904
        $seenCases = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        for ($arrayRow=1; $arrayRow -le ($lastDataRow-$firstDataRow+1); $arrayRow++) {
            $sourceRow = New-Object object[] 16
            $hasRawData = $false
            foreach ($rawIndex in $usedIndexes) {
                $sourceRow[$rawIndex] = $data[$arrayRow,$columnNumbers[$rawIndex]]
                if ($null -ne $sourceRow[$rawIndex] -and [string]$sourceRow[$rawIndex] -ne '') { $hasRawData = $true }
            }
            if ($hasRawData) { Add-CaseToArchive $sourceRow $date1904 $sourceConfig $seenCases $statistics $csvWriter }
        }
        $data = $null
    }
    finally {
        if ($null -ne $range) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range) }
        if ($null -ne $worksheet) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($worksheet) }
        if ($null -ne $workbook) {
            $workbook.Close($false)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
        }
    }
}

function Enable-ArchiveInDashboard([string]$originalPath, [string]$newPath, [string]$csvPath) {
    if ([IO.File]::Exists($newPath)) { throw 'Ziel-Dashboard existiert bereits. Anderen Namen für -UpdatedDashboardPath wählen.' }
    if (-not [IO.File]::Exists($csvPath)) { throw 'Archiv-CSV fehlt; Archivabfrage wird nicht aktiviert.' }
    [IO.File]::Copy($originalPath, $newPath, $false)
    $workbook = $null
    $worksheet = $null
    $query = $null
    $saved = $false
    try {
        $workbook = Open-SafeWorkbook $newPath $false
        $worksheet = $workbook.Worksheets.Item('_Jahresarchiv')
        $query = $worksheet.QueryTables.Item('StatReaktJahresarchiv')
        $query.Connection = 'TEXT;' + $csvPath
        $query.TextFilePlatform = 1252
        $query.TextFileParseType = 1 # xlDelimited
        $query.TextFileStartRow = 2
        $query.FieldNames = $false
        $query.TextFileTextQualifier = 1 # xlTextQualifierDoubleQuote
        $query.TextFileSemicolonDelimiter = $true
        $query.TextFileTabDelimiter = $false
        $query.TextFileCommaDelimiter = $false
        $query.TextFileSpaceDelimiter = $false
        $query.TextFileConsecutiveDelimiter = $false
        $query.TextFileColumnDataTypes = [int[]]@(2,1,1,2,1,1,1,1,1,1,2,2)
        $query.EnableRefresh = $true
        $query.BackgroundQuery = $false
        $query.RefreshOnFileOpen = $false
        if (-not $query.Refresh($false)) { throw 'Excel hat die Archiv-Aktualisierung abgebrochen.' }
        $workbook.Save()
        $saved = $true
    }
    finally {
        if ($null -ne $query) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($query) }
        if ($null -ne $worksheet) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($worksheet) }
        if ($null -ne $workbook) {
            $workbook.Close($false)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
        }
        if (-not $saved -and [IO.File]::Exists($newPath)) { [IO.File]::Delete($newPath) }
    }
}

$temporaryCsv = $null
$csvWriter = $null
try {
    $manifestFullPath = [IO.Path]::GetFullPath($ManifestPath)
    $baseDirectory = [IO.Path]::GetDirectoryName($manifestFullPath)
    $manifest = [IO.File]::ReadAllText($manifestFullPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    $sources = @(Get-Option $manifest 'sources' @())
    $selectedSources = @()
    $seenYears = [Collections.Generic.HashSet[int]]::new()
    foreach ($source in $sources) {
        if ((Get-Option $source 'enabled' $true) -eq $false) { continue }
        $yearValue = Get-Option $source 'year' $null
        if (($yearValue -isnot [int] -and $yearValue -isnot [long]) -or $yearValue -lt 1900 -or $yearValue -gt 9999) {
            throw 'Jede aktivierte Quelle benötigt ein ganzzahliges Kalenderjahr (year).'
        }
        if (-not $seenYears.Add([int]$yearValue)) { throw ('Mehrere aktivierte Quellen für ' + $yearValue + ': genau eine Jahresdatei auswählen.') }
        $standDate = [datetime]::MinValue
        $standText = [string](Get-Option $source 'stand_date' '')
        if (-not [datetime]::TryParseExact($standText, 'yyyy-MM-dd', $invariant,
            [Globalization.DateTimeStyles]::None, [ref]$standDate)) { throw 'Jahresquelle benötigt den tatsächlichen Datenstand stand_date im Format yyyy-MM-dd.' }
        $sourcePath = Get-LocalPath (Get-Option $source 'path' '') $baseDirectory
        if (-not [IO.File]::Exists($sourcePath)) { throw ('Ausgewählte Jahresdatei fehlt: ' + [IO.Path]::GetFileName($sourcePath)) }
        $source.path = $sourcePath
        $selectedSources += $source
    }
    $selectedSources = @($selectedSources | Sort-Object year)
    if ([string]::IsNullOrEmpty($OutputCsv)) { $OutputCsv = Get-Option $manifest 'archive_csv' $defaultCsv }
    $csvFullPath = Get-LocalPath $OutputCsv $baseDirectory
    foreach ($source in $selectedSources) {
        if ($source.path -eq $csvFullPath) { throw 'Archivziel darf keine Jahresquelle überschreiben.' }
    }
    # Clinical archive CSV stays local and outside Git checkouts.
    $ancestor = [IO.DirectoryInfo]::new([IO.Path]::GetDirectoryName($csvFullPath))
    while ($null -ne $ancestor) {
        if ([IO.File]::Exists([IO.Path]::Combine($ancestor.FullName,'.git','HEAD')) -or
            [IO.File]::Exists([IO.Path]::Combine($ancestor.FullName,'.git'))) { throw 'Archiv-CSV außerhalb eines Git-Verzeichnisses speichern.' }
        $ancestor = $ancestor.Parent
    }
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($csvFullPath))
    $temporaryCsv = $csvFullPath + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'
    $strictEncoding = [Text.Encoding]::GetEncoding(1252,[Text.EncoderExceptionFallback]::new(),[Text.DecoderExceptionFallback]::new())
    $csvWriter = [IO.StreamWriter]::new($temporaryCsv,$false,$strictEncoding)
    $csvWriter.NewLine = "`r`n"
    Write-ArchiveRow $csvWriter $archiveHeaders
    $reports = @()
    foreach ($source in $selectedSources) {
        $statistics = [ordered]@{
            year=[int]$source.year; source_basename=[IO.Path]::GetFileName($source.path);
            stand_date=$source.stand_date; source_rows=0; later_duplicate_rows=0;
            missing_case_number_rows=0; invalid_admission_rows=0;
            other_admission_year_rows=0; archive_rows=0; months=[ordered]@{};
            valid_duration_counts=@(0,0,0,0,0); missing_duration_counts=@(0,0,0,0,0);
            negative_duration_counts=@(0,0,0,0,0); duration_sum_seconds=@([long]0,[long]0,[long]0,[long]0,[long]0)
        }
        $extension = [IO.Path]::GetExtension($source.path).ToLowerInvariant()
        if ($extension -eq '.csv') { Export-CsvSource $source $statistics $csvWriter }
        elseif ($extension -eq '.xlsx' -or $extension -eq '.xlsm') { Export-WorkbookSource $source $statistics $csvWriter }
        else { throw 'Quelle muss CSV, XLSX oder XLSM sein; alte XLS-Dateien vorher in Excel als XLSX speichern.' }
        $reports += [pscustomobject]$statistics
    }
    $csvWriter.Dispose(); $csvWriter = $null
    if ([IO.File]::Exists($csvFullPath)) { [IO.File]::Replace($temporaryCsv,$csvFullPath,$null) }
    else { [IO.File]::Move($temporaryCsv,$csvFullPath) }
    $temporaryCsv = $null
    $reportPath = $csvFullPath + '.pruefung.json'
    $reportJson = [pscustomobject]@{
        duration_unit='whole seconds'; case_numbers_exported=$false;
        names_exported=$false; sources=$reports
    } | ConvertTo-Json -Depth 7
    [IO.File]::WriteAllText($reportPath,$reportJson,[Text.UTF8Encoding]::new($false))
    $totalRows = 0
    foreach ($report in $reports) { $totalRows += $report.archive_rows }
    Write-Host ('Jahresarchiv erstellt: ' + $totalRows + ' Fälle; ' + $selectedSources.Count + ' Kalenderjahr(e).')
    Write-Host ('CSV: ' + $csvFullPath)
    Write-Host ('Prüfung ohne Namen/Fallnummern: ' + $reportPath)
    if (-not [string]::IsNullOrEmpty($DashboardPath)) {
        $dashboardFullPath = Get-LocalPath $DashboardPath $baseDirectory
        if ([string]::IsNullOrEmpty($UpdatedDashboardPath)) {
            $UpdatedDashboardPath = [IO.Path]::Combine([IO.Path]::GetDirectoryName($dashboardFullPath),
                [IO.Path]::GetFileNameWithoutExtension($dashboardFullPath) + '_mit_Jahresarchiv' + [IO.Path]::GetExtension($dashboardFullPath))
        }
        $newDashboardFullPath = Get-LocalPath $UpdatedDashboardPath $baseDirectory
        foreach ($source in $selectedSources) {
            if ($source.path -eq $newDashboardFullPath) { throw 'Neues Dashboard darf keine ausgewählte Jahresquelle überschreiben.' }
        }
        if ($newDashboardFullPath -eq $dashboardFullPath) { throw 'Für das aktualisierte Dashboard einen anderen Dateinamen wählen.' }
        Enable-ArchiveInDashboard $dashboardFullPath $newDashboardFullPath $csvFullPath
        Write-Host ('Neue Dashboard-Kopie mit aktivem Archiv: ' + $newDashboardFullPath)
    }
}
finally {
    if ($null -ne $csvWriter) { $csvWriter.Dispose() }
    if ($null -ne $temporaryCsv -and [IO.File]::Exists($temporaryCsv)) { [IO.File]::Delete($temporaryCsv) }
    if ($null -ne $script:excelApplication) {
        $script:excelApplication.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:excelApplication)
        $script:excelApplication = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}
