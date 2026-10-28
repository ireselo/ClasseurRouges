# Engine: Extraction PDF & Generation Excel (.xlsx) - Classeurs Rouges (CO Geneve)

Add-Type -AssemblyName System.Runtime.WindowsRuntime
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$xlsxCsPath = Join-Path $scriptDir "XlsxWriter.cs"
$nativePdfCsPath = Join-Path $scriptDir "NativePdfExtractor.cs"

if (-not ("ClasseursRouges.XlsxWriter" -as [type])) {
    $refs = @("System.dll", "System.Core.dll", "System.IO.Compression.dll", "System.IO.Compression.FileSystem.dll", "System.Drawing.dll")
    Add-Type -Path @($xlsxCsPath, $nativePdfCsPath) -ReferencedAssemblies $refs -ErrorAction Stop
}

$script:asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
})[0]
$script:asTaskAction = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction'
})[0]

function Await-WinRT($WinRtTask, $ResultType) {
    $asTask = $script:asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    return $netTask.Result
}

[Windows.Data.Pdf.PdfDocument,Windows.Data.Pdf,ContentType=WindowsRuntime] | Out-Null
[Windows.Data.Pdf.PdfPageRenderOptions,Windows.Data.Pdf,ContentType=WindowsRuntime] | Out-Null
[Windows.Storage.Streams.InMemoryRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime] | Out-Null
[Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics.Imaging,ContentType=WindowsRuntime] | Out-Null
[Windows.Graphics.Imaging.SoftwareBitmap,Windows.Graphics.Imaging,ContentType=WindowsRuntime] | Out-Null
[Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime] | Out-Null
[Windows.Globalization.Language,Windows.Globalization,ContentType=WindowsRuntime] | Out-Null

function Normalize-BulletinGrade([string]$val) {
    if ([string]::IsNullOrWhiteSpace($val)) { return "" }
    $v = $val.Trim() -replace ',', '.'
    if ($v -match '^(?i)c\.?m\.?$') { return "C.M" }
    if ($v -match '^(?i)disp\.?$') { return "Disp" }
    if ($v -match '^(?i)n\.?e\.?$') { return "N.E" }
    if ($v -match '^\(([1-6])\.([0-9])\)$') { return "($($Matches[1]).$($Matches[2]))" }
    if ($v -match '^\(([1-6])([0-9])\)$') { return "($($Matches[1]).$($Matches[2]))" }
    if ($v -match '^([1-6])\.([0-9])$') { return "$($Matches[1]).$($Matches[2])" }
    if ($v -match '^([1-6])([0-9])$') { return "$($Matches[1]).$($Matches[2])" }
    if ($v -match '^([1-6])1([0-9])$') { return "$($Matches[1]).$($Matches[2])" }
    return $null
}

function Clean-SubjectName([string]$name) {
    if ([string]::IsNullOrWhiteSpace($name)) { return "" }
    $n = $name.Trim() -replace '\s+', ' '
    $n = $n -replace '^[|.\-~:;]+\s*', ''
    $n = $n -replace '\s*[|.\-~:;]+$', ''
    $map = @{
        "Frangais" = "Français"
        "Francais" = "Français"
        "Mathematiques" = "Mathématiques"
        "Geographie" = "Géographie"
        "Activites creatrices et manuelles" = "Activités créatrices et manuelles"
    }
    foreach ($k in $map.Keys) {
        if ($n -ieq $k) { return $map[$k] }
    }
    return $n
}

function Extract-BulletinsFromPdf([string]$pdfPath) {
    if (-not (Test-Path $pdfPath)) {
        throw "Fichier introuvable : $pdfPath"
    }

    $results = @()

    $fs = [System.IO.File]::OpenRead($pdfPath)
    try {
        $winrtStream = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($fs)
        $pdfDoc = Await-WinRT ([Windows.Data.Pdf.PdfDocument]::LoadFromStreamAsync($winrtStream)) ([Windows.Data.Pdf.PdfDocument])
        $pageCount = $pdfDoc.PageCount

        $nativePages = [ClasseursRouges.NativePdfTextExtractor]::ExtractWordsPerPage($pdfPath)

        $lang = New-Object Windows.Globalization.Language("fr-FR")
        $ocrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
        if ($null -eq $ocrEngine) {
            $ocrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
        }
        $tessExe = "C:\Program Files\Tesseract-OCR\tesseract.exe"
        $hasTess = Test-Path $tessExe

        for ($p = 0; $p -lt $pageCount; $p++) {
            $page = $pdfDoc.GetPage([uint32]$p)
            $stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
            $options = New-Object Windows.Data.Pdf.PdfPageRenderOptions
            $targetWidth = 1984
            $scale = $targetWidth / $page.Size.Width
            $options.DestinationWidth = [uint32]$targetWidth
            $options.DestinationHeight = [uint32]($page.Size.Height * $scale)

            $renderTask = $page.RenderToStreamAsync($stream, $options)
            $netAction = $script:asTaskAction.Invoke($null, @($renderTask))
            $netAction.Wait(-1) | Out-Null

            $tempPng = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "agy_bulletin_p$p.png")
            $tempTsvBase = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "agy_bulletin_tsv_p$p")
            $outFs = [System.IO.File]::Create($tempPng)
            $stream.Seek(0)
            $netStream = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($stream)
            $netStream.CopyTo($outFs)
            $outFs.Close()

            $bmpFull = New-Object System.Drawing.Bitmap($tempPng)
            $prevW = 560
            $prevH = [int]($bmpFull.Height * ($prevW / [double]$bmpFull.Width))
            $bmpPrev = New-Object System.Drawing.Bitmap($bmpFull, $prevW, $prevH)
            $msJpg = New-Object System.IO.MemoryStream
            $bmpPrev.Save($msJpg, [System.Drawing.Imaging.ImageFormat]::Jpeg)
            $previewBase64 = "data:image/jpeg;base64," + [Convert]::ToBase64String($msJpg.ToArray())
            $msJpg.Dispose()
            $bmpPrev.Dispose()

            $stream.Seek(0)
            $decoder = Await-WinRT ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
            $softBmp = Await-WinRT ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
            $winrtRes = Await-WinRT ($ocrEngine.RecognizeAsync($softBmp)) ([Windows.Media.Ocr.OcrResult])

            $winrtWords = @()
            foreach ($line in $winrtRes.Lines) {
                foreach ($w in $line.Words) {
                    $winrtWords += [PSCustomObject]@{
                        Text = $w.Text.Trim()
                        X = [int]$w.BoundingRect.X
                        Y = [int]$w.BoundingRect.Y
                        W = [int]$w.BoundingRect.Width
                        H = [int]$w.BoundingRect.Height
                        Source = "WinRT"
                    }
                }
            }

            if ($nativePages -and $p -lt $nativePages.Count -and $nativePages[$p].Count -gt 5) {
                foreach ($nw in $nativePages[$p]) {
                    $winrtWords += [PSCustomObject]@{
                        Text = $nw.Text
                        X = [int]$nw.X
                        Y = [int]$nw.Y
                        W = 40
                        H = 20
                        Source = "Digital"
                    }
                }
            }

            $tessWords = @()
            if ($hasTess) {
                $prevEAP = $ErrorActionPreference
                $ErrorActionPreference = "Continue"
                & $tessExe $tempPng $tempTsvBase --psm 6 tsv 2>$null
                $ErrorActionPreference = $prevEAP
                if (Test-Path "$tempTsvBase.tsv") {
                    $tsvLines = Get-Content "$tempTsvBase.tsv" -Encoding UTF8
                    foreach ($line in ($tsvLines | Select-Object -Skip 1)) {
                        $cols = $line -split "`t"
                        if ($cols.Count -ge 12 -and $cols[0] -eq '5' -and $cols[11].Trim() -ne '') {
                            $tessWords += [PSCustomObject]@{
                                Text = $cols[11].Trim()
                                X = [int]$cols[6]
                                Y = [int]$cols[7]
                                W = [int]$cols[8]
                                H = [int]$cols[9]
                                Source = "Tess"
                            }
                        }
                    }
                    Remove-Item "$tempTsvBase.tsv" -Force -ErrorAction SilentlyContinue
                }
            }

            $sortedWin = $winrtWords | Sort-Object Y, X
            $winRows = @()
            foreach ($w in $sortedWin) {
                $matched = $null
                foreach ($r in $winRows) {
                    if ([math]::Abs($w.Y - $r.AvgY) -le 16) { $matched = $r; break }
                }
                if ($null -eq $matched) {
                    $winRows += [PSCustomObject]@{ AvgY = $w.Y; Words = @($w) }
                } else {
                    $matched.Words += $w
                    $matched.AvgY = [int](($matched.Words | Measure-Object -Property Y -Average).Average)
                }
            }

            $anneeScolaire = ""
            $college = ""
            $eleveNom = ""
            $numEleve = ""
            $classe = ""
            $maitre = ""
            $niveau = ""
            $regroupement = ""
            $moyPrincipales = ""
            $moyGenerale = ""
            $situation = ""
            $orientation = ""
            $absT1 = ""; $absT2 = ""; $absT3 = ""; $absTotal = ""

            foreach ($r in ($winRows | Sort-Object AvgY)) {
                $lineText = (($r.Words | Sort-Object X) | ForEach-Object { $_.Text }) -join " "
                if ($lineText -match 'Ann.*scolaire\s*:\s*(\d{4}-\d{4})') { $anneeScolaire = $Matches[1] }
                if ($lineText -match 'Coll.*ge\s+(.+)$' -and $r.AvgY -lt 350) { $college = $Matches[1].Trim() }
                if ($lineText -match 'Classe\s*:\s*([A-Za-z0-9]+)') { $classe = $Matches[1].Trim() }
                if ($lineText -match "d'.*l.*ve\s*:\s*([A-Za-z0-9]+)") { $numEleve = $Matches[1].Trim() }
                if ($lineText -match '^((?:MME|M\.|MR|MADAME|MONSIEUR)\s+.+)$' -and $r.AvgY -lt 650) { $maitre = $Matches[1].Trim() }
                if ($lineText -match '^(9CO|10CO|11CO)[,\s]+(.*)$' -and $r.AvgY -lt 680) {
                    $niveau = $Matches[1]
                    $regroupement = $Matches[2].Trim()
                }
                if ($r.AvgY -gt 375 -and $r.AvgY -lt 545) {
                    $leftWords = $r.Words | Where-Object { $_.X -lt 1300 -and $_.Text -notmatch '^(9CO|10CO|11CO|BULLETIN|SCOLAIRE|N|NO|d|Classe)$' }
                    if ($leftWords.Count -gt 0) {
                        $namePart = ($leftWords | Sort-Object X | ForEach-Object { $_.Text }) -join " "
                        if ($eleveNom -eq "") { $eleveNom = $namePart } else { $eleveNom += " " + $namePart }
                    }
                }
                if ($lineText -match 'Moyenne\s+des\s+disciplines\s+principales') {
                    $gradeW = $r.Words | Where-Object { $_.X -gt 1400 } | Select-Object -First 1
                    if ($gradeW) { $moyPrincipales = Normalize-BulletinGrade $gradeW.Text }
                }
                if ($lineText -match 'Moyenne\s+g.*n.*rale' -and $r.AvgY -lt 1700) {
                    $gradeW = $r.Words | Where-Object { $_.X -gt 1400 } | Select-Object -First 1
                    if ($gradeW) { $moyGenerale = Normalize-BulletinGrade $gradeW.Text }
                }
                if ($lineText -match 'SITUATION\s*:\s*(.+)$') { $situation = $Matches[1].Trim() }
                if ($lineText -match "D.*cision\s+d'orientation\s*:?\s*(.*)$") {
                    $cand = $Matches[1].Trim() -replace '^[.\s]+|[.\s]+$', ''
                    if ($cand.Length -gt 1) { $orientation = $cand }
                }
                if ($lineText -match 'Absences') {
                    foreach ($w in $r.Words) {
                        if ($w.Text -match '^\d+$') {
                            if ($w.X -ge 1000 -and $w.X -lt 1160) { $absT1 = $w.Text }
                            elseif ($w.X -ge 1160 -and $w.X -lt 1300) { $absT2 = $w.Text }
                            elseif ($w.X -ge 1300 -and $w.X -lt 1430) { $absT3 = $w.Text }
                            elseif ($w.X -ge 1580) { $absTotal = $w.Text }
                        }
                    }
                }
            }

            if ([string]::IsNullOrWhiteSpace($eleveNom) -and $hasTess) {
                try {
                    $rect = New-Object System.Drawing.Rectangle(200, 390, 1050, 165)
                    $crop = $bmpFull.Clone($rect, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
                    $badCols = New-Object bool[] $crop.Width
                    for ($cx = 0; $cx -lt $crop.Width; $cx++) {
                        $darkCount = 0
                        for ($cy = 0; $cy -lt $crop.Height; $cy++) {
                            $c = $crop.GetPixel($cx, $cy)
                            if ($c.R -lt 80 -and $c.G -lt 80 -and $c.B -lt 80) { $darkCount++ }
                        }
                        if ($darkCount -gt 38) {
                            for ($nx = [math]::Max(0, $cx - 4); $nx -le [math]::Min($crop.Width - 1, $cx + 4); $nx++) {
                                $badCols[$nx] = $true
                            }
                        }
                    }
                    for ($cx = 0; $cx -lt $crop.Width; $cx++) {
                        if ($badCols[$cx]) {
                            for ($cy = 0; $cy -lt $crop.Height; $cy++) {
                                $crop.SetPixel($cx, $cy, [System.Drawing.Color]::White)
                            }
                        }
                    }
                    $cropPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "agy_crop_name_p$p.png")
                    $crop.Save($cropPath, [System.Drawing.Imaging.ImageFormat]::Png)
                    $crop.Dispose()
                    $prevEAP2 = $ErrorActionPreference
                    $ErrorActionPreference = "Continue"
                    $rawName = (& $tessExe $cropPath stdout --psm 6 2>$null) -join " "
                    $ErrorActionPreference = $prevEAP2
                    Remove-Item $cropPath -Force -ErrorAction SilentlyContinue
                    $cleanName = ($rawName -replace '[^A-Za-z\-\s]', ' ' -replace '\s+', ' ').Trim()
                    if ($cleanName.Length -ge 2) { $eleveNom = $cleanName }
                } catch {}
            }

            $bmpFull.Dispose()
            Remove-Item $tempPng -Force -ErrorAction SilentlyContinue

            if ($niveau -eq "" -and $classe -match '^0?9') { $niveau = "9CO" }
            elseif ($niveau -eq "" -and $classe -match '^10') { $niveau = "10CO" }
            elseif ($niveau -eq "" -and $classe -match '^11') { $niveau = "11CO" }

            $subjects = @()
            foreach ($r in ($winRows | Where-Object { $_.AvgY -gt 740 -and $_.AvgY -lt 1470 } | Sort-Object AvgY)) {
                $leftWords = $r.Words | Where-Object { $_.X -lt 950 } | Sort-Object X
                if ($leftWords.Count -eq 0) { continue }
                $minX = $leftWords[0].X
                $subjName = Clean-SubjectName (($leftWords | ForEach-Object { $_.Text }) -join " ")
                if ([string]::IsNullOrWhiteSpace($subjName)) { continue }

                $rowY = $r.AvgY
                $combinedRight = @()
                $combinedRight += ($r.Words | Where-Object { $_.X -ge 980 })
                $combinedRight += ($tessWords | Where-Object { [math]::Abs($_.Y - $rowY) -le 18 -and $_.X -ge 980 })

                $t1 = ""; $t2 = ""; $t3 = ""; $moyP = ""; $moyA = ""
                foreach ($cw in $combinedRight) {
                    $norm = Normalize-BulletinGrade $cw.Text
                    if ($null -ne $norm -and $norm -ne "") {
                        if ($cw.X -ge 1000 -and $cw.X -lt 1150) { if ($t1 -eq "" -or $cw.Source -eq "Tess") { $t1 = $norm } }
                        elseif ($cw.X -ge 1150 -and $cw.X -lt 1300) { if ($t2 -eq "" -or $cw.Source -eq "Tess") { $t2 = $norm } }
                        elseif ($cw.X -ge 1300 -and $cw.X -lt 1450) { if ($t3 -eq "" -or $cw.Source -eq "Tess") { $t3 = $norm } }
                        elseif ($cw.X -ge 1460 -and $cw.X -lt 1610) { if ($moyP -eq "" -or $cw.Source -eq "WinRT") { $moyP = $norm } }
                        elseif ($cw.X -ge 1610 -and $cw.X -lt 1800) { if ($moyA -eq "" -or $cw.Source -eq "WinRT") { $moyA = $norm } }
                    }
                }

                $hasAnyGrade = ($t1 -ne "" -or $t2 -ne "" -or $t3 -ne "" -or $moyP -ne "" -or $moyA -ne "")
                $isDomainHeader = ($subjName -match '^(Langues|Arts|Corps et mouvement|Sciences humaines.*|Math.*et sciences.*|Domaines.*|Disciplines)$')
                if ((-not $isDomainHeader) -and ($minX -ge 242 -or $hasAnyGrade)) {
                    $isPrincipale = ($moyP -ne "" -or $subjName -match '^(Fran.*ais|Allemand|Math.*matiques)$')
                    $moyAnnuelle = if ($moyP -ne "") { $moyP } else { $moyA }
                    $subjects += [PSCustomObject]@{
                        discipline = $subjName
                        principale = [bool]$isPrincipale
                        t1 = $t1
                        t2 = $t2
                        t3 = $t3
                        moyenne = $moyAnnuelle
                    }
                }
            }

            $results += [PSCustomObject]@{
                pageNumber = ($p + 1)
                sourceFile = [System.IO.Path]::GetFileName($pdfPath)
                anneeScolaire = $anneeScolaire
                college = $college
                eleveNom = $eleveNom
                numEleve = $numEleve
                classe = $classe
                niveau = $niveau
                regroupement = $regroupement
                maitre = $maitre
                moyPrincipales = $moyPrincipales
                moyGenerale = $moyGenerale
                situation = $situation
                orientation = $orientation
                absT1 = $absT1
                absT2 = $absT2
                absT3 = $absT3
                absTotal = $absTotal
                disciplines = $subjects
                previewImage = $previewBase64
            }
        }
    } finally {
        $fs.Close()
    }

    return ,$results
}

function Get-DatabasePath([string]$baseDir) {
    $dataDir = Join-Path $baseDir "data"
    if (-not (Test-Path $dataDir)) {
        New-Item -ItemType Directory -Path $dataDir -Force | Out-Null
    }
    return Join-Path $dataDir "classeurs_rouges_db.json"
}

function Load-Database([string]$baseDir) {
    $dbPath = Get-DatabasePath $baseDir
    if (-not (Test-Path $dbPath)) {
        return [PSCustomObject]@{
            version = "1.0"
            updatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            students = @()
        }
    }
    $raw = [System.IO.File]::ReadAllText($dbPath, [System.Text.Encoding]::UTF8)
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return [PSCustomObject]@{
            version = "1.0"
            updatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            students = @()
        }
    }
    $obj = $raw | ConvertFrom-Json
    if ($null -eq $obj.students) {
        $obj | Add-Member -NotePropertyName "students" -NotePropertyValue @() -Force
    } else {
        $obj.students = @($obj.students)
    }
    return $obj
}

function Save-Database([string]$baseDir, $db) {
    $dbPath = Get-DatabasePath $baseDir
    $backupDir = Join-Path (Join-Path $baseDir "data") "backups"
    if (-not (Test-Path $backupDir)) {
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    }
    if (Test-Path $dbPath) {
        $stamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
        Copy-Item $dbPath (Join-Path $backupDir "db_backup_$stamp.json") -Force -ErrorAction SilentlyContinue
        $oldBackups = Get-ChildItem $backupDir -Filter "db_backup_*.json" | Sort-Object LastWriteTime -Descending | Select-Object -Skip 15
        foreach ($ob in $oldBackups) { Remove-Item $ob.FullName -Force -ErrorAction SilentlyContinue }
    }
    $db.updatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $json = $db | ConvertTo-Json -Depth 10
    [System.IO.File]::WriteAllText($dbPath, $json, (New-Object System.Text.UTF8Encoding($false)))
}

function Normalize-Identity([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "" }
    return ($s.Trim().ToUpperInvariant() -replace '\s+', ' ')
}

function Read-FileSmartEncoding([string]$filePath) {
    $bytes = [System.IO.File]::ReadAllBytes($filePath)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    }
    $asciiPreview = [System.Text.Encoding]::ASCII.GetString($bytes, 0, [Math]::Min($bytes.Length, 250))
    if ($asciiPreview -match '(?i)encoding\s*=\s*["''](?:ISO-8859-1|windows-1252|latin1)["'']') {
        return [System.Text.Encoding]::GetEncoding(28591).GetString($bytes)
    }
    $utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
    try {
        return $utf8Strict.GetString($bytes)
    } catch {
        return [System.Text.Encoding]::GetEncoding(28591).GetString($bytes)
    }
}

function Normalize-SchoolYear([string]$yr, [string]$noAnnee = "") {
    if (-not [string]::IsNullOrWhiteSpace($yr)) {
        $y = $yr.Trim()
        if ($y -match '^(\d{4})-(\d{2})$') {
            $start = $Matches[1]
            $century = $start.Substring(0, 2)
            return "$start-$century$($Matches[2])"
        }
        if ($y -match '^\d{4}-\d{4}$') { return $y }
    }
    if (-not [string]::IsNullOrWhiteSpace($noAnnee) -and $noAnnee -match '^\d+$') {
        $n = [int]$noAnnee
        $startYr = 1994 + $n
        $endYr = $startYr + 1
        return "$startYr-$endYr"
    }
    return ""
}

function Convert-ExcelToCsvString([string]$excelPath) {
    $xl = $null
    $wb = $null
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $xl.Workbooks.Open($excelPath, 0, $true)
        $ws = $wb.Worksheets.Item(1)
        $used = $ws.UsedRange
        $rowCount = $used.Rows.Count
        $colCount = $used.Columns.Count
        $vals = $used.Value2
        $sb = New-Object System.Text.StringBuilder
        if ($rowCount -eq 1 -and $colCount -eq 1) {
            [void]$sb.AppendLine([string]$vals)
        } else {
            $maxCols = [Math]::Min($colCount, 20)
            for ($r = 1; $r -le $rowCount; $r++) {
                $rowArr = New-Object string[] $maxCols
                for ($c = 1; $c -le $maxCols; $c++) {
                    $cellVal = $vals[$r, $c]
                    $rowArr[$c - 1] = if ($null -eq $cellVal) { "" } else { ([string]$cellVal).Replace("`r", " ").Replace("`n", " ").Trim() }
                }
                [void]$sb.AppendLine(($rowArr -join ","))
            }
        }
        return $sb.ToString()
    } finally {
        if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
        if ($null -ne $xl) {
            try { $xl.Quit() } catch {}
            try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null } catch {}
        }
    }
}

function Get-ClassAndRegroupementFromCode([string]$code) {
    if ([string]::IsNullOrWhiteSpace($code)) {
        return @{ classe = ""; niveau = ""; regroupement = "" }
    }
    $c = $code.Trim()
    $cls = ""
    if ($c -match '(?i)(0?9\d{2}(?:R[123]|TC|MI|DR)|10\d{2}(?:CT|LC|LS|MI|Ac|AT)|11\d{2}(?:CT|LC|LS|MI|Ac|AT))') {
        $cls = $Matches[1]
    } elseif ($c -match '(?i)(0?9\d{2}|10\d{2}|11\d{2})') {
        $cls = $Matches[1]
    }
    $niv = ""
    if ($cls -match '^0?9') { $niv = "9CO" }
    elseif ($cls -match '^10') { $niv = "10CO" }
    elseif ($cls -match '^11') { $niv = "11CO" }

    $reg = ""
    if ($cls -match 'R1$') { $reg = "regroupement 1" }
    elseif ($cls -match 'R2$') { $reg = "regroupement 2" }
    elseif ($cls -match 'R3$') { $reg = "regroupement 3" }
    elseif ($cls -match 'TC$') { $reg = "Tronc commun" }
    elseif ($cls -match 'DR$') { $reg = "Classe DR" }
    elseif ($cls -match 'CT$') { $reg = "Section CT" }
    elseif ($cls -match 'LC$') { $reg = "Section LC" }
    elseif ($cls -match 'LS$') { $reg = "Section LS" }
    elseif ($cls -match 'MI$') { $reg = "Classe MI" }
    elseif ($cls -match '(?i)Ac$') { $reg = "Accueil" }
    elseif ($cls -match '(?i)AT$') { $reg = "Atelier" }

    return @{ classe = $cls; niveau = $niv; regroupement = $reg }
}

function Resolve-CourseCodeInfo([string]$coursCode, [string]$niveauHint = "", [string]$classeHint = "") {
    if ([string]::IsNullOrWhiteSpace($coursCode)) { return $null }
    $c = $coursCode.Trim()

    # Ignore non-graded courses and support/atelier lines requested to be excluded
    if ($c -match '^(?i)(MT|SH|DT|TMBois|EPI\b|Appui|Passerelle|Rattr\.|SA\+)') {
        return $null
    }

    # Extract alphabetic course prefix (e.g. BIO, DMS, AC, AL, AN, BI, EN, EP, FR, FL, GE, HI, LA, MA, MC, MI, MU, AV, PY, SA, EC, EO, FO, CM)
    $prefix = ""
    if ($c -match '^([A-Za-z]+)') {
        $prefix = $Matches[1].ToUpperInvariant()
    }

    $subjectMap = @{
        "FR"  = "Français"
        "FL"  = "Français lecture"
        "AL"  = "Allemand"
        "MA"  = "Mathématiques"
        "AN"  = "Anglais"
        "BI"  = "Biologie"
        "BIO" = "Biologie"
        "PY"  = "Physique"
        "DMS" = "DMS"
        "HI"  = "Histoire"
        "GE"  = "Géographie"
        "LA"  = "Latin"
        "SA"  = "Sciences appliqués"
        "MC"  = "Informatique"
        "MI"  = "Média-Image"
        "AC"  = "Activités créatrices et manuelles"
        "AV"  = "Arts visuels"
        "MU"  = "Musique"
        "EP"  = "Education physique"
        "EN"  = "Education nutritionnelle"
        "EC"  = "EVACOM / TAF"
        "CM"  = "Connaissance du Milieu"
        "EO"  = "Education aux choix et à l'orientation"
        "FO"  = "Education aux choix et à l'orientation"
    }

    $discName = if ($subjectMap.ContainsKey($prefix)) { $subjectMap[$prefix] } else { $c }
    $clsInfo = Get-ClassAndRegroupementFromCode $c
    $effNiv = if (-not [string]::IsNullOrWhiteSpace($niveauHint)) { $niveauHint } else { $clsInfo.niveau }
    $effCls = if (-not [string]::IsNullOrWhiteSpace($classeHint)) { $classeHint } else { $clsInfo.classe }

    # Determine if discipline is principale
    $isPrincipale = $false
    if ($prefix -in @("FR", "AL", "MA")) {
        $isPrincipale = $true
    } elseif ($effNiv -in @("10CO", "11CO")) {
        if ($effCls -match 'LC$' -and $prefix -in @("AN", "SA")) {
            $isPrincipale = $true
        } elseif ($effCls -match 'LS$' -and $prefix -in @("LA", "DMS", "BI", "BIO", "PY", "AN")) {
            $isPrincipale = $true
        } elseif ($prefix -in @("LA", "SA", "DMS")) {
            $isPrincipale = $true
        }
    }

    return [PSCustomObject]@{
        codeCours  = $c
        prefix     = $prefix
        discipline = $discName
        principale = [bool]$isPrincipale
        classe     = $clsInfo.classe
        niveau     = $clsInfo.niveau
        regroupement = $clsInfo.regroupement
    }
}

function Calculate-CountingAverage([string[]]$grades) {
    $sum = 0.0
    $count = 0
    foreach ($g in $grades) {
        if ([string]::IsNullOrWhiteSpace($g)) { continue }
        $s = $g.Trim()
        # Exclude parenthesized grades (Compte="0") and non-numeric codes
        if ($s.StartsWith("(") -or $s -match '^(?i)(C\.M|Disp|N\.E)') { continue }
        $num = 0.0
        if ([double]::TryParse(($s -replace ',', '.'), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$num)) {
            $sum += $num
            $count++
        }
    }
    if ($count -eq 0) { return "" }
    $avg = [Math]::Round(($sum / $count) + 1e-9, 1, [System.MidpointRounding]::AwayFromZero)
    return $avg.ToString("0.0", [System.Globalization.CultureInfo]::InvariantCulture)
}

function Parse-TeachersCsvContent([string]$rawContent) {
    $courseTeachersList = @{}
    $classSubjectTeachersList = @{}
    $classMaitre = @{}
    $detectedCollege = ""
    $detectedYear = ""

    $lines = $rawContent -split "\r?\n"
    $expectingHeaderMaitre = $false
    $currentBlockHeaderMaitre = ""
    $currentBlockClasses = @()
    $currentBlockCourses = @()

    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $cols = $line -split ','
        $c0 = if ($cols.Count -gt 0) { $cols[0].Trim() } else { "" }
        $c1 = if ($cols.Count -gt 1) { $cols[1].Trim() } else { "" }
        $c2 = if ($cols.Count -gt 2) { $cols[2].Trim() } else { "" }
        $c3 = if ($cols.Count -gt 3) { $cols[3].Trim() } else { "" }
        $c4 = if ($cols.Count -gt 4) { $cols[4].Trim() } else { "" }

        if ($c3 -eq "Maitre" -and $c1 -eq "") {
            # Flush previous block's header maitre to detected classes
            if ($currentBlockHeaderMaitre -ne "" -and $currentBlockClasses.Count -gt 0) {
                $topCls = ($currentBlockClasses | Group-Object | Sort-Object Count -Descending | Select-Object -First 1).Name
                if ($topCls -and -not $classMaitre.ContainsKey($topCls)) {
                    $classMaitre[$topCls] = $currentBlockHeaderMaitre
                }
            }
            $expectingHeaderMaitre = $true
            $currentBlockHeaderMaitre = ""
            $currentBlockClasses = @()
            $currentBlockCourses = @()
            continue
        }

        if ($expectingHeaderMaitre) {
            if ($c3 -ne "" -and $c3 -ne "Champ30") {
                $currentBlockHeaderMaitre = $c3
            }
            $expectingHeaderMaitre = $false
            continue
        }

        if ($c3 -eq "Champ30") { continue }

        if ($c3 -match "(?i)(CYCLE D'ORIENTATION|COLL[EÈ]GE)" -or $c4 -match '^\d{4}-\d{2,4}$') {
            if ($c3 -ne "" -and $detectedCollege -eq "") { $detectedCollege = $c3 }
            if ($c4 -ne "" -and $detectedYear -eq "") { $detectedYear = Normalize-SchoolYear $c4 }
            continue
        }

        if ($c0 -eq "Groupe" -and $c1 -eq "LibelleCours") { continue }

        if ($c1 -ne "" -and $c2 -ne "") {
            # If MT<Classe>, record class maitre
            if ($c1 -match '^(?i)MT(.+)$') {
                $mtClsInfo = Get-ClassAndRegroupementFromCode $c1
                $mtCls = if ($mtClsInfo.classe -ne "") { $mtClsInfo.classe } else { $Matches[1].Trim() }
                if ($mtCls -ne "") {
                    $currentBlockClasses += $mtCls
                    if ($currentBlockHeaderMaitre -ne "") {
                        $classMaitre[$mtCls] = $currentBlockHeaderMaitre
                    } else {
                        if (-not $classMaitre.ContainsKey($mtCls)) {
                            $classMaitre[$mtCls] = $c2
                        } elseif ($classMaitre[$mtCls] -notmatch [regex]::Escape($c2)) {
                            $classMaitre[$mtCls] = $classMaitre[$mtCls] + " / " + $c2
                        }
                    }
                }
                continue
            }

            if (-not $courseTeachersList.ContainsKey($c1)) {
                $courseTeachersList[$c1] = New-Object 'System.Collections.Generic.List[string]'
            }
            if (-not $courseTeachersList[$c1].Contains($c2)) {
                $courseTeachersList[$c1].Add($c2)
            }

            $info = Resolve-CourseCodeInfo $c1
            if ($null -ne $info -and $info.classe -ne "" -and $info.discipline -ne "") {
                $currentBlockClasses += $info.classe
                $key = "$($info.classe)|$(Normalize-Identity $info.discipline)"
                if (-not $classSubjectTeachersList.ContainsKey($key)) {
                    $classSubjectTeachersList[$key] = New-Object 'System.Collections.Generic.List[string]'
                }
                if (-not $classSubjectTeachersList[$key].Contains($c2)) {
                    $classSubjectTeachersList[$key].Add($c2)
                }
            }
        }
    }

    if ($currentBlockHeaderMaitre -ne "" -and $currentBlockClasses.Count -gt 0) {
        $topCls = ($currentBlockClasses | Group-Object | Sort-Object Count -Descending | Select-Object -First 1).Name
        if ($topCls -and -not $classMaitre.ContainsKey($topCls)) {
            $classMaitre[$topCls] = $currentBlockHeaderMaitre
        }
    }

    $courseTeachers = @{}
    foreach ($k in $courseTeachersList.Keys) {
        $courseTeachers[$k] = ($courseTeachersList[$k] -join " / ")
    }
    $classSubjectTeachers = @{}
    foreach ($k in $classSubjectTeachersList.Keys) {
        $classSubjectTeachers[$k] = ($classSubjectTeachersList[$k] -join " / ")
    }

    return [PSCustomObject]@{
        college              = $detectedCollege
        anneeScolaire        = $detectedYear
        courseTeachers       = $courseTeachers
        classSubjectTeachers = $classSubjectTeachers
        classMaitre          = $classMaitre
    }
}

function Parse-NotesOrAbsencesXmlContent([string]$rawXml, [string]$sourceFileName = "Document.xml") {
    $isNotes = ($rawXml -match '<notes\b')
    $isAbsences = ($rawXml -match '<absences\b')
    if (-not $isNotes -and -not $isAbsences) {
        return $null
    }

    $noAnnee = ""
    $sigle = ""
    if ($rawXml -match '<(?:notes|absences)\b[^>]*\bNoAnnee="([^"]*)"') { $noAnnee = $Matches[1] }
    if ($rawXml -match '<(?:notes|absences)\b[^>]*\bsigle="([^"]*)"') { $sigle = $Matches[1] }
    $anneeScolaire = Normalize-SchoolYear "" $noAnnee
    $college = if ($sigle -eq "VOIRE") { "CYCLE D'ORIENTATION DES VOIRETS" } else { $sigle }

    $degreMatches = [regex]::Matches($rawXml, '<Degre\b[^>]*\bValeur="([^"]+)"[^>]*>([\s\S]*?)(?=</Degre>|<Degre\b|\z)')
    $blocksToProcess = @()
    if ($degreMatches.Count -gt 0) {
        foreach ($dm in $degreMatches) {
            $dVal = $dm.Groups[1].Value.Trim()
            $niv = if ($dVal -match '^0?9$') { "9CO" } elseif ($dVal -eq "10") { "10CO" } elseif ($dVal -eq "11") { "11CO" } else { "9CO" }
            $blocksToProcess += @{ niveau = $niv; content = $dm.Groups[2].Value }
        }
    } else {
        $blocksToProcess += @{ niveau = "9CO"; content = $rawXml }
    }

    $records = @()
    $pageIdx = 0

    foreach ($blk in $blocksToProcess) {
        $niv = $blk.niveau
        $eleveMatches = [regex]::Matches($blk.content, '<eleve\b([^>]*)>([\s\S]*?)(?=</eleve>|<eleve\b|</eleves>|</Degre>|</notes>|</absences>|\z)')
        foreach ($em in $eleveMatches) {
            $pageIdx++
            $attrs = $em.Groups[1].Value
            $inner = $em.Groups[2].Value

            $noBds = ""; $nom = ""; $prenom = ""
            if ($attrs -match '\bNoBDS="([^"]*)"') { $noBds = $Matches[1].Trim() }
            if ($attrs -match '\bNom="([^"]*)"') { $nom = $Matches[1].Trim() }
            if ($attrs -match '\bPrenom="([^"]*)"') { $prenom = $Matches[1].Trim() }
            $eleveNom = ("$nom $prenom").Trim()

            if ($isNotes) {
                $noteMatches = [regex]::Matches($inner, '<note\b[^>]*\bCours="([^"]+)"[^>]*>([\s\S]*?)(?=</note>|<note\b|\z)')
                $classCandidates = @()
                $rawNotes = @()

                foreach ($nm in $noteMatches) {
                    $coursCode = $nm.Groups[1].Value.Trim()
                    $noteInner = $nm.Groups[2].Value
                    $clsCheck = Get-ClassAndRegroupementFromCode $coursCode
                    if ($clsCheck.classe -ne "") { $classCandidates += $clsCheck.classe }

                    $p1 = ""; $p2 = ""; $p3 = ""
                    $periodeMatches = [regex]::Matches($noteInner, '<Periode\b([^>]*)/?>')
                    foreach ($pm in $periodeMatches) {
                        $pAttrs = $pm.Groups[1].Value
                        $periodeId = ""; $pVal = ""; $pCompte = "1"
                        if ($pAttrs -match '\bid="([^"]*)"') { $periodeId = $Matches[1].Trim() }
                        if ($pAttrs -match '\bValeur="([^"]*)"') { $pVal = $Matches[1].Trim() }
                        if ($pAttrs -match '\bCompte="([^"]*)"') { $pCompte = $Matches[1].Trim() }

                        if ($periodeId -in @("1", "2", "3") -and $pVal -ne "") {
                            $normVal = $pVal -replace ',', '.'
                            $formattedVal = if ($pCompte -eq "0") { "($normVal)" } else { $normVal }
                            if ($periodeId -eq "1") { $p1 = $formattedVal }
                            elseif ($periodeId -eq "2") { $p2 = $formattedVal }
                            elseif ($periodeId -eq "3") { $p3 = $formattedVal }
                        }
                    }

                    $rawNotes += [PSCustomObject]@{
                        coursCode = $coursCode
                        t1 = $p1
                        t2 = $p2
                        t3 = $p3
                    }
                }

                $studentClasse = ""
                $studentRegroupement = ""
                if ($classCandidates.Count -gt 0) {
                    $studentClasse = ($classCandidates | Group-Object | Sort-Object Count -Descending | Select-Object -First 1).Name
                    $cInfo = Get-ClassAndRegroupementFromCode $studentClasse
                    $studentRegroupement = $cInfo.regroupement
                    if ($cInfo.niveau -ne "") { $niv = $cInfo.niveau }
                }

                $disciplines = @()
                foreach ($rn in $rawNotes) {
                    $info = Resolve-CourseCodeInfo $rn.coursCode $niv $studentClasse
                    if ($null -eq $info) { continue }
                    $moy = Calculate-CountingAverage @($rn.t1, $rn.t2, $rn.t3)
                    $disciplines += [PSCustomObject]@{
                        codeCours  = $rn.coursCode
                        discipline = $info.discipline
                        professeur = ""
                        principale = [bool]$info.principale
                        t1         = $rn.t1
                        t2         = $rn.t2
                        t3         = $rn.t3
                        moyenne    = $moy
                    }
                }

                $princGrades = @($disciplines | Where-Object { $_.principale -and $_.moyenne -ne "" } | ForEach-Object { $_.moyenne })
                $allGrades   = @($disciplines | Where-Object { $_.moyenne -ne "" } | ForEach-Object { $_.moyenne })
                $moyPrincipales = Calculate-CountingAverage $princGrades
                $moyGenerale    = Calculate-CountingAverage $allGrades

                $records += [PSCustomObject]@{
                    recordType     = "notes"
                    pageNumber     = $pageIdx
                    sourceFile     = $sourceFileName
                    anneeScolaire  = $anneeScolaire
                    college        = $college
                    eleveNom       = $eleveNom
                    numEleve       = $noBds
                    classe         = $studentClasse
                    niveau         = $niv
                    regroupement   = $studentRegroupement
                    maitre         = ""
                    moyPrincipales = $moyPrincipales
                    moyGenerale    = $moyGenerale
                    situation      = ""
                    orientation    = ""
                    absT1          = ""
                    absT2          = ""
                    absT3          = ""
                    absTotal       = ""
                    disciplines    = $disciplines
                    previewImage   = ""
                }
            } elseif ($isAbsences) {
                $abs1 = ""; $abs2 = ""; $abs3 = ""
                $periodeMatches = [regex]::Matches($inner, '<Periode\b([^>]*)/?>')
                foreach ($pm in $periodeMatches) {
                    $pAttrs = $pm.Groups[1].Value
                    $periodeId = ""; $pVal = ""
                    if ($pAttrs -match '\bid="([^"]*)"') { $periodeId = $Matches[1].Trim() }
                    if ($pAttrs -match '\bValeur="([^"]*)"') { $pVal = $Matches[1].Trim() }
                    if ($periodeId -eq "1") { $abs1 = $pVal }
                    elseif ($periodeId -eq "2") { $abs2 = $pVal }
                    elseif ($periodeId -eq "3") { $abs3 = $pVal }
                }
                $sumAbs = 0
                $hasAnyAbs = $false
                foreach ($av in @($abs1, $abs2, $abs3)) {
                    if ($av -match '^\d+$') { $sumAbs += [int]$av; $hasAnyAbs = $true }
                }
                $absTot = if ($hasAnyAbs) { [string]$sumAbs } else { "" }

                $records += [PSCustomObject]@{
                    recordType     = "absences"
                    pageNumber     = $pageIdx
                    sourceFile     = $sourceFileName
                    anneeScolaire  = $anneeScolaire
                    college        = $college
                    eleveNom       = $eleveNom
                    numEleve       = $noBds
                    classe         = ""
                    niveau         = $niv
                    regroupement   = ""
                    maitre         = ""
                    moyPrincipales = ""
                    moyGenerale    = ""
                    situation      = ""
                    orientation    = ""
                    absT1          = $abs1
                    absT2          = $abs2
                    absT3          = $abs3
                    absTotal       = $absTot
                    disciplines    = @()
                    previewImage   = ""
                }
            }
        }
    }

    return [PSCustomObject]@{
        docType       = if ($isNotes) { "notes" } else { "absences" }
        anneeScolaire = $anneeScolaire
        college       = $college
        records       = $records
    }
}

function Get-CatalogLookup($dictObj, [string]$yr, [string]$baseKey) {
    if ($null -eq $dictObj -or [string]::IsNullOrWhiteSpace($baseKey)) { return $null }
    if (-not [string]::IsNullOrWhiteSpace($yr)) {
        $yrKey = "$($yr.Trim())|$baseKey"
        if ($dictObj -is [hashtable] -and $dictObj.ContainsKey($yrKey)) { return $dictObj[$yrKey] }
        elseif ($null -ne $dictObj.PSObject.Properties[$yrKey]) { return $dictObj.PSObject.Properties[$yrKey].Value }
    }
    if ($dictObj -is [hashtable] -and $dictObj.ContainsKey($baseKey)) { return $dictObj[$baseKey] }
    elseif ($null -ne $dictObj.PSObject.Properties[$baseKey]) { return $dictObj.PSObject.Properties[$baseKey].Value }
    return $null
}

function Apply-TeacherCatalogToBulletin($b, $catalog) {
    if ($null -eq $catalog) { return $b }
    if ([string]::IsNullOrWhiteSpace($b.college) -and -not [string]::IsNullOrWhiteSpace($catalog.college)) {
        $b.college = $catalog.college
    }
    if (-not [string]::IsNullOrWhiteSpace($catalog.anneeScolaire) -and [string]::IsNullOrWhiteSpace($b.anneeScolaire)) {
        $b.anneeScolaire = $catalog.anneeScolaire
    }
    $yr = if ($b.anneeScolaire) { $b.anneeScolaire.Trim() } else { "" }
    $cls = if ($b.classe) { $b.classe.Trim() } else { "" }
    if ($cls -ne "" -and $null -ne $catalog.classMaitre) {
        $cm = Get-CatalogLookup $catalog.classMaitre $yr $cls
        if (-not [string]::IsNullOrWhiteSpace($cm)) {
            $b.maitre = $cm
        }
    }

    if ($null -ne $b.disciplines) {
        foreach ($d in $b.disciplines) {
            if ($null -eq $d.PSObject.Properties["professeur"]) {
                $d | Add-Member -NotePropertyName "professeur" -NotePropertyValue "" -Force
            }
            $prof = ""
            $cCode = if ($null -ne $d.PSObject.Properties["codeCours"] -and $d.codeCours) { $d.codeCours.Trim() } else { "" }
            if ($cCode -ne "" -and $null -ne $catalog.courseTeachers) {
                $prof = Get-CatalogLookup $catalog.courseTeachers $yr $cCode
            }
            if ([string]::IsNullOrWhiteSpace($prof) -and $cls -ne "" -and -not [string]::IsNullOrWhiteSpace($d.discipline) -and $null -ne $catalog.classSubjectTeachers) {
                $fallbackKey = "$cls|$(Normalize-Identity $d.discipline)"
                $prof = Get-CatalogLookup $catalog.classSubjectTeachers $yr $fallbackKey
            }
            if (-not [string]::IsNullOrWhiteSpace($prof)) {
                $d.professeur = $prof
            }
        }
    }
    return $b
}

function Merge-TeacherCatalogIntoDb($db, $newCatalog) {
    if ($null -eq $newCatalog) { return $db }
    if ($null -eq $db.PSObject.Properties["teacherCatalog"] -or $null -eq $db.teacherCatalog) {
        $db | Add-Member -NotePropertyName "teacherCatalog" -NotePropertyValue ([PSCustomObject]@{
            college              = $newCatalog.college
            anneeScolaire        = $newCatalog.anneeScolaire
            courseTeachers       = @{}
            classSubjectTeachers = @{}
            classMaitre          = @{}
        }) -Force
    }
    $cat = $db.teacherCatalog
    if (-not [string]::IsNullOrWhiteSpace($newCatalog.college)) { $cat.college = $newCatalog.college }
    if (-not [string]::IsNullOrWhiteSpace($newCatalog.anneeScolaire)) { $cat.anneeScolaire = $newCatalog.anneeScolaire }
    $yr = if ($newCatalog.anneeScolaire) { $newCatalog.anneeScolaire.Trim() } else { "" }

    $ctHash = @{}
    if ($cat.courseTeachers -is [hashtable]) { foreach ($k in $cat.courseTeachers.Keys) { $ctHash[$k] = $cat.courseTeachers[$k] } }
    elseif ($null -ne $cat.courseTeachers) { foreach ($p in $cat.courseTeachers.PSObject.Properties) { $ctHash[$p.Name] = $p.Value } }
    foreach ($k in $newCatalog.courseTeachers.Keys) {
        $ctHash[$k] = $newCatalog.courseTeachers[$k]
        if ($yr -ne "") { $ctHash["$yr|$k"] = $newCatalog.courseTeachers[$k] }
    }
    $cat.courseTeachers = [PSCustomObject]$ctHash

    $cstHash = @{}
    if ($cat.classSubjectTeachers -is [hashtable]) { foreach ($k in $cat.classSubjectTeachers.Keys) { $cstHash[$k] = $cat.classSubjectTeachers[$k] } }
    elseif ($null -ne $cat.classSubjectTeachers) { foreach ($p in $cat.classSubjectTeachers.PSObject.Properties) { $cstHash[$p.Name] = $p.Value } }
    foreach ($k in $newCatalog.classSubjectTeachers.Keys) {
        $cstHash[$k] = $newCatalog.classSubjectTeachers[$k]
        if ($yr -ne "") { $cstHash["$yr|$k"] = $newCatalog.classSubjectTeachers[$k] }
    }
    $cat.classSubjectTeachers = [PSCustomObject]$cstHash

    $cmHash = @{}
    if ($cat.classMaitre -is [hashtable]) { foreach ($k in $cat.classMaitre.Keys) { $cmHash[$k] = $cat.classMaitre[$k] } }
    elseif ($null -ne $cat.classMaitre) { foreach ($p in $cat.classMaitre.PSObject.Properties) { $cmHash[$p.Name] = $p.Value } }
    foreach ($k in $newCatalog.classMaitre.Keys) {
        $cmHash[$k] = $newCatalog.classMaitre[$k]
        if ($yr -ne "") { $cmHash["$yr|$k"] = $newCatalog.classMaitre[$k] }
    }
    $cat.classMaitre = [PSCustomObject]$cmHash

    # Also update existing students in DB
    foreach ($st in $db.students) {
        foreach ($yrRec in $st.history) {
            Apply-TeacherCatalogToBulletin $yrRec $cat | Out-Null
        }
    }
    return $db
}

function Import-DocumentsBatch([string[]]$filePaths, [string]$baseDir) {
    $db = Load-Database $baseDir
    $pdfBulletins = New-Object 'System.Collections.Generic.List[object]'
    $notesRecords = New-Object 'System.Collections.Generic.List[object]'
    $absencesRecords = New-Object 'System.Collections.Generic.List[object]'
    $teachersParsedCount = 0

    foreach ($fp in $filePaths) {
        if (-not (Test-Path $fp)) { continue }
        $ext = [System.IO.Path]::GetExtension($fp).ToLowerInvariant()
        $origName = [System.IO.Path]::GetFileName($fp) -replace '^agy_upload_', ''

        if ($ext -eq ".pdf") {
            $extracted = Extract-BulletinsFromPdf $fp
            foreach ($eb in $extracted) {
                $eb.sourceFile = $origName
                if ($null -ne $db.PSObject.Properties["teacherCatalog"] -and $null -ne $db.teacherCatalog) {
                    Apply-TeacherCatalogToBulletin $eb $db.teacherCatalog | Out-Null
                }
                $pdfBulletins.Add($eb)
            }
        } elseif ($ext -eq ".xml") {
            $rawXml = Read-FileSmartEncoding $fp
            $parsedXml = Parse-NotesOrAbsencesXmlContent $rawXml $origName
            if ($null -ne $parsedXml) {
                if ($parsedXml.docType -eq "notes") {
                    foreach ($r in $parsedXml.records) { $notesRecords.Add($r) }
                } elseif ($parsedXml.docType -eq "absences") {
                    foreach ($r in $parsedXml.records) { $absencesRecords.Add($r) }
                }
            }
        } elseif ($ext -in @(".csv", ".txt", ".xls", ".xlsx")) {
            $rawCsv = ""
            if ($ext -in @(".xls", ".xlsx")) {
                $rawCsv = Convert-ExcelToCsvString $fp
            } else {
                $rawCsv = Read-FileSmartEncoding $fp
            }
            if ($rawCsv -match '(?i)(LibelleCours|Maitre)') {
                $parsedCsv = Parse-TeachersCsvContent $rawCsv
                $db = Merge-TeacherCatalogIntoDb $db $parsedCsv
                Save-Database $baseDir $db
                $teachersParsedCount++
            }
        }
    }

    # Index absences by NoBDS and by Name+Niveau for fast O(1) lookup
    $absByNum = @{}
    $absByName = @{}
    $usedAbsIndices = @{}
    for ($ai = 0; $ai -lt $absencesRecords.Count; $ai++) {
        $ar = $absencesRecords[$ai]
        $nNum = Normalize-Identity $ar.numEleve
        $nName = Normalize-Identity $ar.eleveNom
        $nYr = if ($ar.anneeScolaire) { $ar.anneeScolaire } else { "" }
        if ($nNum -ne "") { $absByNum["$nYr|$($ar.niveau)|$nNum"] = $ai }
        if ($nName -ne "") { $absByName["$nYr|$($ar.niveau)|$nName"] = $ai }
    }

    $mergedXmlBulletins = New-Object 'System.Collections.Generic.List[object]'

    foreach ($nr in $notesRecords) {
        $normNum = Normalize-Identity $nr.numEleve
        $normName = Normalize-Identity $nr.eleveNom
        $nYr = if ($nr.anneeScolaire) { $nr.anneeScolaire } else { "" }
        $matchedAi = $null
        if ($normNum -ne "" -and $absByNum.ContainsKey("$nYr|$($nr.niveau)|$normNum")) {
            $matchedAi = $absByNum["$nYr|$($nr.niveau)|$normNum"]
        } elseif ($normName -ne "" -and $absByName.ContainsKey("$nYr|$($nr.niveau)|$normName")) {
            $matchedAi = $absByName["$nYr|$($nr.niveau)|$normName"]
        }
        if ($null -ne $matchedAi -and -not $usedAbsIndices.ContainsKey($matchedAi)) {
            $ar = $absencesRecords[$matchedAi]
            $nr.absT1 = $ar.absT1
            $nr.absT2 = $ar.absT2
            $nr.absT3 = $ar.absT3
            $nr.absTotal = $ar.absTotal
            $nr.sourceFile = "$($nr.sourceFile) + $($ar.sourceFile)"
            $usedAbsIndices[$matchedAi] = $true
        }
        if ($null -ne $db.PSObject.Properties["teacherCatalog"] -and $null -ne $db.teacherCatalog) {
            if ([string]::IsNullOrWhiteSpace($nr.anneeScolaire) -and -not [string]::IsNullOrWhiteSpace($db.teacherCatalog.anneeScolaire)) {
                $nr.anneeScolaire = $db.teacherCatalog.anneeScolaire
            }
            Apply-TeacherCatalogToBulletin $nr $db.teacherCatalog | Out-Null
        }
        $mergedXmlBulletins.Add($nr)
    }

    # Any absences records that were not paired with a notes file in this batch:
    for ($ai = 0; $ai -lt $absencesRecords.Count; $ai++) {
        if ($usedAbsIndices.ContainsKey($ai)) { continue }
        $ar = $absencesRecords[$ai]
        # Pre-fill from existing student in DB if present so review shows their class and grades
        $normNum = Normalize-Identity $ar.numEleve
        $normName = Normalize-Identity $ar.eleveNom
        foreach ($st in $db.students) {
            if (($normNum -ne "" -and (Normalize-Identity $st.numEleve) -eq $normNum) -or
                ($normName -ne "" -and (Normalize-Identity $st.eleveNom) -eq $normName)) {
                foreach ($yr in $st.history) {
                    if ($yr.niveau -eq $ar.niveau) {
                        $ar.classe = $yr.classe
                        $ar.regroupement = $yr.regroupement
                        $ar.maitre = $yr.maitre
                        $ar.college = $yr.college
                        if (-not [string]::IsNullOrWhiteSpace($yr.anneeScolaire)) { $ar.anneeScolaire = $yr.anneeScolaire }
                        $ar.moyPrincipales = $yr.moyPrincipales
                        $ar.moyGenerale = $yr.moyGenerale
                        $ar.situation = $yr.situation
                        $ar.orientation = $yr.orientation
                        $ar.disciplines = $yr.disciplines
                        break
                    }
                }
                break
            }
        }
        $mergedXmlBulletins.Add($ar)
    }

    $autoSavedCount = 0
    $xmlArr = $mergedXmlBulletins.ToArray()
    if ($xmlArr.Count -gt 0) {
        $db = Upsert-BulletinRecordsBatch $db $xmlArr
        Save-Database $baseDir $db
        $autoSavedCount = $xmlArr.Count
    }

    $pdfArr = $pdfBulletins.ToArray()
    for ($i = 0; $i -lt $pdfArr.Count; $i++) {
        $pdfArr[$i].pageNumber = ($i + 1)
    }

    return [PSCustomObject]@{
        bulletins           = $pdfArr
        autoSavedCount      = $autoSavedCount
        teachersParsedCount = $teachersParsedCount
        db                  = $db
    }
}

function Upsert-BulletinRecordsBatch($db, $bulletins) {
    $studentList = New-Object 'System.Collections.Generic.List[object]'
    $byNum = @{}
    $byName = @{}
    foreach ($st in $db.students) {
        $studentList.Add($st)
        $nNum = Normalize-Identity $st.numEleve
        $nName = Normalize-Identity $st.eleveNom
        if ($nNum -ne "") { $byNum[$nNum] = $st }
        if ($nName -ne "") { $byName[$nName] = $st }
    }

    foreach ($b in $bulletins) {
        $normNum = Normalize-Identity $b.numEleve
        $normName = Normalize-Identity $b.eleveNom

        if ($null -ne $db.PSObject.Properties["teacherCatalog"] -and $null -ne $db.teacherCatalog) {
            Apply-TeacherCatalogToBulletin $b $db.teacherCatalog | Out-Null
        }

        $matchedStudent = $null
        if ($normNum -ne "" -and $byNum.ContainsKey($normNum)) {
            $matchedStudent = $byNum[$normNum]
        } elseif ($normName -ne "" -and $byName.ContainsKey($normName)) {
            $matchedStudent = $byName[$normName]
        }

        if ($null -eq $matchedStudent) {
            $newId = [Guid]::NewGuid().ToString("N").Substring(0, 10)
            $matchedStudent = [PSCustomObject]@{
                id = $newId
                eleveNom = if ($b.eleveNom) { $b.eleveNom.Trim() } else { "" }
                numEleve = if ($b.numEleve) { $b.numEleve.Trim() } else { "" }
                classeActuelle = if ($b.classe) { $b.classe.Trim() } else { "" }
                niveauActuel = if ($b.niveau) { $b.niveau.Trim() } else { "" }
                anneeActuelle = if ($b.anneeScolaire) { $b.anneeScolaire.Trim() } else { "" }
                commentaires = ""
                history = @()
            }
            $studentList.Add($matchedStudent)
            if ($normNum -ne "") { $byNum[$normNum] = $matchedStudent }
            if ($normName -ne "") { $byName[$normName] = $matchedStudent }
        } else {
            if (-not [string]::IsNullOrWhiteSpace($b.eleveNom)) { $matchedStudent.eleveNom = $b.eleveNom.Trim() }
            if (-not [string]::IsNullOrWhiteSpace($b.numEleve)) {
                $matchedStudent.numEleve = $b.numEleve.Trim()
                if ($normNum -ne "") { $byNum[$normNum] = $matchedStudent }
            }
        }

        $history = @($matchedStudent.history)
        $matchedYear = $null
        foreach ($yr in $history) {
            if ($yr.anneeScolaire -eq $b.anneeScolaire -and $yr.niveau -eq $b.niveau) {
                $matchedYear = $yr
                break
            }
        }

        if ($null -eq $matchedYear) {
            $normalizedDiscs = New-Object 'System.Collections.Generic.List[object]'
            foreach ($newD in $b.disciplines) {
                if ([string]::IsNullOrWhiteSpace($newD.discipline)) { continue }
                $profVal = if ($null -ne $newD.PSObject.Properties["professeur"] -and $newD.professeur) { $newD.professeur } else { "" }
                $codeVal = if ($null -ne $newD.PSObject.Properties["codeCours"] -and $newD.codeCours) { $newD.codeCours } else { "" }
                $normalizedDiscs.Add([PSCustomObject]@{
                    codeCours  = $codeVal
                    discipline = $newD.discipline
                    professeur = $profVal
                    principale = [bool]$newD.principale
                    t1         = $newD.t1
                    t2         = $newD.t2
                    t3         = $newD.t3
                    moyenne    = $newD.moyenne
                })
            }
            $matchedYear = [PSCustomObject]@{
                anneeScolaire = $b.anneeScolaire
                niveau = $b.niveau
                classe = $b.classe
                regroupement = $b.regroupement
                maitre = $b.maitre
                college = $b.college
                moyPrincipales = $b.moyPrincipales
                moyGenerale = $b.moyGenerale
                situation = $b.situation
                orientation = $b.orientation
                absT1 = $b.absT1
                absT2 = $b.absT2
                absT3 = $b.absT3
                absTotal = $b.absTotal
                disciplines = $normalizedDiscs.ToArray()
                updatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            }
            $history += $matchedYear
        } else {
            if (-not [string]::IsNullOrWhiteSpace($b.classe)) { $matchedYear.classe = $b.classe }
            if (-not [string]::IsNullOrWhiteSpace($b.regroupement)) { $matchedYear.regroupement = $b.regroupement }
            if (-not [string]::IsNullOrWhiteSpace($b.maitre)) { $matchedYear.maitre = $b.maitre }
            if (-not [string]::IsNullOrWhiteSpace($b.college)) { $matchedYear.college = $b.college }
            if (-not [string]::IsNullOrWhiteSpace($b.moyPrincipales)) { $matchedYear.moyPrincipales = $b.moyPrincipales }
            if (-not [string]::IsNullOrWhiteSpace($b.moyGenerale)) { $matchedYear.moyGenerale = $b.moyGenerale }
            if (-not [string]::IsNullOrWhiteSpace($b.situation)) { $matchedYear.situation = $b.situation }
            if (-not [string]::IsNullOrWhiteSpace($b.orientation)) { $matchedYear.orientation = $b.orientation }
            if (-not [string]::IsNullOrWhiteSpace($b.absT1)) { $matchedYear.absT1 = $b.absT1 }
            if (-not [string]::IsNullOrWhiteSpace($b.absT2)) { $matchedYear.absT2 = $b.absT2 }
            if (-not [string]::IsNullOrWhiteSpace($b.absT3)) { $matchedYear.absT3 = $b.absT3 }
            if (-not [string]::IsNullOrWhiteSpace($b.absTotal)) { $matchedYear.absTotal = $b.absTotal }
            $matchedYear.updatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

            $existingDiscs = New-Object 'System.Collections.Generic.List[object]'
            foreach ($ed in $matchedYear.disciplines) { $existingDiscs.Add($ed) }
            foreach ($newD in $b.disciplines) {
                if ([string]::IsNullOrWhiteSpace($newD.discipline)) { continue }
                $profVal = if ($null -ne $newD.PSObject.Properties["professeur"] -and $newD.professeur) { $newD.professeur } else { "" }
                $codeVal = if ($null -ne $newD.PSObject.Properties["codeCours"] -and $newD.codeCours) { $newD.codeCours } else { "" }
                $foundD = $null
                foreach ($ed in $existingDiscs) {
                    if ((Normalize-Identity $ed.discipline) -eq (Normalize-Identity $newD.discipline)) {
                        $foundD = $ed
                        break
                    }
                }
                if ($null -eq $foundD) {
                    $existingDiscs.Add([PSCustomObject]@{
                        codeCours  = $codeVal
                        discipline = $newD.discipline
                        professeur = $profVal
                        principale = [bool]$newD.principale
                        t1 = $newD.t1
                        t2 = $newD.t2
                        t3 = $newD.t3
                        moyenne = $newD.moyenne
                    })
                } else {
                    $foundD.principale = [bool]$newD.principale
                    if ($null -eq $foundD.PSObject.Properties["professeur"]) {
                        $foundD | Add-Member -NotePropertyName "professeur" -NotePropertyValue "" -Force
                    }
                    if ($null -eq $foundD.PSObject.Properties["codeCours"]) {
                        $foundD | Add-Member -NotePropertyName "codeCours" -NotePropertyValue "" -Force
                    }
                    if (-not [string]::IsNullOrWhiteSpace($profVal)) { $foundD.professeur = $profVal }
                    if (-not [string]::IsNullOrWhiteSpace($codeVal)) { $foundD.codeCours = $codeVal }
                    if (-not [string]::IsNullOrWhiteSpace($newD.t1)) { $foundD.t1 = $newD.t1 }
                    if (-not [string]::IsNullOrWhiteSpace($newD.t2)) { $foundD.t2 = $newD.t2 }
                    if (-not [string]::IsNullOrWhiteSpace($newD.t3)) { $foundD.t3 = $newD.t3 }
                    if (-not [string]::IsNullOrWhiteSpace($newD.moyenne)) { $foundD.moyenne = $newD.moyenne }
                }
            }
            $matchedYear.disciplines = $existingDiscs.ToArray()
        }

        if ($null -ne $db.PSObject.Properties["teacherCatalog"] -and $null -ne $db.teacherCatalog) {
            Apply-TeacherCatalogToBulletin $matchedYear $db.teacherCatalog | Out-Null
        }

        $matchedStudent.history = @($history | Sort-Object anneeScolaire, niveau)
        $latestYear = $matchedStudent.history[-1]
        $matchedStudent.classeActuelle = $latestYear.classe
        $matchedStudent.niveauActuel = $latestYear.niveau
        $matchedStudent.anneeActuelle = $latestYear.anneeScolaire
    }

    $db.students = @($studentList.ToArray() | Sort-Object eleveNom)
    return $db
}

function Upsert-BulletinRecord($db, $b) {
    $students = @($db.students)
    $normNum = Normalize-Identity $b.numEleve
    $normName = Normalize-Identity $b.eleveNom

    if ($null -ne $db.PSObject.Properties["teacherCatalog"] -and $null -ne $db.teacherCatalog) {
        Apply-TeacherCatalogToBulletin $b $db.teacherCatalog | Out-Null
    }

    $matchedStudent = $null
    foreach ($st in $students) {
        if ($normNum -ne "" -and (Normalize-Identity $st.numEleve) -eq $normNum) {
            $matchedStudent = $st
            break
        }
    }
    if ($null -eq $matchedStudent -and $normName -ne "") {
        foreach ($st in $students) {
            if ((Normalize-Identity $st.eleveNom) -eq $normName) {
                $matchedStudent = $st
                break
            }
        }
    }

    if ($null -eq $matchedStudent) {
        $newId = [Guid]::NewGuid().ToString("N").Substring(0, 10)
        $matchedStudent = [PSCustomObject]@{
            id = $newId
            eleveNom = if ($b.eleveNom) { $b.eleveNom.Trim() } else { "" }
            numEleve = if ($b.numEleve) { $b.numEleve.Trim() } else { "" }
            classeActuelle = if ($b.classe) { $b.classe.Trim() } else { "" }
            niveauActuel = if ($b.niveau) { $b.niveau.Trim() } else { "" }
            anneeActuelle = if ($b.anneeScolaire) { $b.anneeScolaire.Trim() } else { "" }
            commentaires = ""
            history = @()
        }
        $students += $matchedStudent
    } else {
        if (-not [string]::IsNullOrWhiteSpace($b.eleveNom)) { $matchedStudent.eleveNom = $b.eleveNom.Trim() }
        if (-not [string]::IsNullOrWhiteSpace($b.numEleve)) { $matchedStudent.numEleve = $b.numEleve.Trim() }
    }

    $history = @($matchedStudent.history)
    $matchedYear = $null
    foreach ($yr in $history) {
        if ($yr.anneeScolaire -eq $b.anneeScolaire -and $yr.niveau -eq $b.niveau) {
            $matchedYear = $yr
            break
        }
    }

    if ($null -eq $matchedYear) {
        $normalizedDiscs = @()
        foreach ($newD in $b.disciplines) {
            if ([string]::IsNullOrWhiteSpace($newD.discipline)) { continue }
            $profVal = if ($null -ne $newD.PSObject.Properties["professeur"] -and $newD.professeur) { $newD.professeur } else { "" }
            $codeVal = if ($null -ne $newD.PSObject.Properties["codeCours"] -and $newD.codeCours) { $newD.codeCours } else { "" }
            $normalizedDiscs += [PSCustomObject]@{
                codeCours  = $codeVal
                discipline = $newD.discipline
                professeur = $profVal
                principale = [bool]$newD.principale
                t1         = $newD.t1
                t2         = $newD.t2
                t3         = $newD.t3
                moyenne    = $newD.moyenne
            }
        }
        $matchedYear = [PSCustomObject]@{
            anneeScolaire = $b.anneeScolaire
            niveau = $b.niveau
            classe = $b.classe
            regroupement = $b.regroupement
            maitre = $b.maitre
            college = $b.college
            moyPrincipales = $b.moyPrincipales
            moyGenerale = $b.moyGenerale
            situation = $b.situation
            orientation = $b.orientation
            absT1 = $b.absT1
            absT2 = $b.absT2
            absT3 = $b.absT3
            absTotal = $b.absTotal
            disciplines = $normalizedDiscs
            updatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        }
        $history += $matchedYear
    } else {
        if (-not [string]::IsNullOrWhiteSpace($b.classe)) { $matchedYear.classe = $b.classe }
        if (-not [string]::IsNullOrWhiteSpace($b.regroupement)) { $matchedYear.regroupement = $b.regroupement }
        if (-not [string]::IsNullOrWhiteSpace($b.maitre)) { $matchedYear.maitre = $b.maitre }
        if (-not [string]::IsNullOrWhiteSpace($b.college)) { $matchedYear.college = $b.college }
        if (-not [string]::IsNullOrWhiteSpace($b.moyPrincipales)) { $matchedYear.moyPrincipales = $b.moyPrincipales }
        if (-not [string]::IsNullOrWhiteSpace($b.moyGenerale)) { $matchedYear.moyGenerale = $b.moyGenerale }
        if (-not [string]::IsNullOrWhiteSpace($b.situation)) { $matchedYear.situation = $b.situation }
        if (-not [string]::IsNullOrWhiteSpace($b.orientation)) { $matchedYear.orientation = $b.orientation }
        if (-not [string]::IsNullOrWhiteSpace($b.absT1)) { $matchedYear.absT1 = $b.absT1 }
        if (-not [string]::IsNullOrWhiteSpace($b.absT2)) { $matchedYear.absT2 = $b.absT2 }
        if (-not [string]::IsNullOrWhiteSpace($b.absT3)) { $matchedYear.absT3 = $b.absT3 }
        if (-not [string]::IsNullOrWhiteSpace($b.absTotal)) { $matchedYear.absTotal = $b.absTotal }
        $matchedYear.updatedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

        $existingDiscs = @($matchedYear.disciplines)
        foreach ($newD in $b.disciplines) {
            if ([string]::IsNullOrWhiteSpace($newD.discipline)) { continue }
            $profVal = if ($null -ne $newD.PSObject.Properties["professeur"] -and $newD.professeur) { $newD.professeur } else { "" }
            $codeVal = if ($null -ne $newD.PSObject.Properties["codeCours"] -and $newD.codeCours) { $newD.codeCours } else { "" }
            $foundD = $null
            foreach ($ed in $existingDiscs) {
                if ((Normalize-Identity $ed.discipline) -eq (Normalize-Identity $newD.discipline)) {
                    $foundD = $ed
                    break
                }
            }
            if ($null -eq $foundD) {
                $existingDiscs += [PSCustomObject]@{
                    codeCours  = $codeVal
                    discipline = $newD.discipline
                    professeur = $profVal
                    principale = [bool]$newD.principale
                    t1 = $newD.t1
                    t2 = $newD.t2
                    t3 = $newD.t3
                    moyenne = $newD.moyenne
                }
            } else {
                $foundD.principale = [bool]$newD.principale
                if ($null -eq $foundD.PSObject.Properties["professeur"]) {
                    $foundD | Add-Member -NotePropertyName "professeur" -NotePropertyValue "" -Force
                }
                if ($null -eq $foundD.PSObject.Properties["codeCours"]) {
                    $foundD | Add-Member -NotePropertyName "codeCours" -NotePropertyValue "" -Force
                }
                if (-not [string]::IsNullOrWhiteSpace($profVal)) { $foundD.professeur = $profVal }
                if (-not [string]::IsNullOrWhiteSpace($codeVal)) { $foundD.codeCours = $codeVal }
                if (-not [string]::IsNullOrWhiteSpace($newD.t1)) { $foundD.t1 = $newD.t1 }
                if (-not [string]::IsNullOrWhiteSpace($newD.t2)) { $foundD.t2 = $newD.t2 }
                if (-not [string]::IsNullOrWhiteSpace($newD.t3)) { $foundD.t3 = $newD.t3 }
                if (-not [string]::IsNullOrWhiteSpace($newD.moyenne)) { $foundD.moyenne = $newD.moyenne }
            }
        }
        $matchedYear.disciplines = $existingDiscs
    }

    if ($null -ne $db.PSObject.Properties["teacherCatalog"] -and $null -ne $db.teacherCatalog) {
        Apply-TeacherCatalogToBulletin $matchedYear $db.teacherCatalog | Out-Null
    }

    $matchedStudent.history = @($history | Sort-Object anneeScolaire, niveau)
    $latestYear = $matchedStudent.history[-1]
    $matchedStudent.classeActuelle = $latestYear.classe
    $matchedStudent.niveauActuel = $latestYear.niveau
    $matchedStudent.anneeActuelle = $latestYear.anneeScolaire

    $db.students = @($students | Sort-Object eleveNom)
    return $db
}

function New-ExcelCell([string]$val, [int]$styleIdx = 0, [bool]$isNum = $false) {
    $c = New-Object ClasseursRouges.ExcelCell
    $c.Value = if ($null -eq $val) { "" } else { [string]$val }
    $c.StyleIndex = $styleIdx
    $c.IsNumeric = $isNum
    return $c
}

function New-GradeCell([string]$val, [bool]$isBoldMoyenne = $false) {
    $c = New-Object ClasseursRouges.ExcelCell
    if ([string]::IsNullOrWhiteSpace($val)) {
        $c.Value = ""
        $c.StyleIndex = if ($isBoldMoyenne) { 6 } else { 3 }
        $c.IsNumeric = $false
        return $c
    }
    $v = $val.Trim()
    if ($v.StartsWith("(") -or $v -match '^(?i)(C\.M|Disp|N\.E)') {
        $c.Value = $v
        $c.StyleIndex = 5
        $c.IsNumeric = $false
        return $c
    }
    $num = 0.0
    if ([double]::TryParse(($v -replace ',', '.'), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$num)) {
        $c.Value = $num.ToString("0.0", [System.Globalization.CultureInfo]::InvariantCulture)
        $c.IsNumeric = $true
        if ($num -lt 4.0) {
            $c.StyleIndex = if ($isBoldMoyenne) { 7 } else { 4 }
        } else {
            $c.StyleIndex = if ($isBoldMoyenne) { 6 } else { 3 }
        }
        return $c
    }
    $c.Value = $v
    $c.StyleIndex = 0
    $c.IsNumeric = $false
    return $c
}

function Export-CohortExcelReport($db, [string]$targetClasse, [string]$outputPath, [string]$targetAnnee = "TOUTES") {
    $allStudents = @($db.students)
    $cohort = @()
    $filterByYr = (-not [string]::IsNullOrWhiteSpace($targetAnnee) -and $targetAnnee -ne "TOUTES")
    $filterByCls = (-not [string]::IsNullOrWhiteSpace($targetClasse) -and $targetClasse -ne "TOUTES")

    if (-not $filterByYr -and -not $filterByCls) {
        $cohort = $allStudents
    } else {
        foreach ($st in $allStudents) {
            $matched = $false
            foreach ($h in $st.history) {
                $yrOk = (-not $filterByYr) -or ($h.anneeScolaire -eq $targetAnnee)
                $clsOk = (-not $filterByCls) -or ($h.classe -eq $targetClasse)
                if ($yrOk -and $clsOk) {
                    $matched = $true
                    break
                }
            }
            if ($matched) { $cohort += $st }
        }
    }

    $cohort = @($cohort | Sort-Object eleveNom)
    $sheets = New-Object 'System.Collections.Generic.List[ClasseursRouges.ExcelSheet]'

    $shSynth = New-Object ClasseursRouges.ExcelSheet
    $shSynth.Name = "Synthese Cohorte"
    $shSynth.ColumnWidths.AddRange([double[]]@(
        24.0, 12.0, 12.0, 10.0,
        11.0, 9.0, 9.0, 7.0, 7.0, 7.0, 8.0, 18.0,
        11.0, 9.0, 9.0, 7.0, 7.0, 7.0, 8.0, 18.0,
        11.0, 9.0, 9.0, 7.0, 7.0, 7.0, 8.0, 18.0
    ))

    $rTitle = New-Object ClasseursRouges.ExcelRow
    $rTitle.Height = 28.0
    $yrTag = if ($filterByYr) { " - ANNEE $targetAnnee" } else { "" }
    $titleLabel = if (-not $filterByCls) {
        "CLASSEURS ROUGES - SYNTHESE PLURIANNUELLE (9CO / 10CO / 11CO)$yrTag - TOUTES LES CLASSES"
    } else {
        "CLASSEURS ROUGES - SYNTHESE DE LA COHORTE : CLASSE $targetClasse$yrTag (9CO / 10CO / 11CO)"
    }
    $rTitle.Cells.Add((New-ExcelCell $titleLabel 8))
    $shSynth.Rows.Add($rTitle)

    $rGrp = New-Object ClasseursRouges.ExcelRow
    $rGrp.Height = 24.0
    $rGrp.Cells.Add((New-ExcelCell "IDENTITE DE L'ELEVE" 1))
    $rGrp.Cells.Add((New-ExcelCell "" 1))
    $rGrp.Cells.Add((New-ExcelCell "" 1))
    $rGrp.Cells.Add((New-ExcelCell "" 1))
    $shSynth.MergedCells.Add("A2:D2")

    $rGrp.Cells.Add((New-ExcelCell "9e ANNEE (9CO)" 1))
    for ($i = 0; $i -lt 7; $i++) { $rGrp.Cells.Add((New-ExcelCell "" 1)) }
    $shSynth.MergedCells.Add("E2:L2")

    $rGrp.Cells.Add((New-ExcelCell "10e ANNEE (10CO)" 1))
    for ($i = 0; $i -lt 7; $i++) { $rGrp.Cells.Add((New-ExcelCell "" 1)) }
    $shSynth.MergedCells.Add("M2:T2")

    $rGrp.Cells.Add((New-ExcelCell "11e ANNEE (11CO)" 1))
    for ($i = 0; $i -lt 7; $i++) { $rGrp.Cells.Add((New-ExcelCell "" 1)) }
    $shSynth.MergedCells.Add("U2:AB2")
    $shSynth.Rows.Add($rGrp)

    $rSub = New-Object ClasseursRouges.ExcelRow
    $rSub.Height = 24.0
    foreach ($h in @("Nom et Prenom", "N d'eleve", "Classe act.", "Redoubl.")) {
        $rSub.Cells.Add((New-ExcelCell $h 2))
    }
    for ($yrIdx = 0; $yrIdx -lt 3; $yrIdx++) {
        foreach ($h in @("Classe", "Moy. Princ.", "Moy. Gen.", "Abs T1", "Abs T2", "Abs T3", "Abs Tot.", "Situation / Orientation")) {
            $rSub.Cells.Add((New-ExcelCell $h 2))
        }
    }
    $shSynth.Rows.Add($rSub)

    foreach ($st in $cohort) {
        $row = New-Object ClasseursRouges.ExcelRow
        $row.Height = 21.0
        $hist = @($st.history)
        $hasRedoublement = if (($hist | Group-Object niveau | Where-Object { $_.Count -gt 1 }).Count -gt 0) { "Oui" } else { "Non" }

        $row.Cells.Add((New-ExcelCell $st.eleveNom 9))
        $row.Cells.Add((New-ExcelCell $st.numEleve 0))
        $row.Cells.Add((New-ExcelCell $st.classeActuelle 0))
        $row.Cells.Add((New-ExcelCell $hasRedoublement 0))

        foreach ($niv in @("9CO", "10CO", "11CO")) {
            $nivRecords = @($hist | Where-Object { $_.niveau -eq $niv } | Sort-Object anneeScolaire)
            if ($nivRecords.Count -eq 0) {
                for ($k = 0; $k -lt 8; $k++) { $row.Cells.Add((New-ExcelCell "" 0)) }
            } else {
                $rec = $nivRecords[-1]
                $clsLabel = if ($nivRecords.Count -gt 1) {
                    ($nivRecords | ForEach-Object { "$($_.classe) ($($_.anneeScolaire))" }) -join " / "
                } else {
                    "$($rec.classe) ($($rec.anneeScolaire))"
                }
                $sitLabel = $rec.situation
                if (-not [string]::IsNullOrWhiteSpace($rec.orientation)) {
                    $sitLabel = if ($sitLabel) { "$sitLabel -> $($rec.orientation)" } else { $rec.orientation }
                }
                $row.Cells.Add((New-ExcelCell $clsLabel 0))
                $row.Cells.Add((New-GradeCell $rec.moyPrincipales $true))
                $row.Cells.Add((New-GradeCell $rec.moyGenerale $true))
                $row.Cells.Add((New-ExcelCell $rec.absT1 3 ($rec.absT1 -match '^\d+$')))
                $row.Cells.Add((New-ExcelCell $rec.absT2 3 ($rec.absT2 -match '^\d+$')))
                $row.Cells.Add((New-ExcelCell $rec.absT3 3 ($rec.absT3 -match '^\d+$')))
                $row.Cells.Add((New-ExcelCell $rec.absTotal 6 ($rec.absTotal -match '^\d+$')))
                $row.Cells.Add((New-ExcelCell $sitLabel 0))
            }
        }
        $shSynth.Rows.Add($row)
    }
    $sheets.Add($shSynth)

    foreach ($niv in @("9CO", "10CO", "11CO")) {
        $shNiv = New-Object ClasseursRouges.ExcelSheet
        $shNiv.Name = "Detail $niv"

        $discList = New-Object 'System.Collections.Generic.List[string]'
        foreach ($st in $cohort) {
            foreach ($h in $st.history) {
                if ($h.niveau -eq $niv) {
                    foreach ($d in $h.disciplines) {
                        if (-not [string]::IsNullOrWhiteSpace($d.discipline)) {
                            $exists = $false
                            foreach ($ex in $discList) {
                                if ((Normalize-Identity $ex) -eq (Normalize-Identity $d.discipline)) { $exists = $true; break }
                            }
                            if (-not $exists) { $discList.Add($d.discipline) }
                        }
                    }
                }
            }
        }

        $shNiv.ColumnWidths.AddRange([double[]]@(24.0, 12.0, 11.0, 18.0, 9.0, 9.0, 8.0, 8.0, 8.0, 8.0, 18.0))
        foreach ($dName in $discList) {
            $shNiv.ColumnWidths.AddRange([double[]]@(6.5, 6.5, 6.5, 7.5))
        }

        $rNivTitle = New-Object ClasseursRouges.ExcelRow
        $rNivTitle.Height = 26.0
        $rNivTitle.Cells.Add((New-ExcelCell "DETAIL DES NOTES TRIMESTRIELLES ET MOYENNES - $niv" 8))
        $shNiv.Rows.Add($rNivTitle)

        $rNivGrp = New-Object ClasseursRouges.ExcelRow
        $rNivGrp.Height = 24.0
        $rNivGrp.Cells.Add((New-ExcelCell "INFORMATIONS GENERALES & ABSENCES ($niv)" 1))
        for ($i = 1; $i -lt 11; $i++) { $rNivGrp.Cells.Add((New-ExcelCell "" 1)) }
        $shNiv.MergedCells.Add("A2:K2")

        $colCursor = 11
        foreach ($dName in $discList) {
            $startCol = [ClasseursRouges.XlsxWriter]::GetColName($colCursor)
            $endCol = [ClasseursRouges.XlsxWriter]::GetColName($colCursor + 3)
            $rNivGrp.Cells.Add((New-ExcelCell $dName 1))
            $rNivGrp.Cells.Add((New-ExcelCell "" 1))
            $rNivGrp.Cells.Add((New-ExcelCell "" 1))
            $rNivGrp.Cells.Add((New-ExcelCell "" 1))
            $shNiv.MergedCells.Add("${startCol}2:${endCol}2")
            $colCursor += 4
        }
        $shNiv.Rows.Add($rNivGrp)

        $rNivSub = New-Object ClasseursRouges.ExcelRow
        $rNivSub.Height = 22.0
        foreach ($h in @("Nom et Prenom", "Annee scol.", "Classe", "Maitre de classe", "Moy. Princ.", "Moy. Gen.", "Abs T1", "Abs T2", "Abs T3", "Abs Tot.", "Situation")) {
            $rNivSub.Cells.Add((New-ExcelCell $h 2))
        }
        foreach ($dName in $discList) {
            foreach ($th in @("T1", "T2", "T3", "Moy")) {
                $rNivSub.Cells.Add((New-ExcelCell $th 2))
            }
        }
        $shNiv.Rows.Add($rNivSub)

        foreach ($st in $cohort) {
            $nivRecords = @($st.history | Where-Object { $_.niveau -eq $niv } | Sort-Object anneeScolaire)
            foreach ($rec in $nivRecords) {
                $row = New-Object ClasseursRouges.ExcelRow
                $row.Height = 20.0
                $nameDisplay = if ($nivRecords.Count -gt 1) { "$($st.eleveNom) ($($rec.anneeScolaire))" } else { $st.eleveNom }
                $row.Cells.Add((New-ExcelCell $nameDisplay 9))
                $row.Cells.Add((New-ExcelCell $rec.anneeScolaire 0))
                $row.Cells.Add((New-ExcelCell $rec.classe 0))
                $row.Cells.Add((New-ExcelCell $rec.maitre 0))
                $row.Cells.Add((New-GradeCell $rec.moyPrincipales $true))
                $row.Cells.Add((New-GradeCell $rec.moyGenerale $true))
                $row.Cells.Add((New-ExcelCell $rec.absT1 3 ($rec.absT1 -match '^\d+$')))
                $row.Cells.Add((New-ExcelCell $rec.absT2 3 ($rec.absT2 -match '^\d+$')))
                $row.Cells.Add((New-ExcelCell $rec.absT3 3 ($rec.absT3 -match '^\d+$')))
                $row.Cells.Add((New-ExcelCell $rec.absTotal 6 ($rec.absTotal -match '^\d+$')))
                $row.Cells.Add((New-ExcelCell $rec.situation 0))

                foreach ($dName in $discList) {
                    $matchedD = $null
                    foreach ($d in $rec.disciplines) {
                        if ((Normalize-Identity $d.discipline) -eq (Normalize-Identity $dName)) {
                            $matchedD = $d
                            break
                        }
                    }
                    if ($null -eq $matchedD) {
                        $row.Cells.Add((New-ExcelCell "" 3))
                        $row.Cells.Add((New-ExcelCell "" 3))
                        $row.Cells.Add((New-ExcelCell "" 3))
                        $row.Cells.Add((New-ExcelCell "" 6))
                    } else {
                        $row.Cells.Add((New-GradeCell $matchedD.t1 $false))
                        $row.Cells.Add((New-GradeCell $matchedD.t2 $false))
                        $row.Cells.Add((New-GradeCell $matchedD.t3 $false))
                        $row.Cells.Add((New-GradeCell $matchedD.moyenne $true))
                    }
                }
                $shNiv.Rows.Add($row)
            }
        }
        $sheets.Add($shNiv)
    }

    foreach ($st in $cohort) {
        $shSt = New-Object ClasseursRouges.ExcelSheet
        $shSt.Name = $st.eleveNom
        $shSt.ColumnWidths.AddRange([double[]]@(34.0, 16.0, 9.0, 9.0, 9.0, 14.0))

        $rH1 = New-Object ClasseursRouges.ExcelRow
        $rH1.Height = 26.0
        $rH1.Cells.Add((New-ExcelCell "FICHE INDIVIDUELLE DE SUIVI - $($st.eleveNom) (N $($st.numEleve))" 8))
        $shSt.Rows.Add($rH1)

        foreach ($rec in ($st.history | Sort-Object anneeScolaire, niveau)) {
            $rBlank = New-Object ClasseursRouges.ExcelRow
            $rBlank.Height = 10.0
            $shSt.Rows.Add($rBlank)

            $rYrHead = New-Object ClasseursRouges.ExcelRow
            $rYrHead.Height = 24.0
            $yrTitle = "$($rec.niveau) - Annee scolaire $($rec.anneeScolaire) | Classe : $($rec.classe) ($($rec.regroupement)) | Maitre : $($rec.maitre)"
            $rYrHead.Cells.Add((New-ExcelCell $yrTitle 1))
            for ($i = 1; $i -lt 6; $i++) { $rYrHead.Cells.Add((New-ExcelCell "" 1)) }
            $rowIdx1 = $shSt.Rows.Count + 1
            $shSt.MergedCells.Add("A${rowIdx1}:F${rowIdx1}")
            $shSt.Rows.Add($rYrHead)

            $rCols = New-Object ClasseursRouges.ExcelRow
            $rCols.Height = 20.0
            foreach ($ch in @("Discipline", "Categorie", "T1", "T2", "T3", "Moyenne annuelle")) {
                $rCols.Cells.Add((New-ExcelCell $ch 2))
            }
            $shSt.Rows.Add($rCols)

            foreach ($d in $rec.disciplines) {
                $rD = New-Object ClasseursRouges.ExcelRow
                $profStr = if ($null -ne $d.PSObject.Properties["professeur"] -and -not [string]::IsNullOrWhiteSpace($d.professeur)) { $d.professeur.Trim() } else { "" }
                $rD.Height = if ($profStr -ne "") { 29.0 } else { 19.5 }
                $cat = if ($d.principale) { "Principale" } else { "Autre discipline" }
                $styleDisc = if ($d.principale) { 9 } else { 0 }
                $discCellText = if ($profStr -ne "") { "$($d.discipline)`n$profStr" } else { $d.discipline }
                $rD.Cells.Add((New-ExcelCell $discCellText $styleDisc))
                $rD.Cells.Add((New-ExcelCell $cat 0))
                $rD.Cells.Add((New-GradeCell $d.t1 $false))
                $rD.Cells.Add((New-GradeCell $d.t2 $false))
                $rD.Cells.Add((New-GradeCell $d.t3 $false))
                $rD.Cells.Add((New-GradeCell $d.moyenne $true))
                $shSt.Rows.Add($rD)
            }

            $rMP = New-Object ClasseursRouges.ExcelRow
            $rMP.Height = 20.0
            $rMP.Cells.Add((New-ExcelCell "Moyenne des disciplines principales" 9))
            $rMP.Cells.Add((New-ExcelCell "" 0))
            $rMP.Cells.Add((New-ExcelCell "" 0))
            $rMP.Cells.Add((New-ExcelCell "" 0))
            $rMP.Cells.Add((New-ExcelCell "" 0))
            $rMP.Cells.Add((New-GradeCell $rec.moyPrincipales $true))
            $shSt.Rows.Add($rMP)

            $rMG = New-Object ClasseursRouges.ExcelRow
            $rMG.Height = 20.0
            $rMG.Cells.Add((New-ExcelCell "Moyenne generale" 9))
            $rMG.Cells.Add((New-ExcelCell "" 0))
            $rMG.Cells.Add((New-ExcelCell "" 0))
            $rMG.Cells.Add((New-ExcelCell "" 0))
            $rMG.Cells.Add((New-ExcelCell "" 0))
            $rMG.Cells.Add((New-GradeCell $rec.moyGenerale $true))
            $shSt.Rows.Add($rMG)

            $rAbs = New-Object ClasseursRouges.ExcelRow
            $rAbs.Height = 20.0
            $rAbs.Cells.Add((New-ExcelCell "Absences (T1 / T2 / T3 / Total annuel)" 9))
            $rAbs.Cells.Add((New-ExcelCell ("Situation : " + $rec.situation) 0))
            $rAbs.Cells.Add((New-ExcelCell $rec.absT1 3 ($rec.absT1 -match '^\d+$')))
            $rAbs.Cells.Add((New-ExcelCell $rec.absT2 3 ($rec.absT2 -match '^\d+$')))
            $rAbs.Cells.Add((New-ExcelCell $rec.absT3 3 ($rec.absT3 -match '^\d+$')))
            $rAbs.Cells.Add((New-ExcelCell $rec.absTotal 6 ($rec.absTotal -match '^\d+$')))
            $shSt.Rows.Add($rAbs)
        }
        $sheets.Add($shSt)
    }

    [ClasseursRouges.XlsxWriter]::SaveWorkbook($outputPath, $sheets)
    return $outputPath
}

function Export-AllClassesExcelReportsForYear($db, [string]$targetAnnee, [string]$outputFolder) {
    if (-not (Test-Path $outputFolder)) {
        New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
    }
    $filterByYr = (-not [string]::IsNullOrWhiteSpace($targetAnnee) -and $targetAnnee -ne "TOUTES")
    $pairs = New-Object 'System.Collections.Generic.List[object]'
    $seenKeys = @{}

    foreach ($st in @($db.students)) {
        foreach ($h in @($st.history)) {
            $cls = if ($h.classe) { $h.classe.Trim() } else { "" }
            if ([string]::IsNullOrWhiteSpace($cls)) { continue }
            $yr = if ($h.anneeScolaire) { $h.anneeScolaire.Trim() } else { "" }
            if ($filterByYr -and $yr -ne $targetAnnee) { continue }
            $effYr = if ($filterByYr) { $targetAnnee } else { $yr }
            $key = "${effYr}|${cls}"
            if (-not $seenKeys.ContainsKey($key)) {
                $seenKeys[$key] = $true
                $pairs.Add([PSCustomObject]@{
                    anneeScolaire = $effYr
                    classe        = $cls
                })
            }
        }
    }

    $sortedPairs = @($pairs | Sort-Object anneeScolaire, classe)
    $dateStr = (Get-Date).ToString("yyyy-MM-dd")
    $generatedFiles = @()

    foreach ($item in $sortedPairs) {
        $yr = $item.anneeScolaire
        $cls = $item.classe
        $cleanYr = if ([string]::IsNullOrWhiteSpace($yr) -or $yr -eq "TOUTES") { "" } else { ($yr -replace '[\\/:*?"<>|]', '-') + "_" }
        $safeCls = $cls -replace '[\\/:*?"<>|]', '-'
        $fileName = "Classeur_Rouge_${cleanYr}${safeCls}_${dateStr}.xlsx"
        $outPath = Join-Path $outputFolder $fileName
        $yrFilter = if ([string]::IsNullOrWhiteSpace($yr)) { "TOUTES" } else { $yr }
        Export-CohortExcelReport $db $cls $outPath $yrFilter | Out-Null
        $generatedFiles += $outPath
    }

    return ,$generatedFiles
}
