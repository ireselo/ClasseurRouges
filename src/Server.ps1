# Serveur Local & Lanceur d'Application - Classeurs Rouges (CO Geneve)
# Fonctionne 100% hors-ligne (127.0.0.1 uniquement)

$ErrorActionPreference = "Continue"
$baseDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
. (Join-Path $baseDir "src\Engine.ps1")

Add-Type -AssemblyName System.Windows.Forms

function Select-PdfFilesDialog {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Sélectionner des documents à importer (PDF, XML Notes/Absences, Excel/CSV Professeurs)"
    $dlg.Filter = "Documents supportés (*.pdf;*.xml;*.xls;*.xlsx;*.csv;*.txt)|*.pdf;*.xml;*.xls;*.xlsx;*.csv;*.txt|Fichiers Excel Professeurs (*.xls;*.xlsx)|*.xls;*.xlsx|Fichiers XML Notes & Absences (*.xml)|*.xml|Bulletins PDF (*.pdf)|*.pdf|Fichiers CSV (*.csv;*.txt)|*.csv;*.txt|Tous les fichiers (*.*)|*.*"
    $dlg.Multiselect = $true
    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true
    $form.WindowState = [System.Windows.Forms.FormWindowState]::Minimized
    $form.Show()
    $form.Hide()
    $res = $dlg.ShowDialog($form)
    $form.Dispose()
    if ($res -eq [System.Windows.Forms.DialogResult]::OK) {
        return @($dlg.FileNames)
    }
    return @()
}

function Select-SaveExcelDialog([string]$defaultFileName) {
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title = "Enregistrer le rapport Excel (.xlsx)"
    $dlg.Filter = "Classeur Excel (*.xlsx)|*.xlsx"
    $dlg.FileName = $defaultFileName
    $rapportsDir = Join-Path $baseDir "Rapports_Excel"
    if (-not (Test-Path $rapportsDir)) {
        New-Item -ItemType Directory -Path $rapportsDir -Force | Out-Null
    }
    $dlg.InitialDirectory = $rapportsDir
    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true
    $form.WindowState = [System.Windows.Forms.FormWindowState]::Minimized
    $form.Show()
    $form.Hide()
    $res = $dlg.ShowDialog($form)
    $form.Dispose()
    if ($res -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dlg.FileName
    }
    return ""
}

function Select-ExportFolderDialog([string]$description = "Sélectionner le dossier où enregistrer les rapports Excel (1 fichier par classe)") {
    $rapportsDir = Join-Path $baseDir "Rapports_Excel"
    if (-not (Test-Path $rapportsDir)) {
        New-Item -ItemType Directory -Path $rapportsDir -Force | Out-Null
    }
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = $description
    $dlg.SelectedPath = $rapportsDir
    $dlg.ShowNewFolderButton = $true
    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true
    $form.WindowState = [System.Windows.Forms.FormWindowState]::Minimized
    $form.Show()
    $form.Hide()
    $res = $dlg.ShowDialog($form)
    $form.Dispose()
    if ($res -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dlg.SelectedPath
    }
    return ""
}

function Send-JsonResponse($response, $obj, [int]$statusCode = 200) {
    $response.StatusCode = $statusCode
    $response.ContentType = "application/json; charset=utf-8"
    $response.Headers.Add("Cache-Control", "no-store")
    $json = $obj | ConvertTo-Json -Depth 12 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $response.ContentLength64 = $bytes.Length
    $response.OutputStream.Write($bytes, 0, $bytes.Length)
    $response.OutputStream.Close()
}

# Find free localhost port
$port = 18940
$listener = $null
for ($tryP = 18940; $tryP -le 18990; $tryP++) {
    try {
        $l = New-Object System.Net.HttpListener
        $l.Prefixes.Add("http://127.0.0.1:$tryP/")
        $l.Start()
        $listener = $l
        $port = $tryP
        break
    } catch {
        if ($null -ne $l) { try { $l.Close() } catch {} }
    }
}

if ($null -eq $listener) {
    [System.Windows.Forms.MessageBox]::Show("Impossible de démarrer le moteur local.", "Classeurs Rouges", "OK", "Error") | Out-Null
    exit 1
}

$url = "http://127.0.0.1:$port/"
Write-Host "Classeurs Rouges demarre sur $url (100% Hors-Ligne)"

# Launch desktop window (Edge --app mode or default browser)
$edgePaths = @(
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
    "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
)
$edgeExe = $edgePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
if ($edgeExe) {
    $userDataDir = Join-Path $env:LOCALAPPDATA "ClasseursRouges_AppProfile"
    Start-Process -FilePath $edgeExe -ArgumentList "--app=$url --window-size=1420,900 --user-data-dir=`"$userDataDir`" --disable-extensions --no-first-run"
} else {
    Start-Process $url
}

# Main HTTP loop
try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $req = $context.Request
        $res = $context.Response
        $path = $req.Url.AbsolutePath

        try {
            if ($req.HttpMethod -eq "GET" -and ($path -eq "/" -or $path -eq "/index.html")) {
                $uiPath = Join-Path $baseDir "src\ui.html"
                $htmlBytes = [System.IO.File]::ReadAllBytes($uiPath)
                $res.StatusCode = 200
                $res.ContentType = "text/html; charset=utf-8"
                $res.Headers.Add("Cache-Control", "no-store")
                $res.ContentLength64 = $htmlBytes.Length
                $res.OutputStream.Write($htmlBytes, 0, $htmlBytes.Length)
                $res.OutputStream.Close()
                continue
            }

            if ($req.HttpMethod -eq "GET" -and $path -eq "/api/db") {
                $db = Load-Database $baseDir
                Send-JsonResponse $res @{ ok = $true; db = $db }
                continue
            }

            if ($req.HttpMethod -eq "POST" -and $path -eq "/api/pick-pdfs") {
                . (Join-Path $baseDir "src\Engine.ps1")
                $files = Select-PdfFilesDialog
                if ($files.Count -eq 0) {
                    Send-JsonResponse $res @{ ok = $true; cancelled = $true; bulletins = @() }
                    continue
                }
                $batchRes = Import-DocumentsBatch $files $baseDir
                Send-JsonResponse $res @{
                    ok = $true
                    cancelled = $false
                    bulletins = @($batchRes.bulletins)
                    autoSavedCount = $batchRes.autoSavedCount
                    teachersParsedCount = $batchRes.teachersParsedCount
                    db = $batchRes.db
                }
                continue
            }

            if ($req.HttpMethod -eq "POST" -and $path -eq "/api/upload-pdf") {
                . (Join-Path $baseDir "src\Engine.ps1")
                $sr = New-Object System.IO.StreamReader($req.InputStream, [System.Text.Encoding]::UTF8)
                $bodyJson = $sr.ReadToEnd()
                $sr.Close()
                $payload = $bodyJson | ConvertFrom-Json

                $tempFiles = @()
                try {
                    foreach ($item in $payload.files) {
                        $b64 = $item.dataBase64
                        if ($b64 -match '^data:.*?;base64,(.*)$') { $b64 = $Matches[1] }
                        $bytes = [Convert]::FromBase64String($b64)
                        $safeName = [System.IO.Path]::GetFileName($item.fileName)
                        $tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "agy_upload_" + $safeName)
                        [System.IO.File]::WriteAllBytes($tempPath, $bytes)
                        $tempFiles += $tempPath
                    }
                    $batchRes = Import-DocumentsBatch $tempFiles $baseDir
                    Send-JsonResponse $res @{
                        ok = $true
                        bulletins = @($batchRes.bulletins)
                        autoSavedCount = $batchRes.autoSavedCount
                        teachersParsedCount = $batchRes.teachersParsedCount
                        db = $batchRes.db
                    }
                } finally {
                    foreach ($tf in $tempFiles) {
                        Remove-Item $tf -Force -ErrorAction SilentlyContinue
                    }
                }
                continue
            }

            if ($req.HttpMethod -eq "POST" -and $path -eq "/api/save-bulletins") {
                . (Join-Path $baseDir "src\Engine.ps1")
                $sr = New-Object System.IO.StreamReader($req.InputStream, [System.Text.Encoding]::UTF8)
                $bodyJson = $sr.ReadToEnd()
                $sr.Close()
                $payload = $bodyJson | ConvertFrom-Json

                $db = Load-Database $baseDir
                $bulletinsArr = @($payload.bulletins)
                $count = $bulletinsArr.Count
                if ($count -gt 0) {
                    $db = Upsert-BulletinRecordsBatch $db $bulletinsArr
                }
                Save-Database $baseDir $db
                Send-JsonResponse $res @{ ok = $true; savedCount = $count; db = $db }
                continue
            }

            if ($req.HttpMethod -eq "POST" -and $path -eq "/api/delete-student") {
                $sr = New-Object System.IO.StreamReader($req.InputStream, [System.Text.Encoding]::UTF8)
                $bodyJson = $sr.ReadToEnd()
                $sr.Close()
                $payload = $bodyJson | ConvertFrom-Json

                $db = Load-Database $baseDir
                $db.students = @($db.students | Where-Object { $_.id -ne $payload.studentId })
                Save-Database $baseDir $db
                Send-JsonResponse $res @{ ok = $true; db = $db }
                continue
            }

            if ($req.HttpMethod -eq "POST" -and $path -eq "/api/delete-year") {
                $sr = New-Object System.IO.StreamReader($req.InputStream, [System.Text.Encoding]::UTF8)
                $bodyJson = $sr.ReadToEnd()
                $sr.Close()
                $payload = $bodyJson | ConvertFrom-Json

                $db = Load-Database $baseDir
                foreach ($st in $db.students) {
                    if ($st.id -eq $payload.studentId) {
                        $st.history = @($st.history | Where-Object {
                            -not ($_.anneeScolaire -eq $payload.anneeScolaire -and $_.niveau -eq $payload.niveau)
                        })
                        if ($st.history.Count -gt 0) {
                            $latest = $st.history[-1]
                            $st.classeActuelle = $latest.classe
                            $st.niveauActuel = $latest.niveau
                            $st.anneeActuelle = $latest.anneeScolaire
                        }
                    }
                }
                $db.students = @($db.students | Where-Object { $_.history.Count -gt 0 })
                Save-Database $baseDir $db
                Send-JsonResponse $res @{ ok = $true; db = $db }
                continue
            }

            if ($req.HttpMethod -eq "POST" -and $path -eq "/api/end-school-year") {
                $sr = New-Object System.IO.StreamReader($req.InputStream, [System.Text.Encoding]::UTF8)
                $bodyJson = $sr.ReadToEnd()
                $sr.Close()
                $payload = $bodyJson | ConvertFrom-Json

                $repeatingIds = @()
                if ($null -ne $payload.repeatingStudentIds) {
                    $repeatingIds = @($payload.repeatingStudentIds | ForEach-Object { [string]$_ })
                }

                $db = Load-Database $baseDir
                $remainingStudents = @()
                $deletedCount = 0
                $keptCount = 0

                foreach ($st in $db.students) {
                    $has11InHistory = ($null -ne $st.history -and @($st.history | Where-Object { $_.niveau -eq "11CO" }).Count -gt 0)
                    $is11 = ($st.niveauActuel -eq "11CO") -or $has11InHistory -or ($st.classeActuelle -match '^11')

                    if ($is11) {
                        if ($repeatingIds -contains [string]$st.id) {
                            $remainingStudents += $st
                            $keptCount++
                        } else {
                            $deletedCount++
                        }
                    } else {
                        $remainingStudents += $st
                    }
                }

                $db.students = @($remainingStudents)
                Save-Database $baseDir $db
                Send-JsonResponse $res @{
                    ok = $true
                    deletedCount = $deletedCount
                    keptCount = $keptCount
                    db = $db
                }
                continue
            }

            if ($req.HttpMethod -eq "POST" -and $path -eq "/api/export-excel") {
                . (Join-Path $baseDir "src\Engine.ps1")
                $sr = New-Object System.IO.StreamReader($req.InputStream, [System.Text.Encoding]::UTF8)
                $bodyJson = $sr.ReadToEnd()
                $sr.Close()
                $payload = $bodyJson | ConvertFrom-Json

                $targetAnnee = if ($payload.PSObject.Properties["anneeScolaire"] -and $payload.anneeScolaire) { $payload.anneeScolaire } else { "TOUTES" }
                $targetClasse = $payload.classe
                $cleanYr = if ([string]::IsNullOrWhiteSpace($targetAnnee) -or $targetAnnee -eq "TOUTES") { "" } else { "${targetAnnee}_" }
                $cleanCls = if ([string]::IsNullOrWhiteSpace($targetClasse) -or $targetClasse -eq "TOUTES") { "Toutes_Classes" } else { $targetClasse }
                $dateStr = (Get-Date).ToString("yyyy-MM-dd")
                $defaultName = "Classeur_Rouge_${cleanYr}${cleanCls}_${dateStr}.xlsx"

                $savePath = Select-SaveExcelDialog $defaultName
                if ([string]::IsNullOrWhiteSpace($savePath)) {
                    Send-JsonResponse $res @{ ok = $true; cancelled = $true }
                    continue
                }

                $db = Load-Database $baseDir
                Export-CohortExcelReport $db $targetClasse $savePath $targetAnnee | Out-Null

                if ($payload.openAfter -eq $true -and (Test-Path $savePath)) {
                    Start-Process $savePath
                }

                Send-JsonResponse $res @{ ok = $true; cancelled = $false; filePath = $savePath }
                continue
            }

            if ($req.HttpMethod -eq "POST" -and $path -eq "/api/export-all-classes-excel") {
                . (Join-Path $baseDir "src\Engine.ps1")
                $sr = New-Object System.IO.StreamReader($req.InputStream, [System.Text.Encoding]::UTF8)
                $bodyJson = $sr.ReadToEnd()
                $sr.Close()
                $payload = $bodyJson | ConvertFrom-Json

                $targetAnnee = if ($payload.PSObject.Properties["anneeScolaire"] -and $payload.anneeScolaire) { $payload.anneeScolaire } else { "TOUTES" }
                $yrDesc = if ($targetAnnee -eq "TOUTES") { "toutes les années" } else { "l'année scolaire $targetAnnee" }

                $folderPath = Select-ExportFolderDialog "Sélectionner le dossier où générer les classeurs Excel par classe pour $yrDesc"
                if ([string]::IsNullOrWhiteSpace($folderPath)) {
                    Send-JsonResponse $res @{ ok = $true; cancelled = $true }
                    continue
                }

                $db = Load-Database $baseDir
                $generatedFiles = Export-AllClassesExcelReportsForYear $db $targetAnnee $folderPath

                if ($payload.openAfter -eq $true -and $generatedFiles.Count -gt 0 -and (Test-Path $folderPath)) {
                    Start-Process $folderPath
                }

                Send-JsonResponse $res @{
                    ok         = $true
                    cancelled  = $false
                    folderPath = $folderPath
                    count      = $generatedFiles.Count
                    files      = @($generatedFiles)
                }
                continue
            }

            if ($req.HttpMethod -eq "POST" -and $path -eq "/api/quit") {
                Send-JsonResponse $res @{ ok = $true }
                break
            }

            Send-JsonResponse $res @{ ok = $false; error = "Route inconnue : $path" } 404
        } catch {
            Send-JsonResponse $res @{ ok = $false; error = $_.Exception.Message } 500
        }
    }
} finally {
    $listener.Stop()
    $listener.Close()
}
