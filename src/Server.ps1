# Serveur Local & Lanceur d'Application - Classeurs Rouges (CO Geneve)
# Fonctionne 100% hors-ligne (127.0.0.1 uniquement)

$ErrorActionPreference = "Continue"
$baseDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
. (Join-Path $baseDir "src\Engine.ps1")

Add-Type -AssemblyName System.Windows.Forms

function Select-PdfFilesDialog {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Sélectionner un ou plusieurs bulletins scolaires (PDF)"
    $dlg.Filter = "Bulletins PDF (*.pdf)|*.pdf|Tous les fichiers (*.*)|*.*"
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
                $files = Select-PdfFilesDialog
                if ($files.Count -eq 0) {
                    Send-JsonResponse $res @{ ok = $true; cancelled = $true; bulletins = @() }
                    continue
                }
                $allBulletins = @()
                foreach ($f in $files) {
                    $extracted = Extract-BulletinsFromPdf $f
                    $allBulletins += $extracted
                }
                Send-JsonResponse $res @{ ok = $true; cancelled = $false; bulletins = $allBulletins }
                continue
            }

            if ($req.HttpMethod -eq "POST" -and $path -eq "/api/upload-pdf") {
                $sr = New-Object System.IO.StreamReader($req.InputStream, [System.Text.Encoding]::UTF8)
                $bodyJson = $sr.ReadToEnd()
                $sr.Close()
                $payload = $bodyJson | ConvertFrom-Json

                $allBulletins = @()
                foreach ($item in $payload.files) {
                    $b64 = $item.dataBase64
                    if ($b64 -match '^data:.*?;base64,(.*)$') { $b64 = $Matches[1] }
                    $bytes = [Convert]::FromBase64String($b64)
                    $safeName = [System.IO.Path]::GetFileName($item.fileName)
                    $tempPdf = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "agy_upload_" + $safeName)
                    [System.IO.File]::WriteAllBytes($tempPdf, $bytes)
                    try {
                        $extracted = Extract-BulletinsFromPdf $tempPdf
                        foreach ($eb in $extracted) { $eb.sourceFile = $safeName }
                        $allBulletins += $extracted
                    } finally {
                        Remove-Item $tempPdf -Force -ErrorAction SilentlyContinue
                    }
                }
                Send-JsonResponse $res @{ ok = $true; bulletins = $allBulletins }
                continue
            }

            if ($req.HttpMethod -eq "POST" -and $path -eq "/api/save-bulletins") {
                $sr = New-Object System.IO.StreamReader($req.InputStream, [System.Text.Encoding]::UTF8)
                $bodyJson = $sr.ReadToEnd()
                $sr.Close()
                $payload = $bodyJson | ConvertFrom-Json

                $db = Load-Database $baseDir
                $count = 0
                foreach ($b in $payload.bulletins) {
                    $db = Upsert-BulletinRecord $db $b
                    $count++
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
                $sr = New-Object System.IO.StreamReader($req.InputStream, [System.Text.Encoding]::UTF8)
                $bodyJson = $sr.ReadToEnd()
                $sr.Close()
                $payload = $bodyJson | ConvertFrom-Json

                $targetClasse = $payload.classe
                $cleanCls = if ([string]::IsNullOrWhiteSpace($targetClasse) -or $targetClasse -eq "TOUTES") { "Toutes_Classes" } else { $targetClasse }
                $dateStr = (Get-Date).ToString("yyyy-MM-dd")
                $defaultName = "Classeur_Rouge_${cleanCls}_${dateStr}.xlsx"

                $savePath = Select-SaveExcelDialog $defaultName
                if ([string]::IsNullOrWhiteSpace($savePath)) {
                    Send-JsonResponse $res @{ ok = $true; cancelled = $true }
                    continue
                }

                $db = Load-Database $baseDir
                Export-CohortExcelReport $db $targetClasse $savePath | Out-Null

                if ($payload.openAfter -eq $true -and (Test-Path $savePath)) {
                    Start-Process $savePath
                }

                Send-JsonResponse $res @{ ok = $true; cancelled = $false; filePath = $savePath }
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
