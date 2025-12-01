Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# Aggiungi assembly necessario per CommonOpenFileDialog
# -------------------
# Classe per DataGridView
# -------------------
Add-Type @"
public class FileRecord {
    public string FilePath {get; set;}
    public string FileNameDate {get; set;}
    public string ExifDate {get; set;}
    public string MediaDate {get; set;}
    public string FileSystemDate {get; set;}
    public string JsonDate {get; set;}
    public string Status {get; set;}
}
"@

Add-Type @"
using System;
using System.ComponentModel;
using System.Collections.Generic;

public class SortableBindingList<T> : BindingList<T>
{
    public string Info { get { return "Sortable BindingList"; } }

    private bool _isSorted = false;
    private ListSortDirection _sortDirection = ListSortDirection.Ascending;
    private PropertyDescriptor _sortProperty;

    protected override bool SupportsSortingCore
    {
        get { return true; }
    }

    protected override bool IsSortedCore
    {
        get { return _isSorted; }
    }

    protected override PropertyDescriptor SortPropertyCore
    {
        get { return _sortProperty; }
    }

    protected override ListSortDirection SortDirectionCore
    {
        get { return _sortDirection; }
    }

    protected override void ApplySortCore(PropertyDescriptor prop, ListSortDirection direction)
    {
        List<T> items = new List<T>(this.Items);
        items.Sort(delegate(T x, T y)
        {
            object xValue = prop.GetValue(x);
            object yValue = prop.GetValue(y);
            return Comparer<object>.Default.Compare(xValue, yValue);
        });

        if (direction == ListSortDirection.Descending)
        {
            items.Reverse();
        }

        for (int i = 0; i < items.Count; i++)
        {
            this.Items[i] = items[i];
        }

        _sortProperty = prop;
        _sortDirection = direction;
        _isSorted = true;

        this.OnListChanged(new ListChangedEventArgs(ListChangedType.Reset, -1));
    }

    protected override void RemoveSortCore()
    {
        _isSorted = false;
    }
}
"@





# Cache globale JSON per cartelle e sottocartelle
$Global:JsonCacheByDirectory = @{}

$Global:StopRequested = $false

$Global:ImageExtensions = @(".jpg",".jpeg",".png",".heic",".webp",".dng",".nef",".cr2",".cr3",".arw",".rw2",".gif",".tif",".tiff",".bmp")
$Global:VideoExtensions = @(".mp4",".mov",".m4v",".avi",".mpg",".mpeg",".mkv",".mts",".m2ts",".wmv",".mpe", ".m2p", ".m2v", ".mp2")

# Formati senza metadata interni (MPEG Program Stream + Elementary Streams)
$Global:NoMetadataFormats = @(".mpg", ".mpeg", ".mpe", ".m2p", ".m2v", ".mp2", ".avi")


function Initialize-JsonCache {
    param (
        [string]$RootPath
    )

    # Svuota la cache precedente
    $Global:JsonCacheByDirectory.Clear()

    # Enumerazione efficiente di tutti i JSON nella cartella e sottocartelle
    $jsonPaths = [System.IO.Directory]::EnumerateFiles($RootPath, '*.json', [System.IO.SearchOption]::AllDirectories)

    foreach ($fullPath in $jsonPaths) {
        try {
            $dir = [System.IO.Path]::GetDirectoryName($fullPath)
            $fi = Get-Item -LiteralPath $fullPath

            if (-not $Global:JsonCacheByDirectory.ContainsKey($dir)) {
                $Global:JsonCacheByDirectory[$dir] = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
            }

            $Global:JsonCacheByDirectory[$dir].Add($fi)
        }
        catch {
            Write-Warning "Errore accesso JSON: $fullPath — $_"
        }
    }
}

# -------------------
# Funzioni di estrazione date
# -------------------
# -------------------
# Funzione: Estrazione data da JSON Google Photos
# -------------------
# -------------------
# Funzione: Estrazione data da JSON Google Photos con cache
# -------------------
function Get-GoogleJsonDate {
    param (
        [string]$MediaFile
    )
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($MediaFile)
    $dir = [System.IO.Path]::GetDirectoryName($MediaFile)
    # Controlla se la cartella è nella cache
    if (-not $Global:JsonCacheByDirectory.ContainsKey($dir)) {
        return $null
    }

    $jsonList = $Global:JsonCacheByDirectory[$dir]

	
	# 1) MATCH ESATTO 
	$exactMatches = $jsonList | Where-Object { $_.BaseName -eq $baseName }
	if ($exactMatches.Count -gt 0) {
		$jsonFile = $exactMatches[0]
	}
	else {

		# 2) MATCH PER PREFISSO (progressivamente accorciato)
		$searchName = $baseName
		$jsonFile = $null

		while ($searchName.Length -ge 4) {

			# Candidati che iniziano con searchName
			$candidates = $jsonList | Where-Object { $_.BaseName.StartsWith($searchName) }

			if ($candidates.Count -gt 0) {

				# Prendi il più lungo (senza Sort-Object)
				$jsonFile = $candidates[0]
				$maxLength = $jsonFile.BaseName.Length

				foreach ($file in $candidates) {
					if ($file.BaseName.Length -gt $maxLength) {
						$jsonFile = $file
						$maxLength = $file.BaseName.Length
					}
				}

				break  # trovato ? interrompi il ciclo
			}

			# Accorcia il basename di 1 carattere
			$searchName = $searchName.Substring(0, $searchName.Length - 1)
		}

		# Se ancora null ? nessun JSON trovato
		if (-not $jsonFile) { return $null }
	}


    try {
        $content = Get-Content -Raw -Path $jsonFile.FullName | ConvertFrom-Json

        $ts = $null
        if ($content.photoTakenTime -and $content.photoTakenTime.timestamp) {
            $ts = $content.photoTakenTime.timestamp
        }
        elseif ($content.videoTakenTime -and $content.videoTakenTime.timestamp) {
            $ts = $content.videoTakenTime.timestamp
        }
        elseif ($content.creationTime -and $content.creationTime.timestamp) {
            $ts = $content.creationTime.timestamp
        }

        if ($ts) {
            return ([datetime]'1970-01-01').AddSeconds([double]$ts).ToLocalTime()
        }
    }
    catch {
        Write-Warning "Errore parsing JSON: $($jsonFile.FullName) — $_"
    }

    return $null
}



# -------------------
# Funzione: Estrazione data dal nome del file
# -------------------
function Get-DateFromFilename {
    param([string]$FileName)

    $name = [System.IO.Path]::GetFileNameWithoutExtension($FileName)

    # -------------------------------------------------------------
    # FUNZIONE DI SUPPORTO — costruisce DateTime in modo sicuro
    # -------------------------------------------------------------
    function _mkdate {
        param($y, $m, $d, $h = 0, $min = 0, $s = 0, $ms = 0)
        try { return (Get-Date -Year $y -Month $m -Day $d -Hour $h -Minute $min -Second $s -Millisecond $ms) }
        catch { return $null }
    }

    # -------------------------------------------------------------
    # 1) PATTERN COMPLETI (data + ora) — ALTA PRIORITÀ
    # -------------------------------------------------------------
    $fullPatterns = @(
        # IMG/PXL/MVIMG_YYYYMMDD_HHMMSS
        @{
            Regex = '.*?(IMG|PXL|MVIMG)_(\d{4})(\d{2})(\d{2})[_ ]?(\d{2})(\d{2})(\d{2})';
            Map   = { param($m) _mkdate $m[2] $m[3] $m[4] $m[5] $m[6] $m[7] }
        },
        # IMG_YYYYMMDD_HHMMSS
        @{
            Regex = '.*?IMG_(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})';
            Map   = { param($m) _mkdate $m[1] $m[2] $m[3] $m[4] $m[5] $m[6] }
        },
        # YYYY-MM-DD-HH-MM-SS
        @{
            Regex = '(^|[^0-9])(\d{4})-(\d{2})-(\d{2})-(\d{2})-(\d{2})-(\d{2})(?=[^0-9]|$)';
            Map   = { param($m) _mkdate $m[2] $m[3] $m[4] $m[5] $m[6] $m[7] }
        },
        # YYYYMMDD-HHMMSS o YYYYMMDD_HHMMSS
        @{
            Regex = '(^|[^0-9])(\d{4})(\d{2})(\d{2})[-_](\d{2})(\d{2})(\d{2})(?=[^0-9]|$)';
            Map   = { param($m) _mkdate $m[2] $m[3] $m[4] $m[5] $m[6] $m[7] }
        },
        # 14 cifre (YYYYMMDDHHMMSS)
        @{
            Regex = '(^|[^0-9])(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})(?=[^0-9]|$)';
            Map   = { param($m) _mkdate $m[2] $m[3] $m[4] $m[5] $m[6] $m[7] }
        },
        # 17 cifre con millisecondi
        @{
            Regex = '(^|[^0-9])(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})(\d{3})(?=[^0-9]|$)';
            Map   = { param($m) _mkdate $m[2] $m[3] $m[4] $m[5] $m[6] $m[7] $m[8] }
        }
    )

    foreach ($p in $fullPatterns) {
        if ($name -match $p.Regex) {
            $dt = & $p.Map $matches
            if ($dt) { return $dt }
        }
    }

    # -------------------------------------------------------------
    # 2) SOLO DATA (senza ora) — media priorità
    # -------------------------------------------------------------
    $dateOnlyPatterns = @(
        # YYYY-MM-DD
        @{
            Regex = '(^|[^0-9])(\d{4})-(\d{2})-(\d{2})(?=[^0-9]|$)';
            Map   = { param($m) _mkdate $m[2] $m[3] $m[4] }
        },
        # YYYY_MM_DD
        @{
            Regex = '(^|[^0-9])(\d{4})_(\d{2})_(\d{2})(?=[^0-9]|$)';
            Map   = { param($m) _mkdate $m[2] $m[3] $m[4] }
        },
        # YYYYMMDD (8 cifre)
        @{
            Regex = '(^|[^0-9])(\d{4})(\d{2})(\d{2})(?=[^0-9]|$)';
            Map   = { param($m) _mkdate $m[2] $m[3] $m[4] }
        }
    )

    foreach ($p in $dateOnlyPatterns) {
        if ($name -match $p.Regex) {
            $dt = & $p.Map $matches
            if ($dt) { return $dt }
        }
    }

    # -------------------------------------------------------------
    # 3) TIMESTAMP UNIX — bassa priorità
    # -------------------------------------------------------------
    if ($name -match '(^|[^0-9])(\d{13})(?=[^0-9]|$)') {
        try { return ([datetime]'1970-01-01').AddSeconds([double]$matches[2]/1000).ToLocalTime() } catch {}
    }
    if ($name -match '(^|[^0-9])(\d{10})(?=[^0-9]|$)') {
        try { return ([datetime]'1970-01-01').AddSeconds([double]$matches[2]).ToLocalTime() } catch {}
    }

    return $null
}



# -------------------
# EXIF fallback
# -------------------
function Get-ExifDateFallback {
    param([string]$FilePath)
    try {
        $img = [System.Drawing.Image]::FromFile($FilePath)
        $exifTags = @(36867, 36868, 306)
        foreach ($tag in $exifTags) {
            $prop = $img.PropertyItems | Where-Object { $_.Id -eq $tag }
            if ($prop) {
                $dateStr = [System.Text.Encoding]::ASCII.GetString($prop.Value).Trim([char]0).Trim()
                $dt = $null
                try { $dt = [datetime]::ParseExact($dateStr,"yyyy:MM:dd HH:mm:ss",$null) } catch { 
                    try { $dt = [datetime]::Parse($dateStr) } catch { $dt = $null } 
                }
                if ($dt) {
                    $img.Dispose()
#                    $global:txtOutput.AppendText("EXIF [${FilePath}] Tag ${tag}: $dateStr => $dt`r`n")
                    return $dt
                }
            }
        }
        $img.Dispose()
        return $null
    } catch { return $null }
}

# -------------------
# EXIF con ExifTool
# -------------------
function Get-ExifDate {
    param([string]$FilePath, [string]$ExifToolPath)
    $dt = Get-ExifDateFallback $FilePath
    if ($dt -ne $null) { return $dt }
    if (-not $ExifToolPath) { return $null }
    try {
        # $args = @("-charset filename=UTF8", "-j", "-DateTimeOriginal", "-CreateDate", "-AllDates", $FilePath)
		$args = @("-j", "-DateTimeOriginal", "-CreateDate", "-AllDates", $FilePath)
        $output = & $ExifToolPath @args
        if ($output) {
            $json = $output | ConvertFrom-Json
            $obj = $json[0]
            if ($obj.DateTimeOriginal) { $dt = [datetime]::Parse($obj.DateTimeOriginal) }
            elseif ($obj.CreateDate) { $dt = [datetime]::Parse($obj.CreateDate) }
            elseif ($obj.ModifyDate) { $dt = [datetime]::Parse($obj.ModifyDate) }
            return $dt
        }
    } catch { }
    return $null
}

# -------------------
# FFProbe MediaDate
# -------------------
function Get-VideoMediaDate {
    param([string]$FilePath, [string]$FFProbePath)
    if (-not $FFProbePath) { return $null }
    try {
        $ffprobeOutput = & $FFProbePath -v quiet -print_format json -show_entries format_tags=creation_time $FilePath
        if ($ffprobeOutput) {
            $json = $ffprobeOutput | ConvertFrom-Json
            if ($json.format.tags.creation_time) { 
                return [datetime]::Parse($json.format.tags.creation_time)
            }
        }
        return $null
    } catch { return $null }
}

# -------------------
# GUI
# -------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Analisi File Multimediali"
$form.Size = New-Object System.Drawing.Size(1180,648)
$form.StartPosition = "CenterScreen"

# --- Anteprima immagine ---
$picPreview = New-Object System.Windows.Forms.PictureBox
$picPreview.Location = New-Object System.Drawing.Point(880, 96)
$picPreview.Size = New-Object System.Drawing.Size(280, 280)
$picPreview.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$picPreview.BorderStyle = 'FixedSingle'
$picPreview.Anchor = "Top,Right"
$form.Controls.Add($picPreview)


# --- Percorso ---
$txtPath = New-Object System.Windows.Forms.TextBox
$txtPath.Location = New-Object System.Drawing.Point(10,9)
$txtPath.Size = New-Object System.Drawing.Size(700,23)
$txtPath.AllowDrop = $true
$txtPath.Anchor = "Top,Left,Right"
$form.Controls.Add($txtPath)

$btnSelectFolder = New-Object System.Windows.Forms.Button
$btnSelectFolder.Text = "Sfoglia..."
$btnSelectFolder.Location = New-Object System.Drawing.Point(720,9)
$btnSelectFolder.Size = New-Object System.Drawing.Size(150,23)
$btnSelectFolder.Anchor = "Top,Right"
$form.Controls.Add($btnSelectFolder)

# --- Pulsanti + progressbar ---
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Avvia Analisi"
$btnStart.Location = New-Object System.Drawing.Point(10,41)
$btnStart.Size = New-Object System.Drawing.Size(150,27)
$btnStart.Enabled = $false
$btnStart.Anchor = "Top,Left"
$form.Controls.Add($btnStart)

$btnFixFromLog = New-Object System.Windows.Forms.Button
$btnFixFromLog.Text = "Correggi da Log"
$btnFixFromLog.Location = New-Object System.Drawing.Point(170,41)
$btnFixFromLog.Size = New-Object System.Drawing.Size(150,27)
$btnFixFromLog.Enabled = $false
$btnFixFromLog.Anchor = "Top,Left"
$form.Controls.Add($btnFixFromLog)

$btnSaveCsv = New-Object System.Windows.Forms.Button
$btnSaveCsv.Text = "Salva CSV"
$btnSaveCsv.Location = New-Object System.Drawing.Point(330,41)
$btnSaveCsv.Size = New-Object System.Drawing.Size(150,27)
$btnSaveCsv.Enabled = $false
$btnSaveCsv.Anchor = "Top,Left"
$form.Controls.Add($btnSaveCsv)

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = "Stop"
$btnStop.Location = New-Object System.Drawing.Point(490,41)
$btnStop.Size = New-Object System.Drawing.Size(100,27)
$btnStop.Enabled = $false
$btnStop.Anchor = "Top,Left"
$form.Controls.Add($btnStop)

$btnStop.Add_Click({
    $global:StopRequested = $true
    $txtOutput.AppendText("Richiesta di interruzione elaborazione...`r`n")
})

# --- Pulsante Elimina Riga ---
$btnDeleteRow = New-Object System.Windows.Forms.Button
$btnDeleteRow.Text = "Elimina Riga"
$btnDeleteRow.Location = New-Object System.Drawing.Point(600,41)
$btnDeleteRow.Size = New-Object System.Drawing.Size(100,27)
$btnDeleteRow.Enabled = $true
$btnDeleteRow.Anchor = "Top,Left"
$form.Controls.Add($btnDeleteRow)

# --- Gestione click con conferma ---
$btnDeleteRow.Add_Click({
    $selectedCount = $dataGrid.SelectedRows.Count
    if ($selectedCount -eq 0) { return }  # niente selezionato

    $msg = "Sei sicuro di eliminare $selectedCount riga/e selezionata/e?"
    $res = [System.Windows.Forms.MessageBox]::Show($msg,"Conferma eliminazione",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning)

    if ($res -eq [System.Windows.Forms.DialogResult]::Yes) {
        foreach ($row in $dataGrid.SelectedRows) {
            $item = $row.DataBoundItem
            if ($item) {
                $Results.Remove($item)   # rimuove dal BindingList
            }
        }
    }
})



$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(170,75)
$progressBar.Size = New-Object System.Drawing.Size(700,20)
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Anchor = "Top,Left,Right"
$form.Controls.Add($progressBar)

# --- Label ETC ---
$lblETC = New-Object System.Windows.Forms.Label
$lblETC.Location = New-Object System.Drawing.Point(10,75)   # subito sotto la progress bar
$lblETC.Size = New-Object System.Drawing.Size(160,20)
$lblETC.Text = "ETC stimato: --:--:--"
$lblETC.Anchor = "Top,Left,Right"
$form.Controls.Add($lblETC)

# --- DataGrid ---
$dataGrid = New-Object System.Windows.Forms.DataGridView
$dataGrid.Location = New-Object System.Drawing.Point(10,96)
$dataGrid.Size = New-Object System.Drawing.Size(860,415)
$dataGrid.ReadOnly = $false
$dataGrid.AllowUserToAddRows = $false
$dataGrid.AllowUserToDeleteRows = $false
$dataGrid.SelectionMode = "FullRowSelect"
$dataGrid.AutoGenerateColumns = $false
$dataGrid.Anchor = "Top,Bottom,Left,Right"
$form.Controls.Add($dataGrid)

# --- Console log ---
$global:txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = New-Object System.Drawing.Point(10,511)
$txtOutput.Size = New-Object System.Drawing.Size(860,90)
$txtOutput.Multiline = $true
$txtOutput.ScrollBars = "Vertical"
$txtOutput.ReadOnly = $true
$txtOutput.Anchor = "Bottom,Left,Right"
$form.Controls.Add($txtOutput)

# --- Colonne DataGrid ---
$columns = @(
    @{Header="FilePath"; Data="FilePath"; Width=300},
    @{Header="FileNameDate"; Data="FileNameDate"; Width=140},
    @{Header="ExifDate"; Data="ExifDate"; Width=130},
    @{Header="MediaDate"; Data="MediaDate"; Width=130},
    @{Header="FileSystemDate"; Data="FileSystemDate"; Width=130},
    @{Header="JsonDate"; Data="JsonDate"; Width=140},  
    @{Header="Status"; Data="Status"; Width=100}
)

foreach ($col in $columns) {
    $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $column.HeaderText = $col.Header
    $column.DataPropertyName = $col.Data
    $column.Width = $col.Width
    [void]$dataGrid.Columns.Add($column)
}
# Rendi tutte le colonne ordinabili cliccando sull'intestazione
foreach ($col in $dataGrid.Columns) {
    $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
}

$dataGrid.Columns[0].ReadOnly = $true


$ExifToolPath = $null
try { $ExifToolPath = (Get-Command "exiftool.exe" -ErrorAction Stop).Source } catch { }
$FFProbePath = $null
try { $FFProbePath = (Get-Command ".\ffprobe.exe" -ErrorAction Stop).Source } catch { }






# Evento cambio selezione riga -> mostra anteprima
$dataGrid.Add_SelectionChanged({

    if ($dataGrid.SelectedRows.Count -eq 0) {
        if ($picPreview.Image) { $picPreview.Image.Dispose(); $picPreview.Image = $null }
        return
    }

    # Rilascia immagine precedente
    if ($picPreview.Image) {
        $picPreview.Image.Dispose()
        $picPreview.Image = $null
    }

    $record = $dataGrid.SelectedRows[0].DataBoundItem
    if (-not $record) { return }

    $file = $record.FilePath
    if (-not (Test-Path $file)) { return }

    $ext = [System.IO.Path]::GetExtension($file).ToLower()

    # -------------------
    # Immagine
    # -------------------
    if ($Global:ImageExtensions -contains $ext) {
        $txtOutput.AppendText("Preview Immagine $file .`r`n")
        try {
            $img = [System.Drawing.Image]::FromFile($file)
            $clone = $img.Clone()
            $img.Dispose()
            $picPreview.Image = $clone
        }
        catch {
            $picPreview.Image = $null
        }
        return
    }

    # -------------------
    # Video
    # -------------------
        # Se è un video, prova a estrarre la frame a metà video (richiede ffmpeg)
	# if ($Global:VideoExtensions -contains $ext -and $FFProbePath) {
		# try {
			# $txtOutput.AppendText("Preview Video $file in memoria.`r`n")

			# # Rilascia immagine precedente
			# if ($picPreview.Image) {
				# $picPreview.Image.Dispose()
				# $picPreview.Image = $null
			# }

			# $ffmpegPath = $FFProbePath.Replace("ffprobe","ffmpeg")

			# # Calcola durata
			# $duration = & $ffmpegPath -i "`"$file`"" 2>&1 |
				# Select-String "Duration" |
				# ForEach-Object {
					# ($_ -match "Duration: (\d+):(\d+):(\d+.\d+)") | Out-Null
					# $hours = [int]$matches[1]
					# $minutes = [int]$matches[2]
					# $seconds = [double]$matches[3]
					# return ($hours*3600 + $minutes*60 + $seconds)
				# }

			# $time = [math]::Round($duration / 2, 2)

			# # Ottieni dimensioni del video
			# $videoInfo = & $ffmpegPath -i "`"$file`"" 2>&1 |
				# Select-String "Stream.*Video" |
				# ForEach-Object {
					# ($_ -match ", (\d+)x(\d+)[ ,]") | Out-Null
					# return @{ Width = [int]$matches[1]; Height = [int]$matches[2] }
				# }
			# $w = $videoInfo.Width
			# $h = $videoInfo.Height

			# # ProcessStartInfo per ffmpeg raw RGB
			# $startInfo = New-Object System.Diagnostics.ProcessStartInfo
			# $startInfo.FileName = $ffmpegPath
			# $startInfo.Arguments = "-ss $time -i `"$file`" -vframes 1 -f rawvideo -pix_fmt rgb24 -"
			# $startInfo.RedirectStandardOutput = $true
			# $startInfo.UseShellExecute = $false
			# $startInfo.CreateNoWindow = $true

			# $proc = New-Object System.Diagnostics.Process
			# $proc.StartInfo = $startInfo
			# $proc.Start() | Out-Null

			# # Leggi raw RGB
			# $frameSize = $w * $h * 3
			# $buffer = New-Object byte[] $frameSize
			# $read = 0
			# while ($read -lt $frameSize) {
				# $r = $proc.StandardOutput.BaseStream.Read($buffer, $read, $frameSize - $read)
				# if ($r -le 0) { break }
				# $read += $r
			# }
			# $proc.WaitForExit()

			# # Converti raw RGB in Bitmap
			# $bmp = New-Object System.Drawing.Bitmap $w, $h, ([System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
			# $rect = [System.Drawing.Rectangle]::FromLTRB(0, 0, $w, $h)
			# $bmpData = $bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::WriteOnly, $bmp.PixelFormat)
			# [System.Runtime.InteropServices.Marshal]::Copy($buffer, 0, $bmpData.Scan0, $buffer.Length)
			
# # $stride = $bmpData.Stride
# # $rowSize = $w * 3
# # $tmp = New-Object byte[] $stride

# # for ($y = 0; $y -lt $h; $y++) {
    # # for ($x = 0; $x -lt $w; $x++) {
        # # $r = $buffer[($y*$w + $x)*3 + 0]
        # # $g = $buffer[($y*$w + $x)*3 + 1]
        # # $b = $buffer[($y*$w + $x)*3 + 2]
        # # $tmp[$x*3 + 0] = $b
        # # $tmp[$x*3 + 1] = $g
        # # $tmp[$x*3 + 2] = $r
    # # }
    # # [System.Runtime.InteropServices.Marshal]::Copy($tmp, 0, $bmpData.Scan0 + $y*$stride, $rowSize)
# # }
			
			
			# $bmp.UnlockBits($bmpData)

			# # Assegna al PictureBox
			# $picPreview.Image = $bmp
			# $txtOutput.AppendText("Preview Video caricata in memoria.`r`n")
		# }
		# catch {
			# $picPreview.Image = $null
			# $txtOutput.AppendText("Errore preview video in memoria.`r`n")
		# }
		# return
	# }

if ($Global:VideoExtensions -contains $ext -and $FFProbePath) {
    try {
        $txtOutput.AppendText("Preview Video $file in memoria.`r`n")

        # Rilascia immagine precedente
        if ($picPreview.Image) { $picPreview.Image.Dispose(); $picPreview.Image = $null }

        $ffmpegPath = $FFProbePath.Replace("ffprobe","ffmpeg")

        # Calcola durata del video
        $duration = & $ffmpegPath -i "`"$file`"" 2>&1 |
            Select-String "Duration" |
            ForEach-Object {
                ($_ -match "Duration: (\d+):(\d+):(\d+.\d+)") | Out-Null
                $hours = [int]$matches[1]
                $minutes = [int]$matches[2]
                $seconds = [double]$matches[3]
                return ($hours*3600 + $minutes*60 + $seconds)
            }

        $time = [math]::Round($duration / 2, 2)

        # Cattura frame come PNG direttamente in memoria
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $ffmpegPath
        $startInfo.Arguments = "-ss $time -i `"$file`" -vframes 1 -f image2pipe -vcodec png -"
        $startInfo.RedirectStandardOutput = $true
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true

        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo = $startInfo
        $proc.Start() | Out-Null

        $ms = New-Object System.IO.MemoryStream
        $proc.StandardOutput.BaseStream.CopyTo($ms)
        $proc.WaitForExit()

        $ms.Position = 0
        $bmp = [System.Drawing.Image]::FromStream($ms)
        $picPreview.Image = $bmp
        $txtOutput.AppendText("Preview Video caricata in memoria.`r`n")
    }
    catch {
        $picPreview.Image = $null
        $txtOutput.AppendText("Errore preview video in memoria: $($_.Exception.Message)`r`n")
    }
    return
}




    # -------------------
    # Formato non supportato
    # -------------------
    $picPreview.Image = $null
    $txtOutput.AppendText("Formato non supportato per la Preview.`r`n")
})



# # # # # $dataGrid.Add_SelectionChanged({
    # # # # # if ($dataGrid.SelectedRows.Count -eq 0) {
        # # # # # $picPreview.Image = $null
        # # # # # return
    # # # # # }

    # # # # # $record = $dataGrid.SelectedRows[0].DataBoundItem
    # # # # # if (-not $record) { return }

    # # # # # $file = $record.FilePath
    # # # # # if (-not (Test-Path $file)) { return }

    # # # # # $ext = [System.IO.Path]::GetExtension($file).ToLower()

    # # # # # # Se è un'immagine, carica preview
    # # # # # if ($Global:ImageExtensions -contains $ext) {
		# # # # # $txtOutput.AppendText("Preview Immagine.`r`n")
        # # # # # try {
            # # # # # $img = [System.Drawing.Image]::FromFile($file)

            # # # # # # Clona per evitare lock del file
            # # # # # $clone = $img.Clone()
            # # # # # $img.Dispose()

            # # # # # $picPreview.Image = $clone
        # # # # # }
        # # # # # catch {
            # # # # # $picPreview.Image = $null
        # # # # # }
        # # # # # return
    # # # # # }

# # # # # # Se è un video, prova a estrarre il frame a metà video (in memoria)
# # # # # if ($Global:VideoExtensions -contains $ext -and $FFProbePath) {
    # # # # # try {
        # # # # # $txtOutput.AppendText("Preview Video $file in memoria.`r`n")

        # # # # # # Rilascia immagine precedente
        # # # # # if ($picPreview.Image) {
            # # # # # $picPreview.Image.Dispose()
            # # # # # $picPreview.Image = $null
        # # # # # }

        # # # # # $ffmpegPath = $FFProbePath.Replace("ffprobe","ffmpeg")

        # # # # # # Calcola durata
        # # # # # $duration = & $ffmpegPath -i "`"$file`"" 2>&1 |
            # # # # # Select-String "Duration" |
            # # # # # ForEach-Object {
                # # # # # ($_ -match "Duration: (\d+):(\d+):(\d+.\d+)") | Out-Null
                # # # # # $hours = [int]$matches[1]
                # # # # # $minutes = [int]$matches[2]
                # # # # # $seconds = [double]$matches[3]
                # # # # # return ($hours*3600 + $minutes*60 + $seconds)
            # # # # # }

        # # # # # $time = [math]::Round($duration / 2, 2)

        # # # # # # Ottieni dimensioni del video
        # # # # # $videoInfo = & $ffmpegPath -i "`"$file`"" 2>&1 |
            # # # # # Select-String "Stream.*Video" |
            # # # # # ForEach-Object {
                # # # # # ($_ -match ", (\d+)x(\d+)[ ,]") | Out-Null
                # # # # # return @{ Width = [int]$matches[1]; Height = [int]$matches[2] }
            # # # # # }
        # # # # # $w = $videoInfo.Width
        # # # # # $h = $videoInfo.Height

        # # # # # # ProcessStartInfo per ffmpeg raw RGB
        # # # # # $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        # # # # # $startInfo.FileName = $ffmpegPath
        # # # # # $startInfo.Arguments = "-ss $time -i `"$file`" -vframes 1 -f rawvideo -pix_fmt rgb24 -"
        # # # # # $startInfo.RedirectStandardOutput = $true
        # # # # # $startInfo.UseShellExecute = $false
        # # # # # $startInfo.CreateNoWindow = $true

        # # # # # $proc = New-Object System.Diagnostics.Process
        # # # # # $proc.StartInfo = $startInfo
        # # # # # $proc.Start() | Out-Null

        # # # # # # Leggi raw RGB
        # # # # # $frameSize = $w * $h * 3
        # # # # # $buffer = New-Object byte[] $frameSize
        # # # # # $read = 0
        # # # # # while ($read -lt $frameSize) {
            # # # # # $r = $proc.StandardOutput.BaseStream.Read($buffer, $read, $frameSize - $read)
            # # # # # if ($r -le 0) { break }
            # # # # # $read += $r
        # # # # # }
        # # # # # $proc.WaitForExit()

        # # # # # # Converti raw RGB in Bitmap
        # # # # # $bmp = New-Object System.Drawing.Bitmap $w, $h, ([System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
        # # # # # $rect = [System.Drawing.Rectangle]::FromLTRB(0, 0, $w, $h)
        # # # # # $bmpData = $bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::WriteOnly, $bmp.PixelFormat)
        # # # # # [System.Runtime.InteropServices.Marshal]::Copy($buffer, 0, $bmpData.Scan0, $buffer.Length)
        # # # # # $bmp.UnlockBits($bmpData)

        # # # # # # Assegna al PictureBox
        # # # # # $picPreview.Image = $bmp
        # # # # # $txtOutput.AppendText("Preview Video caricata in memoria.`r`n")
    # # # # # }
    # # # # # catch {
        # # # # # $picPreview.Image = $null
        # # # # # $txtOutput.AppendText("Errore preview video in memoria.`r`n")
    # # # # # }
    # # # # # return
# # # # # }


    # # # # # # # Se è un video, prova a estrarre la frame a metà video (richiede ffmpeg)
    # # # # # # if ($Global:VideoExtensions -contains $ext -and $FFProbePath) {
        # # # # # # try {
			# # # # # # $txtOutput.AppendText("Preview Video $file .`r`n")
			
			
            # # # # # # $tempFrame = Join-Path $env:TEMP "preview_frame.jpg"
			
			
			# # # # # # # Cancella file precedente se esiste
			# # # # # # if (Test-Path $tempFrame) { Remove-Item $tempFrame -Force -ErrorAction SilentlyContinue }

					
			# # # # # # # Cattura il frame a metà del video
			# # # # # # $ffmpegPath = $FFProbePath.Replace("ffprobe","ffmpeg")
			# # # # # # $duration = & $ffmpegPath -i $file 2>&1 | 
				# # # # # # Select-String "Duration" | 
				# # # # # # ForEach-Object {
					# # # # # # ($_ -match "Duration: (\d+):(\d+):(\d+.\d+)") | Out-Null
					# # # # # # $hours = [int]$matches[1]
					# # # # # # $minutes = [int]$matches[2]
					# # # # # # $seconds = [double]$matches[3]
					# # # # # # return ($hours*3600 + $minutes*60 + $seconds)
				# # # # # # }

			# # # # # # $time = [math]::Round($duration / 2, 2)  # frame a metà video

			# # # # # # # Estrarre frame a metà del video
			# # # # # # & $ffmpegPath -y -i "`"$file`"" -ss $time -vframes 1 "`"$tempFrame`"" -hide_banner -loglevel error				

            # # # # # # if (Test-Path $tempFrame) {
				# # # # # # $img = [System.Drawing.Image]::FromFile($tempFrame)
				# # # # # # $clone2 = $img.Clone()
				# # # # # # $img.Dispose()
				# # # # # # Remove-Item $tempFrame -Force 
				# # # # # # $picPreview.Image = $clone2
				# # # # # # $txtOutput.AppendText("Preview Video estratta .`r`n")
            # # # # # # }
        # # # # # # }
        # # # # # # catch { $picPreview.Image = $null }
        # # # # # # return
    # # # # # # }

    # # # # # # Altri formati ? no preview
    # # # # # $picPreview.Image = $null
	# # # # # $txtOutput.AppendText("Formato non supportato per la Preview.`r`n")
# # # # # })








# $Results = New-Object System.ComponentModel.BindingList[FileRecord]
# $Results = New-Object SortableBindingList[FileRecord]
$Results = $null
# $dataGrid.DataSource = $Results
$temporaryList = New-Object 'System.Collections.Generic.List[FileRecord]'



# --- Drag&Drop ---
$txtPath.Add_DragEnter({ if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) { $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy } })
$txtPath.Add_DragDrop({
    $files = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
    if ($files.Length -gt 0 -and (Test-Path $files[0] -PathType Container)) {
        $txtPath.Text = $files[0]
        $btnStart.Enabled = $true
        $btnFixFromLog.Enabled = Test-Path (Join-Path $files[0] "FileAnalysisReport.csv")
		$btnSaveCsv.Enabled = Test-Path (Join-Path $txtPath.Text "FileAnalysisReport.csv")

    }
})

# $btnSelectFolder.Add_Click({
    # $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    # $folderBrowser.ShowNewFolderButton = $false
    # if ($folderBrowser.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
        # $txtPath.Text = $folderBrowser.SelectedPath
        # $btnStart.Enabled = $true
        # $btnFixFromLog.Enabled = Test-Path (Join-Path $folderBrowser.SelectedPath "FileAnalysisReport.csv")
		# $btnSaveCsv.Enabled = Test-Path (Join-Path $txtPath.Text "FileAnalysisReport.csv")

    # }
# })

$btnSelectFolder.Add_Click({
    try {
        $shell = New-Object -ComObject Shell.Application

        # Imposta cartella iniziale
        $initialFolder = if (![string]::IsNullOrWhiteSpace($txtPath.Text) -and (Test-Path $txtPath.Text -PathType Container)) {
            $txtPath.Text
        } else {
            # [Environment]::GetFolderPath("MyDocuments")
			[System.Environment+SpecialFolder]::Desktop
        }

        # Converti la cartella iniziale in FolderItem
        $folderItem = $shell.NameSpace($initialFolder)
        if ($folderItem -ne $null) {
            $folder = $shell.BrowseForFolder(0, "Seleziona cartella", 0, $folderItem.Self)
        } else {
            $folder = $shell.BrowseForFolder(0, "Seleziona cartella", 0, 0)
        }

        if ($folder -ne $null) {
            $path = $folder.Self.Path
            $txtPath.Text = $path
            $btnStart.Enabled = $true
            $btnFixFromLog.Enabled = Test-Path (Join-Path $path "FileAnalysisReport.csv")
            $btnSaveCsv.Enabled = Test-Path (Join-Path $path "FileAnalysisReport.csv")
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Errore durante la selezione della cartella: $($_.Exception.Message)","Errore")
    }
})

# $btnSelectFolder.Add_Click({
    # try {
        # $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog

        # # Non bloccare l'utente: root = Desktop
        # $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::Desktop
        # $folderBrowser.ShowNewFolderButton = $false

        # # Imposta percorso iniziale dalla textbox se valido, altrimenti Documenti
        # if (-not [string]::IsNullOrWhiteSpace($txtPath.Text) -and (Test-Path $txtPath.Text -PathType Container)) {
            # $folderBrowser.SelectedPath = $txtPath.Text
        # } else {
            # $folderBrowser.SelectedPath = [Environment]::GetFolderPath("MyDocuments")
        # }

        # if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            # $txtPath.Text = $folderBrowser.SelectedPath
            # $btnStart.Enabled = $true
            # $btnFixFromLog.Enabled = Test-Path (Join-Path $folderBrowser.SelectedPath "FileAnalysisReport.csv")
            # $btnSaveCsv.Enabled = Test-Path (Join-Path $folderBrowser.SelectedPath "FileAnalysisReport.csv")
        # }
    # }
    # catch {
        # [System.Windows.Forms.MessageBox]::Show("Errore durante la selezione della cartella: $($_.Exception.Message)","Errore")
    # }
# })


$txtPath.Add_TextChanged({
    if ((Test-Path $txtPath.Text -PathType Container)) {
        $btnStart.Enabled = $true
        $btnFixFromLog.Enabled = Test-Path (Join-Path $txtPath.Text "FileAnalysisReport.csv")
		$btnSaveCsv.Enabled = Test-Path (Join-Path $txtPath.Text "FileAnalysisReport.csv")

    } else {
        $btnStart.Enabled = $false
        $btnFixFromLog.Enabled = $false
		$btnSaveCsv.Enabled = $false
    }
})

# -------------------
# Funzione Fix-DateFromLogIncremental con ETC
# -------------------
function Fix-DateFromLogIncremental {
    param(
        [string]$CsvFile,           # FileAnalysisReport.csv
        [string]$FixedCsvFile,      # FileAnalysisReport_fixed.csv
        [string]$ExifToolPath
    )
	$Results = $null
	$dataGrid.DataSource = $null
    # $temporaryList.Clear()
	
    if (-not (Test-Path $CsvFile)) { [System.Windows.Forms.MessageBox]::Show("File CSV non trovato.","Errore"); return }
    if (-not $ExifToolPath) { [System.Windows.Forms.MessageBox]::Show("ExifTool non trovato. Impossibile correggere.","Errore"); return }

    # Carica CSV fixed esistente per incrementale
    $processedFixedFiles = @()
    if (Test-Path $FixedCsvFile) {
        $existingFixed = Import-Csv -Path $FixedCsvFile
        # $processedFixedFiles = $existingFixed.FilePath
		
		# Creo un set di tutti i FilePath già presenti con Status = "Fixed"
		$processedFixedFiles = $existingFixed |
			Where-Object { $_.Status -eq "Fixed" } |
			Select-Object -ExpandProperty FilePath		
		
    }

    # Import CSV originale e filtra i record da correggere non ancora presenti nel fixed CSV
    $items = Import-Csv -Path $CsvFile
	# Se la griglia è vuota e il CSV esiste, carica i dati prima di iniziare i fix
	# if ($Results.Count -eq 0){
	 if ($temporaryList.Count -eq 0){
			foreach ($row in $items) {
			$rec = New-Object FileRecord
			$rec.FilePath       = $row.FilePath
			$rec.FileNameDate   = $row.FileNameDate
			$rec.ExifDate       = $row.ExifDate
			$rec.MediaDate      = $row.MediaDate
			$rec.FileSystemDate = $row.FileSystemDate
			$rec.JsonDate       = $row.JsonDate
			$rec.Status         = $row.Status
			# $Results.Add($rec)
			$temporaryList.Add($rec)
		}
	}
	$resultsDict = @{}
	foreach ($r in $temporaryList) { $resultsDict[$r.FilePath] = $r }


	# Aggiorna in lista temporanea i record già processati in precedenza
	foreach ($item in $existingFixed) {
		if ($resultsDict.ContainsKey($item.FilePath)) {
			$temporaryListItem = $resultsDict[$item.FilePath]
			if ($temporaryListItem) {
				$temporaryListItem.ExifDate       = $item.ExifDate
				$temporaryListItem.FileSystemDate = $item.FileSystemDate
				$temporaryListItem.MediaDate      = $item.MediaDate
				$temporaryListItem.JsonDate       = $item.JsonDate
				$temporaryListItem.Status         = "Already Fixed"
			}
		}
		
		# $gridItem = $Results | Where-Object { $_.FilePath -eq $item.FilePath }
		# if ($gridItem) {
			# $gridItem.ExifDate       = $item.ExifDate
			# $gridItem.FileSystemDate = $item.FileSystemDate
			# $gridItem.MediaDate      = $item.MediaDate
			# $gridItem.JsonDate       = $item.JsonDate
			# $gridItem.Status         = "Already Fixed"
		# }
	}


	# Crea un HashSet dai file già fixati
	$processedFixedFilesSet = [System.Collections.Generic.HashSet[string]]::new()
	foreach ($f in $processedFixedFiles) { $processedFixedFilesSet.Add($f) }

	# Filtra i file da fixare usando HashSet
	$toFix = $items | Where-Object { 
		$_.Status -eq "Mismatch" -and -not $processedFixedFilesSet.Contains($_.FilePath)
	}

	# $dataGrid.DataSource = $Results
	# $dataGrid.Refresh()
	# [System.Windows.Forms.Application]::DoEvents()

	# $toFix = $items | Where-Object { $_.Status -eq "Mismatch" -and -not ($processedFixedFiles -contains $_.FilePath) }
	
    if ($toFix.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Nessun file da correggere o tutti già presenti nel CSV fixed.","OK"); return }

    # Progress bar
    $global:progressBar.Minimum = 0
    $global:progressBar.Maximum = $toFix.Count
    $global:progressBar.Value = 0
    $counter = 0
    $startTime = Get-Date
    $lblETC.Text = "ETC stimato: --:--:--"

    $updateInterval = 2  # ETC aggiornamento ogni 5 file
	# $dataGrid.SuspendLayout()      # blocca aggiornamenti GUI
    foreach ($item in $toFix) {
		if ($global:StopRequested) {
			$txtOutput.AppendText("Correzione da log interrotta dall'utente.`r`n")
			break
		}		
        $counter++
        $global:progressBar.Value = [Math]::Min($counter, $global:progressBar.Maximum)
        if ($counter % $updateInterval -eq 0 -or $counter -eq $toFix.Count) {
            $elapsed = (Get-Date) - $startTime
            $avgPerFile = $elapsed.TotalSeconds / $counter
            $remainingFiles = $toFix.Count - $counter
            $etcSeconds = [Math]::Round($avgPerFile * $remainingFiles)
            $etcTime = [TimeSpan]::FromSeconds($etcSeconds)
            $lblETC.Text = "ETC stimato: $($etcTime.ToString('hh\:mm\:ss'))"
#            $txtOutput.AppendText("Fix $counter/$($toFix.Count) - ETC stimato: $($etcTime.ToString('hh\:mm\:ss'))`r`n")
            # $dataGrid.Refresh()
			[System.Windows.Forms.Application]::DoEvents()
        }

        $file = $item.FilePath
        $item.Status = "Fixed"
		if (-not (Test-Path $file)) { $global:txtOutput.AppendText("File non trovato: $file`r`n"); continue }
		# Se non esistono né JsonDate né FileNameDate ? impossibile correggere
		if (
			(-not $item.JsonDate -or $item.JsonDate.Trim() -eq "") -and
			(-not $item.FileNameDate -or $item.FileNameDate.Trim() -eq "")
		) {
			$global:txtOutput.AppendText("Nessuna data disponibile per $file (né JsonDate né FileNameDate)`r`n")
			$item.Status = "Error"
			continue
		}

		# Scegli la data migliore (JsonDate ? FileNameDate)
		$sourceDate = $null

		if ($item.JsonDate -and $item.JsonDate.Trim() -ne "") {
			$sourceDate = $item.JsonDate
		} else {
			$sourceDate = $item.FileNameDate
		}

		# Parse della data scelta
		try { 
			$newDate = [datetime]::Parse($sourceDate)
		}
		catch { 
			try { 
				$newDate = [datetime]::ParseExact($sourceDate,"yyyy-MM-dd",$null)
			} catch { 
				$newDate = $null
			} 
		}


        if (-not $newDate) { $global:txtOutput.AppendText("Impossibile parse della data per $file : $($item.FileNameDate)`r`n"); continue }

        $dateString = $newDate.ToString("yyyy:MM:dd HH:mm:ss")
        # $global:txtOutput.AppendText("Correggo EXIF + filesystem => $file con data $dateString`r`n")
		# try {
			# & $ExifToolPath `
				# -charset "filename=cp1252" `
				# -overwrite_original `
				# "-DateTimeOriginal=$dateString" `
				# "-CreateDate=$dateString" `
				# "-ModifyDate=$dateString" `
				# "-MediaCreateDate=$dateString" `
				# "-TrackCreateDate=$dateString" `
				# "`"$file`"" | Out-Null
			# if ($LASTEXITCODE -ne 0) {
					# throw "ExifTool ha restituito codice errore $LASTEXITCODE"
					# # $item.Status = "Error"
			# }
		# }
		# catch {
			# $global:txtOutput.AppendText("Exiftool ha fallito per $item `r`n")
			# # $item.Status = "Error"
		# }
		
# Carica ExifLibrary solo una volta
		if (-not ("ExifLibrary.ImageFile" -as [type])) {


			# -----------------------------
			# Caricamento sicuro di ExifLibrary.dll
			# -----------------------------
			# Percorso della DLL relativa alla cartella dello script
			$dllPath = Join-Path $PSScriptRoot "ExifLibrary.dll"

			if (-not (Test-Path $dllPath)) { throw "ExifLibrary.dll non trovata in $dllPath" }

			try { Unblock-File -Path $dllPath -ErrorAction SilentlyContinue } catch {}
			try { Add-Type -Path $dllPath } catch { throw "Errore caricamento DLL: $($_.Exception.Message)" }

			# Write-Host "ExifLibrary.dll caricata correttamente."
			$global:txtOutput.AppendText("ExifLibrary.dll caricata correttamente.`r`n")

		}


		$global:txtOutput.AppendText("Correggo EXIF + filesystem => $file con data $dateString`r`n")

		# Controlla estensione
		$ext = [System.IO.Path]::GetExtension($file).ToLower()

		
			if ($ext -in ".jpg", ".jpeg", ".jpe") {
				try {
					# -------------------------
					#   FOTO ? EXIFLIBRARY
					# -------------------------

					# Carica l'immagine JPEG
					$img = [ExifLibrary.ImageFile]::FromFile($file)

					# $newDate deve essere un oggetto [datetime]
					# Imposta i tag EXIF di data/ora
					$img.Properties.Set([ExifLibrary.ExifTag]::DateTimeOriginal, $newDate)
					$img.Properties.Set([ExifLibrary.ExifTag]::DateTimeDigitized, $newDate)
					$img.Properties.Set([ExifLibrary.ExifTag]::DateTime, $newDate)

					# Salva le modifiche
					$img.Save($file)

		        } catch {
					$global:txtOutput.AppendText("Errore EXIFLIBRARY per $item : $($_.Exception.Message)`r`n")
					$item.Status = "Error"
				}
			}
			else{ #if ($ext -notin $Global:NoMetadataFormats) {
				# $global:txtOutput.AppendText("$ext non è in $Global:NoMetadataFormats `r`n")
				# -------------------------
				#   VIDEO / ALTRO ? EXIFTOOL
				# -------------------------
				try {
					# Primo tentativo: scrittura diretta
					& $ExifToolPath `
						-charset "filename=cp1252" `
						-overwrite_original_in_place `
						-m `
						-v1 `
						-api IgnoreMinorErrors=1 `
						"-DateTimeOriginal=$dateString" `
						"-CreateDate=$dateString" `
						"-ModifyDate=$dateString" `
						"-MediaCreateDate=$dateString" `
						"-TrackCreateDate=$dateString" `
						"$file" | Out-Null

					if ($LASTEXITCODE -ne 0) { throw "ExifTool fallito con codice $LASTEXITCODE" }
					$item.ExifDate = $dateString
				}
				catch {
					# Fallback con file temporaneo solo se il primo metodo fallisce
					# Write-Host "Metodo diretto fallito per $file. Uso fallback con copia temporanea..."
					$global:txtOutput.AppendText("Metodo diretto fallito per $file. Uso fallback con copia temporanea...`r`n")

					# Rileva estensione reale
					$realExt = (& $ExifToolPath -s -s -s -FileTypeExtension "$file").Trim()
					if (-not $realExt) { $realExt = "mov" }  # fallback generico

					$tempFile = "$file.temp.$realExt"
					Copy-Item "$file" $tempFile -Force

					# Aggiorna metadati sulla copia
					& $ExifToolPath -overwrite_original `
					  -charset "filename=cp1252" `
					  "-DateTimeOriginal=$dateString" `
					  "-CreateDate=$dateString" `
					  "-ModifyDate=$dateString" `
					  "-MediaCreateDate=$dateString" `
					  "-TrackCreateDate=$dateString" `
					  "$tempFile" | Out-Null

					if ($LASTEXITCODE -ne 0) {
						Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
						$global:txtOutput.AppendText("ExifTool fallito anche sul file temporaneo con codice $LASTEXITCODE `r`n")
						# throw "ExifTool fallito anche sul file temporaneo con codice $LASTEXITCODE"
#**********
# # --- Step 3: Fallback con FFmpeg ---
    # $ok = $true   # stato di esecuzione per il singolo file

    # # Controllo esistenza file
    # if (-not (Test-Path $file)) {
        # $global:txtOutput.AppendText("SKIP: File non trovato: $file`r`n")
        # $ok = $false
    # }

    # if ($ok) {

        # $ext        = [System.IO.Path]::GetExtension($file)
        # $tempFile   = "$file.temp$ext"
        # $backupFile = "$file.bak$ext"

        # # Rimuove eventuale file temp precedente
        # if (Test-Path $tempFile) {
            # Remove-Item -Force $tempFile
        # }

        # # Durata originale
        # $origDur = ffprobe -v error -show_entries format=duration -of default=nk=1:nw=1 "$file"

        # if (-not $origDur) {
            # $global:txtOutput.AppendText("ERRORE: durata originale non letta ? $file`r`n")
            # $ok = $false
        # }
    # }

    # if ($ok) {
		
		
		
        # # FFmpeg scrittura metadati
			# & $FFProbePath.Replace("ffprobe","ffmpeg") `
			# -v error -y -i "$file" `
		    # -metadata ICRD="$dateString" `
            # -metadata creation_time="$dateString" `
            # -metadata date="$dateString" `
            # -metadata com.apple.quicktime.creationdate="$dateString" `
            # -codec copy `
            # "$tempFile"

        # # Controllo temp generato
        # if (-not (Test-Path $tempFile)) {
            # $global:txtOutput.AppendText("ERRORE: temp non generato ? $file`r`n")
            # $ok = $false
        # }
    # }

    # if ($ok) {

        # # Durata del file temporaneo
        # $tempDur = ffprobe -v error -show_entries format=duration -of default=nk=1:nw=1 "$tempFile"

        # # Integrità
        # if (-not $tempDur -or [double]$tempDur -lt ([double]$origDur * 0.95)) {
            # $global:txtOutput.AppendText("ERRORE: temp corrotto ? $file (eliminato)`r`n")
            # Remove-Item -Force $tempFile
            # $ok = $false
        # }
    # }

    # if ($ok) {

        # # Backup solo se non esiste
        # if (-not (Test-Path $backupFile)) {
            # Copy-Item -Force "$file" "$backupFile"
        # }

        # # Se il backup non è stato creato, NON procedere
        # if (-not (Test-Path $backupFile)) {
            # $global:txtOutput.AppendText("ERRORE: impossibile creare backup ? $file`r`n")
            # $ok = $false
        # }
    # }

    # if ($ok) {

        # # Sovrascrizione finale con file valido
        # Move-Item -Force "$tempFile" "$file"
        # $global:txtOutput.AppendText("OK: aggiornato ? $file`r`n")

    # }
#**********
					}
					else {
						# Sostituisci l'originale con la copia temporanea
						Move-Item -Force $tempFile $file
						$item.ExifDate = $dateString
						$item.MediaDate = $dateString 
						# Write-Host "Fallback completato con successo per $file"
						$global:txtOutput.AppendText("Fallback completato con successo per $file`r`n")
					}
				}
	
			}
        try {
            (Get-Item $file).CreationTime = $newDate
            (Get-Item $file).LastWriteTime = $newDate
			$item.FileSystemDate = $dateString
		} catch {
            $global:txtOutput.AppendText("Errore aggiornamento filesystem per $item : $($_.Exception.Message)`r`n")
			$item.Status = "Error"
        }


		# --- AGGIORNA ANCHE IL RECORD NELLA lista temporanea --- 
		if ($resultsDict.ContainsKey($item.FilePath)) {
			$temporaryListItem = $resultsDict[$item.FilePath]
			if ($temporaryListItem) {
				$temporaryListItem.ExifDate       = $item.ExifDate
				$temporaryListItem.FileSystemDate = $item.FileSystemDate
				$temporaryListItem.MediaDate      = $item.MediaDate
				$temporaryListItem.JsonDate       = $item.JsonDate
				$temporaryListItem.Status         = $item.Status
			}
		}


			



		# $gridItem = $Results | Where-Object { $_.FilePath -eq $item.FilePath }
		# if ($gridItem) {
			# $gridItem.ExifDate       = $item.ExifDate
			# $gridItem.FileSystemDate = $item.FileSystemDate
			# $gridItem.MediaDate      = $item.MediaDate
			# $gridItem.JsonDate       = $item.JsonDate
			# $gridItem.Status         = $item.Status
		# }

		# $dataGrid.Refresh()
		# [System.Windows.Forms.Application]::DoEvents()
		

        # Salvataggio incrementale nel CSV fixed
        $obj = [PSCustomObject]@{
            FilePath       = $item.FilePath
            FileNameDate   = $item.FileNameDate
            ExifDate       = $item.ExifDate
            MediaDate      = $item.MediaDate
            FileSystemDate = $item.FileSystemDate
			JsonDate       = $item.JsonDate
            Status         = $item.Status
        }
        $obj | Export-Csv -Path $FixedCsvFile -NoTypeInformation -Encoding UTF8 -Append
    }
	
	# $dataGrid.ResumeLayout()       # riattiva aggiornamenti GUI
	$txtOutput.AppendText("Preparo la lista `r`n")

	$Results = New-Object 'SortableBindingList[FileRecord]'
	foreach ($tItem in $temporaryList) {
		$Results.Add($tItem)
	}

	# $Results = New-Object SortableBindingList[FileRecord] ($temporaryList)
	$dataGrid.DataSource = $Results
	$dataGrid.Refresh()
	$elapsed = (Get-Date) - $startTime
	$txtOutput.AppendText("Correzioni completate in $elapsed .`r`n")
    [System.Windows.Forms.MessageBox]::Show("Correzioni completate in $elapsed .","OK")
    $txtOutput.AppendText("Nuovo CSV incrementale salvato come: $FixedCsvFile`r`n")
}



# -------------------
# Pulsante Start Analisi
# -------------------
$btnStart.Add_Click({
    # $btnStart.Enabled = $false
	
	$global:StopRequested = $false
	$btnStop.Enabled = $true
	$btnStart.Enabled = $false
	$btnFixFromLog.Enabled = $false
	$dataGrid.DataSource = $null
	
    $path = $txtPath.Text
	
	Initialize-JsonCache -RootPath $path
	$txtOutput.AppendText("Cache JSON inizializzata in memoria.`r`n")

    # $ExifToolPath = $null
    # try { $ExifToolPath = (Get-Command "exiftool.exe" -ErrorAction Stop).Source } catch { }
    # $FFProbePath = $null
    # try { $FFProbePath = (Get-Command "ffprobe.exe" -ErrorAction Stop).Source } catch { }

    # $Global:ImageExtensions = @(".jpg",".jpeg",".png",".heic",".webp",".dng",".nef",".cr2",".cr3",".arw",".rw2",".gif",".tif",".tiff",".bmp")
    # $Global:VideoExtensions = @(".mp4",".mov",".m4v",".avi",".mpg",".mpeg",".mkv",".mts",".m2ts",".wmv",".mpe", ".m2p", ".m2v", ".mp2")

# # Formati senza metadata interni (MPEG Program Stream + Elementary Streams)
	# $Global:NoMetadataFormats = @(".mpg", ".mpeg", ".mpe", ".m2p", ".m2v", ".mp2", ".avi")


    $Files = Get-ChildItem -Path $path -File -Recurse | 
        Where-Object { $Global:ImageExtensions -contains $_.Extension.ToLower() -or $Global:VideoExtensions -contains $_.Extension.ToLower() }

    $progressBar.Maximum = $Files.Count
    $progress = 0
    $startTime = Get-Date
    $txtOutput.Clear()
    $Results = $null
	$temporaryList.Clear()
    $lblETC.Text = "ETC stimato: --:--:--"

    $CsvOutput = Join-Path $path "FileAnalysisReport.csv"
    $ProcessedFiles = @()
    if (Test-Path $CsvOutput) {
        $existing = Import-Csv -Path $CsvOutput
        $ProcessedFiles = $existing.FilePath
        foreach ($i in $existing) {
            $obj = New-Object FileRecord
            $obj.FilePath       = $i.FilePath
            $obj.FileNameDate   = $i.FileNameDate
            $obj.ExifDate       = $i.ExifDate
            $obj.MediaDate      = $i.MediaDate
            $obj.FileSystemDate = $i.FileSystemDate
			$obj.JsonDate       = $i.JsonDate
            $obj.Status         = $i.Status
            # [void]($Results.Add($obj))
			[void]($temporaryList.Add($obj))
        }
    }

	# Crea un HashSet dai file processati
	$processedFilesSet = [System.Collections.Generic.HashSet[string]]::new()
	foreach ($f in $ProcessedFiles) { $processedFilesSet.Add($f) }


    $updateInterval = 50  # ETC aggiornamento ogni 5 file

	# $dataGrid.SuspendLayout()      # blocca aggiornamenti GUI

    foreach ($file in $Files) {
		if ($global:StopRequested) {
			$txtOutput.AppendText("Elaborazione analisi interrotta dall'utente.`r`n")
			break
		}

        $progress++
        $global:progressBar.Value = [Math]::Min($progress, $global:progressBar.Maximum)

        # ETC aggiornato ogni $updateInterval file
        if ($progress % $updateInterval -eq 0 -or $progress -eq $Files.Count) {
            $elapsed = (Get-Date) - $startTime
            $avgPerFile = $elapsed.TotalSeconds / $progress
            $remainingFiles = $Files.Count - $progress
            $etcSeconds = [Math]::Round($avgPerFile * $remainingFiles)
            $etcTime = [TimeSpan]::FromSeconds($etcSeconds)
            $lblETC.Text = "ETC stimato: $($etcTime.ToString('hh\:mm\:ss'))"
#            $txtOutput.AppendText("File $progress/$($Files.Count) - ETC stimato: $($etcTime.ToString('hh\:mm\:ss'))`r`n")
            [System.Windows.Forms.Application]::DoEvents()
        }

        # if ($ProcessedFiles -contains $file.FullName) { continue }
        if ($processedFilesSet.Contains($file.FullName)) { continue }

        # Elaborazione file (logica esistente)
        $FileNameDate = Get-DateFromFilename $file.Name
        $ReferenceDate = $null
        $ExifDate = $null
        $VideoDate = $null
        $JsonDate = Get-GoogleJsonDate $file.FullName

        if ($Global:ImageExtensions -contains $file.Extension.ToLower()) {
            $ExifDate = Get-ExifDate $file.FullName $ExifToolPath
            $ReferenceDate = if ($ExifDate) { $ExifDate } else { $file.CreationTime }
        }
        elseif ($Global:VideoExtensions -contains $file.Extension.ToLower()) {
            $VideoDate = Get-VideoMediaDate $file.FullName $FFProbePath
            $ReferenceDate = if ($VideoDate) { $VideoDate } else { $file.CreationTime }
        }

        $CompareDate = if ($JsonDate) { $JsonDate } elseif ($FileNameDate) { $FileNameDate } else { $null }

        $Status = "OK"
        if ($CompareDate -and $ReferenceDate) { if ($CompareDate.Date -ne $ReferenceDate.Date) { $Status = "Mismatch" } }
        elseif ($CompareDate -and -not $ReferenceDate) { $Status = "NoReferenceDate" }
        else { $Status = "NoDateInFilename-JSON" }

        $obj = New-Object FileRecord
        $obj.FilePath       = $file.FullName
        $obj.FileNameDate   = if ($FileNameDate) { $FileNameDate.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
        $obj.ExifDate       = if ($ExifDate) { $ExifDate.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
        $obj.MediaDate      = if ($VideoDate) { $VideoDate.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
        $obj.FileSystemDate = $file.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
        $obj.JsonDate       = if ($JsonDate) { $JsonDate.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
        $obj.Status         = $Status
        # [void]($Results.Add($obj))
		[void]($temporaryList.Add($obj))

        $obj | Select-Object FilePath,FileNameDate,ExifDate,MediaDate,FileSystemDate,JsonDate,Status |
            Export-Csv -Path $CsvOutput -NoTypeInformation -Encoding UTF8 -Append
    }
	# $dataGrid.ResumeLayout()       # riattiva aggiornamenti GUI
	# $dataGrid.Refresh()
	$txtOutput.AppendText("Preparo la lista `r`n")
	
	$Results = New-Object 'SortableBindingList[FileRecord]'
	foreach ($tItem in $temporaryList) {
		$Results.Add($tItem)
	}

	# $Results = New-Object SortableBindingList[FileRecord] ($temporaryList)
	$dataGrid.DataSource = $Results
	$dataGrid.Refresh()


	$btnStop.Enabled = $false
	$btnStart.Enabled = $true
	$btnFixFromLog.Enabled = Test-Path $CsvOutput  # o $FixedCsvFile nel caso di fix
	$btnSaveCsv.Enabled = Test-Path (Join-Path $txtPath.Text "FileAnalysisReport.csv")
	$elapsed = (Get-Date) - $startTime
	$txtOutput.AppendText("Analisi completata in $elapsed .`r`n")
    [System.Windows.Forms.MessageBox]::Show("Analisi completata in $elapsed . Risultati salvati in $CsvOutput", "Completato")
    # $btnStart.Enabled = $true
    # $btnFixFromLog.Enabled = Test-Path $CsvOutput
})


# -------------------
# Pulsante Fix da Log
# -------------------
$btnFixFromLog.Add_Click({
	$global:StopRequested = $false
	$btnStop.Enabled = $true
	$btnStart.Enabled = $false
	$btnFixFromLog.Enabled = $false

    $path = $txtPath.Text
    $CsvFile = Join-Path $path "FileAnalysisReport.csv"
    $FixedCsvFile = Join-Path $path "FileAnalysisReport_fixed.csv"

    try { $ExifToolPath = (Get-Command "exiftool.exe" -ErrorAction Stop).Source } 
    catch { [System.Windows.Forms.MessageBox]::Show("ExifTool non trovato nella PATH.","Errore"); return }

    $txtOutput.AppendText("Avvio correzione dalla lista log...`r`n")
    Fix-DateFromLogIncremental -CsvFile $CsvFile -FixedCsvFile $FixedCsvFile -ExifToolPath $ExifToolPath
	$btnStop.Enabled = $false
	$btnStart.Enabled = $true
	$btnFixFromLog.Enabled = $true
	$btnSaveCsv.Enabled = Test-Path (Join-Path $txtPath.Text "FileAnalysisReport.csv")
})
$btnSaveCsv.Add_Click({
    $path = $txtPath.Text
    $CsvFile = Join-Path $path "FileAnalysisReport.csv"

    if (-not (Test-Path $CsvFile)) {
        [System.Windows.Forms.MessageBox]::Show("Nessun CSV da salvare trovato.","Errore")
        return
    }

    try {
        # Esporta i dati presenti nel DataGridView
        $Results | Select-Object FilePath,FileNameDate,ExifDate,MediaDate,FileSystemDate,JsonDate,Status |
            Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8

        [System.Windows.Forms.MessageBox]::Show("Modifiche salvate nel CSV.","OK")
        $txtOutput.AppendText("CSV aggiornato: $CsvFile`r`n")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Errore durante il salvataggio del CSV: $($_.Exception.Message)","Errore")
    }
})

[void]$form.ShowDialog()
