param(
  [string]$Image = "",
  [string]$Folder = "",
  [string]$Out = "",
  [string]$TextFile = ""
)
# Typenschild-Erkennung mit der in Windows eingebauten Texterkennung (offline, kostenlos).
# -Image <Foto> -Out <Datei>   : liest ein Foto, schreibt SCHLUESSEL=Wert-Zeilen, danach #ROH und den erkannten Text
# -Folder <Ordner> -Out <Datei>: Testlauf ueber alle Fotos im Ordner
# -TextFile <Datei>            : Text wurde schon erkannt (z. B. von Google) - nur noch auswerten
$ErrorActionPreference = "Stop"
$log = $Out + ".log"
function Log($m) { Add-Content -LiteralPath $log -Value ((Get-Date -Format "HH:mm:ss") + " " + $m) -Encoding UTF8 }
Log "start"
trap { Log ("FEHLER: " + $_.Exception.Message + " @ " + $_.InvocationInfo.ScriptLineNumber); break }

# Windows-Texterkennung nur laden, wenn das Foto selbst gelesen wird (nicht bei -TextFile)
if (-not $TextFile) {
  Add-Type -AssemblyName System.Drawing
  Add-Type -AssemblyName System.Runtime.WindowsRuntime
  $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
  $null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
  $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType = WindowsRuntime]
  $null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics, ContentType = WindowsRuntime]
  $null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]

  $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
      $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
  function Await($op, [Type]$t) {
    $task = $asTaskGeneric.MakeGenericMethod($t).Invoke($null, @($op))
    $null = $task.Wait(-1)
    $task.Result
  }
  $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
  if ($null -eq $engine) { throw "Keine Windows-Texterkennung verfuegbar (Sprachpaket fehlt)." }
  Log ("engine: " + $engine.RecognizerLanguage.LanguageTag)
}

function Ocr-Bitmap([System.Drawing.Bitmap]$bmp) {
  $tmp = Join-Path $env:TEMP ("tsocr_" + [guid]::NewGuid().ToString("N") + ".bmp")
  $bmp.Save($tmp, [System.Drawing.Imaging.ImageFormat]::Bmp)
  try {
    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($tmp)) ([Windows.Storage.StorageFile])
    $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    try {
      $dec = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
      $sbmp = Await ($dec.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
      return (Await ($engine.RecognizeAsync($sbmp)) ([Windows.Media.Ocr.OcrResult]))
    } finally { $stream.Dispose() }
  } finally { Remove-Item $tmp -ErrorAction SilentlyContinue }
}

# Woerter anhand ihrer Position wieder zu Zeilen zusammensetzen (Beschriftung + Wert)
function Build-Rows($res) {
  $ws = New-Object System.Collections.ArrayList
  foreach ($ln in $res.Lines) { foreach ($w in $ln.Words) {
    $r = $w.BoundingRect
    [void]$ws.Add([pscustomobject]@{ T = $w.Text; X = $r.X; Y = ($r.Y + $r.Height / 2); H = $r.Height })
  } }
  if ($ws.Count -eq 0) { return ,@() }
  $hs = @($ws | ForEach-Object { $_.H } | Sort-Object)
  $mh = $hs[[int]($hs.Count / 2)]
  $rows = New-Object System.Collections.ArrayList
  foreach ($w in ($ws | Sort-Object Y)) {
    $hit = $null
    foreach ($row in $rows) { if ([Math]::Abs($w.Y - $row.Y) -lt 0.55 * $mh) { $hit = $row; break } }
    if ($hit) { [void]$hit.W.Add($w); $hit.Y = ($hit.Y * ($hit.W.Count - 1) + $w.Y) / $hit.W.Count }
    else { $nw = New-Object System.Collections.ArrayList; [void]$nw.Add($w); [void]$rows.Add([pscustomobject]@{ Y = $w.Y; W = $nw }) }
  }
  return ,@($rows | Sort-Object Y | ForEach-Object { (@($_.W | Sort-Object X | ForEach-Object { $_.T }) -join ' ') })
}

function Score($res) {
  # Punkte fuer echte Typenschild-Woerter, damit kopfueber gelesener Buchstabensalat verliert
  $t = (@($res.Lines | ForEach-Object { $_.Text }) -join "`n")
  $s = 0
  $s += 4 * ([regex]::Matches($t, '(?i)\b(type|typ|input|output|model|serial|ser|volts?|amps?|date|hz|nr|no|batter\w*|charger|ladeger\w*|made|ger(ä|a)te)\b')).Count
  $s += 3 * ([regex]::Matches($t, '\d\s*(V|A|Hz|kW|kVA|Ah)\b')).Count
  foreach ($l in $res.Lines) { foreach ($w in $l.Words) { if ($w.Text -match '^[A-Za-z0-9/.,:\-]{3,}$' -and $w.Text -match '[aeiouAEIOU0-9]') { $s++ } } }
  return $s
}

function Scale-Bitmap([System.Drawing.Image]$src, [int]$target, [bool]$contrast) {
  $long = [Math]::Max($src.Width, $src.Height)
  $f = [double]$target / $long
  $maxDim = [Windows.Media.Ocr.OcrEngine]::MaxImageDimension
  if ($long * $f -gt $maxDim) { $f = $maxDim / $long }
  $w = [int]($src.Width * $f); $h = [int]($src.Height * $f)
  $bmp = New-Object System.Drawing.Bitmap $w, $h, ([System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.Clear([System.Drawing.Color]::White)
  $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  if ($contrast) {
    $c = 1.3; $t = 0.5 * (1 - $c)
    $cm = New-Object System.Drawing.Imaging.ColorMatrix
    $cm.Matrix00 = 0.299 * $c; $cm.Matrix01 = 0.299 * $c; $cm.Matrix02 = 0.299 * $c
    $cm.Matrix10 = 0.587 * $c; $cm.Matrix11 = 0.587 * $c; $cm.Matrix12 = 0.587 * $c
    $cm.Matrix20 = 0.114 * $c; $cm.Matrix21 = 0.114 * $c; $cm.Matrix22 = 0.114 * $c
    $cm.Matrix33 = 1; $cm.Matrix40 = $t; $cm.Matrix41 = $t; $cm.Matrix42 = $t; $cm.Matrix44 = 1
    $ia = New-Object System.Drawing.Imaging.ImageAttributes
    $ia.SetColorMatrix($cm)
    $g.DrawImage($src, (New-Object System.Drawing.Rectangle 0, 0, $w, $h), 0, 0, $src.Width, $src.Height, [System.Drawing.GraphicsUnit]::Pixel, $ia)
  } else {
    $g.DrawImage($src, 0, 0, $w, $h)
  }
  $g.Dispose()
  return $bmp
}

function Read-Plate([string]$path) {
  $src = [System.Drawing.Image]::FromFile($path)
  # EXIF-Drehung des Handys beachten
  try {
    $o = $src.GetPropertyItem(0x0112).Value[0]
    switch ($o) { 3 { $src.RotateFlip('Rotate180FlipNone') } 6 { $src.RotateFlip('Rotate90FlipNone') } 8 { $src.RotateFlip('Rotate270FlipNone') } }
  } catch {}
  $small = Scale-Bitmap $src 1600 $false
  $bestScore = -1; $bestRot = 0; $scores = @()
  foreach ($rot in 0, 90, 180, 270) {
    $copy = $small.Clone()
    switch ($rot) { 90 { $copy.RotateFlip('Rotate90FlipNone') } 180 { $copy.RotateFlip('Rotate180FlipNone') } 270 { $copy.RotateFlip('Rotate270FlipNone') } }
    $res = Ocr-Bitmap $copy; $copy.Dispose()
    $sc = Score $res; $scores += "$rot=$sc"
    if ($sc -gt $bestScore) { $bestScore = $sc; $bestRot = $rot }
  }
  $small.Dispose()
  # Foto in die beste Richtung drehen
  $srcR = New-Object System.Drawing.Bitmap $src
  $src.Dispose()
  switch ($bestRot) { 90 { $srcR.RotateFlip('Rotate90FlipNone') } 180 { $srcR.RotateFlip('Rotate180FlipNone') } 270 { $srcR.RotateFlip('Rotate270FlipNone') } }
  $a = Scale-Bitmap $srcR 2000 $false
  $fa = [double]$srcR.Width / $a.Width
  $bestRes = Ocr-Bitmap $a; $a.Dispose()
  # zweite Variante: groesser und in Graustufen mit etwas mehr Kontrast
  $b = Scale-Bitmap $srcR 2500 $true
  $resB = Ocr-Bitmap $b; $b.Dispose()
  # dritte/vierte Variante: nur der Bereich mit Text (das Schild), stark vergroessert
  $resC = $null; $resC2 = $null
  $xs = New-Object System.Collections.ArrayList; $ys = New-Object System.Collections.ArrayList
  foreach ($ln in $bestRes.Lines) { foreach ($w in $ln.Words) { $r = $w.BoundingRect; [void]$xs.Add($r.X); [void]$xs.Add($r.X + $r.Width); [void]$ys.Add($r.Y); [void]$ys.Add($r.Y + $r.Height) } }
  if ($xs.Count -ge 6) {
    $x0 = ($xs | Measure-Object -Minimum).Minimum; $x1 = ($xs | Measure-Object -Maximum).Maximum
    $y0 = ($ys | Measure-Object -Minimum).Minimum; $y1 = ($ys | Measure-Object -Maximum).Maximum
    $pad = 0.08 * [Math]::Max($x1 - $x0, $y1 - $y0)
    $cx0 = [int][Math]::Max(0, ($x0 - $pad) * $fa); $cy0 = [int][Math]::Max(0, ($y0 - $pad) * $fa)
    $cx1 = [int][Math]::Min($srcR.Width, ($x1 + $pad) * $fa); $cy1 = [int][Math]::Min($srcR.Height, ($y1 + $pad) * $fa)
    $cw = $cx1 - $cx0; $chh = $cy1 - $cy0
    if ($cw -gt 50 -and $chh -gt 50 -and ($cw * $chh) -lt (0.85 * $srcR.Width * $srcR.Height)) {
      $crop = $srcR.Clone((New-Object System.Drawing.Rectangle $cx0, $cy0, $cw, $chh), [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
      $c = Scale-Bitmap $crop 2600 $true; $resC = Ocr-Bitmap $c; $c.Dispose()
      $c2 = Scale-Bitmap $crop 2600 $false; $resC2 = Ocr-Bitmap $c2; $c2.Dispose()
      $crop.Dispose()
      Log ("  Ausschnitt " + $cw + "x" + $chh + " von " + $srcR.Width + "x" + $srcR.Height)
    }
  }
  $srcR.Dispose()
  $rowsA = Build-Rows $bestRes; $rowsB = Build-Rows $resB
  $rowsC = @(); $rowsC2 = @(); $linesC = @()
  if ($resC) { $rowsC = Build-Rows $resC; $rowsC2 = Build-Rows $resC2; $linesC = @($resC.Lines | ForEach-Object { $_.Text }) + @($resC2.Lines | ForEach-Object { $_.Text }) }
  $linesA = @($bestRes.Lines | ForEach-Object { $_.Text }); $linesB = @($resB.Lines | ForEach-Object { $_.Text })
  return @{ Rot = $bestRot; Scores = ($scores -join ' '); A = (@($rowsA) + @($linesA)); B = (@($rowsB) + @($linesB)); RowsA = $rowsA; RowsB = $rowsB; RowsC = $rowsC; RowsC2 = $rowsC2; LinesC = $linesC
           All = (@($rowsA) + @($rowsB) + @($rowsC) + @($rowsC2) + @($linesA) + @($linesB) + @($linesC)) }
}

# ---------------- Werte aus dem Text holen ----------------
function M1([string]$s, [string]$p, [bool]$ci = $true) {
  $o = [System.Text.RegularExpressions.RegexOptions]::None
  if ($ci) { $o = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }
  return [regex]::Match($s, $p, $o)
}
function Num([string]$v) { return ($v -replace ',', '.') }
function Fix-Token([string]$l) {
  $l = [regex]::Replace($l, '\b(\d{1,2})\s?S\s?A\b', '$1.5A')
  # typische Lesefehler in Zahlen: "IOA" -> "10A", "23OV" -> "230V"
  return [regex]::Replace($l, '\b([0-9IOloS]{2,4})([AV])\b', {
    param($m)
    $n = $m.Groups[1].Value
    if ($n -notmatch '[0-9O]') { return $m.Value }
    ($n -replace '[Il]', '1' -replace '[Oo]', '0' -replace 'S', '5') + $m.Groups[2].Value
  })
}
function Norm-Volt([string]$v) {
  $v = ($v -replace '\s', '')
  if ($v -match '^(\d{2,3})-(\d{3})$') {
    $lo = [int]$Matches[1]; $hi = [int]$Matches[2]
    if ($lo -le 230 -and $hi -ge 230 -and $hi -lt 300) { return '230' }
    if ($lo -le 400 -and $hi -ge 400) { return '400' }
    return $v
  }
  if ($v -match '^2[23]\d$') { return '230' }
  if ($v -match '^(38\d|39\d|40\d|41\d)$') { return '400' }
  return $v
}
$Brands = @(
  @('micro\s*power','Micropower'), @('\bIEB\b|ELEKTRONIK\s+Brilon|\bBrilon\b','IEB'), @('AXIMA','AXIMA'),
  @('S\.?\s*p\.?\s*E\.?\s*(ELETTRONICA|INDUSTRIALE)|CREVALCORE|\bS\.P\.E\b','SPE'), @('Exide','Exide'), @('Nil\s*f\s*isk|vNiIfisk','Nilfisk'),
  @('\biaE\b|powertron|Untergr','iaE'), @('Fronius','Fronius'), @('Hoppecke','Hoppecke'), @('EnerSys','EnerSys'),
  @('Jungheinrich','Jungheinrich'), @('\bLinde\b','Linde'), @('\bStill\b','Still'), @('\bBHT\b','BHT'), @('Deltec','Deltec'),
  @('Zivan','Zivan'), @('Mentzer','Mentzer'), @('Benning','Benning'), @('Delta-?Q','Delta-Q'), @('CTEK','CTEK'),
  @('Hawker','Hawker'), @('\bGNB\b','GNB'), @('K(ä|a)rcher','Kärcher'), @('CHARIS|\bIES\b','IES'), @('\bFerro\b','Ferro')
)
$StopLabels = '\s*[•·|].*$|\s+(?:Before|Charger|read|For|S/?N|SN|SER(?:IAL)?|P/N|CP/N|DATE|DAT|Datum|INPUT|IN|OUT|OUTPUT|Ue|Ua|Batterie\S*|CE)\b.*$'

function Parse-Plate($lines) {
  $lines = @($lines | Where-Object { $_ -and $_.Trim() } | ForEach-Object { Fix-Token $_ })
  $t = $lines -join "`n"
  $d = [ordered]@{}
  foreach ($b in $Brands) { if ((M1 $t $b[0]).Success) { $d['HERSTELLER'] = $b[1]; break } }

  # ---- Typ ----
  $ch = M1 $t 'CHARIS\s*HF\s*(\d)'
  if ($ch.Success) { $d['TYP'] = 'Charis HF' + $ch.Groups[1].Value }
  if (-not $d.Contains('TYP')) {
    foreach ($l in $lines) {
      if ((M1 $l 'Batter|Bat\.|Setting|Gewicht|Weight').Success) { continue }
      $m = M1 $l '(?:^|\s)(?:TYPE?|Mod[eo]ll?)\s*[:.]?\s+([A-Za-z0-9].{2,})$'
      if ($m.Success) {
        $v = ([regex]::Replace($m.Groups[1].Value, $StopLabels, '', 'IgnoreCase')).Trim(' ', ':', '.')
        if ($v.Length -ge 3 -and ($v -match '\d' -or $v -cmatch '^[A-Z]{4}')) { $d['TYP'] = $v; break }
      }
    }
  }
  if (-not $d.Contains('TYP')) {
    foreach ($l in $lines) {
      $m = M1 $l '(powertron\s+[A-Z]{1,2})\b'
      if (-not $m.Success) { $m = M1 $l '\bM[.,]\s*(C[A-Z0-9]+(?:\s*NA)?\s+\d{2}-\d{1,2})' }
      if (-not $m.Success) { $m = M1 $l '\b(SP\d{2}S\d{3}S-?\d{3,5})' }
      if ($m.Success) { $d['TYP'] = $m.Groups[1].Value.Trim(); break }
    }
  }

  # ---- Fabrik-/Seriennummer ----
  foreach ($l in $lines) {
    $m = M1 $l 'Ger(ä|a)te-?\s*Nr\.?\s*[:.]?\s*(\d[\d ]{4,}\d)'
    if ($m.Success) { $d['FABRNR'] = ($m.Groups[2].Value -replace '\s+', ' ').Trim(); break }
    $m = M1 $l '(?:S/N|\bSN\b|Se[rn][il]?(?:al)?\.?\s*(?:No|Nr|#)\.?|Sen\.?\s*No\.?|Serien-?\s*Nr\.?|N\S?\s*s(é|e)rie|Fabr?\.?\s*-?\s*Nr\.?|\bSER\b\.?)\s*[:.#]?\s*([A-Z0-9][A-Z0-9\-/]{3,})'
    if ($m.Success -and $m.Groups[2].Value -match '\d') { $d['FABRNR'] = $m.Groups[2].Value; break }
  }
  if (-not $d.Contains('FABRNR')) {
    $m = M1 $t '\*\s*(\d{5,14})\s*\*'
    if ($m.Success) { $d['FABRNR'] = $m.Groups[1].Value }
  }
  if (-not $d.Contains('FABRNR')) {
    $m = M1 $t '\b([A-Z]{3}-SW\d{6,})'
    if ($m.Success) { $d['FABRNR'] = $m.Groups[1].Value }
  }
  if (-not $d.Contains('FABRNR')) {
    foreach ($l in $lines) {
      if ((M1 $l 'P\S?N|PIN|Art|Sales|Order|Prod|REF').Success) { continue }
      if ($l -match '^\s*([1-9]\d{9,13})\s*$') { $d['FABRNR'] = $Matches[1]; break }
    }
  }

  # ---- Baujahr ----
  foreach ($l in $lines) {
    $m = M1 $l '(?:Date|DAT|Datum|Baujahr|Bj|Year|Fab)\.?\s*[:.]?\s*(?:\d{1,2}\s*[./-]\s*)?(?:\d{1,2}\s*[./-]\s*)((?:19|20)?\d{2})(?![\d''\u2019`])'
    if (-not $m.Success) { $m = M1 $l '(?:Baujahr|Bj\.?|Year)\s*[:.]?\s*((?:19|20)\d{2})' }
    if (-not $m.Success) { $m = M1 $l '^\s*\d{2}\s*/\s*((?:19|20)\d{2})\s*$' }
    if (-not $m.Success) { $m = M1 $l '(?<![\d.])\d{2}\.\d{2}\.((?:19|20)\d{2})(?![\d.])' }
    if ($m.Success) {
      $y = $m.Groups[1].Value; if ($y.Length -eq 2) { $y = '20' + $y }
      if ([int]$y -ge 1980 -and [int]$y -le 2035) { $d['BAUJAHR'] = $y; break }
    }
  }

  # ---- Netzseite: Spannung + Strom ----
  # 1) Abschnitts-Schreibweise (Exide): "Input" ... "Volts 230" "Amps 5.0"
  $sec = ''
  foreach ($l in $lines) {
    if ($l -match '^\s*(Input|Eingang)\b') { $sec = 'in' } elseif ($l -match '^\s*(Output|Ausgang)\b') { $sec = 'out' }
    $mv = M1 $l '\bVolt\S*\s*[:.]?\s*(\d{2,3})\b'
    $ma = M1 $l '\bAmp\S*\s*[:.]?\s*(\d{1,3}(?:[.,]\d+)?)\b'
    if ($sec -eq 'in') {
      if ($mv.Success -and -not $d.Contains('SPANNUNG')) { $d['SPANNUNG'] = $mv.Groups[1].Value }
      if ($ma.Success -and -not $d.Contains('STROM')) { $d['STROM'] = Num $ma.Groups[1].Value }
    } elseif ($sec -eq 'out') {
      if ($ma.Success -and -not $d.Contains('LEISTUNG')) { $d['LEISTUNG'] = Num $ma.Groups[1].Value }
    }
  }
  # 2) IEB-Schreibweise: "Ue 230V~ Ie N 6,1 A", "Ua 24V Ia N 50 A"
  foreach ($l in $lines) {
    $m = M1 $l '\bUe\b\D{0,4}(\d{3})\s*V' $false
    if ($m.Success -and -not $d.Contains('SPANNUNG')) { $d['SPANNUNG'] = $m.Groups[1].Value }
    $m = M1 $l '\bIe\s*N?\b\D{0,4}(\d{1,2}(?:[.,]\d+)?)\s*A\b' $false
    if ($m.Success -and -not $d.Contains('STROM')) { $d['STROM'] = Num $m.Groups[1].Value }
    $m = M1 $l '\bIa\s*N?\b\D{0,4}(\d{1,3}(?:[.,]\d+)?)\s*A\b' $false
    if ($m.Success -and -not $d.Contains('LEISTUNG')) { $d['LEISTUNG'] = Num $m.Groups[1].Value }
  }
  # 3) Zeilen mit Input/IN/AC/Netz (alles ab "OUT" abschneiden)
  foreach ($l0 in $lines) {
    $l = [regex]::Split($l0, '(?i)\bOUT')[0]
    if (-not (M1 $l 'input|^\s*IN\b|\bIN\.|\bin\s+AC|netz|eingang|prim|VAC|\bAC\b|1x|3x|~|N\s*PE').Success) { continue }
    $m = M1 $l '(\d{3}(?:\s*-\s*\d{3})?)\s*(?:V|1/)' $false
    if (-not $m.Success) { $m = M1 $l '^\s*IN\.?\s+(\d{3})\b' }
    if (-not $m.Success) { continue }
    if (-not $d.Contains('SPANNUNG')) { $d['SPANNUNG'] = ($m.Groups[1].Value -replace '\s', '') }
    $rest = $l.Substring($m.Index + $m.Length)
    $nv = M1 $rest '\d+\s*V' $false
    if ($nv.Success) { $rest = $rest.Substring(0, $nv.Index) }
    $a = M1 $rest '(\d{1,2}(?:[.,]\d+)?)(?:\s*-\s*\d{1,2}(?:[.,]\d+)?)?\s*A\b' $false
    if ($a.Success -and -not $d.Contains('STROM')) { $d['STROM'] = Num $a.Groups[1].Value }
    if ($d.Contains('STROM')) { break }
  }
  # Strom in eigener Zeile: "3 N PE 12,9 A", "IN. 6 A MAX", "50-60Hz 10A"
  if (-not $d.Contains('STROM')) {
    foreach ($l in $lines) {
      if ((M1 $l 'DC|out|Ah\b').Success) { continue }
      $m = M1 $l '(?:N\s*PE|\bIN\.?|Hz)\D{0,6}?(\d{1,2}(?:[.,]\d+)?)\s*A\b'
      if ($m.Success) { $d['STROM'] = Num $m.Groups[1].Value; break }
    }
  }
  if ($d.Contains('SPANNUNG')) { $d['SPANNUNG'] = Norm-Volt $d['SPANNUNG'] }
  if (-not $d.Contains('SPANNUNG')) {
    if ((M1 $t '3\s*/?\s*N\s*/?\s*PE|3\s*~|3x\s*400').Success) { $d['SPANNUNG'] = '400' }
    elseif ((M1 $t '1\s*/?\s*N\s*/?\s*PE|1x\s*2[23]0').Success) { $d['SPANNUNG'] = '230' }
    else {
      foreach ($l in $lines) {
        if ((M1 $l 'Ah|DC|out').Success) { continue }
        $m = M1 $l '(?<![\d.,])(220-240|100-240|200-240|380-415|230|400)(?![\d.,])' $false
        if ($m.Success) {
          $d['SPANNUNG'] = Norm-Volt $m.Groups[1].Value
          if (-not $d.Contains('STROM')) {
            $rest = $l.Substring($m.Index + $m.Length) -replace '^\s*V', ''
            $nv = M1 $rest '\d+\s*V' $false
            if ($nv.Success) { $rest = $rest.Substring(0, $nv.Index) }
            $a = M1 $rest '(\d{1,2}(?:[.,]\d+)?)(?:\s*-\s*\d{1,2}(?:[.,]\d+)?)?\s*A\b' $false
            if ($a.Success) { $d['STROM'] = Num $a.Groups[1].Value }
          }
          break
        }
      }
    }
  }

  # ---- Ladestrom (Ausgang) -> Feld "Nennleistung" in A ----
  if (-not $d.Contains('LEISTUNG')) {
    foreach ($l in $lines) {
      $m = M1 $l '\b(\d{1,2})\s*V\s*(?:DC|[^\d\s]{1,2})?\s*(\d{1,3}(?:[.,]\d+)?)(?:\s*-\s*(\d{1,3}))?\s*A\b' $false
      if ($m.Success -and [int]$m.Groups[1].Value -le 96) {
        $v = $m.Groups[2].Value; if ($m.Groups[3].Success) { $v = $m.Groups[3].Value }
        $d['LEISTUNG'] = Num $v; break
      }
    }
  }
  if (-not $d.Contains('LEISTUNG')) {
    foreach ($l in $lines) {
      $m = M1 $l '\bOUT\S*\D{0,12}?(\d{1,3})\s*A\b'
      if ($m.Success) { $d['LEISTUNG'] = $m.Groups[1].Value; break }
    }
  }
  if (-not $d.Contains('LEISTUNG')) {
    foreach ($l in $lines) { if ($l -match '^\s*(\d{2,3})\s*A\s*$' -and [int]$Matches[1] -ge 20) { $d['LEISTUNG'] = $Matches[1]; break } }
  }
  if ($d.Contains('LEISTUNG')) { $d['EINHEIT'] = 'A' }

  # ---- Schutzklasse ----
  $m = M1 $t '(?:\bcl|class|Schutzklasse|\bSK)\s*\.?\s*([I1l]{1,3}|[23])\b'
  if ($m.Success) { $g = $m.Groups[1].Value; if ($g -eq '2' -or $g -eq '3') { $d['SCHUTZKLASSE'] = $g } else { $d['SCHUTZKLASSE'] = [string]$g.Length } }
  if (-not $d.Contains('SCHUTZKLASSE')) { $d['SCHUTZKLASSE'] = '1' }
  return (Merge-Results $d ([ordered]@{}))
}

function Merge-Results($dA, $dB) {
  # Variante mit mehr Treffern zuerst, fehlende Werte aus der anderen ergaenzen
  $p = $dA; $s = $dB
  if ($dB.Count -gt $dA.Count) { $p = $dB; $s = $dA }
  $m = [ordered]@{}
  foreach ($k in 'HERSTELLER','TYP','FABRNR','BAUJAHR','SCHUTZKLASSE','LEISTUNG','EINHEIT','SPANNUNG','STROM') {
    if ($p.Contains($k)) { $m[$k] = $p[$k] } elseif ($s.Contains($k)) { $m[$k] = $s[$k] }
  }
  return $m
}
function Format-Result($d) { (@($d.Keys | ForEach-Object { $_ + '=' + $d[$_] }) -join "`r`n") }

# Google liefert Beschriftung und Wert oft in getrennten Zeilen ("S/N" / "123456").
# Stehen mehrere Beschriftungen untereinander, gehoeren die folgenden Werte der Reihe nach dazu.
$LabelRx = '^\s*(TYPE?|Typ|MODEL|S\s*/\s*N|SN|C?P\s*/\s*N|Ser(?:ial)?\.?\s*(?:No|Nr|#+)?\.?|Art\.?\s*No\.?|Volts?|Amps?|Ue|Ua|[lI]e\s*N?|[iI]a\s*N?|Date|DAT\.?|Baujahr)\s*[:.#]*\s*$'
function Join-Labels($lines) {
  $out = New-Object System.Collections.ArrayList
  $i = 0
  while ($i -lt $lines.Count) {
    $labs = @()
    while ($i -lt $lines.Count -and $lines[$i] -match $LabelRx) { $labs += $lines[$i].Trim(); $i++ }
    if ($labs.Count -eq 0) {
      $l = $lines[$i]
      # "3 N PE" / "50 Hz" und darunter "12,9 A"
      if ($i + 1 -lt $lines.Count -and $l -match '(N\s*/?\s*PE|Hz)\s*$' -and $lines[$i + 1] -match '^\s*\d{1,2}(?:[.,]\d+)?\s*A\s*$') { $l = $l + ' ' + $lines[$i + 1].Trim(); $i++ }
      # "Geraete-Nr 630 912033" und darunter "0419"
      if ($i + 1 -lt $lines.Count -and $l -match 'Ger(ä|a)te-?\s*Nr' -and $lines[$i + 1] -match '^\s*\d{3,5}\s*$') { $l = $l + ' ' + $lines[$i + 1].Trim(); $i++ }
      [void]$out.Add($l); $i++; continue
    }
    $k = 0
    while ($k -lt $labs.Count -and ($i + $k) -lt $lines.Count) {
      $v = $lines[$i + $k]
      if ($v -match $LabelRx -or $v -match '^\s*(Input|Output|IN|OUT|Eingang|Ausgang)\b' -or $v -notmatch '\d') { break }
      $k++
    }
    for ($j = 0; $j -lt $labs.Count; $j++) {
      $lab = $labs[$j] -creplace '^\s*le', 'Ie' -creplace '^\s*ia', 'Ia'
      if ($j -lt $k) {
        $v = $lines[$i + $j].Trim()
        # Einheit steht manchmal in der naechsten Zeile ("50" / "A ...")
        if ($k -eq 1 -and $v -match '^\d+(?:[.,]\d+)?$' -and ($i + 1) -lt $lines.Count -and $lines[$i + 1] -match '^\s*A\b' -and $lab -match '^(Ie|Ia)') { $v += ' A' }
        [void]$out.Add($lab + ' ' + $v)
      } else { [void]$out.Add($lab) }
    }
    $i += $k
  }
  return @($out)
}

function Run-One([string]$path) {
  Log ("bild " + (Split-Path $path -Leaf))
  $r = Read-Plate $path
  $res = Parse-Plate $r.All
  return @{ R = $r; Res = $res }
}

if ($Folder) {
  Get-ChildItem -LiteralPath $Folder -File | Where-Object { $_.Extension -match '^\.(jpe?g|png)$' } | Sort-Object Name | ForEach-Object {
    $x = Run-One $_.FullName
    $part = "=== " + $_.Name + " | rot " + $x.R.Rot + " | " + $x.R.Scores + "`r`n" + (Format-Result $x.Res) + "`r`n#ROWS`r`n" + (($x.R.RowsA) -join "`r`n") + "`r`n#ROWSB`r`n" + (($x.R.RowsB) -join "`r`n") + "`r`n#ROWSC`r`n" + (($x.R.RowsC) -join "`r`n") + "`r`n"
    [System.IO.File]::AppendAllText($Out, $part, (New-Object System.Text.UTF8Encoding $false))
  }
} elseif ($TextFile) {
  $lines = @([System.IO.File]::ReadAllLines($TextFile, [System.Text.Encoding]::UTF8) | Where-Object { $_.Trim() })
  Log ("text " + $lines.Count + " Zeilen")
  $joined = Join-Labels $lines
  $res = Parse-Plate $joined
  # Ladestrom in eigener Zeile unter OUTPUT ("125 A 6 kW")
  if (-not $res.Contains('LEISTUNG')) {
    $o = $false
    foreach ($l in $joined) {
      if ($l -match '^\s*(Output|OUT)\b') { $o = $true; continue }
      if ($o -and $l -match '^\s*(\d{2,3})\s*A\b' -and [int]$Matches[1] -ge 20) { $res['LEISTUNG'] = $Matches[1]; $res['EINHEIT'] = 'A'; break }
    }
  }
  # Plausibilitaet: Ladestrom kann nicht die Netzspannung sein
  if ($res.Contains('LEISTUNG') -and $res.Contains('SPANNUNG') -and $res['LEISTUNG'] -eq $res['SPANNUNG']) { $res.Remove('LEISTUNG'); $res.Remove('EINHEIT') }
  if ($res.Contains('TYP')) { $res['TYP'] = ($res['TYP'] -replace '^\s*MODEL\s+', '').Trim() }
  # Exide-Ladegeraete: Typ steht ohne Beschriftung ganz oben ("SP24S030S-77002")
  if (-not $res.Contains('TYP')) {
    foreach ($l in $lines) { if ($l -match '^\s*(SP\d{2}[A-Z0-9 ]{3,}-\d{4,5})\s*$') { $res['TYP'] = $Matches[1]; break } }
  }
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine((Format-Result $res))
  [void]$sb.AppendLine("#ROH")
  foreach ($l in $lines) { [void]$sb.AppendLine($l) }
  [System.IO.File]::WriteAllText($Out, $sb.ToString(), (New-Object System.Text.UTF8Encoding $true))
} else {
  $x = Run-One $Image
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine((Format-Result $x.Res))
  [void]$sb.AppendLine("#ROH")
  foreach ($l in (@($x.R.RowsA) + @('') + @($x.R.RowsC) + @('') + @($x.R.RowsB))) { [void]$sb.AppendLine($l) }
  [System.IO.File]::WriteAllText($Out, $sb.ToString(), (New-Object System.Text.UTF8Encoding $true))
}
Log "fertig"
