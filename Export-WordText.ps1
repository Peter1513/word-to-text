<#
.SYNOPSIS
    Extract all text from Word files (.docx/.docm/.dotx/.dotm; .doc/.rtf/.odt via Word or LibreOffice).

.DESCRIPTION
    Default output is Markdown: Word heading styles -> #/##, bullet/numbered lists -> -/1.
    (nested), tables -> Markdown tables, blank line between paragraphs. -Format Txt gives
    plain text (tables tab-separated), -Format Json structured data.

    .docx family is parsed directly from the ZIP/XML (no Word needed): body paragraphs,
    tables (cells tab-separated), text boxes, content controls, tracked insertions
    (deletions skipped), plus headers, footers, footnotes and endnotes.
    Legacy formats (.doc/.rtf/.odt) are converted to a temporary .docx via Microsoft Word COM
    (Windows, if installed), otherwise LibreOffice (soffice), then parsed the same way.

    Works on Windows PowerShell 5.1 and PowerShell 7+.

.EXAMPLE
    .\Export-WordText.ps1 -Path C:\Docs -OutDir C:\Docs\md
    One .md per Word file.

.EXAMPLE
    .\Export-WordText.ps1 -Path C:\Docs -Recurse -Combine C:\all.md
    All documents in one Markdown file (each under '# <file name>').

.EXAMPLE
    .\Export-WordText.ps1 -Path a.docx, b.doc -Format Json -Combine out.json

.NOTES
    Exit codes: 0 = all ok, 1 = some files failed, 2 = no Word files found / bad arguments.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('FullName')]
    [string[]] $Path,

    [switch] $Recurse,

    # One output file per input, written into this folder.
    [string] $OutDir,

    # All output into this single file. If neither -OutDir nor -Combine: print to console.
    [string] $Combine,

    [ValidateSet('Txt', 'Md', 'Json')]
    [string] $Format = 'Md',

    # Body only: skip headers, footers, footnotes/endnotes.
    [switch] $NoExtras,

    # Write UTF-8 with BOM (helps old Notepad / Excel recognise Chinese text).
    [switch] $Bom,

    [switch] $Overwrite
)

begin {
    Set-StrictMode -Version 2.0
    $ErrorActionPreference = 'Stop'
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

    $script:DocxExt   = @('.docx', '.docm', '.dotx', '.dotm')
    $script:LegacyExt = @('.doc', '.dot', '.rtf', '.odt')
    $script:AllExt    = $script:DocxExt + $script:LegacyExt
    $script:Encoding  = New-Object System.Text.UTF8Encoding($Bom.IsPresent)
    $script:WordApp   = $null
    $inputs = New-Object System.Collections.Generic.List[string]

    # ------------------------------------------------------------------ docx parsing
    # Markdown mode: list lines are prefixed with $TightList, table rows with $TightTable,
    # so Join-Blocks keeps consecutive items of the same kind on adjacent lines and puts
    # a blank line between everything else.
    $script:Md = $false; $script:HeadingShift = 0
    $script:TightList = [string][char]1; $script:TightTable = [string][char]2
    $script:StyleInfo = @{}; $script:NumFmt = @{}

    function Get-Child($node, [string] $name) {
        if ($null -eq $node) { return $null }
        foreach ($c in $node.ChildNodes) { if ($c.LocalName -eq $name) { return $c } }
        return $null
    }
    function Get-WAttr($node, [string] $name) {
        if ($null -eq $node -or $null -eq $node.Attributes) { return $null }
        foreach ($a in $node.Attributes) { if ($a.LocalName -eq $name) { return $a.Value } }
        return $null
    }

    function Import-Styles($zip) {
        # styleId -> @{ Level = heading level or 0; NumId; Ilvl } (inherits via basedOn)
        $script:StyleInfo = @{}
        $xml = Read-ZipXml $zip 'word/styles.xml'
        if (-not $xml) { return }
        $raw = @{}
        foreach ($s in $xml.DocumentElement.ChildNodes) {
            if ($s.LocalName -ne 'style') { continue }
            $pPr = Get-Child $s 'pPr'; $numPr = Get-Child $pPr 'numPr'
            $raw[(Get-WAttr $s 'styleId')] = @{
                Name    = [string](Get-WAttr (Get-Child $s 'name') 'val')
                BasedOn = Get-WAttr (Get-Child $s 'basedOn') 'val'
                Outline = Get-WAttr (Get-Child $pPr 'outlineLvl') 'val'
                NumId   = Get-WAttr (Get-Child $numPr 'numId') 'val'
                Ilvl    = Get-WAttr (Get-Child $numPr 'ilvl') 'val'
            }
        }
        foreach ($id in @($raw.Keys)) {
            $level = 0; $numId = $null; $ilvl = $null; $cur = $id; $depth = 0
            while ($cur -and $raw.ContainsKey($cur) -and $depth -lt 10) {
                $r = $raw[$cur]
                if (-not $level) {
                    # built-in names are stored in English even in Chinese Word ("heading 1")
                    if ($r.Name -match '^heading\s*([1-9])$') { $level = [int]$Matches[1] }
                    elseif ($r.Name -eq 'Title') { $level = 1 }
                    elseif ($r.Outline -match '^[0-8]$') { $level = [int]$r.Outline + 1 }
                }
                if (-not $numId -and $r.NumId) { $numId = $r.NumId; $ilvl = $r.Ilvl }
                $cur = $r.BasedOn; $depth++
            }
            $script:StyleInfo[$id] = @{ Level = $level; NumId = $numId; Ilvl = $ilvl }
        }
    }

    function Import-Numbering($zip) {
        # numId -> @{ ilvl -> numFmt ('bullet', 'decimal', 'none', ...) }
        $script:NumFmt = @{}
        $xml = Read-ZipXml $zip 'word/numbering.xml'
        if (-not $xml) { return }
        $abstract = @{}
        foreach ($n in $xml.DocumentElement.ChildNodes) {
            if ($n.LocalName -ne 'abstractNum') { continue }
            $levels = @{}
            foreach ($lvl in $n.ChildNodes) {
                if ($lvl.LocalName -eq 'lvl') { $levels[(Get-WAttr $lvl 'ilvl')] = Get-WAttr (Get-Child $lvl 'numFmt') 'val' }
            }
            $abstract[(Get-WAttr $n 'abstractNumId')] = $levels
        }
        foreach ($n in $xml.DocumentElement.ChildNodes) {
            if ($n.LocalName -ne 'num') { continue }
            $aid = Get-WAttr (Get-Child $n 'abstractNumId') 'val'
            if ($abstract.ContainsKey($aid)) { $script:NumFmt[(Get-WAttr $n 'numId')] = $abstract[$aid] }
        }
    }

    function Get-ListPrefix($numId, $ilvl) {
        if (-not $numId -or $numId -eq '0') { return $null }
        if (-not $ilvl) { $ilvl = '0' }
        $fmt = 'bullet'
        if ($script:NumFmt.ContainsKey($numId) -and $script:NumFmt[$numId].ContainsKey($ilvl)) { $fmt = $script:NumFmt[$numId][$ilvl] }
        if ($fmt -eq 'none') { return $null }
        $indent = '   ' * [int]$ilvl
        if ($fmt -eq 'bullet') { return "$indent- " } else { return "${indent}1. " }
    }

    function ConvertTo-MdParagraph([System.Xml.XmlNode] $p, [string] $text) {
        $pPr = Get-Child $p 'pPr'
        $styleId = Get-WAttr (Get-Child $pPr 'pStyle') 'val'
        $style = if ($styleId -and $script:StyleInfo.ContainsKey($styleId)) { $script:StyleInfo[$styleId] } else { @{ Level = 0; NumId = $null; Ilvl = $null } }
        $level = $style.Level
        $outline = Get-WAttr (Get-Child $pPr 'outlineLvl') 'val'
        if ($outline -match '^[0-8]$') { $level = [int]$outline + 1 }
        $text = $text.Trim(' ', "`t")

        if ($level -gt 0) {
            $hashes = '#' * [Math]::Min(6, $level + $script:HeadingShift)
            return "$hashes " + ($text -replace '\s*\n\s*', ' ')
        }
        $numPr = Get-Child $pPr 'numPr'
        $numId = Get-WAttr (Get-Child $numPr 'numId') 'val'; $ilvl = Get-WAttr (Get-Child $numPr 'ilvl') 'val'
        if (-not $numPr) { $numId = $style.NumId; $ilvl = $style.Ilvl }
        $prefix = Get-ListPrefix $numId $ilvl
        if ($prefix) { return $script:TightList + $prefix + ($text -replace '\s*\n\s*', ' ') }

        # plain paragraph: escape text that Markdown would read as structure
        if ($text -match '^(\d+)([.)])(\s.*)?$') { $text = $Matches[1] + '\' + $Matches[2] + $Matches[3] }
        elseif ($text -match '^(#|[-+*]\s|>|```|~~~|---|===|\+\+\+|\|)') { $text = '\' + $text }
        return ($text -replace "`n", "  `n")   # Word line break -> Markdown hard break
    }

    function Test-BlockMarker([string] $l) {
        # ordinal char compare: culture-aware StartsWith() ignores control chars and would match everything
        return ($l.Length -gt 0 -and ($l[0] -eq $script:TightList[0] -or $l[0] -eq $script:TightTable[0]))
    }

    function Join-Blocks($lines) {
        if (-not $script:Md) { return ($lines -join "`n") }
        $sb = New-Object System.Text.StringBuilder
        $prev = $null; $first = $true
        foreach ($l in $lines) {
            if ($l -eq '') { continue }
            # kinds as char codes: string -eq is culture-aware and treats \u0001 and \u0002 as equal
            $kind = if ((Test-BlockMarker $l)) { [int]$l[0] } else { 0 }
            if ($kind -ne 0) { $l = $l.Substring(1) }
            if (-not $first) { [void]$sb.Append($(if ($kind -ne 0 -and $kind -eq $prev) { "`n" } else { "`n`n" })) }
            [void]$sb.Append($l); $prev = $kind; $first = $false
        }
        return $sb.ToString()
    }

    function Get-ParagraphLines([System.Xml.XmlNode] $p) {
        $sb = New-Object System.Text.StringBuilder
        $nested = New-Object System.Collections.Generic.List[string]
        $walk = {
            param($elem)
            foreach ($c in $elem.ChildNodes) {
                if ($c.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                switch ($c.LocalName) {
                    't'             { [void]$sb.Append($c.InnerText) }
                    'tab'           { [void]$sb.Append("`t") }
                    'br'            { [void]$sb.Append("`n") }
                    'cr'            { [void]$sb.Append("`n") }
                    'noBreakHyphen' { [void]$sb.Append('-') }
                    'txbxContent'   { foreach ($l in (Get-BlockLines $c)) { $nested.Add($l) } }
                    # deleted revisions, duplicate VML fallback of text boxes, field codes
                    { $_ -in 'del', 'delText', 'Fallback', 'instrText' } { }
                    default         { & $walk $c }
                }
            }
        }
        & $walk $p
        $lines = New-Object System.Collections.Generic.List[string]
        $text = $sb.ToString()
        if (-not $script:Md) { $lines.Add($text) }
        elseif ($text.Trim()) { $lines.Add((ConvertTo-MdParagraph $p $text)) }   # empty paragraphs dropped
        foreach ($l in $nested) { $lines.Add($l) }
        return ,$lines
    }

    function Get-MdTableLines([System.Xml.XmlNode] $tbl) {
        $grid = New-Object System.Collections.Generic.List[object]
        foreach ($tr in $tbl.ChildNodes) {
            if ($tr.LocalName -ne 'tr') { continue }
            $cells = New-Object System.Collections.Generic.List[string]
            foreach ($tc in $tr.ChildNodes) {
                if ($tc.LocalName -ne 'tc') { continue }
                $parts = foreach ($l in (Get-BlockLines $tc)) {
                    if ((Test-BlockMarker $l)) { $l = $l.Substring(1) }
                    if ($l.Trim()) { ($l -replace '\|', '\|') -replace '\s*\n\s*', ' ' }
                }
                $cells.Add((@($parts) -join '<br>'))
            }
            $grid.Add($cells)
        }
        $rows = New-Object System.Collections.Generic.List[string]
        if ($grid.Count -eq 0) { return ,$rows }
        $width = ($grid | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
        if ($width -lt 1) { return ,$rows }
        for ($i = 0; $i -lt $grid.Count; $i++) {
            $cells = @($grid[$i]); while ($cells.Count -lt $width) { $cells += '' }   # merged cells -> pad
            $rows.Add($script:TightTable + '| ' + ($cells -join ' | ') + ' |')
            if ($i -eq 0) { $rows.Add($script:TightTable + '|' + ((@(1..$width) | ForEach-Object { ' --- ' }) -join '|') + '|') }
        }
        return ,$rows
    }

    function Get-TableLines([System.Xml.XmlNode] $tbl) {
        if ($script:Md) { return ,(Get-MdTableLines $tbl) }
        $rows = New-Object System.Collections.Generic.List[string]
        foreach ($tr in $tbl.ChildNodes) {
            if ($tr.LocalName -ne 'tr') { continue }
            $cells = foreach ($tc in $tr.ChildNodes) {
                if ($tc.LocalName -ne 'tc') { continue }
                ((Get-BlockLines $tc) | Where-Object { $_.Trim() }) -join ' '
            }
            $rows.Add((@($cells) -join "`t"))
        }
        return ,$rows
    }

    function Get-BlockLines([System.Xml.XmlNode] $container) {
        $out = New-Object System.Collections.Generic.List[string]
        foreach ($child in $container.ChildNodes) {
            $lines = switch ($child.LocalName) {
                'p'   { Get-ParagraphLines $child }
                'tbl' { Get-TableLines $child }
                { $_ -in 'sdt', 'sdtContent', 'customXml', 'ins' } { Get-BlockLines $child }
                default { $null }
            }
            # $null check, not truthiness: a list holding one empty paragraph is "falsy"
            if ($null -ne $lines) { foreach ($l in $lines) { $out.Add([string]$l) } }
        }
        return ,$out
    }

    function Read-ZipXml($zip, [string] $name) {
        $entry = $zip.GetEntry($name)
        if (-not $entry) { return $null }
        $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
        try {
            $doc = New-Object System.Xml.XmlDocument
            $doc.PreserveWhitespace = $true
            $doc.LoadXml($reader.ReadToEnd())
            return $doc
        } finally { $reader.Dispose() }
    }

    function Get-PartText($zip, [string] $name) {
        $xml = Read-ZipXml $zip $name
        if (-not $xml) { return '' }
        $root = $xml.DocumentElement
        $lines = New-Object System.Collections.Generic.List[string]
        if ($root.LocalName -in 'footnotes', 'endnotes') {
            foreach ($note in $root.ChildNodes) {
                if ($note.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                # skip separator / continuationSeparator pseudo-notes
                if ($note.Attributes -and ($note.Attributes | Where-Object { $_.LocalName -eq 'type' })) { continue }
                foreach ($l in (Get-BlockLines $note)) { $lines.Add($l) }
            }
        } else {
            $body = $root.GetElementsByTagName('body', $root.NamespaceURI) | Select-Object -First 1
            if (-not $body) { $body = $root }
            foreach ($l in (Get-BlockLines $body)) { $lines.Add($l) }
        }
        return (Join-Blocks $lines).Trim()
    }

    function Get-SortedParts($zip, [string] $kind) {
        $zip.Entries |
            Where-Object { $_.FullName -match "^word/$kind(\d*)\.xml$" } |
            Sort-Object { if ($_.FullName -match "$kind(\d+)") { [int]$Matches[1] } else { 0 } } |
            ForEach-Object { $_.FullName }
    }

    function Read-Docx([string] $file) {
        try { $zip = [System.IO.Compression.ZipFile]::OpenRead($file) }
        catch { throw "corrupt or not a .docx file: $($_.Exception.InnerException.Message)" }
        try {
            if (-not $zip.GetEntry('word/document.xml')) { throw 'word/document.xml not found (not a Word file?)' }
            if ($script:Md) { Import-Styles $zip; Import-Numbering $zip }
            [pscustomobject]@{
                File    = $file
                Body    = Get-PartText $zip 'word/document.xml'
                Headers = @(Get-SortedParts $zip 'header' | ForEach-Object { Get-PartText $zip $_ })
                Footers = @(Get-SortedParts $zip 'footer' | ForEach-Object { Get-PartText $zip $_ })
                Notes   = @('footnotes', 'endnotes' | ForEach-Object { Get-PartText $zip "word/$_.xml" } | Where-Object { $_ })
            }
        } finally { $zip.Dispose() }
    }

    # ------------------------------------------------------------------ legacy formats
    function Get-WordApp {
        if ($script:WordApp) { return $script:WordApp }
        if (-not $IsWindowsOrPS5) { return $null }
        try {
            $script:WordApp = New-Object -ComObject Word.Application
            $script:WordApp.Visible = $false
            $script:WordApp.DisplayAlerts = 0
        } catch { $script:WordApp = $null }
        return $script:WordApp
    }

    function Read-Legacy([string] $file) {
        # Convert to a temporary .docx (Word first, LibreOffice second), then reuse Read-Docx,
        # so legacy files get the same structure (headings, lists, tables) as .docx.
        $tmp = Join-Path ([IO.Path]::GetTempPath()) ("wordtext_" + [guid]::NewGuid())
        New-Item -ItemType Directory $tmp | Out-Null
        try {
            $dst = Join-Path $tmp 'converted.docx'
            $word = Get-WordApp
            if ($word) {
                # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument)
                $doc = $word.Documents.Open($file, $false, $true, $false, 'x-no-password')
                try { $doc.SaveAs2($dst, 16) }          # 16 = wdFormatDocumentDefault (.docx)
                finally { $doc.Close($false) }
            } else {
                $soffice = Get-Command soffice, soffice.exe, libreoffice -ErrorAction SilentlyContinue | Select-Object -First 1
                if (-not $soffice -and $IsWindowsOrPS5) {
                    $guess = Join-Path $env:ProgramFiles 'LibreOffice\program\soffice.exe'
                    if (Test-Path -LiteralPath $guess) { $soffice = Get-Item $guess }
                }
                if (-not $soffice) { throw "$([IO.Path]::GetExtension($file)) needs Microsoft Word or LibreOffice; neither found" }
                $exe = if ($soffice.Source) { $soffice.Source } else { $soffice.FullName }
                # file:///C:/... on Windows, file:///tmp/... elsewhere ([Uri]"/tmp/x" is not absolute on Linux)
                $p = (Join-Path $tmp 'profile') -replace '\\', '/'
                if (-not $p.StartsWith('/')) { $p = '/' + $p }
                $profileUri = (New-Object System.UriBuilder('file', '', -1, $p)).Uri.AbsoluteUri
                $outDir = Join-Path $tmp 'lo'
                $argLine = ('"-env:UserInstallation={0}" --headless --convert-to "docx:MS Word 2007 XML" --outdir "{1}" "{2}"' -f $profileUri, $outDir, $file)
                $sp = @{ FilePath = $exe; ArgumentList = $argLine; PassThru = $true
                         RedirectStandardOutput = (Join-Path $tmp 'so.log'); RedirectStandardError = (Join-Path $tmp 'se.log') }
                if ($IsWindowsOrPS5) { $sp.WindowStyle = 'Hidden' }   # -WindowStyle exists only on Windows
                $proc = Start-Process @sp
                if (-not $proc.WaitForExit(120000)) {   # LibreOffice can hang on broken files
                    try { $proc.Kill() } catch { }
                    throw 'LibreOffice timed out after 120s'
                }
                $lo = Join-Path $outDir ([IO.Path]::GetFileNameWithoutExtension($file) + '.docx')
                if (-not (Test-Path -LiteralPath $lo)) { throw 'LibreOffice produced no output' }
                Move-Item -LiteralPath $lo -Destination $dst
            }
            $r = Read-Docx $dst
            $r.File = $file
            return $r
        } finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
    }

    # ------------------------------------------------------------------ output
    function Format-Result($r) {
        $parts = New-Object System.Collections.Generic.List[string]
        $parts.Add($r.Body)
        if (-not $NoExtras) {
            foreach ($pair in @(@('HEADERS', 'Headers', $r.Headers), @('FOOTERS', 'Footers', $r.Footers), @('NOTES', 'Notes', $r.Notes))) {
                $items = @($pair[2] | Where-Object { $_ -and $_.Trim() })
                if (-not $items.Count) { continue }
                if ($script:Md) { $parts.Add("`n---`n`n**$($pair[1])**`n`n" + ($items -join "`n`n")) }
                else            { $parts.Add("`n----- $($pair[0]) -----`n" + ($items -join "`n")) }
            }
        }
        return (($parts -join "`n").Trim() + "`n")
    }

    function Write-TextFile([string] $file, [string] $text) {
        $dir = Split-Path $file -Parent
        if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory $dir -Force | Out-Null }
        [IO.File]::WriteAllText($file, $text, $script:Encoding)
    }

    $IsWindowsOrPS5 = ($PSVersionTable.PSVersion.Major -le 5) -or ((Get-Variable IsWindows -ValueOnly -ErrorAction SilentlyContinue) -eq $true)
}

process {
    foreach ($p in $Path) { $inputs.Add($p) }
}

end {
    # ---- argument checks
    if ($OutDir -and $Combine) { Write-Error 'Use either -OutDir or -Combine, not both.' -ErrorAction Continue; exit 2 }
    if ($OutDir -and $Format -eq 'Json') { Write-Error '-Format Json needs -Combine (or console output).' -ErrorAction Continue; exit 2 }
    if ($Combine) {
        $Combine = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Combine)
        if ((Test-Path -LiteralPath $Combine) -and -not $Overwrite) { Write-Error "$Combine exists (use -Overwrite)." -ErrorAction Continue; exit 2 }
    }
    if ($OutDir) { $OutDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir) }

    # ---- collect files
    $files = New-Object System.Collections.Generic.List[string]
    $baseDir = $null
    foreach ($raw in $inputs) {
        if (Test-Path -LiteralPath $raw -PathType Container) {
            $dir = (Resolve-Path -LiteralPath $raw).ProviderPath
            if ($inputs.Count -eq 1) { $baseDir = $dir }
            Get-ChildItem -LiteralPath $dir -File -Recurse:$Recurse |
                Where-Object { $script:AllExt -contains $_.Extension.ToLower() } |
                ForEach-Object { $files.Add($_.FullName) }
        } elseif (Test-Path -LiteralPath $raw -PathType Leaf) {
            $files.Add((Resolve-Path -LiteralPath $raw).ProviderPath)
        } else {
            Write-Warning "not found: $raw"
        }
    }
    # skip Word lock files (~$x.docx), de-duplicate, stable order
    $files = @($files | Where-Object { -not (Split-Path $_ -Leaf).StartsWith('~$') } | Sort-Object -Unique)
    if ($files.Count -eq 0) { Write-Error 'No Word files found.' -ErrorAction Continue; exit 2 }

    $toConsole = -not ($OutDir -or $Combine)
    $script:Md = ($Format -eq 'Md')
    # combined/console Markdown puts '# <file name>' above each doc, so its headings move down one level
    $script:HeadingShift = if ($script:Md -and -not $OutDir) { 1 } else { 0 }
    $ext = '.' + $Format.ToLower()
    $combined = New-Object System.Text.StringBuilder
    $jsonItems = New-Object System.Collections.Generic.List[object]
    $written = @{}
    $failures = 0

    try {
        for ($i = 0; $i -lt $files.Count; $i++) {
            $f = $files[$i]
            $tag = "[$($i + 1)/$($files.Count)]"
            try {
                $fx = [IO.Path]::GetExtension($f).ToLower()
                $r = if ($script:DocxExt -contains $fx) { Read-Docx $f }
                     elseif ($script:LegacyExt -contains $fx) { Read-Legacy $f }
                     else { throw "unsupported file type: $fx" }

                if ($Format -eq 'Json') {
                    $jsonItems.Add([pscustomobject]@{ file = $f; body = [string]$r.Body
                        headers = [string[]]@($r.Headers); footers = [string[]]@($r.Footers); notes = [string[]]@($r.Notes) })
                } else {
                    $text = Format-Result $r
                    $name = Split-Path $f -Leaf
                    if ($OutDir) {
                        $rel = if ($baseDir -and $f.StartsWith($baseDir, [StringComparison]::OrdinalIgnoreCase)) {
                            $f.Substring($baseDir.Length).TrimStart('\', '/') } else { $name }
                        $target = Join-Path $OutDir ([IO.Path]::ChangeExtension($rel, $ext))
                        # a.doc + a.docx in one folder would collide -> keep original extension
                        if ($written.ContainsKey($target.ToLower())) { $target = Join-Path $OutDir ($rel + $ext) }
                        if ((Test-Path -LiteralPath $target) -and -not $Overwrite) { throw "$target exists (use -Overwrite)" }
                        Write-TextFile $target $text
                        $written[$target.ToLower()] = $true
                    } else {
                        $sep = if ($Format -eq 'Md') { "# $name`n`n" } else { "===== $f =====`n" }
                        [void]$combined.Append($sep + $text + "`n")
                    }
                }
                if (-not $toConsole) { Write-Host "$tag ok    $f" }
            } catch {
                $failures++
                Write-Host "$tag FAIL  $f : $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    } finally {
        if ($script:WordApp) {
            $script:WordApp.Quit()
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($script:WordApp)
        }
    }

    $payload = if ($Format -eq 'Json') { (ConvertTo-Json -InputObject $jsonItems.ToArray() -Depth 5) + "`n" } else { $combined.ToString() }
    if ($Combine) { Write-TextFile $Combine $payload }
    elseif ($toConsole) { [Console]::Out.Write($payload) }

    $ok = $files.Count - $failures
    Write-Host "done: $ok ok, $failures failed, $($files.Count) total"
    if ($failures) { exit 1 } else { exit 0 }
}
