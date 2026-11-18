<#  Self-contained tests (no Pester needed).  Run:  .\tests\Test-ExportWordText.ps1
    Exit code 0 = all passed.  #>
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

$Script = Join-Path $PSScriptRoot '..\Export-WordText.ps1' | Resolve-Path
$Shell  = (Get-Process -Id $PID).Path          # same pwsh / powershell.exe running the tests
$Tmp    = Join-Path ([IO.Path]::GetTempPath()) ("wt_test_" + [guid]::NewGuid())
New-Item -ItemType Directory $Tmp | Out-Null
$script:Failed = 0; $script:Passed = 0; $script:LastRun = $null

function Assert($cond, [string] $msg) {
    if ($cond) { $script:Passed++; Write-Host "  PASS  $msg" -ForegroundColor Green }
    else       { $script:Failed++; Write-Host "  FAIL  $msg" -ForegroundColor Red
                 if ($script:LastRun) { Write-Host "        last run output:`n$($script:LastRun)" } }
}
function Run([string[]] $argList) {
    # Windows PowerShell 5.1 turns native stderr into terminating errors under 'Stop' when
    # redirected with 2>&1, so relax it just for the child process call.
    $saved = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
    try { $out = & $Shell -NoProfile -ExecutionPolicy Bypass -File $Script @argList 2>&1 | Out-String }
    finally { $ErrorActionPreference = $saved }
    $script:LastRun = $out
    return [pscustomobject]@{ Code = $LASTEXITCODE; Out = $out }
}
function ReadUtf8($p) { [IO.File]::ReadAllText($p, [Text.Encoding]::UTF8) }

$W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
$Document = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document $W><w:body>
  <w:p><w:r><w:t>Hello</w:t></w:r><w:r><w:t xml:space="preserve"> 世界</w:t></w:r></w:p>
  <w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:t>C</w:t></w:r></w:p>
  <w:p/>
  <w:p><w:del><w:r><w:delText>DELETED</w:delText></w:r></w:del><w:ins><w:r><w:t>Inserted</w:t></w:r></w:ins></w:p>
  <w:tbl>
    <w:tr><w:tc><w:p><w:r><w:t>R1C1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>R1C2</w:t></w:r></w:p></w:tc></w:tr>
    <w:tr><w:tc><w:p><w:r><w:t>R2C1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>R2C2</w:t></w:r></w:p></w:tc></w:tr>
  </w:tbl>
  <w:p><w:r><mc:AlternateContent>
    <mc:Choice><w:drawing><w:txbxContent><w:p><w:r><w:t>TextBox</w:t></w:r></w:p></w:txbxContent></w:drawing></mc:Choice>
    <mc:Fallback><w:pict><w:txbxContent><w:p><w:r><w:t>TextBox</w:t></w:r></w:p></w:txbxContent></w:pict></mc:Fallback>
  </mc:AlternateContent></w:r></w:p>
  <w:sdt><w:sdtContent><w:p><w:r><w:t>InContentControl</w:t></w:r></w:p></w:sdtContent></w:sdt>
  <w:p><w:r><w:fldChar/><w:instrText> PAGE </w:instrText></w:r></w:p>
</w:body></w:document>
"@
$Parts = @{
    '[Content_Types].xml' = '<Types/>'
    'word/document.xml'   = $Document
    'word/header1.xml'    = "<w:hdr $W><w:p><w:r><w:t>MyHeader</w:t></w:r></w:p></w:hdr>"
    'word/footer1.xml'    = "<w:ftr $W><w:p><w:r><w:t>MyFooter</w:t></w:r></w:p></w:ftr>"
    'word/footnotes.xml'  = "<w:footnotes $W><w:footnote w:type=`"separator`" w:id=`"-1`"><w:p><w:r><w:t>SEP</w:t></w:r></w:p></w:footnote><w:footnote w:id=`"1`"><w:p><w:r><w:t>MyFootnote</w:t></w:r></w:p></w:footnote></w:footnotes>"
}

# Markdown fixture: styleId "1" with name "heading 1" mimics Chinese-UI Word; numbering 1 = bullets, 2 = decimal
$MdDocument = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document $W><w:body>
  <w:p><w:pPr><w:pStyle w:val="1"/></w:pPr><w:r><w:t>總結</w:t></w:r></w:p>
  <w:p><w:r><w:t># 不是標題</w:t></w:r></w:p>
  <w:p/>
  <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>項目A</w:t></w:r></w:p>
  <w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>子項目</w:t></w:r></w:p>
  <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>步驟一</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="ListNumber"/></w:pPr><w:r><w:t>步驟二</w:t></w:r></w:p>
  <w:tbl><w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
  <w:p><w:pPr><w:pStyle w:val="MyHeading"/></w:pPr><w:r><w:t>細節</w:t></w:r></w:p>
  <w:tbl>
    <w:tr><w:tc><w:p><w:r><w:t>欄1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>欄2</w:t></w:r></w:p></w:tc></w:tr>
    <w:tr><w:tc><w:p><w:r><w:t>a|b</w:t></w:r></w:p><w:p><w:r><w:t>第二段</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr>
    <w:tr><w:tc><w:p><w:r><w:t>只有一格</w:t></w:r></w:p></w:tc></w:tr>
  </w:tbl>
  <w:p><w:r><w:t>1. 不是清單</w:t></w:r></w:p>
  <w:p><w:r><w:t>第一行</w:t><w:br/><w:t>第二行</w:t></w:r></w:p>
</w:body></w:document>
"@
$MdStyles = @"
<w:styles $W>
  <w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/></w:style>
  <w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 2"/></w:style>
  <w:style w:type="paragraph" w:styleId="MyHeading"><w:name w:val="My Heading"/><w:basedOn w:val="2"/></w:style>
  <w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr></w:style>
</w:styles>
"@
$MdNumbering = @"
<w:numbering $W>
  <w:abstractNum w:abstractNumId="10"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum>
  <w:abstractNum w:abstractNumId="20"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="10"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="20"/></w:num>
</w:numbering>
"@
$MdParts = @{
    '[Content_Types].xml' = '<Types/>'; 'word/document.xml' = $MdDocument
    'word/styles.xml' = $MdStyles; 'word/numbering.xml' = $MdNumbering
    'word/header1.xml' = $Parts['word/header1.xml']
}

function New-Docx([string] $file, [hashtable] $parts = $Parts) {
    $dir = Split-Path $file -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory $dir -Force | Out-Null }
    $zip = [IO.Compression.ZipFile]::Open($file, 'Create')
    try {
        foreach ($k in $parts.Keys) {
            $w = New-Object IO.StreamWriter($zip.CreateEntry($k).Open(), (New-Object Text.UTF8Encoding($false)))
            $w.Write($parts[$k]); $w.Dispose()
        }
    } finally { $zip.Dispose() }
}

try {
    $src = Join-Path $Tmp 'src'
    0..19 | ForEach-Object { New-Docx (Join-Path $src ('file{0:D2}.docx' -f $_)) }
    New-Docx (Join-Path $src 'sub\nested.docx'.Replace('\', [IO.Path]::DirectorySeparatorChar))
    [IO.File]::WriteAllText((Join-Path $src '~$file00.docx'), 'lock')     # Word lock file
    [IO.File]::WriteAllText((Join-Path $src 'notes.txt'), 'ignore me')

    Write-Host "`n[1] per-file output"
    $out = Join-Path $Tmp 'out'
    $r = Run @('-Path', $src, '-OutDir', $out, '-Format', 'Txt')
    Assert ($r.Code -eq 0) "exit code 0 (got $($r.Code))"
    Assert (@(Get-ChildItem $out -Filter *.txt).Count -eq 20) '20 .txt files, lock file + .txt input ignored'
    $t = ReadUtf8 (Join-Path $out 'file05.txt')
    $lines = $t -split "`n"
    Assert ($lines[0] -eq 'Hello 世界') 'runs joined, Chinese preserved'
    Assert ($lines[1] -eq "A`tB" -and $lines[2] -eq 'C') 'tab and line break'
    Assert ($lines[3] -eq '') 'empty paragraph kept'
    Assert ($t.Contains('Inserted') -and -not $t.Contains('DELETED')) 'tracked insert kept, deletion dropped'
    Assert ($t.Contains("R1C1`tR1C2") -and $t.Contains("R2C1`tR2C2")) 'table rows tab-separated'
    Assert (([regex]::Matches($t, 'TextBox')).Count -eq 1) 'text box extracted once (VML fallback skipped)'
    Assert ($t.Contains('InContentControl')) 'content control text'
    Assert (-not $t.Contains('PAGE')) 'field codes skipped'
    Assert ($t.Contains('MyHeader') -and $t.Contains('MyFooter') -and $t.Contains('MyFootnote')) 'header/footer/footnote'
    Assert (-not $t.Contains('SEP')) 'footnote separator skipped'
    $bytes = [IO.File]::ReadAllBytes((Join-Path $out 'file05.txt'))
    Assert ($bytes[0] -ne 0xEF) 'no BOM by default'

    Write-Host "`n[2] no overwrite by default"
    $r = Run @('-Path', $src, '-OutDir', $out, '-Format', 'Txt')
    Assert ($r.Code -eq 1) "existing outputs -> failures, exit 1 (got $($r.Code))"
    $r = Run @('-Path', $src, '-OutDir', $out, '-Overwrite', '-Bom', '-Format', 'Txt')
    Assert ($r.Code -eq 0) '-Overwrite works'
    Assert ([IO.File]::ReadAllBytes((Join-Path $out 'file05.txt'))[0] -eq 0xEF) '-Bom writes BOM'

    Write-Host "`n[3] recurse keeps folder structure"
    $out2 = Join-Path $Tmp 'out2'
    $r = Run @('-Path', $src, '-Recurse', '-OutDir', $out2, '-NoExtras', '-Format', 'Txt')
    Assert ($r.Code -eq 0) 'exit 0'
    $nested = Join-Path (Join-Path $out2 'sub') 'nested.txt'
    Assert (Test-Path $nested) 'sub/nested.txt created'
    Assert (-not (ReadUtf8 $nested).Contains('MyHeader')) '-NoExtras drops headers'

    Write-Host "`n[4] combine to JSON"
    $json = Join-Path $Tmp 'all.json'
    $r = Run @('-Path', $src, '-Combine', $json, '-Format', 'Json')
    $data = ReadUtf8 $json | ConvertFrom-Json
    Assert ($r.Code -eq 0 -and @($data).Count -eq 20) '20 JSON records'
    Assert ($data[0].headers[0] -eq 'MyHeader' -and $data[0].body.Contains('世界')) 'JSON fields + Unicode'

    Write-Host "`n[5] partial failure"
    [IO.File]::WriteAllText((Join-Path $src 'broken.docx'), 'xx')
    $all = Join-Path $Tmp 'all.txt'
    $r = Run @('-Path', $src, '-Combine', $all, '-Format', 'Txt')
    Assert ($r.Code -eq 1) "exit 1 when a file fails (got $($r.Code))"
    Assert ($r.Out -match 'FAIL.*broken\.docx') 'failure reported with file name'
    Assert (([regex]::Matches((ReadUtf8 $all), '=====')).Count -eq 40) 'other 20 files still extracted'

    Write-Host "`n[5b] brackets / Chinese in file names (PowerShell wildcard chars)"
    $br = Join-Path $Tmp 'brackets'
    New-Docx (Join-Path $br 'Report [週報] 20260901.docx')
    $brOut = Join-Path $Tmp 'brout'
    $r = Run @('-Path', $br, '-OutDir', $brOut, '-Format', 'Txt')
    $brFile = Join-Path $brOut 'Report [週報] 20260901.txt'
    Assert ($r.Code -eq 0 -and [IO.File]::Exists($brFile)) 'bracket file name extracted'
    $r = Run @('-Path', $br, '-OutDir', $brOut, '-Format', 'Txt')
    Assert ($r.Code -eq 1) "existing bracket-named output not overwritten (got $($r.Code))"

    Write-Host "`n[6] bad input"
    $empty = Join-Path $Tmp 'empty'; New-Item -ItemType Directory $empty | Out-Null
    Assert ((Run @('-Path', $empty)).Code -eq 2) 'no Word files -> exit 2'
    Assert ((Run @('-Path', $src, '-OutDir', $out, '-Combine', $all)).Code -eq 2) '-OutDir + -Combine -> exit 2'

    Write-Host "`n[8] Markdown output (default format)"
    $mdSrc = Join-Path $Tmp 'mdsrc'
    New-Docx (Join-Path $mdSrc 'report.docx') $MdParts
    $mdOut = Join-Path $Tmp 'mdout'
    $r = Run @('-Path', $mdSrc, '-OutDir', $mdOut)
    $mdFile = Join-Path $mdOut 'report.md'
    Assert ($r.Code -eq 0 -and [IO.File]::Exists($mdFile)) 'default format writes .md'
    $expected = @(
        '# 總結', '', '\# 不是標題', '',
        '- 項目A', '   - 子項目', '1. 步驟一', '1. 步驟二', '', '| x |', '| --- |', '',
        '## 細節', '',
        '| 欄1 | 欄2 |', '| --- | --- |', '| a\|b<br>第二段 | c |', '| 只有一格 |  |', '',
        '1\. 不是清單', '', '第一行  ', '第二行', '', '---', '', '**Headers**', '', 'MyHeader', ''
    ) -join "`n"
    $actual = ReadUtf8 $mdFile
    Assert ($actual -eq $expected) 'headings (zh-TW style id, basedOn), nested/ordered lists, table, escaping, line break, extras'
    if ($actual -ne $expected) { Write-Host "---- expected ----`n$expected`n---- actual ----`n$actual" }

    $mdAll = Join-Path $Tmp 'md_all.md'
    $r = Run @('-Path', $mdSrc, '-Combine', $mdAll)
    $c = ReadUtf8 $mdAll
    Assert ($c.StartsWith("# report.docx`n`n## 總結") -and $c.Contains("`n### 細節")) 'combined: file name is H1, doc headings shifted down'

    $soffice = Get-Command soffice, soffice.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    $hasWord = $false
    if ($IsWindows -or $PSVersionTable.PSVersion.Major -le 5) {
        try { $w = New-Object -ComObject Word.Application; $w.Quit(); $hasWord = $true } catch {}
    }
    if ($soffice -or $hasWord) {
        Write-Host "`n[7] legacy .doc"
        $legacyDir = Join-Path $Tmp 'legacy'; New-Item -ItemType Directory $legacyDir | Out-Null
        if ($hasWord) {
            $w = New-Object -ComObject Word.Application
            $d = $w.Documents.Add(); $d.Content.Text = 'Legacy 舊格式 text'
            $d.SaveAs2((Join-Path $legacyDir 'legacy.doc'), 0); $d.Close(); $w.Quit()
        } else {
            $txt = Join-Path $legacyDir 'legacy.txt'
            [IO.File]::WriteAllText($txt, "Legacy 舊格式 text`n", (New-Object Text.UTF8Encoding($false)))
            & $soffice.Source --headless --convert-to doc --outdir $legacyDir $txt 2>&1 | Out-Null
            Remove-Item $txt
        }
        $lo = Join-Path $Tmp 'legacy.txt'
        $r = Run @('-Path', (Join-Path $legacyDir 'legacy.doc'), '-Combine', $lo, '-Format', 'Txt')
        Assert ($r.Code -eq 0 -and (ReadUtf8 $lo).Contains('Legacy 舊格式 text')) '.doc extracted'
    } else { Write-Host "`n[7] legacy .doc  SKIPPED (no Word / LibreOffice)" }
}
finally { Remove-Item $Tmp -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host "`n$script:Passed passed, $script:Failed failed"
exit [int]($script:Failed -gt 0)
