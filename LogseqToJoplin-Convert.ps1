param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$InputDir,

    [Parameter(Mandatory=$false)]
    [ValidateSet('Logseq', 'Other')]
    [string]$DirectoryType = 'Logseq'
)

# --- Global State ---
$pageMap = @{}
$blockMap = @{}

#region "Logseq Content Processing"

function Index-Logseq-Content {
    param(
        [string]$LogseqDir,
        [hashtable]$pageMap,
        [hashtable]$blockMap
    )
    Write-Host "--- Indexing Logseq content... ---"
    $blockIdPattern = [regex]'(?i)^\s*id::\s*([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})'
    
    $filesToIndex = Get-ChildItem -Path $LogseqDir -Recurse -File | Where-Object { $_.Extension -in '.md', '.org' }
    
    $bakDir = Join-Path $LogseqDir "logseq\bak"
    if (Test-Path $bakDir) {
        $filesToIndex += Get-ChildItem -Path $bakDir -Recurse -File | Where-Object { $_.Extension -in '.md', '.org' }
    }

    foreach ($file in $filesToIndex) {
        $pageName = $file.BaseName
        if (-not $pageMap.ContainsKey($pageName)) {
            $pageMap[$pageName] = $file.FullName
        }

        $lines = Get-Content $file.FullName -ErrorAction SilentlyContinue
        for ($i = 0; $i -lt $lines.Length; $i++) {
            $line = $lines[$i]
            if (($i + 1) -lt $lines.Length) {
                $nextLine = $lines[$i+1]
                $match = $blockIdPattern.Match($nextLine)
                if ($match.Success) {
                    $blockId = $match.Groups[1].Value
                    $blockContent = $line.TrimStart(' ', '-', '*')
                    if (-not $blockMap.ContainsKey($blockId)) {
                        $blockMap[$blockId] = $blockContent
                    }
                }
            }
        }
    }
    Write-Host "Indexed $($pageMap.Count) pages and $($blockMap.Count) blocks."
}

function Resolve-LogseqPlaceholders {
    param(
        [string]$text,
        [hashtable]$blockMap
    )
    $embedPattern = '\{\{embed\s+(.*?)\}\}'
    $resolvedText = [regex]::Replace($text, $embedPattern, { param($m) $m.Groups[1].Value })

    $blockRefPattern = '\(\((.*?)\)\)'
    $maxDepth = 10
    $currentDepth = 0
    while (($match = [regex]::Match($resolvedText, $blockRefPattern)).Success -and $currentDepth -lt $maxDepth) {
        $blockId = $match.Groups[1].Value
        if ($blockMap.ContainsKey($blockId)) {
            $replacement = $blockMap[$blockId]
            $resolvedText = $resolvedText.Remove($match.Index, $match.Length).Insert($match.Index, $replacement)
        } else {
            $resolvedText = $resolvedText.Remove($match.Index, $match.Length).Insert($match.Index, "((NOT_FOUND: $blockId))")
        }
        $currentDepth++
    }

    $pageRefPattern = '\[\[([^\]]+)\]\]'
    $resolvedText = [regex]::Replace($resolvedText, $pageRefPattern, { param($m) $m.Groups[1].Value })

    return $resolvedText
}

#endregion

#region "ENEX XML Creation Functions"

function Create-Enex-Note {
    param(
        [string]$title,
        [string]$content,
        [hashtable]$resources,
        [string]$createdAt,
        [string]$updatedAt,
        [array]$tags
    )
    # Ensure title and content are properly escaped for XML
    $escapedTitle = [System.Security.SecurityElement]::Escape($title)
    
    $noteXml = @"
    <note>
        <title>$escapedTitle</title>
        <content>
            <![CDATA[<?xml version="1.0" encoding="UTF-8" standalone="no"?>
            <!DOCTYPE en-note SYSTEM "http://xml.evernote.com/pub/enml2.dtd">
            <en-note>
                $content
            </en-note>
            ]]>
        </content>
        <created>$createdAt</created>
        <updated>$updatedAt</updated>
"@
    foreach ($tag in $tags) {
        $escapedTag = [System.Security.SecurityElement]::Escape($tag)
        $noteXml += "        <tag>$escapedTag</tag>`n"
    }
    foreach ($hash in $resources.Keys) {
        $noteXml += $resources[$hash]
    }
    $noteXml += "    </note>"
    return $noteXml
}

function Process-Images-For-Enex {
    param(
        [string]$htmlContent,
        [string]$fileDirectory,
        [ref]$resources,
        [string]$mediaDirectory # New parameter for extracted media
    )
    $imagePattern = '<img src="(.*?)"'
    $updatedContent = [regex]::Replace($htmlContent, $imagePattern, {
        param($match)
        $imagePath = $match.Groups[1].Value

        if ($imagePath.StartsWith("http")) { return $match.Value }

        $absoluteImagePath = $imagePath
        if (-not [System.IO.Path]::IsPathRooted($imagePath)) {
            # Prefer mediaDirectory if provided, else fallback to fileDirectory
            $basePath = if ($mediaDirectory) { $mediaDirectory } else { $fileDirectory }
            $absoluteImagePath = Join-Path -Path $basePath -ChildPath $imagePath
        }

        if (Test-Path $absoluteImagePath) {
            try {
                $md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
                $fileStream = [System.IO.File]::OpenRead($absoluteImagePath)
                $hashBytes = $md5.ComputeHash($fileStream)
                $fileStream.Close()
                $hashString = [System.BitConverter]::ToString($hashBytes).Replace("-", "").ToLower()

                if (-not $resources.Value.ContainsKey($hashString)) {
                    $fileInfo = Get-Item $absoluteImagePath
                    $extension = $fileInfo.Extension.TrimStart('.')
                    $mimeType = switch ($extension.ToLower()) {
                        'jpg'  { 'image/jpeg' }
                        'jpeg' { 'image/jpeg' }
                        'png'  { 'image/png' }
                        'gif'  { 'image/gif' }
                        'svg'  { 'image/svg+xml' }
                        'bmp'  { 'image/bmp' }
                        'webp' { 'image/webp' }
                        default { 'application/octet-stream' }
                    }
                    
                    $fileBytes = [System.IO.File]::ReadAllBytes($absoluteImagePath)
                    $base64String = [System.Convert]::ToBase64String($fileBytes)

                    $resourceXml = @"
        <resource>
            <data encoding="base64">$base64String</data>
            <mime>$mimeType</mime>
            <resource-attributes>
                <file-name>$($fileInfo.Name)</file-name>
            </resource-attributes>
        </resource>
"@
                    $resources.Value[$hashString] = $resourceXml
                    Write-Host "Processed resource: $($fileInfo.Name) -> $hashString"
                }
                
                return "<en-media type=`"$($mimeType)`" hash=`"$hashString`" />"
            } catch {
                Write-Warning "Could not process image: $absoluteImagePath. Error: $_"
                return $match.Value
            }
        } else {
            Write-Warning "Image not found at path: $absoluteImagePath"
            return $match.Value
        }
    })
    return $updatedContent
}

#endregion

# --- Main Script ---

# Check for pandoc
if (-not (Get-Command pandoc -ErrorAction SilentlyContinue)) {
    Write-Host "pandoc is not installed. Please install it first." -ForegroundColor Red
    exit 1
}

# Normalize and check folder
$InputDir = (Resolve-Path $InputDir).Path.TrimEnd('\')
if (-not (Test-Path $InputDir)) {
    Write-Host "Error: Folder '$InputDir' does not exist or is empty." -ForegroundColor Red
    exit 1
}

# Output file
$outputFile = "${InputDir}_evernote_export.enex"
if (Test-Path $outputFile) {
    Write-Host "Output file '$outputFile' already exists. Removing it."
    Remove-Item -Force $outputFile
}

# --- Start ENEX file ---
$enexHeader = @"
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE en-export SYSTEM "http://xml.evernote.com/pub/evernote-export3.dtd">
<en-export export-date="$((Get-Date).ToString("yyyyMMdd'T'HHmmss'Z'"))" application="LogseqToJoplinScript" version="2.0">
"@
Add-Content -Path $outputFile -Value $enexHeader -Encoding UTF8

# --- Index Logseq content if applicable ---
if ($DirectoryType -eq 'Logseq') {
    Index-Logseq-Content -LogseqDir $InputDir -pageMap $pageMap -blockMap $blockMap
}

# --- Process files ---
$extensions = @('.md', '.org')
if ($DirectoryType -eq 'Other') {
    $extensions += '.docx', '.odt'
}
$files = Get-ChildItem -Path $InputDir -Recurse -File | Where-Object { $_.Extension -in $extensions }

$total = $files.Count
$success = 0
$failed = 0

Write-Host "`n--- Processing files for Evernote ENEX format...`n"

foreach ($file in $files) {
    Write-Host "------------------------------------------------------------"
    Write-Host "Processing [$($success+$failed+1)/$total]: $($file.FullName.Substring($InputDir.Length))"

    try {
        # --- Determine Title and Tags from Folder Path ---
        $parts = $file.BaseName -split '___'
        $noteTitle = $parts[-1]
        $tags = @()
        if ($parts.Count -gt 1) {
            $tagPath = $parts[0..($parts.Count-2)] -join '/'
            $tags += $tagPath
        }
        
        Write-Host "Note title: $noteTitle"
        if ($tags.Count -gt 0) {
            Write-Host "Tags: $($tags -join ', ')"
        }

        # --- Read and Convert Content ---
        $fromFormat = switch ($file.Extension.ToLower()) {
            '.md'   { 'markdown' }
            '.org'  { 'org' }
            '.docx' { 'docx' }
            '.odt'  { 'odt' }
        }

        $htmlContent = ""
        $tempHtmlFile = [System.IO.Path]::GetTempFileName()
        $mediaDir = $null

        try {
            if ($DirectoryType -eq 'Logseq') {
                $content = Get-Content $file.FullName -Raw
                $content = Resolve-LogseqPlaceholders -text $content -blockMap $blockMap
                
                $tmpFile = [System.IO.Path]::GetTempFileName()
                Set-Content -Path $tmpFile -Value $content -Encoding UTF8
                
                & pandoc $tmpFile -f $fromFormat -t html --wrap=none -o "$tempHtmlFile"
                Remove-Item $tmpFile -Force
            } else { # For 'Other' type, convert directly
                if ($file.Extension.ToLower() -in @('.docx', '.odt')) {
                    $mediaDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.IO.Path]::GetRandomFileName())
                    New-Item -ItemType Directory -Path $mediaDir | Out-Null
                    & pandoc $file.FullName -f $fromFormat -t html --wrap=none --extract-media=$mediaDir -o "$tempHtmlFile"
                } else {
                    & pandoc $file.FullName -f $fromFormat -t html --wrap=none -o "$tempHtmlFile"
                }
            }

            if ($LASTEXITCODE -ne 0) {
                throw "Pandoc conversion failed for $($file.Name)"
            }
            
            $htmlContent = Get-Content $tempHtmlFile -Raw
            Remove-Item $tempHtmlFile -Force

            # --- Process Images ---
            $noteResources = @{}
            $htmlContent = Process-Images-For-Enex -htmlContent $htmlContent -fileDirectory $file.DirectoryName -resources ([ref]$noteResources) -mediaDirectory $mediaDir

            # --- Create and Append ENEX Note ---
            $createdAt = $file.CreationTime.ToString("yyyyMMdd'T'HHmmss'Z'")
            $updatedAt = $file.LastWriteTime.ToString("yyyyMMdd'T'HHmmss'Z'")
            $noteXml = Create-Enex-Note -title $noteTitle -content $htmlContent -resources $noteResources -createdAt $createdAt -updatedAt $updatedAt -tags $tags
            
            Add-Content -Path $outputFile -Value $noteXml -Encoding UTF8

            $success++
            Write-Host "Successfully converted note: $noteTitle" -ForegroundColor Green

        } finally {
            if ($mediaDir -and (Test-Path $mediaDir)) { Remove-Item $mediaDir -Recurse -Force }
        }

    } catch {
        $failed++
        Write-Host "Error processing $($file.FullName): $_" -ForegroundColor Red
    }
}

# --- End ENEX file ---
Add-Content -Path $outputFile -Value "</en-export>" -Encoding UTF8

Write-Host "`n=== Evernote ENEX Export Summary ==="
Write-Host "Total files processed: $total"
Write-Host "Successfully converted: $success"
Write-Host "Failed conversions: $failed"
Write-Host "Output location: $outputFile"
Write-Host ""
Write-Host "To import into Joplin (or Evernote):"
Write-Host "1. In the application, go to File > Import > ENEX"
Write-Host "2. Choose the created ENEX file: $outputFile"
Write-Host ""

