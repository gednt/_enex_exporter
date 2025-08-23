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

        $content = Get-Content $file.FullName -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
        if ($content) {
            if ($file.Extension.ToLower() -eq '.org') {
                Index-Org-Blocks -content $content -blockMap $blockMap
            } else {
                Index-Markdown-Blocks -content $content -blockMap $blockMap
            }
        }
    }
    Write-Host "Indexed $($pageMap.Count) pages and $($blockMap.Count) blocks."
}

function Index-Org-Blocks {
    param(
        [string]$content,
        [hashtable]$blockMap
    )
    
    # Split content into lines for processing
    $lines = $content -split "`r?`n"
    
    for ($i = 0; $i -lt $lines.Length; $i++) {
        $line = $lines[$i]
        
        # Look for :PROPERTIES: blocks that contain :id:
        if ($line -match '^\s*:PROPERTIES:\s*$') {
            $blockContent = ""
            $blockId = $null
            $startIdx = $i - 1  # The line before :PROPERTIES: is usually the content
            
            # Find the block ID in the properties
            $j = $i + 1
            while ($j -lt $lines.Length -and $lines[$j] -notmatch '^\s*:END:\s*$') {
                if ($lines[$j] -match '^\s*:id:\s*([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})\s*$') {
                    $blockId = $matches[1]
                    break
                }
                $j++
            }
            
            # If we found a block ID, extract the content
            if ($blockId -and $startIdx -ge 0) {
                # Get the content line (usually the line before :PROPERTIES:)
                $contentLine = $lines[$startIdx].TrimStart(' ', '*', '-').Trim()
                
                # For nested blocks, we might want to include child content too
                $childContent = @()
                $k = $j + 1  # Start after :END:
                $baseIndent = 0
                
                # Calculate base indentation level for org-mode
                if ($lines[$startIdx] -match '^(\s*\*+)') {
                    $baseIndent = $matches[1].Length
                }
                
                # Collect child content until we hit a same-level or higher-level heading
                while ($k -lt $lines.Length) {
                    $currentLine = $lines[$k]
                    
                    # Check if this is a heading line
                    if ($currentLine -match '^\s*\*+\s+' -or $currentLine -match '^\s*:PROPERTIES:\s*$') {
                        if ($currentLine -match '^(\s*\*+)') {
                            $currentIndent = $matches[1].Length
                            if ($currentIndent -le $baseIndent -and $currentLine.Trim() -ne '') {
                                break
                            }
                        }
                    }
                    
                    # Skip property lines and empty lines
                    if ($currentLine.Trim() -ne '' -and 
                        $currentLine -notmatch '^\s*:.*:\s*$' -and
                        $currentLine -notmatch '^\s*:PROPERTIES:\s*$' -and
                        $currentLine -notmatch '^\s*:END:\s*$') {
                        $childContent += $currentLine.TrimStart(' ', '*', '-').Trim()
                    }
                    $k++
                }
                
                # Combine main content with child content
                if ($childContent.Count -gt 0) {
                    $blockContent = $contentLine + "`n" + ($childContent -join "`n")
                } else {
                    $blockContent = $contentLine
                }
                
                if ($blockContent.Trim() -ne '' -and -not $blockMap.ContainsKey($blockId)) {
                    $blockMap[$blockId] = $blockContent.Trim()
                    Write-Host "Indexed block: $blockId -> $($blockContent.Substring(0, [Math]::Min(50, $blockContent.Length)))..."
                }
            }
        }
    }
}

function Index-Markdown-Blocks {
    param(
        [string]$content,
        [hashtable]$blockMap
    )
    
    # Pattern for markdown block IDs (usually at end of line)
    $blockIdPattern = [regex]'(?m)^(.*?)\s*\^\s*([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})\s*$'
    
    $matches = $blockIdPattern.Matches($content)
    foreach ($match in $matches) {
        $blockId = $match.Groups[2].Value
        $blockContent = $match.Groups[1].Value.Trim()
        
        if ($blockContent -ne '' -and -not $blockMap.ContainsKey($blockId)) {
            $blockMap[$blockId] = $blockContent
            Write-Host "Indexed block: $blockId -> $($blockContent.Substring(0, [Math]::Min(50, $blockContent.Length)))..."
        }
    }
}

function Resolve-LogseqPlaceholders {
    param(
        [string]$text,
        [hashtable]$blockMap
    )
    
    Write-Host "Resolving Logseq placeholders..."
    
    # Remove BOM if present (UTF-8 BOM is EF BB BF)
    if ($text.StartsWith([char]0xFEFF)) {
        $text = $text.Substring(1)
    }
    
    # First, resolve embed blocks {{embed ((block-id))}}
    $embedPattern = '\{\{embed\s+\(\(([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})\)\)\}\}'
    $resolvedText = [regex]::Replace($text, $embedPattern, {
        param($match)
        $blockId = $match.Groups[1].Value
        if ($blockMap.ContainsKey($blockId)) {
            $replacement = $blockMap[$blockId]
            Write-Host "Resolved embed block: $blockId -> $($replacement.Substring(0, [Math]::Min(30, $replacement.Length)))..."
            return $replacement
        } else {
            Write-Warning "Embed block not found: $blockId"
            return "**[Embed Block Not Found: $blockId]**"
        }
    })
    
    # Then resolve regular block references ((block-id))
    $blockRefPattern = '\(\(([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})\)\)'
    $maxDepth = 10
    $currentDepth = 0
    
    while ($currentDepth -lt $maxDepth) {
        $match = [regex]::Match($resolvedText, $blockRefPattern)
        if (-not $match.Success) {
            break
        }
        
        $blockId = $match.Groups[1].Value
        if ($blockMap.ContainsKey($blockId)) {
            $replacement = $blockMap[$blockId]
            Write-Host "Resolved block reference: $blockId -> $($replacement.Substring(0, [Math]::Min(30, $replacement.Length)))..."
            $resolvedText = $resolvedText.Remove($match.Index, $match.Length).Insert($match.Index, $replacement)
        } else {
            Write-Warning "Block reference not found: $blockId"
            $resolvedText = $resolvedText.Remove($match.Index, $match.Length).Insert($match.Index, "**[Block Not Found: $blockId]**")
        }
        $currentDepth++
    }
    
    if ($currentDepth -eq $maxDepth) {
        Write-Warning "Maximum recursion depth reached while resolving block references. Some references may remain unresolved."
    }

    # Handle page references [[Page Name]]
    $pageRefPattern = '\[\[([^\]]+)\]\]'
    $resolvedText = [regex]::Replace($resolvedText, $pageRefPattern, {
        param($match)
        $pageName = $match.Groups[1].Value
        # Convert to a simple link format for better readability
        return "**$pageName**"
    })

    # Handle other Logseq-specific syntax
    # Remove or convert property blocks (but keep them single-line if they're meant to be visible)
    $resolvedText = [regex]::Replace($resolvedText, '(?m)^\s*:PROPERTIES:.*?:END:\s*$', '', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    
    # Handle cloze deletions {{cloze text}}
    $resolvedText = [regex]::Replace($resolvedText, '\{\{cloze\s+(.*?)\}\}', '**$1**')
    
    # Clean up any remaining double embeds that might have been missed
    $resolvedText = [regex]::Replace($resolvedText, '\{\{embed\s+(.*?)\}\}', '$1')
    
    # Clean up org-mode specific syntax that doesn't translate well
    # Remove collapsed property
    $resolvedText = [regex]::Replace($resolvedText, '(?m)^\s*:collapsed:\s*true\s*$', '')
    
    # Remove heading property
    $resolvedText = [regex]::Replace($resolvedText, '(?m)^\s*:heading:\s*\d+\s*$', '')
    
    # Remove id property lines that might have leaked through
    $resolvedText = [regex]::Replace($resolvedText, '(?m)^\s*:id:\s*[a-f0-9-]+\s*$', '')
    
    # Clean up any trailing whitespace and multiple newlines
    $resolvedText = [regex]::Replace($resolvedText, '\n\s*\n\s*\n+', "`n`n")
    $resolvedText = $resolvedText.Trim()

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
                $content = Get-Content $file.FullName -Raw -Encoding UTF8
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
            
            $htmlContent = Get-Content $tempHtmlFile -Raw -Encoding UTF8
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

