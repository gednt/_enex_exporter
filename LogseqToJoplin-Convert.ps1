param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$InputDir,

    [Parameter(Mandatory=$false)]
    [ValidateSet('Logseq', 'Other')]
    [string]$DirectoryType = 'Logseq',

    [Parameter(Mandatory=$false)]
    [switch]$CreateFolderStructure
)

# --- Global State ---
$pageMap = @{}
$blockMap = @{}

#region "Logseq Content Processing"

function Extract-Tags-From-Content {
    param(
        [string]$content
    )
    
    $tags = @()
    
    # Extract [[tag]] style tags (including hierarchical ones like [[Tag1/subtag/tag]])
    $bracketTagPattern = '\[\[([^\]]+)\]\]'
    $bracketMatches = [regex]::Matches($content, $bracketTagPattern)
    foreach ($match in $bracketMatches) {
        $raw = $match.Groups[1].Value.Trim()
        # Heuristic: treat [[...]] as a tag only when it's hierarchical (contains '/')
        # or is a single-token (no whitespace). This avoids treating normal page links like
        # [[Some Page Name]] as tags. Strip any leading '#' if present.
        if (-not $raw) { continue }
        if ($raw -match '/') {
            $tag = $raw
        } elseif (-not ($raw -match '\s')) {
            $tag = $raw
        } else {
            continue
        }
        $tag = $tag.TrimStart('#').Trim()
        if ($tag -and $tag -notin $tags) { $tags += $tag }
    }
    
    # Extract #hashtag style tags (including hierarchical ones like #tag1/subtag/tag)
    $hashtagPattern = '#([a-zA-Z0-9_/-]+)'
    $hashtagMatches = [regex]::Matches($content, $hashtagPattern)
    foreach ($match in $hashtagMatches) {
        $tag = $match.Groups[1].Value.Trim()
        if ($tag -and $tag -notin $tags) {
            $tags += $tag
        }
    }
    
    # Extract tags from properties blocks (for org-mode and markdown)
    $tagPropertyPattern = '(?m)^\s*#?\+?tags?::\s*(.+)$'
    $tagPropertyMatches = [regex]::Matches($content, $tagPropertyPattern)
    foreach ($match in $tagPropertyMatches) {
        $tagLine = $match.Groups[1].Value
        # Parse tags that might be in [[tag]] format
        $tagRefs = [regex]::Matches($tagLine, '\[\[([^\]]+)\]\]')
        foreach ($tagRef in $tagRefs) {
            $tag = $tagRef.Groups[1].Value.Trim()
            if ($tag -and $tag -notin $tags) {
                $tags += $tag
            }
        }
        # Also parse plain comma-separated tags
        $plainTags = $tagLine -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -notmatch '\[\[' }
        foreach ($plainTag in $plainTags) {
            if ($plainTag -and $plainTag -notin $tags) {
                $tags += $plainTag
            }
        }
    }
    
    # Extract org-mode header tags
    $orgTagPattern = '(?m)^\s*#\+tags:\s*(.+)$'
    $orgTagMatches = [regex]::Matches($content, $orgTagPattern)
    foreach ($match in $orgTagMatches) {
        $tagLine = $match.Groups[1].Value
        $orgTags = $tagLine -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        foreach ($orgTag in $orgTags) {
            if ($orgTag -and $orgTag -notin $tags) {
                $tags += $orgTag
            }
        }
    }
    
    return $tags
}

function Convert-Tag-To-Path {
    param(
        [string]$tag
    )
    
    # Convert hierarchical tags (Tag1/subtag/tag) to folder path
    $path = $tag -replace '/', '\'
    
    # Clean up path components to be filesystem-safe
    $pathParts = $path -split '\\' | ForEach-Object {
        $part = $_.Trim()
        # Remove or replace invalid filesystem characters
        $part = $part -replace '[<>:"|?*]', '-'
        $part = $part -replace '[\s]+', ' '  # Normalize spaces
        return $part
    } | Where-Object { $_ }
    
    return $pathParts -join '\'
}

function Fix-Asset-References {
    param(
        [string]$content,
        [string]$sourceDir,
        [string]$destDir,
        [string]$outputDir,
        [string]$top,
        [string]$extractedMediaDir
    )

    # Destination assets folder for this specific file (per-file assets)
    $destAssets = Join-Path $destDir 'assets'

    # Helper to compute relative path from one dir to target
    function Get-Relative([string]$fromDir, [string]$toPath) {
        try {
            $from = (Resolve-Path $fromDir).ProviderPath
            $to = (Resolve-Path $toPath).ProviderPath
            $fromUri = New-Object System.Uri($from + [System.IO.Path]::DirectorySeparatorChar)
            $toUri = New-Object System.Uri($to)
            $rel = $fromUri.MakeRelativeUri($toUri).ToString()
            $rel = [System.Uri]::UnescapeDataString($rel)
            return $rel -replace '/','\\'
        } catch {
            return $toPath
        }
    }

    # Resolve a referenced url/path to a new relative path from destDir
    function Resolve-NewUrl([string]$url) {
        if (-not $url) { return $url }
        $u = $url.Trim()
        if ($u -match '^(https?:\\/\\/|file:|\\/|[A-Za-z]:\\)') { return $u }

        # Normalize leading ./
        $candidate = $u -replace '^[\\./]+',''

        # Candidate absolute path relative to sourceDir
        $abs = $candidate
        if (-not [System.IO.Path]::IsPathRooted($abs)) { $abs = Join-Path $sourceDir $candidate }

        $fileName = [System.IO.Path]::GetFileName($candidate)

        # Priority: extractedMediaDir (from pandoc) -> sourceDir (relative to page) -> input assets under source tree -> search outputDir
        $sourceMatch = $null
        if ($extractedMediaDir -and (Test-Path $extractedMediaDir)) {
            $sourceMatch = Get-ChildItem -Path $extractedMediaDir -Recurse -File -Filter $fileName -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($sourceMatch) { $abs = $sourceMatch.FullName }
        }
        if (-not $sourceMatch -and (Test-Path $abs)) { $sourceMatch = Get-Item -Path $abs -ErrorAction SilentlyContinue }
        if (-not $sourceMatch) {
            # Try to find in original input assets nearby
            $candidateUnderInput = Get-ChildItem -Path $sourceDir -Recurse -File -Filter $fileName -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($candidateUnderInput) { $sourceMatch = $candidateUnderInput; $abs = $sourceMatch.FullName }
        }
        if (-not $sourceMatch) {
            # Last resort: any file under outputDir with same name (rare but possible)
            $candidateOut = Get-ChildItem -Path $outputDir -Recurse -File -Filter $fileName -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($candidateOut) { $sourceMatch = $candidateOut; $abs = $sourceMatch.FullName }
        }
        if (-not $sourceMatch) { return $u }

        # Ensure destination assets folder exists and copy the asset in (do not overwrite if exists)
        if (-not (Test-Path $destAssets)) { New-Item -ItemType Directory -Path $destAssets -Force | Out-Null }
        $destFile = Join-Path $destAssets $fileName
        if (-not (Test-Path $destFile)) {
            try { Copy-Item -Path $abs -Destination $destFile -Force -ErrorAction SilentlyContinue } catch { }
        }

        # Return a consistent local assets path: assets/<url-encoded-filename>
        $fileNameOnly = [System.IO.Path]::GetFileName($destFile)
        try {
            $encoded = [System.Uri]::EscapeDataString($fileNameOnly)
        } catch { $encoded = $fileNameOnly }
        return "assets/$encoded"
    }

    $fixed = $content

    # Markdown links/images: capture optional leading '!' (image), alt text and url
    $mdPattern = '(\\?)(?<bang>!?)\[(?<alt>[^\]]*)\]\((?<url>[^)]+)\)'
    $fixed = [regex]::Replace($fixed, $mdPattern, {
        param($m)
        # $m.Groups[1] is the optional backslash before bang, we want to skip it (remove it)
        $bang = $m.Groups['bang'].Value
        $alt = $m.Groups['alt'].Value
        $url = $m.Groups['url'].Value

        # Clean alt text: remove stray backslashes used for escaping and trim
        $cleanAlt = $alt -replace '\\',''
        $cleanAlt = $cleanAlt.Trim()

        $newUrl = Resolve-NewUrl $url
        return "$bang[$cleanAlt]($newUrl)"
    })

    # HTML img src
    $imgPattern = '<img\s+[^>]*src\s*=\s*["''](?<url>[^"'']+)["''][^>]*>'
    $fixed = [regex]::Replace($fixed, $imgPattern, {
        param($m)
        $url = $m.Groups['url'].Value
        $newUrl = Resolve-NewUrl $url
        return ($m.Value -replace [regex]::Escape($url), [System.Text.RegularExpressions.Regex]::Escape($newUrl)) -replace '\\Q','' -replace '\\E',''
    })

    # Logseq/wiki image links: ![[...]]
    $wikiImgPattern = '!\[\[([^\]]+)\]\]'
    $fixed = [regex]::Replace($fixed, $wikiImgPattern, {
        param($m)
        $inner = $m.Groups[1].Value
        $newInner = Resolve-NewUrl $inner
        return "![[${newInner}]]"
    })

    # Wiki links that point to assets
    $wikiLinkPattern = '\[\[([^\]]+)\]\]'
    $fixed = [regex]::Replace($fixed, $wikiLinkPattern, {
        param($m)
        $inner = $m.Groups[1].Value
        if ($inner -match '(?i)assets[\\/]|\.(png|jpg|jpeg|gif|svg|webp|bmp)$') {
            $newInner = Resolve-NewUrl $inner
            return "[[${newInner}]]"
        }
        return $m.Value
    })

    # Cleanup: remove any escaping backslash that was left before image tokens (e.g. \![...])
    try {
        $fixed = $fixed -replace '\\\\!\\\[\\\[', '![['    # \![[ -> ![[
        $fixed = $fixed -replace '\\\\!\\\[', '!['          # \![ -> ![
    } catch { }

    return $fixed
}

function Create-Folder-Structure-From-Tags {
    param(
        [string]$inputDir,
        [string]$outputDir
    )
    
    if ($DirectoryType -eq 'Logseq') {
        Write-Host "--- Creating folder structure from tags (Logseq mode)... ---"
    } else {
        Write-Host "--- Creating folder structure preserving original hierarchy (Other mode)... ---"
    }
    
    # Get all files to process
    $extensions = @('.md', '.org')
    if ($DirectoryType -eq 'Other') {
        $extensions += @('.docx', '.odt')
    }
    $files = Get-ChildItem -Path $inputDir -Recurse -File | Where-Object { $_.Extension -in $extensions }
    
    # Copy any 'assets' directories found under input to the corresponding top-level output folder
    try {
        $assetDirs = Get-ChildItem -Path $inputDir -Recurse -Directory | Where-Object { $_.Name -ieq 'assets' }
        foreach ($ad in $assetDirs) {
            $rel = $ad.FullName.Substring($inputDir.Length).TrimStart('\', '/')
            $parts = $rel -split '[\\/]'
            $top = if ($parts.Length -gt 1) { $parts[0] } else { '' }
            $destBase = if ($top -and $top -ne '') { Join-Path $outputDir $top } else { $outputDir }
            $destAssets = Join-Path $destBase 'assets'

            # Recreate destination assets folder (clean copy)
            if (Test-Path $destAssets) { Remove-Item -Path $destAssets -Recurse -Force -ErrorAction SilentlyContinue }
            New-Item -ItemType Directory -Path $destAssets -Force | Out-Null

            # Copy contents preserving original formats (don't convert)
            $sourcePattern = Join-Path $ad.FullName '*'
            Copy-Item -Path $sourcePattern -Destination $destAssets -Recurse -Force -ErrorAction SilentlyContinue

            Write-Host "Copied assets: $($ad.FullName) -> $($destAssets.Substring($outputDir.Length + 1))"
        }
    } catch {
        Write-Warning "Failed to copy assets folders: $_"
    }

    $processedFiles = 0
    $totalFiles = $files.Count

    if ($DirectoryType -eq 'Logseq') {
        # First pass: build map of top-level input folders -> tags used under them
        $topTagMap = @{}
        foreach ($file in $files) {
            try {
                $relativePath = $file.FullName.Substring($inputDir.Length).TrimStart('\', '/')
                $parts = $relativePath -split '[\\/]'
                $top = if ($parts.Length -gt 1) { $parts[0] } else { '' }

                $content = Get-Content $file.FullName -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
                if ($content) {
                    $tags = Extract-Tags-From-Content -content $content
                    if (-not $topTagMap.ContainsKey($top)) { $topTagMap[$top] = [System.Collections.Generic.HashSet[string]]::new() }
                    foreach ($t in $tags) { $topTagMap[$top].Add($t) | Out-Null }
                }
            } catch { }
        }

        # Create tag folders under each top-level root (or root when top is empty)
        foreach ($top in $topTagMap.Keys) {
            foreach ($t in $topTagMap[$top]) {
                $folderPath = Convert-Tag-To-Path -tag $t
                $base = if ($top -and $top -ne '') { Join-Path $outputDir $top } else { $outputDir }
                $fullFolder = if ($folderPath -and $folderPath -ne '') { Join-Path $base $folderPath } else { $base }
                if (-not (Test-Path $fullFolder)) {
                    New-Item -ItemType Directory -Path $fullFolder -Force | Out-Null
                    $display = if ($base.Length -lt $outputDir.Length) { $fullFolder } else { $fullFolder.Substring($outputDir.Length + 1) }
                    Write-Host "Created tag folder: $($fullFolder.Substring($outputDir.Length + 1))"
                }
            }
        }

        # Second pass: write files into their corresponding top/tag folders (bottom level)
        foreach ($file in $files) {
            $processedFiles++
            Write-Host "Processing [$processedFiles/$totalFiles]: $($file.Name)"
            try {
                $relativePath = $file.FullName.Substring($inputDir.Length).TrimStart('\', '/')
                $parts = $relativePath -split '[\\/]'
                $top = if ($parts.Length -gt 1) { $parts[0] } else { '' }

                $content = Get-Content $file.FullName -Raw -Encoding UTF8
                $tags = Extract-Tags-From-Content -content $content

                # Resolve placeholders for Logseq
                $processedContent = Resolve-LogseqPlaceholders -text $content -blockMap $blockMap

                # Convert format if needed (org -> md) and extract media so we can copy referenced files into per-file assets
                $extractedMediaDir = $null
                if ($file.Extension.ToLower() -eq '.org') {
                    $tmpOrg = [System.IO.Path]::GetTempFileName()
                    $tmpMd = [System.IO.Path]::GetTempFileName()
                    $mediaDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.IO.Path]::GetRandomFileName())
                    try {
                        New-Item -ItemType Directory -Path $mediaDir | Out-Null
                        Set-Content -Path $tmpOrg -Value $processedContent -Encoding UTF8
                        & pandoc $tmpOrg -f org -t markdown --wrap=none --extract-media="$mediaDir" -o $tmpMd
                        if ($LASTEXITCODE -eq 0) { $processedContent = Get-Content $tmpMd -Raw -Encoding UTF8; $extractedMediaDir = $mediaDir }
                    } finally {
                        Remove-Item $tmpOrg -Force -ErrorAction SilentlyContinue
                        Remove-Item $tmpMd -Force -ErrorAction SilentlyContinue
                        # Do not remove $mediaDir here; we'll use it when copying assets per file and remove later
                    }
                }

                if ($tags.Count -eq 0) {
                    $destBase = if ($top -and $top -ne '') { Join-Path $outputDir $top } else { $outputDir }
                    if (-not (Test-Path $destBase)) { New-Item -ItemType Directory -Path $destBase -Force | Out-Null }
                    $outPath = Join-Path $destBase "$($file.BaseName).md"
                    # Fix asset links so they point to the per-file assets folder; pass extractedMediaDir when available
                    $fixedContent = Fix-Asset-References -content $processedContent -sourceDir $file.DirectoryName -destDir (Split-Path -Parent $outPath) -outputDir $outputDir -top $top -extractedMediaDir $extractedMediaDir
                    Set-Content -Path $outPath -Value $fixedContent -Encoding UTF8
                    Write-Host "  → Created: $($outPath.Substring($outputDir.Length + 1))" -ForegroundColor Green
                } else {
                    foreach ($t in $tags) {
                        $folderPath = Convert-Tag-To-Path -tag $t
                        $destBase = if ($top -and $top -ne '') { Join-Path $outputDir $top } else { $outputDir }
                        $destDir = if ($folderPath -and $folderPath -ne '') { Join-Path $destBase $folderPath } else { $destBase }
                        if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
                        $outPath = Join-Path $destDir "$($file.BaseName).md"
                        if (Test-Path $outPath) {
                            $existing = Get-Content $outPath -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
                            if ($existing -ne $processedContent) {
                                $suffix = 1
                                do { $candidate = Join-Path $destDir ("$($file.BaseName)_$suffix.md"); $suffix++ } while (Test-Path $candidate)
                                $outPath = $candidate
                            }
                        }
                        # Fix asset links for files placed under tag folders; pass extractedMediaDir when available
                        $fixedContent = Fix-Asset-References -content $processedContent -sourceDir $file.DirectoryName -destDir $destDir -outputDir $outputDir -top $top -extractedMediaDir $extractedMediaDir
                        Set-Content -Path $outPath -Value $fixedContent -Encoding UTF8
                        Write-Host "  → Created: $($outPath.Substring($outputDir.Length + 1))" -ForegroundColor Green
                    }
                }
                # Cleanup extracted media directory for this file if it exists
                if ($extractedMediaDir -and (Test-Path $extractedMediaDir)) {
                    try { Remove-Item -Path $extractedMediaDir -Recurse -Force -ErrorAction SilentlyContinue } catch { }
                }
            } catch {
                Write-Warning "Error processing $($file.Name): $_"
            }
        }

    } else {
        # Other mode: preserve original folder structure (existing behavior)
        foreach ($file in $files) {
            $processedFiles++
            Write-Host "Processing [$processedFiles/$totalFiles]: $($file.Name)"
            try {
                $content = Get-Content $file.FullName -Raw -Encoding UTF8
                $relativePath = $file.FullName.Substring($inputDir.Length).TrimStart('\', '/')
                $relativeDir = [System.IO.Path]::GetDirectoryName($relativePath)
                if ($relativeDir) {
                    $fullOutputDir = Join-Path $outputDir $relativeDir
                    if (-not (Test-Path $fullOutputDir)) { New-Item -ItemType Directory -Path $fullOutputDir -Force | Out-Null }
                    $outputPath = Join-Path $fullOutputDir "$($file.BaseName).md"
                } else {
                    $outputPath = Join-Path $outputDir "$($file.BaseName).md"
                }

                # Convert formats if needed
                if ($file.Extension.ToLower() -eq '.org') {
                    $tmpOrg = [System.IO.Path]::GetTempFileName()
                    $tmpMd = [System.IO.Path]::GetTempFileName()
                    try {
                        Set-Content -Path $tmpOrg -Value $content -Encoding UTF8
                        & pandoc $tmpOrg -f org -t markdown --wrap=none -o $tmpMd
                        if ($LASTEXITCODE -eq 0) { $content = Get-Content $tmpMd -Raw -Encoding UTF8 } else { Write-Warning "Failed to convert org file: $($file.Name)" }
                    } finally { Remove-Item $tmpOrg -Force -ErrorAction SilentlyContinue; Remove-Item $tmpMd -Force -ErrorAction SilentlyContinue }
                } elseif ($file.Extension.ToLower() -in @('.docx', '.odt')) {
                    $tmpMd = [System.IO.Path]::GetTempFileName()
                    try {
                        $fromFormat = if ($file.Extension.ToLower() -eq '.docx') { 'docx' } else { 'odt' }
                        & pandoc $file.FullName -f $fromFormat -t markdown --wrap=none -o $tmpMd
                        if ($LASTEXITCODE -eq 0) { $content = Get-Content $tmpMd -Raw -Encoding UTF8 } else { Write-Warning "Failed to convert $($file.Extension) file: $($file.Name)"; continue }
                    } finally { Remove-Item $tmpMd -Force -ErrorAction SilentlyContinue }
                }

                Set-Content -Path $outputPath -Value $content -Encoding UTF8
                Write-Host "  → Created: $($outputPath.Substring($outputDir.Length + 1))" -ForegroundColor Green
            } catch {
                Write-Warning "Error processing $($file.Name): $_"
            }
        }
    }
    
    # Cleanup: remove any empty directories left behind (deepest first)
    try {
        $dirs = Get-ChildItem -Path $outputDir -Recurse -Directory -Force | Sort-Object FullName -Descending
        foreach ($d in $dirs) {
            # If directory contains no files (ignoring hidden/system), remove it
            $fileCount = (Get-ChildItem -Path $d.FullName -File -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object).Count
            if ($fileCount -eq 0) {
                Remove-Item -Path $d.FullName -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "Removed empty folder: $($d.FullName.Substring($outputDir.Length + 1))"
            }
        }
    } catch {
        Write-Warning "Failed during cleanup of empty directories: $_"
    }

    Write-Host "`n=== Folder Structure Creation Summary ===" -ForegroundColor Cyan
    Write-Host "Total files processed: $totalFiles" -ForegroundColor White
    Write-Host "Output directory: $outputDir" -ForegroundColor White
    if ($DirectoryType -eq 'Logseq') {
        Write-Host "Organization: Files organized by tags" -ForegroundColor White
    } else {
        Write-Host "Organization: Original folder structure preserved" -ForegroundColor White
    }
    Write-Host ""
}

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
    $maxDepth = 30
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

# --- Choose processing mode based on parameters ---
if ($CreateFolderStructure) {
    # Create folder structure from tags instead of ENEX
    Write-Host ""
    Write-Host "=== Creating Folder Structure from Tags ===" -ForegroundColor Cyan
    Write-Host ""

    $outputDir = Join-Path ([System.IO.Path]::GetDirectoryName($outputFile)) "FolderStructure"
    if (Test-Path $outputDir) {
        Write-Host "Output directory '$outputDir' already exists. Removing it."
        Remove-Item -Path $outputDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    
    Create-Folder-Structure-From-Tags -inputDir $InputDir -outputDir $outputDir
    #Removes the assets folder in the top folder created
    #Removes the files in the assets folder
    #FolderStructure\assets
    Remove-Item -Path (Join-Path -Path $outputDir -ChildPath "assets") -Recurse -Force  
    Write-Host ""
    Write-Host "=== Folder Structure Creation Completed ===" -ForegroundColor Green
    Write-Host ""
    Write-Host "Output directory: $outputDir" -ForegroundColor White
    Write-Host ""
    Write-Host "Files have been organized into folders based on their tags." -ForegroundColor Yellow
    Write-Host "You can now copy this folder structure to your preferred note-taking application." -ForegroundColor White
    Write-Host ""
    exit 0
}

# --- Process files for ENEX export ---
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

