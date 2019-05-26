<#

.SYNOPSIS
Fixes fonts with broken name tables.

.DESCRIPTION

FTAutofix accepts a comma-delimited list of TrueType/OpenType fonts and/or folders and attempts to automatically fix broken name tables.
It requires one name field to hold the full font name and guesses its way from there.


.PARAMETER Inputs
Comma-delimited list of files and/or Directories to process (can take mixed).
Alias: -i

.PARAMETER Recurse
Recurse subdirectories.
Alias: -r

.EXAMPLE
FTAutofix X:\font.ttf

.EXAMPLE
FTAutofix X:\Fonts -r

.LINK
https://github.com/line0/FontToys

#>
#requires -version 3

function FTAutofix {
  [CmdletBinding()]
  param
  (
    [Parameter(Position = 0, Mandatory = $true, HelpMessage = 'Comma-delimited list of files and/or Directories to process (can take mixed).')]
    [alias("i")]
    [string[]]$Inputs,
    [Parameter(Mandatory = $false, HelpMessage = 'Recurse subdirectories.')]
    [alias("r")]
    [switch]$Recurse = $false,
    [Parameter(Mandatory = $false, HelpMessage = 'Name table entries searched for the full font name.')]
    [alias("nt")]
    [object[]]$NameTableEntries = (@(4, 1, 0, 0), @(4, 3, 0, 0)),
    [Parameter(Mandatory = $false, HelpMessage = 'Retrieve the font name from the file name instead of a name table entry.')]
    [alias("f")]
    [switch]$UseFilename = $false,
    [Parameter(Mandatory = $false, HelpMessage = 'Regex pattern to be used for matching family and style name parts.')]
    [alias("p")]
    [regex]$MatchPattern,
    [Parameter(Mandatory = $false, HelpMessage = 'A single family name that overrides all detected family names.')]
    [alias("fo")]
    [string]$FamilyOverride
  )

  # if there are no other properties, $_.#text is resolved into a string, which is inconvenient so we undo it
  $fontWordsXml = [xml](Get-Content (Join-Path (Split-Path -parent $PSCommandPath) "StyleWords.xml"))
  $knownStyleWords = @($fontWordsXml.fontWords.styleWords.word | Where-Object {$_ -is [System.Xml.XmlElement]}) + @(
    $fontWordsXml.fontWords.styleWords.word | Where-Object {$_ -is [string]} | ForEach-Object {@{"#text" = $_}})
  $knownFamilyWords = @($fontWordsXml.fontWords.familyWords.word | Where-Object {$_ -is [System.Xml.XmlElement]}) + @(
    $fontWordsXml.fontWords.familyWords.word | Where-Object {$_ -is [string]} | ForEach-Object {@{"#text" = $_}})

  try {
    $fonts = Get-Files $Inputs -match '.[o|t]tf$' -matchDesc 'OpenType/TrueType Font' -acceptFolders -recurse:$Recurse
  } catch {
    if($_.Exception.WasThrownFromThrowStatement -eq $true) {
      Write-Host $_.Exception.Message -ForegroundColor Red
      break
    } else {
      throw $_.Exception
    }
  }

  $overallActivity = "FontToys Autofix"
  Write-Progress -Activity $overallActivity -Id 0 -PercentComplete 0 -Status "Step 1/2: Reading and fixing fonts"

  $readActivity = "Reading $($fonts.Count) fonts.."
  Write-Progress -Activity $readActivity -Id 1 -PercentComplete 0 -Status "Font 1/$($fonts.Count)"

  $tableView = @(
    @{label = "File"; Expression = { $_.Path.Name}; width = 50},
    @{label = "Family Name"; Expression = { $_._FamilyName}; width = 50},
    @{label = "Style Name"; Expression = { $_._StyleName}; width = 32}
    @{label = "Selection Flags"; Expression = { $_.GetFsFlags()}; width = 32}
    @{label = "Weight"; Expression = { $_.GetWeight()}; width = 12}
    @{label = "Width"; Expression = { $_.GetWidth()}; width = 16}
  )

  $fonts | FixFont -NameTableEntries $NameTableEntries -useFilename $UseFilename -matchPattern $MatchPattern -familyOverride $FamilyOverride -OutVariable fontList |
  Format-Table -Property $tableView | ForEach-Object {
    $_
    $fntReadCnt = $fontList.Count / $fonts.Count
    Write-Progress -Activity $readActivity -Id 1 -PercentComplete (100 * $fntReadCnt) -Status "File $($fontList.Count+1)/$($fonts.Count): $($fontList[-1].Path.Name)"
    Write-Progress -Activity $overallActivity -Id 0 -PercentComplete (50 * $fntReadCnt) -Status "Step 1/2: Reading and fixing fonts"
  }

  $familyList = $fontList | Group-Object -Property _FamilyName

  $familyActivity = "Writing $($familyList.Count) families.."
  Write-Progress -Activity $overallActivity -Id 0 -PercentComplete 50 -Status "Step 2/2: Writing fixed fonts"
  $doneCnt = 0

  foreach($group in $familyList) {
    Write-Progress -Activity $familyActivity -Id 1 -PercentComplete (100 * $doneCnt / $familyList.Count) -Status $group.Name

    $familyPath = New-Item (Join-Path $group.Group[0].Path.DirectoryName $group.Name) -ItemType Directory -Force

    $styleActivity = "Writing $($group.Group.Count) styles.."
    $group.Group | ForEach-Object {$styleDoneCnt = 0} {
      Write-Progress -Activity $styleActivity -Id 2 -PercentComplete (100 * $styleDoneCnt / $group.Group.Count) -Status $_.GetNames(2, 1, 0, 0).Name
      $percentOverall = 50 + 50 * (($doneCnt / $familyList.Count) + 1 * $styleDoneCnt / ($familyList.Count * $group.Group.Count))
      Write-Progress -Activity $overallActivity -Id 0 -PercentComplete $percentOverall -Status "Step 2/2: Writing fixed fonts"
      Export-Font $_ -OutPath $familyPath.FullName -Format ttf
      $styleDoneCnt++
    }
    $doneCnt++
  }
}

filter FixFont {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory = $True, ValueFromPipeline = $true)]
    [System.IO.FileSystemInfo]$_,
    [object[]]$NameTableEntries = (@(4, 1, 0, 0), @(4, 3, 0, 0)),
    [bool]$useFilename = $false,
    [regex]$matchPattern,
    [string]$familyOverride
  )

  $fontData = Import-Font -InFile $_.FullName

  $fontName = if($useFilename) {
    $fontData.Path.BaseName
  } else {
    ($NameTableEntries | ForEach-Object {$fontData.GetNames($_[0], $_[1], $_[2], $_[3]).Name} | Where-Object {$_ -ne $null} |
    Sort-Object -Property Length -Descending | Select-Object -First 1 # longest font name is probably of the highest quality
    ) -replace ':.*','' # sometimes the full font name is suffixed by version like "Aarcover Plain:001.001"
  }

  $styleWords = @()
  if($matchPattern) {
    $matches = Select-String -InputObject $fontName -pattern $matchPattern | Select-Object -ExpandProperty Matches
    if ($matches.Groups.Count -lt 3) {
      throw "Your match pattern '$($matchPattern)' didn't produce at least 2 groups."
    } else {
      $familyName = $matches.Groups[1].Value
      $styleName = ReplaceStyleWords $matches.Groups[1].Value -StyleWords $knownStyleWords
      $styleWords = @(($styleName -replace "_", " " -split "\s+") | Where-Object {$_})
    }
  } else {
    $fontName = ReplaceStyleWords $fontName -StyleWords ($knownStyleWords + $knownFamilyWords) -ProtectBeginning $true

    # split font name into words at boundaries denoted by underscores, dashes and spaces
    $fontWords = @(($fontName.Trim() -replace "-", " " -replace "_", " " -split "\s+") | Where-Object {$_ -match '[^\s]'})
    Write-Debug "Font Words: $($fontWords -join ', ')"

    $familyWords = @()
    # filter out any style words that just perform replacements
    $matchingknownStyleWords = $knownStyleWords | Where-Object {$_."#text" -or ($_.match -and -not $_.replace)}

    # determine the boundary between family name and style name using incredibly crude heuristics
    $firstStyleWordIndex = -1
    for ($($i = 1; $previousWordIsStyleWord = $false); $i -lt $fontWords.Count; $i++) {
      # filter out any style word not meant to go into this position in the style word order
      $applicableknownStyleWords = $matchingknownStyleWords | Where-Object {
        (!$_.onlyLast -or ($fontWords.Count - $i) -le ($_.onlyLast)) -and (!$_.notFirst -or $i -gt $_.notFirst)
      }

      $currentWordIsStyleWord = IsStyleWord $fontWords[$i] -StyleWords $applicableknownStyleWords
      if (!$previousWordIsStyleWord -and $currentWordIsStyleWord) {
        $firstStyleWordIndex = $i
      } elseif (!$currentWordIsStyleWord) {
        if ($previousWordIsStyleWord -and ($fontWords[$i] -in $knownFamilyWords."#text")) {
          $familyWords += $fontWords[$i]
          $fontWords[$i] = [string]::Empty
        } else {
          $firstStyleWordIndex = -1
        }
      }
      $previousWordIsStyleWord = $currentWordIsStyleWord
    }
    $fontWords[0] = UpperFirst $fontWords[0] # start font with an uppercase character

    # generate the family name, decamelize it and remove it from the style words list
    $familyWords = if ($firstStyleWordIndex -eq -1) {
      $fontWords + $familyWords
    } else {
      $fontWords[0..($firstStyleWordIndex - 1)] + $familyWords
    }
    if ($firstStyleWordIndex -gt -1) {
      # Collect and TitleCase all style words
      $styleWords = @($fontWords[$firstStyleWordIndex..($fontWords.Count - 1)] | Where-Object {$_} | UpperFirst -FixCase)
    }
    $familyName = @($familyWords | ForEach-Object {UpperFirst $_}) -join " " -creplace '(\p{Ll}+)([\p{Lu}\p{Lt}])', '$1 $2'

    Write-Debug "Family Name: $familyName"
    $fontData | Add-Member -Type NoteProperty -Name _FontWords -Value $fontWords # only for debugging purposes
  }

  if ($styleWords.Length -gt 0) {
    $knownStyleWords | Where-Object {$_.'#text'} | ForEach-Object {
      # conform case of all known style words
      $isMatch = if ($_.match -and $styleWords -cmatch "^$($_.match)`$") {
        $styleWords = $styleWords -creplace "^$($_.match)`$", $_.'#text' # a bit wasteful doing the matching twice, but style words of this configuration are rare at the time of writing
        $true
      } elseif (!$_.match -and 0 -le ($i = [Array]::FindIndex($styleWords, [Predicate[string]] {param($word) $word -eq $_.'#text'}))) {
        $styleWords[$i] = $_.'#text'
        $true
      }
      if ($isMatch) {
        ImportMetrics -Font $fontData -StyleWord $_
      }
    }
  }

  $fontData | Add-Member -Type NoteProperty -Name _StyleName -Value ($styleWords -join " ") -PassThru |
  Add-Member -Type NoteProperty -Name _FamilyName -Value $(if ($familyOverride) {
      $familyOverride
    } else {
      $familyName
    })

  $fontData.SetFamily($fontData._FamilyName, $fontData._StyleName)
  return $fontData
}

function UpperFirst(
  [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
  [string]$str,
  [switch]$FixCase = $false
) {
  Process {
    if($FixCase -and ($str -ceq $str.ToLower() -or $str -ceq $str.ToUpper())) {
      return (Get-Culture).TextInfo.ToTitleCase($str)
    } else {
      return $str.Substring(0, 1).ToUpper() + $str.Substring(1)
    }
  }
}

function ReplaceStyleWords(
  [Parameter(Position = 0, Mandatory = $true)]
  [string]$FontName,
  [Parameter(Mandatory = $true)]
  [System.Object[]]$StyleWords,
  [Parameter(Mandatory = $false)]
  [bool]$ProtectBeginning = $false
) {
  $protectBeginningLookahead = if($ProtectBeginning) {
    "(?<=.)"
  }

  $StyleWords | Where-Object {$_.match -and $_.replace} | ForEach-Object {
    $prevFontName = $FontName
    # Protect the very beginning of the font name if configured and don't match partial words
    # Add separation markers for every match to denote word boundaries
    $FontName = $FontName -creplace "$($protectBeginningLookahead)($($_.match))(?!\p{Ll})", "_$($_.replace)_"

    if ($prevFontName -ne $FontName) {
      Write-Debug "Replaced: $($_.match) -> $($_.replace)"
    }
  }
  $StyleWords | Where-Object {$_.separate -eq 1 -and -not ($_.match -and $_.replace)} | ForEach-Object {
    $styleWord = $_
    # match using "match" tag attribute if specified, otherwise match the element content
    $match = if($styleWord.match) {
      $styleWord.match
    } else {
      [Regex]::Escape($styleWord."#text")
    }
    if (!$match) {
      return;
    }

    # Add separation markers for every match to denote word boundaries
    $regexOptions = if (!$styleWord.caseSensitive -or $styleWord.caseSensitive -eq 0) {
      [Text.RegularExpressions.RegexOptions]::IgnoreCase
    } else {
      [Text.RegularExpressions.RegexOptions]::None
    }
    [regex]::Matches($FontName, $match, $regexOptions) | ForEach-Object {
      if (($ProtectBeginning -and $_.Index -eq 0) -or $FontName[$_.Index + $_.Length] -cmatch "\p{Ll}") {
        # Protect the very beginning of the font name if configured and don't match partial words
        return;
      }
      $term = $FontName.Substring($_.Index, $_.Length)
      $FontName = $FontName.Substring(0, $_.Index) + '_' + $term + '_' + $FontName.Substring($_.Index + $_.Length)
      Write-Debug "Separated: $FontName"
    }
  }
  return $FontName
}

function IsStyleWord(
  [Parameter(Position = 0, Mandatory = $true)]
  [string]$Word,
  [Parameter(Mandatory = $true)]
  [System.Object[]]$StyleWords
) {
  foreach ($styleWord in $StyleWords) {
    if ($styleWord.match) {
      if ($Word -cmatch "^$($styleWord.match)`$") {
        return $true
      }
    } elseif ($styleWord.'#text') {
      $match = [Regex]::Escape($styleWord.'#text')
      if ($Word -match "^$match`$") {
        return $true
      }
    }
  }
  return $false
}


function ImportMetrics(
  [Parameter(Mandatory = $true)]
  [object]$Font,
  [Parameter(Mandatory = $true)]
  [object]$StyleWord
) {
  if($StyleWord.weight) {
    $Font.SetWeight([int]$StyleWord.weight)
  }
  if($StyleWord.width) {
    $Font.SetWidth([int]$StyleWord.width)
  }
  if($StyleWord.fsSelection) {
    $Font.AddFsFlags([int]$StyleWord.fsSelection)
  }
}

Export-ModuleMember FTAutofix