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

function FTAutofix
{
[CmdletBinding()]
param
(
[Parameter(Position=0, Mandatory=$true, HelpMessage='Comma-delimited list of files and/or Directories to process (can take mixed).')]
[alias("i")]
[string[]]$Inputs,
[Parameter(Mandatory=$false, HelpMessage='Recurse subdirectories.')]
[alias("r")]
[bool]$Recurse=$false,
[Parameter(Mandatory=$false, HelpMessage='Name table entry that holds the full font name.')]
[alias("nt")]
[int[]]$NameTableEntry=@(4,1,0,0)
)
    $knownStyleWords = ([xml](Get-Content (Join-Path (Split-Path -parent $PSCommandPath) "StyleWords.xml"))).styleWords.word

    try { $fonts = Get-Files $Inputs -match '.[o|t]tf$' -matchDesc 'OpenType/TrueType Font' -acceptFolders -recurse:$Recurse }
    catch
    {
        if($_.Exception.WasThrownFromThrowStatement -eq $true)
        {
            Write-Host $_.Exception.Message -ForegroundColor Red
            break
        }
        else {throw $_.Exception}
    }

    $overallActivity = "FontToys Autofix"
    Write-Progress -Activity $overallActivity -Id 0 -PercentComplete 0 -Status "Step 1/2: Reading and fixing fonts"

    $readActivity = "Reading $($fonts.Count) fonts.."
    Write-Progress -Activity $readActivity -Id 1 -PercentComplete 0 -Status "Font 1/$($fonts.Count)"

    $fontList = @()
    $doneCnt = 0

    $tableView=@(
        @{label="File"; Expression={ $_.Path.Name}},
        @{label="Family Name"; Expression={ $_._FamilyName}},
        @{label="Style Name"; Expression={ $_._StyleName}}

    )

    $fonts | FixFont -nt $NameTableEntry -OutVariable fontList | Format-Table -Property $tableView | %{
        $_
        $fntReadCnt = $fontList.Count/$fonts.Count
        Write-Progress -Activity $readActivity -Id 1 -PercentComplete (100*$fntReadCnt) -Status "File $($fontList.Count+1)/$($fonts.Count): $($fontList[-1].Path.Name)"
        Write-Progress -Activity $overallActivity -Id 0 -PercentComplete (50*$fntReadCnt) -Status "Step 1/2: Reading and fixing fonts"
    }

    $familyList = $fontList | Group-Object -Property _FamilyName


    $familyActivity = "Writing $($familyList.Count) families.."
    Write-Progress -Activity $overallActivity -Id 0 -PercentComplete 50 -Status "Step 2/2: Writing fixed fonts"
    $doneCnt = 0

    foreach($group in $familyList)
    {
        Write-Progress -Activity $familyActivity -Id 1 -PercentComplete (100*$doneCnt/$familyList.Count) -Status $group.Name

        $familyPath = New-Item (Join-Path $group.Group[0].Path.DirectoryName $group.Name) -ItemType Directory -Force
    
        $styleActivity = "Writing $($group.Group.Count) styles.."
        $group.Group | %{$styleDoneCnt=0}{
            Write-Progress -Activity $styleActivity -Id 2 -PercentComplete (100*$styleDoneCnt/$group.Group.Count) -Status $_.GetNames(2,1,0,0).Name
            $percentOverall = 50+50*(($doneCnt/$familyList.Count)+1*$styleDoneCnt/($familyList.Count*$group.Group.Count))
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
    [Parameter(Mandatory=$True, ValueFromPipeline=$true)][System.IO.FileSystemInfo]$_,
    [Parameter(Mandatory=$True)][int[]]$nt
    )
    $nt = $nt[0..3] + @(0)*(4-$nt[0..3].Count) # pad NameTableEntry array to exactly 4 Ints
    
    $fontData = Import-Font -InFile $_
    $knownStyleWords | ?{$_.search -eq 1 -or $_.separate} | %{$fontName = $fontData.GetNames($nt[0],$nt[1],$nt[2],$nt[3]).Name}{
        $match = if($_.match) {$_.match} else {$_."#text"}
        $rep = if($_.replaceString) {$_.replaceString} else {$_."#text"}
        $fontName = $fontName -creplace "^(.+)($match)","`$1_$($rep)_"
        # $fontName = $fontName -creplace "^(.+)(?<!\s)($match)(?!\s)","`$1_$($rep)_"
    }
    $fontWords = $fontName.Trim() -replace "-"," " -replace "_"," " -split "\s+"

    for ($($i=1;$wordMatch = $false); $i -le $fontWords.Count -and -not $wordMatch; $i++)
    {
        $wordMatch = $fontWords[$i] -in $knownStyleWords."#text"+($knownStyleWords | ?{$_ -is [type]"String"})
        $splitIndex = $i
    }

    $fontData | Add-Member -Type NoteProperty -Name _FontWords -Value $fontWords
    $fontData | Add-Member -Type NoteProperty -Name _FamilyName -Value ($fontWords[0..($splitIndex-1)] -join " ")


    $styleWords = @()
    if($wordMatch)
    {
        foreach ($styleWord in $fontWords[$splitIndex..($fontWords.Count-1)]) {
            $knownStyleWords | ?{$_.replaceString} | %{
                $styleWord = $styleWord -replace "^$($_."#text")`$",$_.replaceString  # should probably turn this into a simple assignment for speed
            }
            $knownStyleWords | ?{$_.'#text' -eq $styleWord} | %{
                if($_.weight) { $fontData.SetWeight([int]$_.weight) }
                if($_.width) { $fontData.SetWidth([int]$_.width) }
                if($_.fsSelection) { $fontData.AddFsFlags([int]$_.fsSelection) }
            }
            $styleWords += $styleWord
        }
    }

    $fontData | Add-Member -Type NoteProperty -Name _StyleName -Value ($styleWords -join " ")
    $fontData.SetFamily($fontData._FamilyName,$fontData._StyleName)
    return $fontData
}

Export-ModuleMember FTAutofix