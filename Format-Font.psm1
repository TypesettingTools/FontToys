filter Format-Font([string[]]$Tables = @("name", "OS/2")) {
  if(!$_ -or $_ -isnot [PSCustomObject]) {
    throw "Input missing or not an Import-Font object"
  }

  $nameDesc = ([xml](Get-Content (Join-Path (Split-Path -parent $PSCommandPath) "NameDesc.xml"))).descriptions

  $nameTableView = @(
    @{label = "Name"; Expression = { GetDesc $_ name $nameDesc}},
    @{label = "Platform"; Expression = { GetDesc $_ platform $nameDesc}},
    @{label = "Encoding"; Expression = {GetDesc $_ enc $nameDesc}},
    @{label = "Language"; Expression = { GetDesc $_ lang $nameDesc}},
    @{label = "Text"; Expression = { $_.Name}}
  )
  return $_.GetNames() | Format-Table -Property $nameTableView -AutoSize
}

function GetDesc($record, $prop, $nameDesc) {
  $propID = $record."$($prop)ID"
  $desc = $nameDesc."$($prop)s".$prop | Where-Object {
    [int]$_.id -eq $propID -and ($prop -notmatch "enc|lang" -or $_.PlatformID -eq $record.platformID)
  }

  if ($prop -eq "lang" -and $record.PlatformID -eq 3) {
    $desc = @{desc = [System.Globalization.Cultureinfo]::GetCultureInfo([int]$propID)}
  }
  return "{0}{1}: {2}" -f (" " * ($nameDesc."$($prop)s".maxDigits - ([string]$propID).length)), $propID, $desc.desc
}

Export-ModuleMember Format-Font