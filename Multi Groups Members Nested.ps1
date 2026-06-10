CLS
$ErrorActionPreference = 'Stop'
Import-Module ActiveDirectory -ErrorAction Stop

function Show-And-Wait {
  param([string]$Text,[int]$Seconds=1)
  Write-Host $Text -ForegroundColor Yellow
  Start-Sleep -Seconds $Seconds
}

try {
  $excelTest = New-Object -ComObject Excel.Application
  $excelTest.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excelTest)
} catch {
  Write-Error "Excel is not installed or Excel COM is unavailable."
  return
}

function Coalesce { param([Parameter(ValueFromRemainingArguments=$true)]$Values) foreach ($v in $Values) { if ($null -ne $v -and $v -ne '') { return $v } }; return $null }
function Get-DomainFromDN { param([string]$DN) if(-not $DN){return $null}; $parts=$DN -split ','; $dcs=$parts|?{$_ -like 'DC=*'}|%{$_ -replace '^DC='}; if($dcs.Count){$dcs -join '.'} }
function Join-IdNumber { param($Obj) $p=@(); if($Obj.extensionAttribute15){$p+=[string]$Obj.extensionAttribute15}; if($Obj.extensionAttribute14){$p+=[string]$Obj.extensionAttribute14}; if($p.Count){$p -join ''} }

function Get-GroupMembersAllDNs {
  param([string]$GroupDN,[string]$Server)
  $path = if ($Server) { "LDAP://$Server/$GroupDN" } else { "LDAP://$GroupDN" }
  $de = New-Object System.DirectoryServices.DirectoryEntry($path)
  $ds = New-Object System.DirectoryServices.DirectorySearcher($de)
  $ds.SearchScope='Base'; $ds.PageSize=1000
  $all = New-Object System.Collections.Generic.List[string]; $low=0; $step=1500
  while ($true) {
    $high=$low+$step-1; $attr="member;range=$low-$high"
    $ds.PropertiesToLoad.Clear(); [void]$ds.PropertiesToLoad.Add($attr)
    $res=$ds.FindOne(); if(-not $res){break}
    $key=$null; foreach($k in $res.Properties.PropertyNames){ if($k -like 'member;range=*'){ $key=$k; break } }
    if(-not $key){break}
    $vals=$res.Properties[$key]; if($vals){ foreach($v in $vals){ [void]$all.Add([string]$v) } }
    if ($key -match 'member;range=\d+-\*$') { break }
    if(-not $vals -or $vals.Count -lt $step){ break }
    $low=$high+1
  }
  return $all.ToArray()
}

$script:__ObjCache=@{}
function Get-ADObjectCached {
  param([string]$DN,[string]$Server)
  if($script:__ObjCache.ContainsKey($DN)){ return $script:__ObjCache[$DN] }
  $props = 'objectClass','sAMAccountName','displayName','mail','name','extensionAttribute10','extensionAttribute14','extensionAttribute15','userPrincipalName','enabled','department','title'
  try{$obj=Get-ADObject -Identity $DN -Server $Server -Properties $props -ErrorAction Stop; $script:__ObjCache[$DN]=$obj; return $obj}
  catch{$script:__ObjCache[$DN]=$null; return $null}
}

function New-SafeName { param([string]$Name,[hashtable]$Used,[int]$MaxLen=31)
  $n=($Name -replace '[:\\/\?\*\[\]]','_').Trim(); if($n.Length -gt $MaxLen){$n=$n.Substring(0,$MaxLen)}
  $base=$n; $i=1
  while($Used.ContainsKey($n) -or [string]::IsNullOrWhiteSpace($n)){
    $suf="~$i"; $trim=[Math]::Max(1,$MaxLen-$suf.Length); $n=($base.Substring(0,[Math]::Min($base.Length,$trim)))+$suf; $i++ }
  $Used[$n]=$true; return $n
}

function Get-RootFullControlNames {
  param([string]$GroupDN)
  try { $acl = Get-Acl -Path ("AD:\" + $GroupDN) } catch { return @() }
  $genericAll = [System.DirectoryServices.ActiveDirectoryRights]::GenericAll
  $targets = foreach ($ace in $acl.Access) {
    if ($ace.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow) { continue }
    $rights = $ace.ActiveDirectoryRights
    if ((($rights -band $genericAll) -ne 0) -and (([int]$rights -bxor [int]$genericAll) -eq 0)) { $ace.IdentityReference.Value }
  }
  if (-not $targets) { return @() }
  $out = New-Object System.Collections.Generic.List[string]
  foreach ($val in ($targets | Sort-Object -Unique)) {
    $sam = if ($val -like '*\*') { ($val -split '\\')[-1] } else { $val }
    [void]$out.Add($sam)
  }
  return $out.ToArray()
}

function Build-GroupSheets {
  param([string]$RootGroupDN,[string]$RootDomain)
  $sheets=@{}; $used=@{}; $dnToSheet=@{}
  $seen = New-Object 'System.Collections.Generic.HashSet[string]'
  $q    = New-Object 'System.Collections.Generic.Queue[object]'

  $root = Get-ADGroup -Identity $RootGroupDN -Server $RootDomain -Properties sAMAccountName,displayName
  $rootName = Coalesce $root.sAMAccountName $root.Name
  $rootDisp = Coalesce $root.displayName $rootName
  $rootSheet = New-SafeName -Name $rootDisp -Used $used
  $dnToSheet[$root.DistinguishedName] = $rootSheet
  [void]$seen.Add($root.DistinguishedName)
  $q.Enqueue([PSCustomObject]@{ DN=$root.DistinguishedName; Domain=$RootDomain; Sheet=$rootSheet; IsRoot=$true })

  while ($q.Count -gt 0) {
    $node = $q.Dequeue()
    $dn=$node.DN; $dom=$node.Domain; $name=$node.Sheet; $isRoot=[bool]$node.IsRoot
    $memberDns = Get-GroupMembersAllDNs -GroupDN $dn -Server $dom
    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($childDN in $memberDns) {
      $childDom = Get-DomainFromDN $childDN; if (-not $childDom) { continue }
      $o = Get-ADObjectCached -DN $childDN -Server $childDom; if (-not $o) { continue }
      $type = $o.objectClass; if ($type -is [System.Array]) { $type = $type[-1] }
      $disp = Coalesce $o.displayName $o.sAMAccountName $o.Name
      $user = $o.sAMAccountName
      $idn  = Join-IdNumber $o

      if ($type -eq 'group') {
        $rows.Add([PSCustomObject]@{
          DisplayName=$disp; Email=''; Username=$user; 'ID Number'=$idn; Department=''; Title='';
          ExtensionAttribute10=''; ObjectType='group'; Domain=$childDom; UserPrincipalName=''
        })
        if (-not $dnToSheet.ContainsKey($childDN)) { $dnToSheet[$childDN] = New-SafeName -Name $disp -Used $used }
        if (-not $seen.Contains($childDN)) {
          [void]$seen.Add($childDN)
          $q.Enqueue([PSCustomObject]@{ DN=$childDN; Domain=$childDom; Sheet=$dnToSheet[$childDN]; IsRoot=$false })
        }
      } else {
        $rows.Add([PSCustomObject]@{
          DisplayName=$disp; Email=$o.mail; Username=$user; 'ID Number'=$idn; Department=$o.department; Title=$o.title;
          ExtensionAttribute10=$o.extensionAttribute10; ObjectType=$type; Domain=$childDom; UserPrincipalName=$o.userPrincipalName
        })
      }
    }

    $fc=@(); if ($isRoot) { $fc = Get-RootFullControlNames -GroupDN $dn }
    if ($isRoot) {
      $max=[Math]::Max($rows.Count,$fc.Count); $final=New-Object System.Collections.Generic.List[object]
      for($i=0;$i -lt $max;$i++){
        if($i -lt $rows.Count){
          $r=$rows[$i] | Select-Object DisplayName,Email,Username,'ID Number',Department,Title,ExtensionAttribute10,ObjectType,Domain,UserPrincipalName
          $r | Add-Member -NotePropertyName FullControl -NotePropertyValue ($(if($i -lt $fc.Count){$fc[$i]} else {''})) -Force
          $final.Add($r)
        } else {
          $final.Add([PSCustomObject]@{DisplayName='';Email='';Username='';'ID Number'='';Department='';Title='';
            ExtensionAttribute10='';ObjectType='';Domain='';UserPrincipalName='';FullControl=$fc[$i]})
        }
      }
      $sheets[$name]=$final.ToArray()
    } else {
      $sheets[$name]=$rows.ToArray()
    }
  }
  return $sheets
}

function Write-Workbook {
  param([hashtable]$Sheets,[string]$Path)
  $headersBase=@('DisplayName','Email','Username','ID Number','Department','Title','ExtensionAttribute10','ObjectType','Domain','UserPrincipalName')
  $excel=$null; $wb=$null
  try{
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.ScreenUpdating=$false; $excel.DisplayAlerts=$false
    $wb=$excel.Workbooks.Add()
    $want=@{}; foreach($k in $Sheets.Keys){ $want[$k.Substring(0,[Math]::Min(31,$k.Length))]=$true }

    foreach($kv in $Sheets.GetEnumerator()){
      $sheet=$wb.Worksheets.Add(); $sheet.Name=$kv.Key.Substring(0,[Math]::Min(31,$kv.Key.Length))
      $rows=$kv.Value; if(-not $rows -or $rows.Count -eq 0){ continue }
      $hasFC = ($rows[0].PSObject.Properties.Name -contains 'FullControl')
      $headers = if($hasFC){ $headersBase + 'FullControl' } else { $headersBase }

      $hdr=New-Object 'object[,]' 1,$headers.Count
      for($i=0;$i -lt $headers.Count;$i++){ $hdr[0,$i]=$headers[$i] }
      $sheet.Range($sheet.Cells.Item(1,1),$sheet.Cells.Item(1,$headers.Count)).Value2=$hdr

      $rCount=$rows.Count; $cCount=$headers.Count
      $data=New-Object 'object[,]' $rCount,$cCount
      for($r=0;$r -lt $rCount;$r++){
        $item=$rows[$r]
        for($c=0;$c -lt $cCount;$c++){
          $v=$item.($headers[$c])
          if($null -eq $v){ $data[$r,$c]='' }
          elseif($v -is [DateTime]){ $data[$r,$c]=$v.ToString('yyyy-MM-dd HH:mm:ss') }
          else{ $data[$r,$c]=[string]$v }
        }
      }
      $dest=$sheet.Range($sheet.Cells.Item(2,1),$sheet.Cells.Item($rCount+1,$cCount))
      $dest.Value2=$data
      $sheet.Range($sheet.Cells.Item(1,1),$sheet.Cells.Item(1,$cCount)).EntireColumn.AutoFit() | Out-Null
    }
    $all=@(); foreach($s in $wb.Worksheets){ $all+=$s }
    foreach($s in $all){ if(-not $want.ContainsKey($s.Name)){ try{$s.Delete()}catch{} } }
    $wb.SaveAs($Path)
  } finally {
    if ($wb)    { $wb.Close($false) | Out-Null }
    if ($excel) { $excel.Quit() }
    if ($wb)    { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
    if ($excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

function Resolve-GroupDN { param([string]$Identity,[string]$Domain)
  try { return (Get-ADGroup -Server $Domain -Identity $Identity -Properties DistinguishedName).DistinguishedName } catch {}
  try { $e=$Identity.Replace("'", "''"); $g=Get-ADGroup -Server $Domain -Filter "Name -eq '$e' -or SamAccountName -eq '$e' -or mail -eq '$e'" -Properties DistinguishedName; if($g){return $g.DistinguishedName} } catch {}
  return $null
}

#Start Script
Show-And-Wait "Enter Group Names one per line | Save close Notepad to continue"
$TXTfile = [System.IO.Path]::GetTempFileName()
Start-Process notepad.exe -ArgumentList $TXTfile -Wait
$groupNames = Get-Content -Path $TXTfile -ErrorAction Stop | % { $_.Trim() } | ? { $_ -and -not $_.StartsWith('#') }
if (-not $groupNames) { Write-Error "No valid group names."; return }

$forestDomains=@(); try { $forestDomains=(Get-ADForest).Domains } catch { Write-Warning "Forest domains unavailable. Type domains manually." }

Write-Host -ForegroundColor Yellow "Enter a name for New Folder"
$folderName = Read-Host "Folder name"
if (-not $folderName) { $folderName = "AD-Groups-Export" }
$outFolder = Join-Path $env:TEMP $folderName
if (-not (Test-Path $outFolder)) { New-Item -ItemType Directory -Path $outFolder | Out-Null }

$mapping = New-Object System.Collections.Generic.List[object]
foreach ($g in $groupNames) {
  if ($forestDomains.Count -gt 0) {
    Write-Host ""
    Write-Host "Available forest domains:" -ForegroundColor Cyan
    $i=1; foreach ($d in $forestDomains) { Write-Host ("{0}. {1}" -f $i,$d); $i++ }
  }
  Write-Host -ForegroundColor Yellow "Enter a domain for each group:"
  do {
    $dom = Read-Host ("Enter domain for group '{0}'"-f $g) 
    if ([string]::IsNullOrWhiteSpace($dom)) { Write-Host "Domain is required." -ForegroundColor Red; continue }
    if ($forestDomains.Count -gt 0 -and ($forestDomains -notcontains $dom)) {
      Write-Host "Domain not in forest list. Type again exactly as shown." -ForegroundColor Red
      $dom = $null
    }
  } until ($dom)
  $mapping.Add([pscustomobject]@{ Group=$g; Domain=$dom.Trim() })
}

foreach ($row in $mapping) {
  $script:__ObjCache.Clear()
  Write-Host ("Processing: {0}  [Domain: {1}]" -f $row.Group,$row.Domain) -ForegroundColor Green
  $dn = Resolve-GroupDN -Identity $row.Group -Domain $row.Domain
  if (-not $dn) { Write-Warning ("Group '{0}' not found in {1}. Skipping." -f $row.Group,$row.Domain); continue }
  $sheets = Build-GroupSheets -RootGroupDN $dn -RootDomain $row.Domain

  $fileNameRaw = "{0}@{1}.xlsx" -f $row.Group, $row.Domain
  $fileName = ($fileNameRaw -replace '[\\/:*?"<>|]','_')
  $wbPath = Join-Path $outFolder $fileName
  Write-Workbook -Sheets $sheets -Path $wbPath
  Write-Host ("Saved: {0}" -f $wbPath)
}

Start-Process "explorer.exe" -ArgumentList $outFolder
Write-Host ("Done. Output folder: {0}" -f $outFolder) -ForegroundColor Green