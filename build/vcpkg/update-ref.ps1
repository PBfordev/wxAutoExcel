# ========================
# Script sets REF and SHA512 in wxautoexcel\portfile.cmake
# to the latest commit in the wxAutoExcel master branch.
# ========================

# ---- Constants ----
$PortfilePath = "wxautoexcel\portfile.cmake"
$GitHubUser = "PBfordev"
$GitHubRepo = "wxAutoExcel"
$GitHubURL = "https://github.com/$GitHubUser/$GitHubRepo"
$Branch = "master"

# ---- Functions ----
function Get-LatestCommit {

    Write-Host "Fetching latest commit for $GitHubURL..."
    $result = git ls-remote $GitHubURL $Branch 2>$null
    if (-not $result) {
        Write-Error "Failed to fetch latest commit from $RepoURL"
        exit 1
    }
    return $result.Split("")[0].Split("`t")[0]
}

function Get-SHA512ForCommit {
    param (
        [string]$Commit
    )
    
    Write-Host "Downloading source archive to compute SHA512..."
    $tempFile = "$env:TEMP\$GitHubRepo-$Commit.tar.gz"
    Invoke-WebRequest -Uri "$GitHubURL/archive/$Commit.tar.gz" -OutFile $tempFile -UseBasicParsing
    $sha512 = Get-FileHash -Algorithm SHA512 -Path $tempFile | Select-Object -ExpandProperty Hash

    Remove-Item $tempFile -Force
    return $sha512.ToLower()
}

function Update-Portfile {
    param (
        [string]$NewCommit,
        [string]$NewHash
    )

    if (-not (Test-Path $PortfilePath)) {
        Write-Error "portfile.cmake not found at $Path."
        exit 1
    }
    Write-Host "Updating portfile..."
    
    $content = Get-Content $PortfilePath -Raw
    $content = [regex]::Replace($content, '(^\s*REF\s+)(.+)$', '${1}' + $NewCommit, [System.Text.RegularExpressions.RegexOptions]::Multiline)
    $content = [regex]::Replace($content, 'SHA512\s+[0-9a-fA-F]{128}', "SHA512 $NewHash")
    [System.IO.File]::WriteAllText($PortfilePath, $content)
    
    Write-Host "Updated REF to '$NewCommit' and SHA512 to '$NewHash' in file '$((Resolve-Path $PortfilePath).Path)'."
}

# ---- Run ----
$commit = Get-LatestCommit
$sha512 = Get-SHA512ForCommit -Commit $commit
Update-Portfile -NewCommit $commit -NewHash $sha512