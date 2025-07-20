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
    
    $result = git ls-remote $GitHubURL $Branch 2>$null
    if (-not $result) {
        Write-Error "Failed to fetch latest commit from $RepoURL"
        exit 1
    }
    return $result.Split("`n")[0].Split("`t")[0]
}

function Get-SHA512ForCommit {
    param (
        [string]$Commit
    )
    
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
        Write-Error "portfile.cmake not found at $Path"
        exit 1
    }

    $content = Get-Content $PortfilePath -Raw

    $content = [regex]::Replace($content, '(?<=REF ).*', "$NewCommit")
    $content = [regex]::Replace($content, 'SHA512\s+[0-9a-fA-F]{128}', "SHA512 $NewHash")

    Set-Content -Path $PortfilePath -Value $content -Encoding UTF8    
    Write-Host "`n Updated REF and SHA512 in $((Resolve-Path $PortfilePath).Path)"
}

# ---- Run ----
Write-Host "Fetching latest commit for $RepoURL..."
$commit = Get-LatestCommit

Write-Host "Downloading source archive to compute SHA512..."
$sha512 = Get-SHA512ForCommit -Commit $commit

Write-Host "Updating portfile..."
Update-Portfile -NewCommit $commit -NewHash $sha512