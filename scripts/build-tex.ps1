param(
  [Parameter(Mandatory = $true)]
  [string]$TexFile,

  [string]$AuxDir = "build\\aux",

  [string]$PdfDir = "build\\pdf"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

function Resolve-RepoPath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$PathValue,

    [Parameter(Mandatory = $true)]
    [string]$BaseDir
  )

  if ([System.IO.Path]::IsPathRooted($PathValue)) {
    return [System.IO.Path]::GetFullPath($PathValue)
  }

  return [System.IO.Path]::GetFullPath((Join-Path $BaseDir $PathValue))
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$texPath = Resolve-RepoPath -PathValue $TexFile -BaseDir $repoRoot

if (-not (Test-Path -LiteralPath $texPath -PathType Leaf)) {
  throw "TeX file not found: $texPath"
}

$auxPath = Resolve-RepoPath -PathValue $AuxDir -BaseDir $repoRoot
$pdfPath = Resolve-RepoPath -PathValue $PdfDir -BaseDir $repoRoot

New-Item -ItemType Directory -Force -Path $auxPath, $pdfPath | Out-Null

$texDir = Split-Path -Parent $texPath
$texName = Split-Path -Leaf $texPath
$pdfName = [System.IO.Path]::ChangeExtension($texName, ".pdf")
$auxPdfPath = Join-Path $auxPath $pdfName
$finalPdfPath = Join-Path $pdfPath $pdfName

Push-Location $texDir
try {
  & latexmk `
    -pdf `
    -cd `
    -interaction=nonstopmode `
    -halt-on-error `
    -emulate-aux-dir `
    "-auxdir=$auxPath" `
    "-outdir=$auxPath" `
    "-out2dir=$pdfPath" `
    $texName

  if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
  }

  if ((Test-Path -LiteralPath $finalPdfPath) -and (Test-Path -LiteralPath $auxPdfPath)) {
    Remove-Item -LiteralPath $auxPdfPath -Force
  }
}
finally {
  Pop-Location
}
