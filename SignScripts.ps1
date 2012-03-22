$cert = dir cert:\CurrentUser\my -CodeSigningCert | Select-Object -First 1
$scripts = "C:\Users\ME\Powershell\Scripts"

Get-ChildItem $scripts -Filter *.ps1 -Recurse -ErrorAction SilentlyContinue |
  ForEach-Object {
    (Get-Content -Path $_.FullName) |
      Set-Content -Path $_.FullName -Encoding UTF8
  }

if ($cert) { dir $scripts -Filter *.ps1 -ea 0 |
  Set-AuthenticodeSignature -Certificate $cert
} else {
  Write-Warning 'You do not have a digital certificate for code signing.'
}
