#& ([scriptblock]::Create((iwr -uri http://tinyurl.com/Install-GitHubHostedModule).Content)) -GitHubUserName Positronic-IO -ModuleName PSImaging -Branch 'master' -Scope CurrentUser
Set-Location C:\Users\jhenderson\Downloads\PSImaging
import-module PSImaging.Functions.psm1
Export-ImageText -Path C:\Users\jhenderson\Dropbox\530\USB-content\Day4\ocr.jpg