$exclude = @("venv", "deskbotTeste.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "deskbotTeste.zip" -Force