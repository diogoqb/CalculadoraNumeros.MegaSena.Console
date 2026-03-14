# Configura a licença EPPlus para uso não comercial
$env:EPPlusLicenseContext = "NonCommercial"

# Executa o programa
Start-Process -FilePath ".\CalculadoraNumeros.MegaSena.Console.exe" -NoNewWindow -Wait