#Requires -Version 5.0
"Running PowerShell $($PSVersionTable.PSVersion)."
<#
Programa: configuraProxyTerminal
Objetivo: Configura o proxy do terminal
Autor: Samuel Diniz Casimiro - P_7029
Setor: Ditec/Coaus/Satus
Versões: 1.0 - Criação do script
#>

<#
[system.net.webrequest]::defaultwebproxy = new-object system.net.webproxy('http://proxy_internet:80')
[system.net.webrequest]::defaultwebproxy.credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
[system.net.webrequest]::defaultwebproxy.BypassProxyOnLocal = $true
#>
#netsh winhttp import proxy source=ie
netsh winhttp show proxy
netsh winhttp set proxy "proxy_internet:80" bypass-list="*.camara.gov.br;*.camara.leg.br;localhost"
$Wcl = new-object System.Net.WebClient
$Wcl.Headers.Add("user-agent", "PowerShell Script")
$Wcl.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
#[System.AppContext]::SetSwitch("System.Net.Http.UseSocketsHttpHandler", $false)