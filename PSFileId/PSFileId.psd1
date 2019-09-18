#
# Manifesto de módulo para o módulo 'PSFileId'
#
# Gerado por: Samuel Diniz Casimiro
#
# Gerado em: 21/08/2019
#

@{

# Arquivo de módulo de script ou módulo binário associado a este manifesto.
RootModule = 'PSFileId.psm1'

# Número da versão deste módulo.
ModuleVersion = '1.0.2'

# PSEditions com suporte
# CompatiblePSEditions = @()

# ID usada para identificar este módulo de forma exclusiva
GUID = 'fc3a5c2b-10e3-43d3-947d-5b2a976974bc'

# Autor deste módulo
Author = 'Samuel Diniz Casimiro'

# Empresa ou fornecedor deste módulo
CompanyName = 'Câmara dos Deputados'

# Instrução de direitos autorais para este módulo
Copyright = '(c) 2019 Samuel Diniz Casimiro. Todos os direitos reservados.'

# Descrição da funcionalidade fornecida por este módulo
Description = 'Compile the Windows API GetFileInformationByHandleEx function and exposes it by a PowerShell function called Get-ItemId.'

# A versão mínima do mecanismo do Windows PowerShell exigida por este módulo
PowerShellVersion = '5.1'

# Nome do host do Windows PowerShell exigido por este módulo
# PowerShellHostName = ''

# A versão mínima do host do Windows PowerShell exigida por este módulo
PowerShellHostVersion = '5.1.14393.2214'

# Minimum version of Microsoft .NET Framework required by this module. Este pré-requisito é válido somente para a edição PowerShell Desktop.
# DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module. Este pré-requisito é válido somente para a edição PowerShell Desktop.
# CLRVersion = ''

# Arquitetura de processador (None, X86, Amd64, IA64) exigida por este módulo
# ProcessorArchitecture = ''

# Módulos que devem ser importados para o ambiente global antes da importação deste módulo
# RequiredModules = @()

# Assemblies que devem ser carregados antes da importação deste módulo
# RequiredAssemblies = @()

# Arquivos de script (.ps1) executados no ambiente do chamador antes da importação deste módulo.
# ScriptsToProcess = @()

# Arquivos de tipo (.ps1xml) a serem carregados durante a importação deste módulo
# TypesToProcess = @()

# Arquivos de formato (.ps1xml) a serem carregados na importação deste módulo
# FormatsToProcess = @()

# Módulos para importação como módulos aninhados do módulo especificado em RootModule/ModuleToProcess
# NestedModules = @()

# Funções a serem exportadas deste módulo. Para melhor desempenho, não use curingas e não exclua a entrada. Use uma matriz vazia se não houver nenhuma função a ser exportada.
FunctionsToExport = @('Get-ItemId')

# Cmdlets a serem exportados deste módulo. Para melhor desempenho, não use curingas e não exclua a entrada. Use uma matriz vazia se não houver nenhum cmdlet a ser exportado.
CmdletsToExport = @()

# Variáveis a serem exportadas deste módulo
VariablesToExport = '*'

# Aliases a serem exportados deste módulo. Para melhor desempenho, não use curingas e não exclua a entrada. Use uma matriz vazia se não houver nenhum alias a ser exportado.
AliasesToExport = @()

# Recursos DSC a serem exportados deste módulo
# DscResourcesToExport = @()

# Lista de todos os módulos empacotados com este módulo
# ModuleList = @()

# Lista de todos os arquivos incluídos neste módulo
# FileList = @()

# Dados privados para passar para o módulo especificado em RootModule/ModuleToProcess. Também podem conter uma tabela de hash PSData com metadados adicionais do módulo usados pelo PowerShell.
PrivateData = @{

    PSData = @{

        # Tags aplicadas a este módulo. Elas ajudam na descoberta de módulos em galerias online.
        # Tags = @()

        # Uma URL para a licença deste módulo.
        # LicenseUri = ''

        # Uma URL para o site principal deste projeto.
        # ProjectUri = ''

        # Uma URL para um ícone representando este módulo.
        # IconUri = ''

        # ReleaseNotes deste módulo
        ReleaseNotes = 'Based in the PSBasicInfo module written by Vasily Larionov available at https://www.powershellgallery.com/packages/PSBasicInfo/1.0.3'

    } # Fim da tabela de hash PSData

} # Fim da tabela de hash PrivateData

# URI de HelpInfo deste módulo
# HelpInfoURI = ''

# Prefixo padrão dos comandos exportados deste módulo. Substitua o prefixo padrão usando Import-Module -Prefix.
# DefaultCommandPrefix = ''

}