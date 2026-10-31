
Clear-Host

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "       EXPORTADOR DE XML FISCAL" -ForegroundColor Cyan
Write-Host "          by: Dudu/Carlinhos" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# ==========================================
# CONFIGURACOES DO BANCO
# ==========================================

$server = Read-Host "Informe o servidor SQL"
$database = Read-Host "Informe o banco de dados"
$sqlUser = Read-Host "Informe o usuario SQL"

Write-Host "Digite a senha SQL (nao sera exibida):" -ForegroundColor Yellow
$sqlPass = Read-Host -AsSecureString

# SqlCredential exige uma SecureString somente para leitura
$sqlPass.MakeReadOnly()

# ==========================================
# NUMERACAO DOS DOCUMENTOS
# ==========================================

do {
    $numeroInicialTexto = Read-Host "Informe a numeracao inicial"
    $numeroInicial = 0
    $validoInicial = [int]::TryParse($numeroInicialTexto, [ref]$numeroInicial)

    if (!$validoInicial -or $numeroInicial -lt 0) {
        Write-Host "Informe um numero inteiro valido." -ForegroundColor Red
    }
} while (!$validoInicial -or $numeroInicial -lt 0)

do {
    $numeroFinalTexto = Read-Host "Informe a numeracao final"
    $numeroFinal = 0
    $validoFinal = [int]::TryParse($numeroFinalTexto, [ref]$numeroFinal)

    if (!$validoFinal -or $numeroFinal -lt $numeroInicial) {
        Write-Host "A numeracao final deve ser maior ou igual a inicial." -ForegroundColor Red
    }
} while (!$validoFinal -or $numeroFinal -lt $numeroInicial)

# ==========================================
# MODELO FISCAL
# ==========================================

Write-Host ""
Write-Host "Selecione o modelo fiscal:" -ForegroundColor Cyan
Write-Host "1 - NFC-e (65)"
Write-Host "2 - NF-e (55)"

do {
    $opcaoModelo = Read-Host "Opcao"

    switch ($opcaoModelo) {
        "1" { $modelo = 65 }
        "2" { $modelo = 55 }
        default {
            Write-Host "Opcao invalida." -ForegroundColor Red
            $modelo = $null
        }
    }
} while ($null -eq $modelo)

# ==========================================
# PASTA DE DESTINO
# ==========================================

Write-Host ""
$outputFolder = Read-Host "Informe a pasta de destino (Enter para C:\Moura_\FALTANTES)"

if ([string]::IsNullOrWhiteSpace($outputFolder)) {
    $outputFolder = "C:\Moura_\FALTANTES"
}

if (!(Test-Path -LiteralPath $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
}

# ==========================================
# ARQUIVOS EXISTENTES
# ==========================================

Write-Host ""
Write-Host "Se um XML ja existir na pasta:" -ForegroundColor Cyan
Write-Host "S - Sobrescrever"
Write-Host "N - Preservar o arquivo existente"

do {
    $opcaoArquivo = (Read-Host "Deseja sobrescrever arquivos existentes? (S/N)").Trim().ToUpper()
} while ($opcaoArquivo -notin @("S", "N"))

$sobrescrever = ($opcaoArquivo -eq "S")

# ==========================================
# CONSULTA SQL
# ==========================================

$query = @"
SELECT
    NF.Formulario AS Codigo,
    TRY_CONVERT(XML, NF.XML_NFe) AS XmlNFe
FROM Fiscal_NF NF
WHERE NF.Formulario BETWEEN @Inicial AND @Final
  AND NF.Modelo = @Modelo
ORDER BY NF.Formulario
"@

$connectionString = "Server=$server;Database=$database;TrustServerCertificate=True;"

Add-Type -AssemblyName System.Data

$connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
$credential = New-Object System.Data.SqlClient.SqlCredential($sqlUser, $sqlPass)
$connection.Credential = $credential

$totalEncontrados = 0
$totalExportados = 0
$totalSemXML = 0
$totalIgnorados = 0

try {

    Write-Host ""
    Write-Host "Conectando ao banco de dados..." -ForegroundColor Yellow

    $connection.Open()

    Write-Host "Conectado com sucesso!" -ForegroundColor Green

    $command = $connection.CreateCommand()
    $command.CommandText = $query
    $command.CommandTimeout = 0

    $paramInicial = $command.Parameters.Add("@Inicial", [System.Data.SqlDbType]::Int)
    $paramInicial.Value = $numeroInicial

    $paramFinal = $command.Parameters.Add("@Final", [System.Data.SqlDbType]::Int)
    $paramFinal.Value = $numeroFinal

    $paramModelo = $command.Parameters.Add("@Modelo", [System.Data.SqlDbType]::Int)
    $paramModelo.Value = $modelo

    Write-Host ""
    Write-Host "Consultando documentos fiscais..." -ForegroundColor Yellow

    $reader = $command.ExecuteReader()

    while ($reader.Read()) {

        $totalEncontrados++

        $codigo = $reader["Codigo"].ToString().Trim()

        if ($reader.IsDBNull($reader.GetOrdinal("XmlNFe"))) {
            $totalSemXML++
            Write-Host "XML inexistente ou invalido: $codigo" -ForegroundColor DarkYellow
            continue
        }

        $xml = $reader.GetSqlXml($reader.GetOrdinal("XmlNFe")).Value

        if ([string]::IsNullOrWhiteSpace($xml)) {
            $totalSemXML++
            continue
        }

        $safeCodigo = $codigo

        foreach ($char in [System.IO.Path]::GetInvalidFileNameChars()) {
            $safeCodigo = $safeCodigo.Replace($char, "_")
        }

        $filePath = Join-Path $outputFolder ($safeCodigo + ".xml")

        if ((Test-Path -LiteralPath $filePath) -and !$sobrescrever) {
            $totalIgnorados++
            Write-Host "Arquivo ja existe, ignorado: $safeCodigo.xml" -ForegroundColor DarkYellow
            continue
        }

        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)

        [System.IO.File]::WriteAllText(
            $filePath,
            $xml,
            $utf8NoBom
        )

        $totalExportados++

        Write-Progress `
            -Activity "Exportando XMLs" `
            -Status "Processados: $totalEncontrados | Exportados: $totalExportados"
    }

    $reader.Close()

    Write-Progress -Activity "Exportando XMLs" -Completed

    # ==========================================
    # RESULTADO FINAL
    # ==========================================

    Write-Host ""
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "           EXPORTACAO CONCLUIDA" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan

    Write-Host "Servidor: $server"
    Write-Host "Banco: $database"
    Write-Host "Modelo: $modelo"
    Write-Host "Numeracao: $numeroInicial ate $numeroFinal"
    Write-Host ""

    Write-Host "Documentos encontrados: $totalEncontrados" -ForegroundColor White
    Write-Host "XMLs exportados: $totalExportados" -ForegroundColor Green
    Write-Host "Sem XML valido: $totalSemXML" -ForegroundColor Yellow
    Write-Host "Arquivos ignorados: $totalIgnorados" -ForegroundColor DarkYellow

    Write-Host ""
    Write-Host "Pasta de destino: $outputFolder" -ForegroundColor Cyan

}
catch {

    Write-Host ""
    Write-Host "ERRO NA EXPORTACAO:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red

}
finally {

    if ($null -ne $reader) {
        $reader.Dispose()
    }

    if ($null -ne $command) {
        $command.Dispose()
    }

    if ($connection.State -eq 'Open') {
        $connection.Close()
    }

    $connection.Dispose()
}

Write-Host ""
Read-Host "Pressione ENTER para finalizar"
