[CmdletBinding()]
param(
    [string]$TempRoot = (Join-Path (Get-Location) '_tmp_test_outputs'),
    [string]$FixturePath = ''
)

$ErrorActionPreference = 'Stop'
$bridge = Join-Path $PSScriptRoot 'word_control.js'
$integrationBridge = Join-Path $PSScriptRoot 'test_word_control_integration.js'
$pureBridge = Join-Path $PSScriptRoot 'test_word_control_pure.cjs'
$nodeExecutable = (Get-Command node -ErrorAction Stop).Source
$createdRoot = -not (Test-Path -LiteralPath $TempRoot)
$runDir = Join-Path $TempRoot (Get-Date -Format 'yyyyMMdd_HHmmss_fff')
$wordPidsBefore = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

function Invoke-WordControl {
    param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Arguments)
    $previousPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $output = & cscript //nologo $bridge @Arguments 2>&1
        $exitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousPreference
    }
    if ($exitCode -ne 0) {
        throw "word-control failed: $($Arguments -join ' ')`n$($output -join [Environment]::NewLine)"
    }
    return ($output -join [Environment]::NewLine)
}

function Assert-WordControlFailure {
    param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Arguments)
    $previousPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $output = & cscript //nologo $bridge @Arguments 2>&1
        $exitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousPreference
    }
    if ($exitCode -eq 0) {
        throw "word-control unexpectedly succeeded: $($Arguments -join ' ')"
    }
    return ($output -join [Environment]::NewLine)
}

function Read-Utf8Json {
    param([Parameter(Mandatory = $true)][string]$Path)
    return [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8) | ConvertFrom-Json
}

function Wait-ForOwnedWordExit {
    param([int]$TimeoutSeconds = 20)
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    $stableSince = $null
    $newPids = @()
    do {
        $newPids = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
            Where-Object { $wordPidsBefore -notcontains $_.Id } |
            Select-Object -ExpandProperty Id)
        if ($newPids.Count -eq 0) {
            if ($null -eq $stableSince) {
                $stableSince = [DateTime]::UtcNow
            }
            elseif (([DateTime]::UtcNow - $stableSince).TotalSeconds -ge 1.5) {
                return @()
            }
        }
        else {
            $stableSince = $null
        }
        Start-Sleep -Milliseconds 250
    } while ([DateTime]::UtcNow -lt $deadline)
    return $newPids
}

New-Item -ItemType Directory -Path $runDir -Force | Out-Null

try {
    $pureOutput = & $nodeExecutable $pureBridge
    if ($LASTEXITCODE -ne 0) { throw 'pure regression checks failed' }
    $pureResult = ($pureOutput -join [Environment]::NewLine) | ConvertFrom-Json
    if (-not $pureResult.ok) { throw 'pure regression checks did not report success' }
    $equationInput = Join-Path $runDir 'equation.txt'
    [IO.File]::WriteAllText($equationInput, '\frac{a+b}{c} + \alpha', [Text.UTF8Encoding]::new($false))
    $convertedOutput = Join-Path $runDir 'converted-equation.json'
    $null = Invoke-WordControl convert-equation --input $equationInput --format latex --output $convertedOutput
    $converted = Read-Utf8Json $convertedOutput
    if (-not $converted.ok -or $converted.linear -notmatch 'a\+b' -or $converted.linear -notmatch [char]0x03B1) {
        throw 'equation conversion output was not as expected'
    }

    $unsupportedInput = Join-Path $runDir 'unsupported-equation.txt'
    [IO.File]::WriteAllText($unsupportedInput, '\matrix{a&b}', [Text.UTF8Encoding]::new($false))
    $null = Assert-WordControlFailure convert-equation --input $unsupportedInput --format latex
    $null = Assert-WordControlFailure replace-paragraph --index 1 --input $equationInput --yes

    $smokeDoc = Join-Path $runDir 'word-control-smoke.docx'
    $smoke = (Invoke-WordControl smoke --path $smokeDoc --yes) | ConvertFrom-Json
    if (-not $smoke.ok -or $smoke.table_count -ne 1 -or $smoke.equation_count -ne 1) {
        throw 'hidden Word smoke test returned unexpected object counts'
    }
    if (-not (Test-Path -LiteralPath $smokeDoc -PathType Leaf)) {
        throw 'hidden Word smoke test did not create its DOCX output'
    }
    $null = Assert-WordControlFailure smoke --path $smokeDoc --yes

    $commandIntegration = 'skipped-existing-word-session'
    if ($wordPidsBefore.Count -eq 0) {
        $smokePids = @(Wait-ForOwnedWordExit)
        if ($smokePids.Count -gt 0) {
            throw "hidden smoke test left Word process(es): $($smokePids -join ', ')"
        }
        $previousPreference = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            $integrationOutput = & cscript //nologo $integrationBridge --bridge $bridge --dir $runDir 2>&1
            $integrationExitCode = $LASTEXITCODE
        }
        finally {
            $ErrorActionPreference = $previousPreference
        }

        $newWordPids = @(Wait-ForOwnedWordExit)
        $integrationText = $integrationOutput -join [Environment]::NewLine
        if ($integrationExitCode -ne 0) {
            throw "isolated command integration failed; remaining Word process(es): $($newWordPids -join ', ')`n$integrationText"
        }
        if ($newWordPids.Count -gt 0) {
            throw "isolated integration test left Word process(es): $($newWordPids -join ', ')`n$integrationText"
        }
        $integration = $integrationText | ConvertFrom-Json
        if (-not $integration.ok -or -not $integration.advanced_tables) {
            throw 'isolated command integration returned an unsuccessful result'
        }
        $commandIntegration = 'passed'
    }

    $fixtureDocument = 'not-requested'
    if ($FixturePath) {
        if (-not (Test-Path -LiteralPath $FixturePath -PathType Leaf)) {
            throw "fixture document not found: $FixturePath"
        }
        $fixtureHashBefore = (Get-FileHash -LiteralPath $FixturePath -Algorithm SHA256).Hash
        $fixtureSource = Get-Item -LiteralPath $FixturePath
        $fixtureCopy = Join-Path $runDir ("fixture-copy" + $fixtureSource.Extension)
        Copy-Item -LiteralPath $FixturePath -Destination $fixtureCopy
        $fixtureCopyItem = Get-Item -LiteralPath $fixtureCopy
        if ($fixtureCopyItem.IsReadOnly) {
            $fixtureCopyItem.IsReadOnly = $false
        }

        if ($wordPidsBefore.Count -gt 0) {
            $fixtureDocument = 'skipped-existing-word-session'
        }
        else {
            $beforeFixturePids = @(Wait-ForOwnedWordExit)
            if ($beforeFixturePids.Count -gt 0) {
                throw "cannot start fixture test while task-owned Word process(es) remain: $($beforeFixturePids -join ', ')"
            }

            $previousPreference = $ErrorActionPreference
            $ErrorActionPreference = 'Continue'
            try {
                $fixtureOutput = & cscript //nologo $integrationBridge --bridge $bridge --dir $runDir --fixture-copy $fixtureCopy 2>&1
                $fixtureExitCode = $LASTEXITCODE
            }
            finally {
                $ErrorActionPreference = $previousPreference
            }

            $fixtureWordPids = @(Wait-ForOwnedWordExit -TimeoutSeconds 60)
            $fixtureHashAfter = (Get-FileHash -LiteralPath $FixturePath -Algorithm SHA256).Hash
            if ($fixtureHashAfter -ne $fixtureHashBefore) {
                throw 'the original fixture document hash changed during isolated testing'
            }

            $fixtureText = $fixtureOutput -join [Environment]::NewLine
            if ($fixtureExitCode -ne 0) {
                throw "fixture integration failed; remaining Word process(es): $($fixtureWordPids -join ', ')`n$fixtureText"
            }
            if ($fixtureWordPids.Count -gt 0) {
                throw "fixture integration left Word process(es): $($fixtureWordPids -join ', ')`n$fixtureText"
            }

            $fixtureDocument = $fixtureText | ConvertFrom-Json
            if (-not $fixtureDocument.ok -or -not $fixtureDocument.advanced_table_operations) {
                throw 'fixture integration returned an unsuccessful result'
            }
            $fixtureDocument | Add-Member -NotePropertyName original_hash_unchanged -NotePropertyValue $true
            $fixtureDocument | Add-Member -NotePropertyName source_bytes -NotePropertyValue $fixtureSource.Length
        }
    }

    # WSH eval and ConvertFrom-Json both accept some non-JSON control characters.
    # Validate every generated JSON artifact with the standard Node JSON parser.
    $strictOutput = & $nodeExecutable $pureBridge --validate-json-dir $runDir
    if ($LASTEXITCODE -ne 0) { throw 'strict JSON validation failed' }
    $strictResult = ($strictOutput -join [Environment]::NewLine) | ConvertFrom-Json
    if (-not $strictResult.ok) { throw 'strict JSON validation did not report success' }

    [pscustomobject]@{
        ok = $true
        pure_regression_groups = $pureResult.pure_regression_groups
        strict_json_files = $strictResult.strict_json_files
        equation_conversion = 'passed'
        unsupported_equation_rejection = 'passed'
        paragraph_replacement_disabled = 'passed'
        hidden_word_smoke = 'passed'
        overwrite_guard = 'passed'
        guarded_command_integration = $commandIntegration
        advanced_table_operations = $commandIntegration
        fixture_document = $fixtureDocument
    } | ConvertTo-Json -Compress -Depth 5
}
finally {
    if (Test-Path -LiteralPath $runDir) {
        Remove-Item -LiteralPath $runDir -Recurse -Force
    }
    if ($createdRoot -and (Test-Path -LiteralPath $TempRoot)) {
        $remaining = @(Get-ChildItem -LiteralPath $TempRoot -Force)
        if ($remaining.Count -eq 0) {
            Remove-Item -LiteralPath $TempRoot
        }
    }
}
