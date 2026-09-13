[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Bridge,
    [Parameter(Mandatory = $true)][string]$Directory,
    [Parameter(Mandatory = $true)][string]$Fixture
)
$ErrorActionPreference = 'Stop'
if (@(Get-Process WINWORD -ErrorAction SilentlyContinue).Count) { throw 'Save/output regression requires no existing Word session' }
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class WordControlSaveCancellation {
    public delegate void BeforeSave(object document, ref bool saveAsUi, ref bool cancel);
    public delegate void BeforeClose(object document, ref bool cancel);
    public delegate void AfterOpen(object document);
    public static int Calls;
    public static int CloseCalls;
    public const string LateText = "late close sentinel ";
    public static readonly Guid Events = new Guid("00020A01-0000-0000-C000-000000000046");
    public static readonly BeforeSave Handler = CancelSave;
    public static readonly BeforeClose CloseHandler = AmendBeforeClose;
    public static readonly AfterOpen OpenHandler = ObserveOpen;
    public static readonly System.Collections.Generic.List<int> OpenSecurity = new System.Collections.Generic.List<int>();
    public static string OpenError;
    static void ObserveOpen(object document) {
        object app = null;
        try {
            app = document.GetType().InvokeMember("Application", System.Reflection.BindingFlags.GetProperty, null, document, null);
            OpenSecurity.Add(Convert.ToInt32(app.GetType().InvokeMember("AutomationSecurity", System.Reflection.BindingFlags.GetProperty, null, app, null)));
        } catch (Exception error) { OpenError = error.Message; }
        finally { if (app != null) Marshal.ReleaseComObject(app); }
    }
    static void CancelSave(object document, ref bool saveAsUi, ref bool cancel) { Calls++; cancel = true; }
    static void AmendBeforeClose(object document, ref bool cancel) {
        object content = document.GetType().InvokeMember("Content", System.Reflection.BindingFlags.GetProperty, null, document, null);
        try {
            content.GetType().InvokeMember("InsertBefore", System.Reflection.BindingFlags.InvokeMethod, null, content, new object[] { LateText });
            CloseCalls++;
        } finally { Marshal.ReleaseComObject(content); }
    }
}
'@

function Invoke-Bridge([string[]]$Arguments, [switch]$ExpectFailure) {
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = (Get-Command cscript.exe -ErrorAction Stop).Source
    $quoted = @($Bridge) + $Arguments | ForEach-Object {
        if ($_.Contains('"') -or $_.EndsWith('\')) { throw 'Unsupported test argument' }
        '"' + $_ + '"'
    }
    $info.Arguments = '//nologo ' + ($quoted -join ' ')
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $process = [Diagnostics.Process]::Start($info)
    try {
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        $deadline = [DateTime]::UtcNow.AddSeconds(120)
        while (-not $process.HasExited) {
            # Pump the STA so Word's synchronous events reach our C# sink.
            [Windows.Forms.Application]::DoEvents()
            if ([DateTime]::UtcNow -gt $deadline) { $process.Kill(); throw 'Save/output command timed out' }
            Start-Sleep -Milliseconds 20
        }
        $outputText = $stdout.GetAwaiter().GetResult()
        $errorText = $stderr.GetAwaiter().GetResult()
        if ($ExpectFailure) {
            if ($process.ExitCode -eq 0) { throw "Command unexpectedly succeeded: $($Arguments -join ' ')" }
            return $errorText
        }
        if ($process.ExitCode -ne 0) { throw "Command failed: $($Arguments -join ' '): $errorText" }
        return $outputText | ConvertFrom-Json
    }
    finally { $process.Dispose() }
}

function Read-SharedHash([string]$Path) {
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try { (Get-FileHash -InputStream $stream -Algorithm SHA256).Hash }
    finally { $stream.Dispose() }
}

function Get-DocumentFootprint($Document) {
    $stories = @()
    foreach ($firstStory in $Document.StoryRanges) {
        $story = $firstStory
        while ($null -ne $story) {
            $stories += @([int]$story.StoryType, [string]$story.Text)
            $story = $story.NextStoryRange
        }
    }
    $bookmarks = @($Document.Bookmarks | ForEach-Object { @($_.Name, $_.Range.Start, $_.Range.End) })
    $cells = @($Document.Tables | ForEach-Object { foreach ($cell in $_.Range.Cells) {
        @([string]$cell.Range.Text, $cell.Range.Font.Name, $cell.Range.Font.Size, $cell.Range.Font.Bold,
          $cell.Shading.Texture, $cell.Shading.BackgroundPatternColor)
        foreach ($edge in -1,-2,-3,-4) { $border = $cell.Borders.Item($edge); @($border.LineStyle, $border.LineWidth, $border.Color) }
    } })
    @($stories, $bookmarks, $cells, $Document.Tables.Count, $Document.OMaths.Count, $Document.Comments.Count,
      $Document.Revisions.Count, $Document.Sections.Count, $Document.InlineShapes.Count, $Document.Shapes.Count,
      [bool]$Document.TrackRevisions) | ConvertTo-Json -Depth 8 -Compress
}

function Test-MacroDisabledOpen($WordApplication, $ExistingDocument, [string]$FixturePath, [string]$OutputDirectory) {
    $openSource = Join-Path $OutputDirectory 'open-guard-source.docx'
    if (Test-Path -LiteralPath $openSource) { throw 'Open-security fixture already exists' }
    Copy-Item -LiteralPath $FixturePath -Destination $openSource
    $diskBefore = Read-SharedHash $openSource
    $existingBefore = Get-DocumentFootprint $ExistingDocument
    $existingSaved = [bool]$ExistingDocument.Saved
    $securityBefore = [int]$WordApplication.AutomationSecurity
    $restoredModes = @()
    $openDoc = $null
    $openAttached = $false
    try {
        [Runtime.InteropServices.ComEventsHelper]::Combine($WordApplication, [WordControlSaveCancellation]::Events, 4, [WordControlSaveCancellation]::OpenHandler)
        $openAttached = $true
        foreach ($mode in 1,2,3) {
            $WordApplication.AutomationSecurity = $mode
            $callsBefore = [WordControlSaveCancellation]::OpenSecurity.Count
            $result = Invoke-Bridge @('open', '--path', $openSource)
            $WordApplication.Visible = $false
            if ([WordControlSaveCancellation]::OpenError -or [WordControlSaveCancellation]::OpenSecurity.Count -ne $callsBefore + 1 -or
                [WordControlSaveCancellation]::OpenSecurity[$callsBefore] -ne 3) { throw 'Word opened without the macro guard at its DocumentOpen event' }
            $restoredModes += [int]$WordApplication.AutomationSecurity
            if (-not $result.ok -or $result.opened -ne $true -or $result.automation_security_restored -ne $true -or
                $result.errors.Count -ne 0 -or $WordApplication.AutomationSecurity -ne $mode) { throw 'Open did not confirm success and security restoration' }
            $openDoc = $WordApplication.ActiveDocument
            if ($openDoc.FullName -ne $openSource -or $WordApplication.Documents.Count -ne 2) { throw 'Open returned a different document or lost the existing document' }
            if ((Get-DocumentFootprint $ExistingDocument) -ne $existingBefore -or [bool]$ExistingDocument.Saved -ne $existingSaved) { throw 'Open changed the existing document' }
            $footprint = Get-DocumentFootprint $openDoc
            if ($mode -eq 1) { $openedBefore = $footprint } elseif ($footprint -ne $openedBefore) { throw 'Security mode changed the opened document content or formatting' }
            $openDoc.Close([ref][object]0)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($openDoc)
            $openDoc = $null
            if ((Read-SharedHash $openSource) -ne $diskBefore) { throw 'Open changed source bytes' }
        }
        [pscustomobject]@{ok=$true; observed_modes=@([WordControlSaveCancellation]::OpenSecurity.ToArray());
            restored_modes=$restoredModes; existing_document_unchanged=$true; source_bytes_unchanged=$true}
    } finally {
        if ($openAttached) { [void][Runtime.InteropServices.ComEventsHelper]::Remove($WordApplication, [WordControlSaveCancellation]::Events, 4, [WordControlSaveCancellation]::OpenHandler) }
        if ($openDoc) { $openDoc.Close([ref][object]0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($openDoc) }
        $WordApplication.AutomationSecurity = $securityBefore
    }
}

$source = Join-Path $Directory 'save-guards-source.docx'
$backup = Join-Path $Directory 'save-guards-backup.docx'
$pdf = Join-Path $Directory 'save-guards-preview.pdf'
foreach ($target in @($source, $backup, $pdf)) { if (Test-Path -LiteralPath $target) { throw "Test output already exists: $target" } }
Copy-Item -LiteralPath $Fixture -Destination $source
$word = $null
$doc = $null
$attached = $false
$closeAttached = $false
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    # Word's optional COM arguments are ref object, not ref string.
    $doc = $word.Documents.Open([ref][object][string]$source)
    $version = [string]$word.Version
    $diskBefore = Read-SharedHash $source
    $doc.Content.InsertBefore('unsaved sentinel ')
    $before = Get-DocumentFootprint $doc
    [Runtime.InteropServices.ComEventsHelper]::Combine($word, [WordControlSaveCancellation]::Events, 8, [WordControlSaveCancellation]::Handler)
    $attached = $true
    foreach ($command in 'save-active','close-active') {
        $argsForCommand = @($command, '--expect-path', $source, '--yes')
        if ($command -eq 'close-active') { $argsForCommand += '--save' }
        $callsBefore = [WordControlSaveCancellation]::Calls
        $null = Invoke-Bridge $argsForCommand -ExpectFailure
        if ([WordControlSaveCancellation]::Calls -le $callsBefore) { throw 'Word did not raise the save-cancellation event' }
        if ($word.Documents.Count -ne 1 -or $doc.Saved) { throw 'Cancelled save closed or marked the document saved' }
        if ((Get-DocumentFootprint $doc) -ne $before -or (Read-SharedHash $source) -ne $diskBefore) { throw 'Cancelled save changed content, objects, formatting or disk bytes' }
    }
    [void][Runtime.InteropServices.ComEventsHelper]::Remove($word, [WordControlSaveCancellation]::Events, 8, [WordControlSaveCancellation]::Handler)
    $attached = $false
    $saved = Invoke-Bridge @('save-active', '--expect-path', $source, '--yes')
    if (-not $saved.ok -or -not $doc.Saved) { throw 'Save did not recover after event removal' }
    $diskBefore = Read-SharedHash $source
    $before = Get-DocumentFootprint $doc
    $selectionBefore = @($word.Selection.Range.StoryType, $word.Selection.Start, $word.Selection.End) -join ':'
    foreach ($command in 'save-copy','export-pdf') {
        $target = if ($command -eq 'save-copy') { $backup } else { $pdf }
        [IO.File]::WriteAllText($target, 'prior output must survive', [Text.UTF8Encoding]::new($false))
        $previousHash = Read-SharedHash $target
        $lock = [IO.File]::Open($target, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        try { $null = Invoke-Bridge @($command, '--path', $target, '--overwrite', '--expect-path', $source, '--yes') -ExpectFailure }
        finally { $lock.Dispose() }
        if ((Read-SharedHash $target) -ne $previousHash) { throw 'Locked old output was damaged' }
        $result = Invoke-Bridge @($command, '--path', $target, '--overwrite', '--expect-path', $source, '--yes')
        if (-not $result.ok -or $result.cleanup_warning) { throw 'Output replacement did not complete cleanly' }
        if ($command -eq 'save-copy' -and (Read-SharedHash $target) -ne $diskBefore) { throw 'Saved backup differs from source bytes' }
        if ($command -eq 'export-pdf' -and [Text.Encoding]::ASCII.GetString([IO.File]::ReadAllBytes($target), 0, 5) -ne '%PDF-') { throw 'Export is not a PDF' }
        if ((Get-DocumentFootprint $doc) -ne $before -or (Read-SharedHash $source) -ne $diskBefore -or -not $doc.Saved) { throw 'Backup/export changed the source document' }
        if ((@($word.Selection.Range.StoryType, $word.Selection.Start, $word.Selection.End) -join ':') -ne $selectionBefore) { throw 'Backup/export changed selection' }
    }
    foreach ($target in @($source, "$source`:preview.pdf", (Join-Path $Directory '*.pdf'))) {
        $null = Invoke-Bridge @('export-pdf', '--path', $target, '--overwrite', '--expect-path', $source, '--yes') -ExpectFailure
        if ((Read-SharedHash $source) -ne $diskBefore -or (Get-DocumentFootprint $doc) -ne $before) { throw 'Rejected PDF path changed source' }
    }
    if (@(Get-ChildItem -LiteralPath $Directory -Directory -Filter '.word-control-*').Count) { throw 'Staging directories leaked' }
    $doc.Content.InsertBefore('close-save sentinel ')
    $expected = Get-DocumentFootprint $doc
    # An add-in can introduce an edit after the bridge's save check, during close.
    [Runtime.InteropServices.ComEventsHelper]::Combine($word, [WordControlSaveCancellation]::Events, 6, [WordControlSaveCancellation]::CloseHandler)
    $closeAttached = $true
    $closed = Invoke-Bridge @('close-active', '--save', '--expect-path', $source, '--yes')
    if (-not $closed.ok -or $word.Documents.Count -ne 0) { throw 'Confirmed save did not close the document' }
    [void][Runtime.InteropServices.ComEventsHelper]::Remove($word, [WordControlSaveCancellation]::Events, 6, [WordControlSaveCancellation]::CloseHandler)
    $closeAttached = $false
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)
    $doc = $word.Documents.Open([ref][object][string]$source)
    if ([WordControlSaveCancellation]::CloseCalls -ne 1 -or -not ([string]$doc.Content.Text).StartsWith([WordControlSaveCancellation]::LateText)) { throw 'Close discarded the edit introduced after save confirmation' }
    $prefix = $doc.Range([ref][object]0, [ref][object]([WordControlSaveCancellation]::LateText.Length))
    $prefix.Text = ''
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($prefix)
    if ((Get-DocumentFootprint $doc) -ne $expected) { throw 'Saved/reopened document differs from expected content or structure' }
    $macroDisabledOpen = Test-MacroDisabledOpen $word $doc $Fixture $Directory
    [pscustomobject]@{ok=$true; word_version=$version; cancellation_events=[WordControlSaveCancellation]::Calls;
        locked_outputs_preserved=2; pdf_paths_rejected=3; source_scope_unchanged=$true; saved_readback=$true; late_close_edit_preserved=$true;
        macro_disabled_open=$macroDisabledOpen} | ConvertTo-Json -Depth 5
}
finally {
    if ($attached) { [void][Runtime.InteropServices.ComEventsHelper]::Remove($word, [WordControlSaveCancellation]::Events, 8, [WordControlSaveCancellation]::Handler) }
    if ($closeAttached) { [void][Runtime.InteropServices.ComEventsHelper]::Remove($word, [WordControlSaveCancellation]::Events, 6, [WordControlSaveCancellation]::CloseHandler) }
    if ($doc) { $doc.Close([ref][object]0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) }
    if ($word) { $word.Quit([ref][object]0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
