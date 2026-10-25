# Copyright (c) 2026 NCEPU-Senior. Author: XiaoN.
<#
.SYNOPSIS
Initialize one Visio instance in an isolated Windows PowerShell 5.1 process.
.DESCRIPTION
Creates only an application and a uniquely titled, task-owned blank document.
It never opens a file, closes existing documents, or calls Application.Quit.
The caller must wait for exit 0, verify the recorded HWND and guard ID/title,
then reconnect to the running instance and close only that guard when finished.
The receipt is updated before each COM stage. A timeout does not establish that
this worker stopped: callers must retain its process handle and must not restart
or kill it automatically.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$ReceiptPath,
    [Parameter(Mandatory=$true)][string]$GuardToken
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$ReceiptPath=[IO.Path]::GetFullPath($ReceiptPath)
if($GuardToken -notmatch '^sft-[0-9a-f]{32}$') { throw 'GuardToken must be sft- followed by 32 lowercase hex digits.' }
if(Test-Path -LiteralPath $ReceiptPath) { throw "Refusing to overwrite an existing worker receipt: $ReceiptPath" }
if(-not (Test-Path -LiteralPath ([IO.Path]::GetDirectoryName($ReceiptPath)) -PathType Container)) { throw 'Worker receipt parent directory must already exist.' }
$state=[ordered]@{
    status='starting';stage='initialize';worker_pid=$PID;guard_token=$GuardToken
    guard_title=('Senior-FigTrans startup guard '+$GuardToken)
    guard_document_id=$null;hwnd=$null;windows_pid=$null;error=$null
    started_utc=[DateTime]::UtcNow.ToString('o');updated_utc=$null
}
function Write-State([string]$StageName) {
    $state.stage=$StageName
    $state.updated_utc=[DateTime]::UtcNow.ToString('o')
    $state | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ReceiptPath -Encoding UTF8
}
$hostApp=$null; $startupGuard=$null
try {
    Write-State 'create_application'
    $hostApp=New-Object -ComObject Visio.Application
    Write-State 'read_application_identity'
    $state.hwnd=[long]$hostApp.WindowHandle32
    if($state.hwnd -eq 0) { throw 'Visio did not expose a nonzero application HWND.' }
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; namespace SeniorFigTransBootstrap { public static class Window { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid); } }'
    [uint32]$nativePid=0
    $threadId=[SeniorFigTransBootstrap.Window]::GetWindowThreadProcessId([IntPtr]$state.hwnd,[ref]$nativePid)
    if($threadId -eq 0 -or $nativePid -eq 0) { throw 'Cannot identify the Windows process for the created Visio window.' }
    $state.windows_pid=$nativePid
    Write-State 'create_startup_guard'
    $startupGuard=$hostApp.Documents.Add('')
    $state.guard_document_id=[int]$startupGuard.ID
    Write-State 'title_startup_guard'
    $startupGuard.Title=[string]$state.guard_title
    if([string]$startupGuard.Title -cne [string]$state.guard_title) { throw 'Startup guard title verification failed.' }
    $state.status='ready'
    Write-State 'ready_for_reattachment'
    # Deliberately leave the application and blank guard alive. The caller owns
    # their recorded identity and will reconnect from another client process.
    exit 0
} catch {
    $state.status='failed'
    $state.error=[ordered]@{message=$_.Exception.Message;hresult=('{0:X8}' -f [int]$_.Exception.HResult);script_stack=$_.ScriptStackTrace}
    Write-State $state.stage
    exit 1
}
