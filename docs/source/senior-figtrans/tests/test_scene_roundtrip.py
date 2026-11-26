"""Exercise ownership and endpoint guards without creating an Office process.

Copyright (c) 2026 NCEPU-Senior. Author: XiaoN.
"""
import os
from pathlib import Path
import shutil
import subprocess
import unittest


ROOT = Path(__file__).resolve().parents[1]
POWERSHELL = shutil.which("powershell.exe") if os.name == "nt" else None


@unittest.skipUnless(POWERSHELL, "Windows PowerShell is required by the native scene helper")
class SceneRoundtripTests(unittest.TestCase):
    def run_guard(self, body):
        source = str(ROOT / "scripts" / "render_scene.ps1").replace("'", "''")
        code = r"""
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$tokens=$null; $errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile('__SOURCE__',[ref]$tokens,[ref]$errors)
if($errors.Count) { throw ($errors | Out-String) }
$names=@('Assert-VisioHostIdentity','Assert-VisioStartupGuard','P-ConnectorTargets','V-ConnectorTargets')
$ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -in $names},$true) | ForEach-Object { Invoke-Expression $_.Extent.Text }
function Expect-Failure([scriptblock]$Action) {
    $failed=$false
    try { & $Action | Out-Null } catch { $failed=$true }
    if(-not $failed) { throw 'Expected ownership or endpoint rejection' }
}
""".replace("__SOURCE__", source) + body
        run = subprocess.run([POWERSHELL, "-NoProfile", "-NonInteractive", "-Command", code],
                             capture_output=True, text=True, timeout=30)
        self.assertEqual(run.returncode, 0, run.stdout + run.stderr)

    def test_changed_hwnd_requires_the_same_verified_windows_pid(self):
        self.run_guard(r"""
$expected=[pscustomobject]@{hwnd=100;windows_pid=2048}
Assert-VisioHostIdentity $expected ([pscustomobject]@{hwnd=200;windows_pid=2048})
Expect-Failure { Assert-VisioHostIdentity $expected ([pscustomobject]@{hwnd=100;windows_pid=4096}) }
Expect-Failure { Assert-VisioHostIdentity $expected ([pscustomobject]@{hwnd=0;windows_pid=2048}) }
Expect-Failure { Assert-VisioHostIdentity $expected ([pscustomobject]@{hwnd=200;windows_pid=0}) }
""")

    def test_guard_id_and_exact_generated_title_are_both_required(self):
        self.run_guard(r"""
$token='sft-0123456789abcdef0123456789abcdef'
$title='Senior-FigTrans startup guard '+$token
$receipt=[pscustomobject]@{guard_document_id=0;guard_title=$title}
Assert-VisioStartupGuard ([pscustomobject]@{ID=0;Title=$title}) $receipt $token
Expect-Failure { Assert-VisioStartupGuard ([pscustomobject]@{ID=1;Title=$title}) $receipt $token }
Expect-Failure { Assert-VisioStartupGuard ([pscustomobject]@{ID=0;Title='unrelated drawing'}) $receipt $token }
Expect-Failure { Assert-VisioStartupGuard ([pscustomobject]@{ID=0;Title=$title}) $receipt 'sft-ffffffffffffffffffffffffffffffff' }
""")

    def test_powerpoint_checks_targets_after_both_ends_are_attached(self):
        self.run_guard(r"""
$a=[pscustomobject]@{ID=2;Name='input'}; $b=[pscustomobject]@{ID=3;Name='output'}
$format=[pscustomobject]@{BeginConnected=-1;EndConnected=-1;BeginConnectedShape=$a;EndConnectedShape=$b}
$shape=[pscustomobject]@{Name='flow';ConnectorFormat=$format}
$result=P-ConnectorTargets $shape $a $b
if($result.begin -cne 'input' -or $result.end -cne 'output') { throw 'Incorrect endpoint evidence' }
$format.EndConnectedShape=$a
Expect-Failure { P-ConnectorTargets $shape $a $b }
$format.EndConnectedShape=$b; $format.BeginConnected=0
Expect-Failure { P-ConnectorTargets $shape $a $b }
""")

    def test_visio_checks_both_named_endpoints_and_actual_target_ids(self):
        self.run_guard(r"""
$a=[pscustomobject]@{ID=2;NameU='input'}; $b=[pscustomobject]@{ID=3;NameU='output'}
$owner=[pscustomobject]@{ID=4}
$begin=[pscustomobject]@{FromCell=[pscustomobject]@{Name='BeginX'};FromSheet=$owner;ToSheet=$a}
$end=[pscustomobject]@{FromCell=[pscustomobject]@{Name='EndX'};FromSheet=$owner;ToSheet=$b}
$links=[pscustomobject]@{Count=2;Entries=@($begin,$end)}
$links | Add-Member ScriptMethod Item { param($index) return $this.Entries[$index-1] }
$shape=[pscustomobject]@{ID=4;NameU='flow';Connects=$links}
$result=V-ConnectorTargets $shape $a $b
if($result.begin -cne 'input' -or $result.end -cne 'output') { throw 'Incorrect endpoint evidence' }
$end.ToSheet=$a
Expect-Failure { V-ConnectorTargets $shape $a $b }
$end.ToSheet=$b; $end.FromCell.Name='BeginX'
Expect-Failure { V-ConnectorTargets $shape $a $b }
$end.FromCell.Name='EndX'; $links.Count=1
Expect-Failure { V-ConnectorTargets $shape $a $b }
""")


if __name__ == "__main__":
    unittest.main()
