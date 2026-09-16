<#
  工程表ツール（Gantt-chart アドイン）の manifest.xml を GitHub の最新版に合わせる。

    auto-update.ps1          … 今すぐ manifest.xml を最新版にする（自動更新タスクからはこの形で呼ばれる）
    auto-update.ps1 -Setup   … 上記を「毎日 9:00 ＋ ログオン時」に自動実行するタスクを登録する

  対象の manifest.xml は、このスクリプトと同じフォルダーにあるもの。
  taskpane（html/js/css）は GitHub Pages から毎回読み込まれるので更新不要。
  ローカルに置く必要がある manifest.xml だけが自動更新の対象。
#>
param([switch]$Setup)

$ErrorActionPreference = 'Stop'

$AddinId  = 'dfc3fd23-ff19-4d33-b312-a15d117dd27d'
$RawUrl   = 'https://raw.githubusercontent.com/manh25xx08xx99-sys/Gantt-chart/main/manifest.xml'
$RegKey   = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer'
$TaskName = 'GanttAddin_AutoUpdateManifest'
$Manifest = Join-Path $PSScriptRoot 'manifest.xml'

function Update-Manifest {
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  $tmp = "$Manifest.download"
  try {
    Invoke-WebRequest -Uri $RawUrl -OutFile $tmp -UseBasicParsing -Headers @{ 'Cache-Control' = 'no-cache' }

    # 壊れたファイルや別物で上書きしないよう、XMLとして読めるか・想定のアドインかを確認する
    $xml = [xml](Get-Content -LiteralPath $tmp -Raw -Encoding UTF8)
    if ($xml.OfficeApp.Id -ne $AddinId) {
      throw "ダウンロードした manifest.xml のIdが想定と違います: $($xml.OfficeApp.Id)"
    }

    $changed = $true
    if (Test-Path -LiteralPath $Manifest) {
      $changed = (Get-FileHash -LiteralPath $tmp).Hash -ne (Get-FileHash -LiteralPath $Manifest).Hash
    }
    if ($changed) {
      Move-Item -LiteralPath $tmp -Destination $Manifest -Force
      Write-Host "manifest.xml を更新しました（バージョン $($xml.OfficeApp.Version)）。"
    } else {
      Write-Host "manifest.xml は既に最新です（バージョン $($xml.OfficeApp.Version)）。"
    }

    # 登録が消えた場合やフォルダーを移動した場合にも追従できるよう、毎回パスを入れ直す
    if (-not (Test-Path -LiteralPath $RegKey)) { New-Item -Path $RegKey -Force | Out-Null }
    New-ItemProperty -Path $RegKey -Name $AddinId -Value $Manifest -PropertyType String -Force | Out-Null
  }
  finally {
    if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
  }
}

function Register-DailyTask {
  $argLine = '-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "{0}"' -f $PSCommandPath
  try {
    $action   = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argLine
    $triggers = @((New-ScheduledTaskTrigger -Daily -At '09:00'), (New-ScheduledTaskTrigger -AtLogOn))
    # PCが消えていた時間帯の分も、次に使えるようになった時点で実行する
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries `
      -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $triggers -Settings $settings `
      -Description '工程表ツールの manifest.xml を GitHub の最新版に合わせる' -Force | Out-Null
  } catch {
    # ScheduledTasks モジュールが使えない環境向けのフォールバック（毎日のみ）
    & schtasks.exe /create /tn $TaskName /tr "`"powershell.exe`" $argLine" /sc daily /st 09:00 /f | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "自動更新タスクの登録に失敗しました。" }
  }
  Write-Host "自動更新タスク『$TaskName』を登録しました（毎日 9:00 ＋ ログオン時）。"
}

if ($Setup) {
  # ネットに繋がらない等で更新できなくても、タスク登録だけは進める
  try { Update-Manifest } catch { Write-Host "※今回の更新確認は失敗しました: $($_.Exception.Message)" }
  Register-DailyTask
} else {
  Update-Manifest
}
