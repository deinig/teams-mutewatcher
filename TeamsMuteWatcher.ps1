# TeamsMuteWatcher.ps1 (v1.8, PS 5.1 kompatibel)
# LED an = Teams muted, LED aus = Teams unmuted (UIAutomation, titelbasiert, deutsche Tooltip-/Textregeln)

# Call
# powershell.exe -ExecutionPolicy Bypass -File "C:\DEV\Tools\TeamsMuteWatcher.ps1" `
#  -WebhookMuteOn  "http://192.168.1.12/relay/0?turn=off" `
#  -WebhookMuteOff "http://192.168.1.12/relay/0?turn=on" 

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][string]$WebhookMuteOn,
  [Parameter(Mandatory=$true)][string]$WebhookMuteOff,
  [int][ValidateRange(100,5000)]$PollIntervalMs = 300,
  [int][ValidateRange(0,3000)]$MuteDebounceMs   = 150,
  [int][ValidateRange(0,3000)]$UnmuteDebounceMs = 150,
  [int][ValidateRange(1,120)]$HttpTimeoutSec = 5,
  [int][ValidateRange(1,10)]$MaxRetries = 3,
  [string]$LogFile = "",
  [switch]$DebugWindows,     # listet Top-Level-Fenster
  [switch]$DebugCandidates   # loggt den zuerst passenden Button (Name/HelpText/Quelle)
)

function Write-Log { param([string]$msg,[string]$level="INFO")
  $line = "$(Get-Date -Format o) [$level] $msg"
  Write-Host $line
  if ($LogFile) { Add-Content -Path $LogFile -Value $line }
}

function Invoke-WebhookGet {
  param([string]$Url,[int]$MaxRetriesLocal=$MaxRetries)
  $delay = 300
  for ($i=0; $i -lt $MaxRetriesLocal; $i++) {
    try {
      $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -Method GET -TimeoutSec $HttpTimeoutSec
      if ($resp.StatusCode -ge 200 -and $resp.StatusCode -lt 300) { return $true }
      Write-Log "Webhook $Url -> HTTP $($resp.StatusCode)" "WARN"
    } catch {
      Write-Log "Webhook $Url Fehler: $($_.Exception.Message)" "WARN"
    }
    if ($delay -gt 4000) { $delay = 4000 }
    Start-Sleep -Milliseconds $delay
    $delay *= 2
  }
  return $false
}

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes | Out-Null

# Titel für Teams/Meetings (de/en)
$teamsTitleRegex = '(?i)\b(teams|microsoft\s+teams|besprechung|meeting)\b'

# Button-Namen/Varianten – inkl. „Mikro“
$muteNameRegex = '(?i)\b(mikro|mikrofon|stumm|stummschalten|stummschaltung|wieder\s+aktivieren|mute|unmute|micro|microphone)\b'

$condTopTrue = [System.Windows.Automation.Condition]::TrueCondition
$condButton  = New-Object System.Windows.Automation.PropertyCondition `
  ([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)

# Harte deutsche Tooltip-/Aktions-Texte
$HELP_UNMUTED_TO_MUTED = 'Mikrofon stummschalten'                 # Aktion zeigt: aktuell UNMUTED
$HELP_MUTED_TO_UNMUTED = 'Stummschaltung des Mikrofons aufheben'  # Aktion zeigt: aktuell MUTED

function Infer-State-From-Text([string]$name,[string]$help) {
  $n = (($name | Out-String).Trim()).ToLower()
  $h = (($help | Out-String).Trim()).ToLower()

  # Exakte deutsche Formulierungen (Name ODER HelpText):
  # Aktion „… aufheben“ oder „Mikrofon wieder aktivieren“ => aktuell MUTED
  if ($h -eq $HELP_MUTED_TO_UNMUTED.ToLower() -or
      $n -eq $HELP_MUTED_TO_UNMUTED.ToLower() -or
      $n -eq 'mikrofon wieder aktivieren' -or
      $n -match '\bwieder\s+aktivieren\b') {
    return $true
  }

  # Aktion „Mikrofon stummschalten“ => aktuell UNMUTED
  if ($h -eq $HELP_UNMUTED_TO_MUTED.ToLower() -or
      $n -eq $HELP_UNMUTED_TO_MUTED.ToLower() -or
      $n -match '\bstummschalten\b') {
    return $false
  }

  # Eventuelle Zustandsformulierung
  if ($n -match 'ist\s+stumm' -or $n -match 'stummgeschaltet' -or
      $h -match 'ist\s+stumm' -or $h -match 'stummgeschaltet') {
    return $true
  }

  return $null
}

function Get-TeamsMuteToggleState {
  # Rückgabe: $true (muted), $false (unmuted), $null (kein Meeting/Element)
  try {
    $desktop = [System.Windows.Automation.AutomationElement]::RootElement
    if (-not $desktop) { return $null }

    $tops = $desktop.FindAll([System.Windows.Automation.TreeScope]::Children, $condTopTrue)
    if ($DebugWindows) {
      for ($i=0; $i -lt $tops.Count; $i++) {
        $n = $tops.Item($i).Current.Name
        $off = $tops.Item($i).Current.IsOffscreen
        Write-Log "TopWindow: `"$n`" Offscreen=$off" "DEBUG"
      }
    }

    $teamsWindowFound = $false

    for ($i=0; $i -lt $tops.Count; $i++) {
      $w = $tops.Item($i)
      $title = $w.Current.Name
      if ([string]::IsNullOrWhiteSpace($title)) { continue }
      if ($title -notmatch $teamsTitleRegex) { continue }

      $teamsWindowFound = $true

      # Subtree durchsuchen (auch Offscreen)
      $buttons = $w.FindAll([System.Windows.Automation.TreeScope]::Subtree, $condButton)
      for ($b=0; $b -lt $buttons.Count; $b++) {
        $btn = $buttons.Item($b)
        $name = $btn.Current.Name
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if ($name -notmatch $muteNameRegex) { continue }

        $help = ""
        try { $help = $btn.Current.HelpText } catch { $help = "" }
        if ($DebugCandidates) { Write-Log "Kandidat: Name='$name' Help='$help' Title='$title'" "DEBUG" }

        # 1) TogglePattern
        $toggle = $null
        try { $toggle = $btn.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern) -as [System.Windows.Automation.TogglePattern] } catch { $toggle = $null }
        if ($toggle) {
          switch ($toggle.Current.ToggleState) {
            ([System.Windows.Automation.ToggleState]::On)  { if ($DebugCandidates){Write-Log "Quelle: TogglePattern => muted" "DEBUG"};  return $true  }
            ([System.Windows.Automation.ToggleState]::Off) { if ($DebugCandidates){Write-Log "Quelle: TogglePattern => unmuted" "DEBUG"}; return $false }
          }
        }

        # 2) LegacyIAccessible (Pressed/Checked)
        $legacy = $null
        try { $legacy = $btn.GetCurrentPattern([System.Windows.Automation.LegacyIAccessiblePattern]::Pattern) -as [System.Windows.Automation.LegacyIAccessiblePattern] } catch { $legacy = $null }
        if ($legacy) {
          $stateFlags = [int]$legacy.Current.State
          $isChecked = ($stateFlags -band 0x10) -ne 0  # CHECKED
          $isPressed = ($stateFlags -band 0x08) -ne 0  # PRESSED
          if ($DebugCandidates){ Write-Log "Quelle: LegacyIAccessible state=$stateFlags (Checked=$isChecked Pressed=$isPressed)" "DEBUG" }
          if ($isChecked -or $isPressed) { return $true } else { return $false }
        }

        # 3) Tooltip/Name (de) – inkl. „Mikrofon wieder aktivieren“
        $fromText = Infer-State-From-Text $name $help
        if ($fromText -ne $null) {
          if ($DebugCandidates){ Write-Log "Quelle: Text/Tooltip => $fromText" "DEBUG" }
          return $fromText
        }

        # Falls nicht eindeutig → nächsten Kandidaten prüfen
      }
    }

    # If no Teams window is found, assume the session has ended and mute the microphone
    if (-not $teamsWindowFound) {
      Write-Log "Teams window not found. Sending mute command." "INFO"
      if (Invoke-WebhookGet -Url $WebhookMuteOn) {
        Write-Log "MUTE=ON (Teams closed)"
      } else {
        Write-Log "MUTE=ON Webhook failed (Teams closed)" "ERROR"
      }
    }

    return $null
  } catch {
    Write-Log "UIA-Fehler: $($_.Exception.Message)" "WARN"
    return $null
  }
}

# ---- Hauptloop ----
Write-Log "TeamsMuteWatcher gestartet. Intervall=${PollIntervalMs}ms, Debounce Mute=$MuteDebounceMs ms / Unmute=$UnmuteDebounceMs ms"
Write-Log "MUTE-ON  -> $WebhookMuteOn"
Write-Log "MUTE-OFF -> $WebhookMuteOff"

$stateMuted     = $false
$lastRaw        = $null
$lastChangeTime = Get-Date

$raw = Get-TeamsMuteToggleState
$lastRaw = $raw
if ($raw -is [bool]) {
  $stateMuted = $raw
  Write-Log "Initialer Teams-Mute: $stateMuted"
} else {
  Write-Log "Kein Teams-Meeting (noch) erkannt."
  # Send mute command if no Teams window is found on startup
  Write-Log "Teams window not found on startup. Sending mute command." "INFO"
  if (Invoke-WebhookGet -Url $WebhookMuteOn) {
    Write-Log "MUTE=ON (Startup)"
  } else {
    Write-Log "MUTE=ON Webhook failed (Startup)" "ERROR"
  }
}

while ($true) {
  try {
    $raw = Get-TeamsMuteToggleState  # $true/$false/$null

    if ($raw -ne $lastRaw) {
      $lastRaw = $raw
      $lastChangeTime = Get-Date
    }

    if ($raw -is [bool]) {
      $elapsed = ((Get-Date) - $lastChangeTime).TotalMilliseconds
      if ($raw) { $needed = $MuteDebounceMs } else { $needed = $UnmuteDebounceMs }

      if ($elapsed -ge $needed -and $stateMuted -ne $raw) {
        $stateMuted = $raw
        if ($stateMuted) {
          if (Invoke-WebhookGet -Url $WebhookMuteOn) { Write-Log "MUTE=ON (Teams)" }
          else { Write-Log "MUTE=ON Webhook fehlgeschlagen" "ERROR" }
        } else {
          if (Invoke-WebhookGet -Url $WebhookMuteOff) { Write-Log "MUTE=OFF (Teams)" }
          else { Write-Log "MUTE=OFF Webhook fehlgeschlagen" "ERROR" }
        }
      }
    } else {
      # If no Teams window is found, ensure mute is on
      if (-not $stateMuted) {
        Write-Log "No Teams window detected. Sending mute command." "INFO"
        if (Invoke-WebhookGet -Url $WebhookMuteOn) {
          Write-Log "MUTE=ON (No Teams window)"
          $stateMuted = $true
        } else {
          Write-Log "MUTE=ON Webhook failed (No Teams window)" "ERROR"
        }
      }
    }
  } catch {
    Write-Log "Loop-Fehler: $($_.Exception.Message)" "WARN"
  }
  Start-Sleep -Milliseconds $PollIntervalMs
}
