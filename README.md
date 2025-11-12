# TeamsMuteWatcher

TeamsMuteWatcher is a PowerShell script that monitors the mute state of Microsoft Teams and triggers webhooks based on:

- The current microphone mute state
- The presence of an active Teams window (i.e. “in a meeting”)

It is designed to integrate with external systems (e.g. IoT devices, Shelly plugs, custom lights) to provide visual or acoustic feedback for your meeting status.

Typical use cases:

- Control Shelly power plugs to switch on/off status lights so your family can see whether you are in a meeting (Teams window detected).
- Drive an “On Air” light on your desk based on the mute state in Teams.

> Note: The original version of this script and README was generated with GitHub Copilot and then refined manually.

## Features
- Detects the mute state of Microsoft Teams.
- Sends webhooks when the mute state changes:
  - `WebhookMuteOn`: Triggered when the microphone is muted.
  - `WebhookMuteOff`: Triggered when the microphone is unmuted.
- Sends webhooks based on the presence of a Teams window:
  - `WebHookTeamsWindowDetected`: Triggered when a Teams window is detected.
  - `WebHookNoTeamsWindowDetected`: Triggered when no Teams window is detected.
- Configurable polling interval and debounce times.
- Logs events to the console and optionally to a file.

## Requirements
- Windows PowerShell 5.1 or later.
- Microsoft Teams installed.
- UIAutomation assemblies.

## Installation
1. Clone this repository or download the `TeamsMuteWatcher.ps1` file.
2. Place the script in a directory of your choice.

## Usage
Run the script using PowerShell with the required parameters:

```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\Path\To\TeamsMuteWatcher.ps1" `
  -WebhookMuteOn "http://example.com/mute-on" `
  -WebhookMuteOff "http://example.com/mute-off" `
  -WebHookTeamsWindowDetected "http://example.com/teams-detected" `
  -WebHookNoTeamsWindowDetected "http://example.com/teams-not-detected" `
  -PollIntervalMs 300 `
  -MuteDebounceMs 150 `
  -UnmuteDebounceMs 150
```

### Parameters
- **`WebhookMuteOn`** (required): URL to trigger when the microphone is muted.
- **`WebhookMuteOff`** (required): URL to trigger when the microphone is unmuted.
- **`WebHookTeamsWindowDetected`** (optional): URL to trigger when a Teams window is detected.
- **`WebHookNoTeamsWindowDetected`** (optional): URL to trigger when no Teams window is detected.
- **`PollIntervalMs`** (optional): Polling interval in milliseconds (default: 300).
- **`MuteDebounceMs`** (optional): Debounce time for mute events in milliseconds (default: 150).
- **`UnmuteDebounceMs`** (optional): Debounce time for unmute events in milliseconds (default: 150).
- **`LogFile`** (optional): Path to a log file for event logging.
- **`DebugWindows`** (optional): Enables logging of top-level window names.
- **`DebugCandidates`** (optional): Enables logging of button candidates for mute detection.

## Example
```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\DEV\Tools\TeamsMuteWatcher.ps1" `
  -WebhookMuteOn "http://192.168.1.12/relay/0?turn=off" `
  -WebhookMuteOff "http://192.168.1.12/relay/0?turn=on" `
  -WebHookTeamsWindowDetected "http://192.168.1.12/relay/0?teams=detected" `
  -WebHookNoTeamsWindowDetected "http://192.168.1.12/relay/0?teams=not_detected" `
  -PollIntervalMs 300 `
  -MuteDebounceMs 150 `
  -UnmuteDebounceMs 150
```

## Logging
The script logs events to the console and optionally to a file if the `LogFile` parameter is specified. Log entries include timestamps and event details.

## License
This project is licensed under the MIT License. See the LICENSE file for details.

## Contributing
Contributions are welcome! Feel free to submit issues or pull requests to improve the script.

## Disclaimer
This script is provided as-is without any guarantees. Use it at your own risk.
