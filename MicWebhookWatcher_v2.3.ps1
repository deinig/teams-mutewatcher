# MicWebhookWatcher_v2.3.ps1
# Reagiert auf Mikrofon-Nutzung (Windows 11) und triggert Webhooks (HTTP GET).
# Änderung ggü. v2.2: KEINE EnumAudioEndpoints/IMMDeviceCollection mehr (E_NOINTERFACE-Fix).
# Stattdessen werden die beiden Default-Capture-Endpunkte (Multimedia & Communications) geprüft.

# Call:
# powershell.exe -ExecutionPolicy Bypass -File "C:\DEV\Tools\MicWebhookWatcher_v2.3.ps1" `
#  -WebhookOn "http://192.168.1.11/relay/0?turn=on" `
#  -WebhookOff "http://192.168.1.11/relay/0?turn=off"

param(
    [Parameter(Mandatory=$true)][string]$WebhookOn,
    [Parameter(Mandatory=$true)][string]$WebhookOff,

    [int]$PollIntervalMs = 300,
    [int]$OnDebounceMs   = 150,
    [int]$OffDebounceMs  = 500,

    [int]$HttpTimeoutSec = 5,
    [int]$MaxRetries     = 3,
    [switch]$SkipTlsCheck,

    [string]$LogFile = "",
    [int]$MaxLogBytes = 1048576
)

function Write-Log {
    param([string]$msg, [string]$level="INFO")
    $line = "$(Get-Date -Format o) [$level] $msg"
    Write-Host $line
    if ($LogFile) {
        try {
            if (Test-Path $LogFile) {
                $len = (Get-Item $LogFile).Length
                if ($len -gt $MaxLogBytes) { Move-Item -Force $LogFile "$LogFile.bak" -ErrorAction SilentlyContinue }
            }
            Add-Content -Path $LogFile -Value $line
        } catch { }
    }
}

if ($SkipTlsCheck) {
    add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public static class TrustAllCerts {
    public static void Enable() {
        ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, errors) => true;
    }
}
"@
    [TrustAllCerts]::Enable()
    Write-Log "TLS-Zertifikatsprüfung deaktiviert" "WARN"
}

# --------- C# Interop nur mit GetDefaultAudioEndpoint + IAudioSessionManager2 ----------
$cs = @"
using System;
using System.Runtime.InteropServices;

[ComVisible(true)]
public static class MicProbe
{
    enum EDataFlow { eRender=0, eCapture=1, eAll=2 }
    enum ERole { eConsole=0, eMultimedia=1, eCommunications=2 }
    enum AudioSessionState { Inactive=0, Active=1, Expired=2 }

    [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    class MMDeviceEnumeratorCom {}

    [ComImport, Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMMDeviceEnumerator
    {
        int NotImpl1(); // EnumAudioEndpoints (nicht verwendet)
        int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppDevice);
        int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string pwstrId, out IMMDevice ppDevice);
        int RegisterEndpointNotificationCallback(IntPtr pClient);
        int UnregisterEndpointNotificationCallback(IntPtr pClient);
    }

    [ComImport, Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMMDevice
    {
        int Activate(ref Guid iid, int dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
        int OpenPropertyStore(int stgmAccess, out IntPtr ppProperties);
        int GetId([MarshalAs(UnmanagedType.LPWStr)] out string ppstrId);
        int GetState(out uint pdwState);
    }

    [ComImport, Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IAudioSessionManager2
    {
        int NotImpl1();
        int NotImpl2();
        int GetSessionEnumerator(out IAudioSessionEnumerator SessionEnum);
        int RegisterSessionNotification(IntPtr SessionNotification);
        int UnregisterSessionNotification(IntPtr SessionNotification);
        int NotImpl3();
        int NotImpl4();
    }

    [ComImport, Guid("E2F5BB11-0570-40CA-ACDD-3AA01277DEE8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IAudioSessionEnumerator
    {
        int GetCount(out int SessionCount);
        int GetSession(int SessionIndex, out IAudioSessionControl Session);
    }

    [ComImport, Guid("F4B1A599-7266-4319-A8CA-E70ACB11E8CD"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IAudioSessionControl
    {
        int GetState(out AudioSessionState pRetVal);
        int GetDisplayName([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);
        int SetDisplayName([MarshalAs(UnmanagedType.LPWStr)] string Value, ref Guid EventContext);
        int GetIconPath([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);
        int SetIconPath([MarshalAs(UnmanagedType.LPWStr)] string Value, ref Guid EventContext);
        int GetGroupingParam(out Guid pRetVal);
        int SetGroupingParam(ref Guid Override, ref Guid EventContext);
        int RegisterAudioSessionNotification(IntPtr NewNotifications);
        int UnregisterAudioSessionNotification(IntPtr NewNotifications);
    }

    static readonly Guid IID_IAudioSessionManager2 = new Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F");

    static bool DeviceHasActiveSessions(IMMDevice dev)
    {
        object mgrObj = null;
        IAudioSessionManager2 mgr = null;
        IAudioSessionEnumerator sesEnum = null;

        try
        {
            int CLSCTX_ALL = 23;
            Guid iid = IID_IAudioSessionManager2;
            int hr = dev.Activate(ref iid, CLSCTX_ALL, IntPtr.Zero, out mgrObj);
            if (hr != 0 || mgrObj == null) return false;

            mgr = (IAudioSessionManager2)mgrObj;
            hr = mgr.GetSessionEnumerator(out sesEnum);
            if (hr != 0 || sesEnum == null) return false;

            int sc = 0;
            sesEnum.GetCount(out sc);
            for (int s = 0; s < sc; s++)
            {
                IAudioSessionControl ctrl = null;
                try
                {
                    sesEnum.GetSession(s, out ctrl);
                    if (ctrl != null)
                    {
                        AudioSessionState st;
                        ctrl.GetState(out st);
                        if (st == AudioSessionState.Active)
                            return true;
                    }
                }
                finally
                {
                    if (ctrl != null) Marshal.ReleaseComObject(ctrl);
                }
            }
            return false;
        }
        finally
        {
            if (sesEnum != null) Marshal.ReleaseComObject(sesEnum);
            if (mgr != null) Marshal.ReleaseComObject(mgr);
            if (mgrObj != null) Marshal.ReleaseComObject(mgrObj);
        }
    }

    public static bool AnyActiveDefaultCaptureSession()
    {
        IMMDeviceEnumerator en = null;
        IMMDevice devMultimedia = null;
        IMMDevice devComms = null;

        try
        {
            en = (IMMDeviceEnumerator)new MMDeviceEnumeratorCom();

            // Default Multimedia Capture
            int hr = en.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eMultimedia, out devMultimedia);
            bool multiActive = false;
            if (hr == 0 && devMultimedia != null)
                multiActive = DeviceHasActiveSessions(devMultimedia);

            // Default Communications Capture (kann dasselbe Gerät sein)
            hr = en.GetDefaultAudioEndpoint(EDataFlow.eCapture, ERole.eCommunications, out devComms);
            bool commActive = false;
            if (hr == 0 && devComms != null)
                commActive = DeviceHasActiveSessions(devComms);

            return multiActive || commActive;
        }
        finally
        {
            if (devComms != null) Marshal.ReleaseComObject(devComms);
            if (devMultimedia != null) Marshal.ReleaseComObject(devMultimedia);
            if (en != null) Marshal.ReleaseComObject(en);
        }
    }
}
"@

try {
    Add-Type -TypeDefinition $cs -Language CSharp -ErrorAction Stop | Out-Null
} catch {
    Write-Log "C#-Interop konnte nicht kompiliert werden: $($_.Exception.Message)" "ERROR"
    throw
}

function Test-MicActive { [MicProbe]::AnyActiveDefaultCaptureSession() }

function Invoke-WebhookGet {
    param([string]$Url)

    $attempt = 0
    $delay = 300
    while ($attempt -lt $MaxRetries) {
        try {
            $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -Method GET -TimeoutSec $HttpTimeoutSec
            if ($resp.StatusCode -ge 200 -and $resp.StatusCode -lt 300) { return $true }
            Write-Log "Webhook $Url -> HTTP $($resp.StatusCode)" "WARN"
        } catch {
            Write-Log "Webhook $Url Fehler: $($_.Exception.Message)" "WARN"
        }
        $attempt++
        Start-Sleep -Milliseconds $delay
        $delay = [Math]::Min($delay * 2, 4000)
    }
    return $false
}

# --- Debounce State Machine ---
$state            = $false
$lastRaw          = $false
$lastChangeTime   = Get-Date
$debouncedPending = $false

$initial = Test-MicActive
$state   = $initial
$lastRaw = $initial
$lastChangeTime = Get-Date

Write-Log "MicWatcher gestartet. Intervall=${PollIntervalMs}ms, OnDebounce=${OnDebounceMs}ms, OffDebounce=${OffDebounceMs}ms"
Write-Log "ON  -> $WebhookOn"
Write-Log "OFF -> $WebhookOff"

if ($state) { Invoke-WebhookGet -Url $WebhookOn | Out-Null; Write-Log "Startup MIC=ON" "INFO" }
else        { Invoke-WebhookGet -Url $WebhookOff | Out-Null; Write-Log "Startup MIC=OFF" "INFO" }

$script:stopRequested = $false
$null = Register-EngineEvent PowerShell.Exiting -Action { $script:stopRequested = $true }

try {
    while (-not $script:stopRequested) {
        $raw = Test-MicActive

        if ($raw -ne $lastRaw) {
            $lastRaw = $raw
            $lastChangeTime = Get-Date
            $debouncedPending = $true
        }

        if ($debouncedPending) {
            $elapsed = ((Get-Date) - $lastChangeTime).TotalMilliseconds
            if ($raw) { $needed = $OnDebounceMs } else { $needed = $OffDebounceMs }
            if ($elapsed -ge $needed) {
                if ($state -ne $raw) {
                    $state = $raw
                    if ($state) {
                        if (Invoke-WebhookGet -Url $WebhookOn) { Write-Log "MIC=ON" "INFO" }
                        else { Write-Log "MIC=ON Webhook fehlgeschlagen" "ERROR" }
                    } else {
                        if (Invoke-WebhookGet -Url $WebhookOff) { Write-Log "MIC=OFF" "INFO" }
                        else { Write-Log "MIC=OFF Webhook fehlgeschlagen" "ERROR" }
                    }
                }
                $debouncedPending = $false
            }
        }

        Start-Sleep -Milliseconds $PollIntervalMs
    }
} finally {
    Write-Log "MicWatcher beendet." "INFO"
}
