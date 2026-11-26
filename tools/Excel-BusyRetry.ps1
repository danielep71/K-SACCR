# Retry the calls Excel rejects while it is busy (#773).
#
# A COM call to Excel made while Excel is busy (recalculating, drawing charts,
# finishing a macro) is rejected with RPC_E_CALL_REJECTED (0x80010001), and
# the caller is told to try again later. Without a message filter the calling
# script takes that as an error: a quick check stopped its sheet review with
# "Call was rejected by callee" right after Test_RunAll. Register-ExcelBusyRetry
# registers an OLE message filter on the calling thread, which retries such a
# call every RetryDelayMs milliseconds for up to TimeoutMs milliseconds, and
# then lets it fail as before. Only a call rejected as "retry later" is
# retried: a call refused outright still fails at once.
#
# Dot-source this file and call Register-ExcelBusyRetry before creating
# Excel.Application; call Unregister-ExcelBusyRetry when Excel is closed.
# The filter needs a single-threaded apartment, which Windows PowerShell
# gives the script; on another thread Register-ExcelBusyRetry returns $false
# and registers nothing.

if (-not ("KSaccr.ExcelBusyRetry" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace KSaccr
{
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
     Guid("00000016-0000-0000-C000-000000000046")]
    public interface IOleMessageFilter
    {
        [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
        [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
        [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
    }

    public sealed class ExcelBusyRetry : IOleMessageFilter
    {
        public const int SERVERCALL_ISHANDLED = 0;
        public const int SERVERCALL_REJECTED = 1;
        public const int SERVERCALL_RETRYLATER = 2;
        public const int PENDINGMSG_WAITDEFPROCESS = 2;
        public const int CANCEL_CALL = -1;

        public static int TimeoutMs = 120000;
        public static int RetryDelayMs = 250;
        public static int Retries;
        public static int Abandoned;

        [DllImport("ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

        // What to do with a rejected call: the delay before retrying it, in
        // milliseconds, or CANCEL_CALL to let it fail. dwTickCount is the time
        // since the call was first made.
        public static int Decide(int tickCount, int rejectType)
        {
            if (rejectType == SERVERCALL_RETRYLATER && tickCount < TimeoutMs)
            {
                return RetryDelayMs;
            }
            return CANCEL_CALL;
        }

        public static int Register()
        {
            IOleMessageFilter previous;
            return CoRegisterMessageFilter(new ExcelBusyRetry(), out previous);
        }

        public static int Unregister()
        {
            IOleMessageFilter previous;
            return CoRegisterMessageFilter(null, out previous);
        }

        int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
        {
            return SERVERCALL_ISHANDLED;
        }

        int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            int decision = Decide(dwTickCount, dwRejectType);
            if (decision == CANCEL_CALL)
            {
                Abandoned++;
            }
            else
            {
                Retries++;
            }
            return decision;
        }

        int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        {
            return PENDINGMSG_WAITDEFPROCESS;
        }
    }

    // The visible windows of a process, to tell what keeps Excel busy (#773):
    // a dialog shows as class #32770, with its text in its child controls.
    public static class ProcessWindows
    {
        private delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumProc callback, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxCount);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder text, int maxCount);

        public static uint OfWindow(long hWnd)
        {
            uint processId;
            GetWindowThreadProcessId(new IntPtr(hWnd), out processId);
            return processId;
        }

        private static string Text(IntPtr hWnd)
        {
            StringBuilder text = new StringBuilder(512);
            GetWindowText(hWnd, text, text.Capacity);
            return text.ToString().Replace("\r", " ").Replace("\n", " ").Trim();
        }

        private static string ClassOf(IntPtr hWnd)
        {
            StringBuilder name = new StringBuilder(256);
            GetClassName(hWnd, name, name.Capacity);
            return name.ToString();
        }

        // The visible dialogs of a process: classes #32770 (Win32) and NUIDialog
        // (Office), whose text the window functions cannot read.
        public static long[] Dialogs(uint processId)
        {
            List<long> found = new List<long>();
            EnumWindows(delegate (IntPtr hWnd, IntPtr lParam)
            {
                uint owner;
                GetWindowThreadProcessId(hWnd, out owner);
                string name = ClassOf(hWnd);
                if (owner == processId && IsWindowVisible(hWnd) && (name == "#32770" || name == "NUIDialog"))
                {
                    found.Add(hWnd.ToInt64());
                }
                return true;
            }, IntPtr.Zero);
            return found.ToArray();
        }

        public static string Describe(uint processId)
        {
            List<string> found = new List<string>();
            EnumWindows(delegate (IntPtr hWnd, IntPtr lParam)
            {
                uint owner;
                GetWindowThreadProcessId(hWnd, out owner);
                if (owner != processId || !IsWindowVisible(hWnd))
                {
                    return true;
                }
                string line = ClassOf(hWnd) + " '" + Text(hWnd) + "'";
                if (ClassOf(hWnd) == "#32770")
                {
                    List<string> parts = new List<string>();
                    EnumChildWindows(hWnd, delegate (IntPtr child, IntPtr unused)
                    {
                        string part = Text(child);
                        if (part.Length > 0)
                        {
                            parts.Add(part);
                        }
                        return true;
                    }, IntPtr.Zero);
                    line += " [" + string.Join(" | ", parts.ToArray()) + "]";
                }
                found.Add(line);
                return true;
            }, IntPtr.Zero);
            return found.Count == 0 ? "none" : string.Join("; ", found.ToArray());
        }
    }
}
"@
}

# A picture of a window, for a dialog whose text cannot be read (#773). It is
# off unless a caller asks for it with Enable-ExcelDialogPictures: capturing
# the image of another program's window is the kind of behaviour an
# antivirus's behavioural protection flags, and Norton blocked
# Verify-SACCR-Branch.ps1 as IDP.Generic on the owner's runner while this type
# was compiled on every run (2026-11-26). It needs System.Drawing, which
# Windows PowerShell has: elsewhere it is not compiled and
# Save-ExcelDialogPictures saves nothing.
$script:ExcelDialogPicturesRequested = $false

function Enable-ExcelDialogPictures {
    # Compile the picture type; $true when Save-ExcelDialogPictures can save.
    $script:ExcelDialogPicturesRequested = $true
    if ("KSaccr.WindowPicture" -as [type]) {
        return $true
    }
    if ($PSVersionTable.PSEdition -ne "Desktop") {
        return $false
    }
    try {
        Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @"
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace KSaccr
{
    public static class WindowPicture
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
        [DllImport("user32.dll")]
        private static extern bool PrintWindow(IntPtr hWnd, IntPtr hdc, uint flags);

        // PW_RENDERFULLCONTENT (2) draws windows that Office renders itself.
        public static bool Save(long hWnd, string path)
        {
            RECT rect;
            if (!GetWindowRect(new IntPtr(hWnd), out rect))
            {
                return false;
            }
            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;
            if (width <= 0 || height <= 0)
            {
                return false;
            }
            using (Bitmap picture = new Bitmap(width, height))
            {
                using (Graphics canvas = Graphics.FromImage(picture))
                {
                    IntPtr hdc = canvas.GetHdc();
                    try
                    {
                        PrintWindow(new IntPtr(hWnd), hdc, 2);
                    }
                    finally
                    {
                        canvas.ReleaseHdc(hdc);
                    }
                }
                picture.Save(path, ImageFormat.Png);
            }
            return true;
        }
    }
}
"@
        return $true
    }
    catch {
        return $false
    }
}

function Save-ExcelDialogPictures {
    # Save a picture of each visible dialog of the process in Folder, as
    # busy-dialog-N.png; returns one line for the log.
    param([uint32]$ProcessId, [string]$Folder)
    if ($ProcessId -eq 0) {
        return "Excel dialogs: the process is not known"
    }
    if (-not $script:ExcelDialogPicturesRequested) {
        return "Excel dialogs: not pictured (off; Verify-SACCR-Branch.ps1 -DialogPictures turns it on)"
    }
    if (-not ("KSaccr.WindowPicture" -as [type])) {
        return "Excel dialogs: no picture taken here (System.Drawing is not available)"
    }
    try {
        New-Item -ItemType Directory -Path $Folder -Force | Out-Null
        $saved = New-Object System.Collections.Generic.List[string]
        $index = 0
        foreach ($handle in [KSaccr.ProcessWindows]::Dialogs($ProcessId)) {
            $index++
            $file = Join-Path $Folder ("busy-dialog-" + $index + ".png")
            if ([KSaccr.WindowPicture]::Save($handle, $file)) {
                $saved.Add((Split-Path -Leaf $file))
            }
        }
        if ($saved.Count -eq 0) {
            return "Excel dialogs: none pictured"
        }
        return ("Excel dialogs: pictured in " + ($saved -join ", "))
    }
    catch {
        return ("Excel dialogs: not pictured (" + $_.Exception.Message + ")")
    }
}

function Register-ExcelBusyRetry {
    # Register the filter on this thread; $true when it is registered.
    if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
        return $false
    }
    return ([KSaccr.ExcelBusyRetry]::Register() -eq 0)
}

function Unregister-ExcelBusyRetry {
    # Remove the filter from this thread; nothing happens when none is registered.
    if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq [System.Threading.ApartmentState]::STA) {
        [void][KSaccr.ExcelBusyRetry]::Unregister()
    }
}

function Get-ExcelProcessId {
    # The process of an Excel instance, read through its main window while it
    # still answers; 0 when it cannot be read.
    param($Excel)
    try {
        return [KSaccr.ProcessWindows]::OfWindow([long]$Excel.Hwnd)
    }
    catch {
        return 0
    }
}

function Get-ExcelWindowsSummary {
    # One line for the log: the visible windows of the process, a dialog with
    # its text, so that what keeps Excel busy can be named (#773).
    param([uint32]$ProcessId)
    if ($ProcessId -eq 0) {
        return "Excel windows: the process is not known"
    }
    try {
        return ("Excel windows: " + [KSaccr.ProcessWindows]::Describe($ProcessId))
    }
    catch {
        return ("Excel windows: not read (" + $_.Exception.Message + ")")
    }
}

function Get-ExcelBusyRetrySummary {
    # One line for the log: calls retried while Excel was busy, and abandoned.
    return ("Excel busy: {0} call(s) retried, {1} abandoned after {2} s" -f
            [KSaccr.ExcelBusyRetry]::Retries, [KSaccr.ExcelBusyRetry]::Abandoned,
            ([KSaccr.ExcelBusyRetry]::TimeoutMs / 1000))
}
