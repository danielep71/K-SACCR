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
using System.Runtime.InteropServices;

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
}
"@
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

function Get-ExcelBusyRetrySummary {
    # One line for the log: calls retried while Excel was busy, and abandoned.
    return ("Excel busy: {0} call(s) retried, {1} abandoned after {2} s" -f
            [KSaccr.ExcelBusyRetry]::Retries, [KSaccr.ExcelBusyRetry]::Abandoned,
            ([KSaccr.ExcelBusyRetry]::TimeoutMs / 1000))
}
