using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Esquire.Common.Util
{
    public static class WindowHelper
    {
        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(String lpClassName, String lpWindowName);

        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        public static Boolean BringToFront(string className)
        {
            // Get a handle to the Calculator application.
            IntPtr handle = FindWindow(className, null);

            // Verify that Calculator is a running process.
            if (handle == IntPtr.Zero)
            {
                return false;
            }

            // Make Calculator the foreground application
            SetForegroundWindow(handle);
            return true;
        }
    }
}
