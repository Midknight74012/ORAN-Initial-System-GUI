using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ORAN_Initial_System_GUI
{
    public class MapT_Drive
    {
        Form1 form;
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct NETRESOURCE
        {
            public int dwScope;
            public int dwType;
            public int dwDisplayType;
            public int dwUsage;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpLocalName;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpRemoteName;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpComment;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpProvider;
        }
        // Insert/replace this method inside the Form1 class
        [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
        private static extern int WNetAddConnection2(ref NETRESOURCE lpNetResource, string lpPassword, string lpUsername, int dwFlags);

        [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
        private static extern int WNetCancelConnection2(string lpName, int dwFlags, bool fForce);

        // Common values used with WNetAddConnection2 / WNetCancelConnection2
        private const int RESOURCETYPE_DISK = 1;
        private const int CONNECT_UPDATE_PROFILE = 0x00000001;
        private const int CONNECT_INTERACTIVE = 0x00000008; // optional
        private const int CONNECT_PROMPT = 0x00000010;      // optional
        public bool EnsureTDriveMapped(string remoteShare = @"\\REDACTED_IP\Shared_Folder", string localDrive = "T:", string username = null, string password = null)
        {
            // Quick check: if T: already exists, do nothing.
            try
            {
                foreach (var d in System.IO.DriveInfo.GetDrives())
                {
                    if (string.Equals(d.Name.TrimEnd('\\'), localDrive, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            catch
            {
                // ignore drive enumeration errors and attempt mapping below
            }

            var netResource = new NETRESOURCE
            {
                dwType = RESOURCETYPE_DISK,
                lpLocalName = localDrive,
                lpRemoteName = remoteShare
            };

            int result = WNetAddConnection2(ref netResource, password, username, CONNECT_UPDATE_PROFILE);

            if (result == 0)
                return true;

            // Non-zero result => mapping failed. Short diagnostic on UI thread.
            MessageBox.Show($"Failed to map {localDrive} to {remoteShare} (error {result}).\nIf the share requires credentials, call EnsureTDriveMapped with username/password.",
                        "Map Drive Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }
    }
}