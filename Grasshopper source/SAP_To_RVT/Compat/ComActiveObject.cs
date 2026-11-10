using System;
using System.Runtime.InteropServices;

namespace BIM_Integration
{
    /// <summary>
    /// Attaches to a running COM application (ETABS, SAP2000, SAFE, STAAD) by ProgID on both .NET Framework and .NET.
    /// </summary>
    public static class ComActiveObject
    {
        /// <summary>
        /// Returns the running COM object registered under the given ProgID.
        /// </summary>
        /// <param name="progId">COM ProgID, for example "CSI.ETABS.API.ETABSObject".</param>
        /// <returns>The running COM object.</returns>
        public static object Get(string progId)
        {
#if NETFRAMEWORK
            return Marshal.GetActiveObject(progId);
#else
            // Marshal.GetActiveObject was removed from .NET Core, so the running-object-table lookup is done directly.
            Guid classId = GetClassId(progId);
            GetActiveObject(ref classId, IntPtr.Zero, out object runningObject);
            return runningObject;
#endif
        }

#if !NETFRAMEWORK
        private static Guid GetClassId(string progId)
        {
            try
            {
                CLSIDFromProgIDEx(progId, out Guid classId);
                return classId;
            }
            catch (EntryPointNotFoundException)
            {
                // CLSIDFromProgIDEx is missing on very old Windows versions.
                CLSIDFromProgID(progId, out Guid classId);
                return classId;
            }
        }

        [DllImport("ole32.dll", PreserveSig = false)]
        private static extern void CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid classId);

        [DllImport("ole32.dll", PreserveSig = false)]
        private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid classId);

        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(ref Guid classId, IntPtr reserved, [MarshalAs(UnmanagedType.Interface)] out object runningObject);
#endif
    }
}
