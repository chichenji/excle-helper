using System;
using System.Runtime.InteropServices;

namespace ExcleHelper
{
    /// <summary>
    /// Publishes the live add-in instance to the Running Object Table as
    /// "ExcleHelper.Automation", so scripts / VBA can reach the same connected object:
    ///   Set o = GetObject("ExcleHelper.Automation")   ' VBA
    ///   $o = [Runtime.InteropServices.Marshal]::GetActiveObject('GridKit.Automation')
    /// Excel's COMAddIn.Object is null for non-VSTO add-ins, so this is the
    /// scriptable surface instead.
    /// </summary>
    internal static class Rot
    {
        [ComImport]
        [Guid("00000010-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IRunningObjectTable
        {
            int Register(int grfFlags,
                         [MarshalAs(UnmanagedType.Interface)] object punkObject,
                         [MarshalAs(UnmanagedType.Interface)] IMoniker pmkObjectName);
            void Revoke(int dwRegister);
        }

        [ComImport]
        [Guid("0000000F-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMoniker
        {
        }

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable pprot);

        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        private static extern int CreateItemMoniker(
            [MarshalAs(UnmanagedType.LPWStr)] string lpszDelim,
            [MarshalAs(UnmanagedType.LPWStr)] string lpszItem,
            out IMoniker ppmk);

        [DllImport("ole32.dll")]
        private static extern int CreateClassMoniker(ref Guid rclsid, out IMoniker ppmk);

        [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
        private static extern int MkParseDisplayName(
            IntPtr pbc,
            [MarshalAs(UnmanagedType.LPWStr)] string szUserName,
            out int pchEaten,
            out IMoniker ppmk);

        internal const string Name = "ExcleHelper.Automation";
        internal const string ProgId = "ExcleHelper.Connect";
        internal static readonly Guid Clsid = new Guid("5E1C7A93-2B48-4D16-9F77-8C3A5B1E6D24");
        private static int _cookieClass;
        private static int _cookieItem;

        internal static bool Publish(object instance)
        {
            IRunningObjectTable rot = null;
            try
            {
                if (GetRunningObjectTable(0, out rot) != 0 || rot == null) return false;

                // GetActiveObject(progid) resolves the CLSID and then looks up the ROT by
                // a CLASS moniker, so that is what has to be registered.
                Guid clsid = Clsid;
                IMoniker mk;
                int hr = CreateClassMoniker(ref clsid, out mk);
                Connect.Trace("Rot: CreateClassMoniker hr=0x" + hr.ToString("X8"));
                if (hr == 0 && mk != null)
                {
                    _cookieClass = rot.Register(0, instance, mk);
                    Marshal.ReleaseComObject(mk);
                    Connect.Trace("Rot: class moniker cookie=" + _cookieClass);
                }

                IMoniker item;
                if (CreateItemMoniker("!", Name, out item) == 0 && item != null)
                {
                    _cookieItem = rot.Register(0, instance, item);
                    Marshal.ReleaseComObject(item);
                    Connect.Trace("Rot: item moniker cookie=" + _cookieItem);
                }
                return _cookieClass != 0 || _cookieItem != 0;
            }
            catch (Exception ex)
            {
                Connect.Trace("Rot: publish exception " + ex.Message);
                return false;
            }
            finally
            {
                if (rot != null) { try { Marshal.ReleaseComObject(rot); } catch { } }
            }
        }

        internal static void Unpublish()
        {
            if (_cookieClass == 0 && _cookieItem == 0) return;
            IRunningObjectTable rot = null;
            try
            {
                if (GetRunningObjectTable(0, out rot) == 0 && rot != null)
                {
                    if (_cookieClass != 0) rot.Revoke(_cookieClass);
                    if (_cookieItem != 0) rot.Revoke(_cookieItem);
                }
            }
            catch { }
            finally
            {
                if (rot != null) { try { Marshal.ReleaseComObject(rot); } catch { } }
            }
            _cookieClass = 0;
            _cookieItem = 0;
        }
    }
}
