using System;
using System.Runtime.InteropServices;

namespace ExcleHelper
{
    // ---- Office / VSTO-free interop definitions -------------------------
    // These are declared locally so the add-in has ZERO external assembly
    // dependencies (no VSTO runtime, no PIA, no interop DLLs).

    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
        ext_cm_Solution = 4,
        ext_cm_UISetup = 5
    }

    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
        ext_dm_UISetupComplete = 2,
        ext_dm_SolutionClosed = 3
    }

    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        // extensibility.tlb declares this as a DUAL interface (interface X : IDispatch),
        // so Excel calls the methods through vtable slots 7..11 and passes a real
        // SAFEARRAY(VARIANT)* for `custom`. Declaring it InterfaceIsIDispatch makes the
        // CLR expose only dispinterface forwarding and the call dies inside the CLR
        // with 0x80131531 before any managed code runs.
        [DispId(1)]
        void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object Application,
            [MarshalAs(UnmanagedType.I4)] ext_ConnectMode ConnectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(2)]
        void OnDisconnection(
            [MarshalAs(UnmanagedType.I4)] ext_DisconnectMode RemoveMode,
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(4)]
        void OnStartupComplete(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    // Office.IRibbonExtensibility — carried as a dispinterface: Office invokes
    // GetCustomUI through IDispatch (DispId 1), so InterfaceIsIDispatch is the
    // shape that actually reaches managed code.
    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        string GetCustomUI(string RibbonID);
    }

    // Office.IRibbonControl
    [ComImport]
    [Guid("000C0395-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonControl
    {
        [DispId(1)] string Id { get; }
        [DispId(2)] string Tag { get; }
        [DispId(3)] object Context { get; }
    }

    // Office.IRibbonUI (for Invalidate / InvalidateControl)
    [ComImport]
    [Guid("000C0397-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonUI
    {
        [DispId(1)] void Invalidate();
        [DispId(2)] void InvalidateControl(string ControlID);
    }
}
