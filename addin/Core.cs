using System;
using System.Runtime.InteropServices;

namespace WordAIAssistant
{
    // --- COM-интерфейсы офисной расширяемости (канонические сигнатуры) ---

    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"),
     InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDTExtensibility2
    {
        void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object Application, int ConnectMode,
                          [In, MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
                          [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        void OnDisconnection(int RemoveMode,
                          [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    [ComImport, Guid("000C0396-0000-0000-C000-000000000046"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IRibbonExtensibility
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string RibbonID);
    }

    public static class Constants
    {
        public const string ProgId = "WordAIAssistant.Connect";
        public const string BridgeBase = "http://127.0.0.1:8765";
        public const string AppDataDirName = "WordAIAssistant";
    }
}
