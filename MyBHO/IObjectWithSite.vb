Imports System.Runtime.InteropServices

<ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
 Guid("FC4801A3-2BA9-11CF-A229-00AA003D7352")>
Public Interface IObjectWithSite
    Sub SetSite(<[In](), MarshalAs(UnmanagedType.IUnknown)> ByVal pUnkSite As Object)

    Sub GetSite(ByRef riid As Guid,
        <Out(), MarshalAs(UnmanagedType.IUnknown)> ByRef ppvSite As Object)
End Interface
