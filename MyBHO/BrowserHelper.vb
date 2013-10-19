Imports System.Runtime.InteropServices
Imports Microsoft.Win32
' '' ''Imports SHDocVw
' '' ''Imports mshtml


''' <summary>
''' Set the GUID of this class and specify that this class is ComVisible.
''' A BHO must implement the interface IObjectWithSite. 
''' </summary>
<ComVisible(True), ClassInterface(ClassInterfaceType.None),
Guid("C42D40F0-BEBF-418D-8EA1-18D99AC2AB17")> Public Class BrowserHelper
    Implements IObjectWithSite

    ' To register a BHO, a new key should be created under this key.
    Private Const RegistryKey As String = "Software\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects"

#Region "Com Register/UnRegister Methods"
    ''' <summary>
    ''' When this class is registered to COM, add a new key to the BHORegistryKey 
    ''' to make IE use this BHO.
    ''' On 64bit machine, if the platform of this assembly and the installer is x86,
    ''' 32 bit IE can use this BHO. If the platform of this assembly and the installer
    ''' is x64, 64 bit IE can use this BHO.
    ''' </summary>
    <ComRegisterFunction()>
    Public Shared Sub RegisterBHO(ByVal t As Type)
        Dim key As RegistryKey = Registry.LocalMachine.OpenSubKey(RegistryKey, True)
        If key Is Nothing Then
            key = Registry.LocalMachine.CreateSubKey(RegistryKey)
        End If

        ' 32 digits separated by hyphens, enclosed in braces: 
        ' {00000000-0000-0000-0000-000000000000}
        Dim bhoKeyStr As String = t.GUID.ToString("B")

        Dim bhoKey As RegistryKey = key.OpenSubKey(bhoKeyStr, True)

        ' Create a new key.
        If bhoKey Is Nothing Then
            bhoKey = key.CreateSubKey(bhoKeyStr)
        End If

        ' NoExplorer:dword = 1 prevents the BHO to be loaded by Explorer
        Dim name As String = "NoExplorer"
        Dim value As Object = CObj(1)
        bhoKey.SetValue(name, value)
        key.Close()
        bhoKey.Close()
    End Sub

    ''' <summary>
    ''' When this class is unregistered from COM, delete the key.
    ''' </summary>
    <ComUnregisterFunction()>
    Public Shared Sub UnregisterBHO(ByVal t As Type)
        Dim key As RegistryKey = Registry.LocalMachine.OpenSubKey(RegistryKey, True)
        Dim guidString As String = t.GUID.ToString("B")
        If key IsNot Nothing Then
            key.DeleteSubKey(guidString, False)
        End If
    End Sub
#End Region

#Region "iObjectWithSite implementation"

    Public Sub GetSite(ByRef riid As Guid, ByRef ppvSite As Object) Implements IObjectWithSite.GetSite

    End Sub

    Public Sub SetSite(pUnkSite As Object) Implements IObjectWithSite.SetSite

    End Sub
#End Region



End Class
