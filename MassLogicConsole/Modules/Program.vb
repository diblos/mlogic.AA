Imports System.Reflection
Imports System.IO
Imports System.Security.Cryptography

Module Program
    Sub mainLoader()


        'Dim resource1 As String = "MassLogicConsole.BlackBerryWorkspacesSDK.dll"
        Dim resource2 As String = "MassLogicConsole.EPPlus.dll"
        'EmbeddedAssembly.Load(resource1, "BlackBerryWorkspacesSDK.dll")
        EmbeddedAssembly.Load(resource2, "EPPlus.dll")

        AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf CurrentDomain_AssemblyResolve

        '============================================================================================

        'Application.EnableVisualStyles()
        'Application.SetCompatibleTextRenderingDefault(False)
        'Application.Run(New frmTest())

        '============================================================================================
        'AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf CurrentDomain_AssemblyResolve
        'Application.EnableVisualStyles()
        'Application.SetCompatibleTextRenderingDefault(False)
        'Application.Run(New Form1())
        '============================================================================================
    End Sub

    Private Function CurrentDomain_AssemblyResolve(ByVal sender As Object, ByVal args As ResolveEventArgs) As Assembly
        Return EmbeddedAssembly.[Get](args.Name)
    End Function

End Module
