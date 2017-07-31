Imports System.Reflection
Imports System.IO
Imports System.Security.Cryptography

Module Program
    Sub mainLoader()

        Dim resource As String = "MassLogicService.EPPlus.dll"
        EmbeddedAssembly.Load(resource, "EPPlus.dll")

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
