﻿Imports System.ServiceProcess

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class VRMService
    Inherits System.ServiceProcess.ServiceBase

    'UserService overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    ' The main entry point for the process
    <MTAThread()>
    <System.Diagnostics.DebuggerNonUserCode()>
    Shared Sub Main()
        '#If DEBUG Then
        '        Dim myService As New VRMService()
        '        myService.OnDebug()
        '        System.Threading.Thread.Sleep(System.Threading.Timeout.Infinite)
        '#Else
        '                    'Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        '                    '    More than one NT Service may run within the same process. To add
        '                    '    another service to this process, change the following line to
        '                    '    create a second service object. For example,

        '                    '    ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}

        '                    'ServicesToRun = New System.ServiceProcess.ServiceBase() {New Service1}

        '                    'System.ServiceProcess.ServiceBase.Run(ServicesToRun)
        '#End If


        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ' More than one NT Service may run within the same process. To add
        ' another service to this process, change the following line to
        ' create a second service object. For example,
        '
        '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
        '
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New VRMService}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)

    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    ' NOTE: The following procedure is required by the Component Designer
    ' It can be modified using the Component Designer.  
    ' Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Timer1 = New System.Timers.Timer()
        CType(Me.Timer1, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'Timer1
        '
        Me.Timer1.Enabled = False
        Me.Timer1.Interval = 5000.0R
        '
        'Service1
        '
        Me.ServiceName = "VRMService"
        CType(Me.Timer1, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub

    Friend WithEvents Timer1 As Timers.Timer
End Class
