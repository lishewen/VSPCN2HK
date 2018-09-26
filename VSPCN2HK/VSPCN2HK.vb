Imports System
Imports System.ComponentModel.Design
Imports System.Globalization
Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.VisualBasic
Imports Microsoft.VisualStudio.Editor
Imports Microsoft.VisualStudio.Shell
Imports Microsoft.VisualStudio.Shell.Interop
Imports Microsoft.VisualStudio.Text
Imports Microsoft.VisualStudio.Text.Editor
Imports Microsoft.VisualStudio.TextManager.Interop
Imports Task = System.Threading.Tasks.Task

''' <summary>
''' Command handler
''' </summary>
Public NotInheritable Class VSPCN2HK

    ''' <summary>
    ''' Command ID.
    ''' </summary>
    Public Const CommandId As Integer = 256
    Public Const CommandId2 As Integer = 257

    ''' <summary>
    ''' Command menu group (command set GUID).
    ''' </summary>
    Public Shared ReadOnly CommandSet As New Guid("9e5bd1c0-d70c-4966-a8f8-34efbe984913")

    ''' <summary>
    ''' VS Package that provides this command, not null.
    ''' </summary>
    Private ReadOnly package As AsyncPackage

    ''' <summary>
    ''' Initializes a new instance of the <see cref="VSPCN2HK"/> class.
    ''' Adds our command handlers for menu (the commands must exist in the command table file)
    ''' </summary>
    ''' <param name="package">Owner package, not null.</param>
    Private Sub New(package As AsyncPackage, commandService As OleMenuCommandService)
        If package Is Nothing Then
            Throw New ArgumentNullException("package")
        End If

        If commandService Is Nothing Then
            Throw New ArgumentNullException(NameOf(commandService))
        End If

        Me.package = package

        Dim menuCommandId = New CommandID(CommandSet, CommandId)
        Dim menuCommand = New MenuCommand(AddressOf MenuItemCallback, menuCommandId)
        commandService.AddCommand(menuCommand)
        Dim menuCommandId2 = New CommandID(CommandSet, CommandId2)
        Dim menuCommand2 = New MenuCommand(AddressOf MenuItemToCNCallback, menuCommandId2)
        commandService.AddCommand(menuCommand2)
    End Sub

    ''' <summary>
    ''' Gets the instance of the command.
    ''' </summary>
    Public Shared Property Instance As VSPCN2HK

    ''' <summary>
    ''' Get service provider from the owner package.
    ''' </summary>
    Private ReadOnly Property ServiceProvider As Microsoft.VisualStudio.Shell.IAsyncServiceProvider
        Get
            Return package
        End Get
    End Property

    ''' <summary>
    ''' Initializes the singleton instance of the command.
    ''' </summary>
    ''' <param name="package">Owner package, Not null.</param>
    Public Shared Async Function InitializeAsync(package As AsyncPackage) As Task
        ' Switch to the main thread - the call to AddCommand in VSPCN2HK's constructor requires
        ' the UI thread.
        Await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken)

        Dim commandService As OleMenuCommandService = Await package.GetServiceAsync(GetType(IMenuCommandService))
        Instance = New VSPCN2HK(package, commandService)
    End Function

    Private Sub MenuItemCallback(sender As Object, e As EventArgs)
        ConvAsync(VbStrConv.TraditionalChinese)
    End Sub

    Private Sub MenuItemToCNCallback(ByVal sender As Object, ByVal e As EventArgs)
        ConvAsync(VbStrConv.SimplifiedChinese)
    End Sub

    Private Async Sub ConvAsync(sc As VbStrConv)
        ThreadHelper.ThrowIfNotOnUIThread()

        Dim service As IVsTextManager = Await package.GetServiceAsync(GetType(SVsTextManager))
        Dim ppView As IVsTextView = Nothing
        Dim fMustHaveFocus As Integer = 1
        service.GetActiveView(fMustHaveFocus, Nothing, ppView)
        Dim data As IVsUserData = TryCast(ppView, IVsUserData)
        If (data Is Nothing) Then
            MsgBox("No text view is currently open", MsgBoxStyle.ApplicationModal, Nothing)
        Else
            Dim pvtData As Object = Nothing
            Dim guidIWpfTextViewHost As Guid = DefGuidList.guidIWpfTextViewHost
            data.GetData(guidIWpfTextViewHost, pvtData)
            Dim host As IWpfTextViewHost = DirectCast(pvtData, IWpfTextViewHost)
            If host.TextView.Selection.IsEmpty Then
                Dim span As New Span(0, host.TextView.TextBuffer.CurrentSnapshot.Length)
                host.TextView.TextBuffer.Replace(span, StrConv(host.TextView.TextBuffer.CurrentSnapshot.GetText, sc, 0))
            Else
                host.TextView.TextBuffer.Replace(host.TextView.Selection.SelectedSpans.Item(0), StrConv(host.TextView.Selection.SelectedSpans.Item(0).GetText, sc, 0))
            End If
        End If
    End Sub
End Class
