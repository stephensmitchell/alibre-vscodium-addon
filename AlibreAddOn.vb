Imports System.IO
Imports System.Reflection
Imports System.Diagnostics
Imports System.Threading
Imports AlibreAddOn
Imports AlibreX
Imports IronPython.Hosting
Imports Microsoft.Scripting.Hosting
Namespace AlibreAddOnAssembly
    Public Module AlibreAddOn
        Private Property AlibreRoot As IADRoot
        Private _parentWinHandle As IntPtr
        Private _addOnHandle As AddOnRibbon
        Public Sub AddOnLoad(ByVal hwnd As IntPtr, ByVal pAutomationHook As IAutomationHook, ByVal unused As IntPtr)
            AlibreRoot = CType(pAutomationHook.Root, IADRoot)
            _parentWinHandle = hwnd
            _addOnHandle = New AddOnRibbon(AlibreRoot, _parentWinHandle)
        End Sub
        Public Sub AddOnUnload(ByVal hwnd As IntPtr, ByVal forceUnload As Boolean, ByRef cancel As Boolean, ByVal reserved1 As Integer, ByVal reserved2 As Integer)
            Try
                If _addOnHandle IsNot Nothing Then
                End If
            Catch ex As Exception
                Debug.WriteLine($"Error cleaning up: {ex.Message}")
            End Try
            _addOnHandle = Nothing
            AlibreRoot = Nothing
        End Sub
        Public Function GetRoot() As IADRoot
            Return AlibreRoot
        End Function
        Public Sub AddOnInvoke(ByVal hwnd As IntPtr, ByVal pAutomationHook As IntPtr, ByVal sessionName As String, ByVal isLicensed As Boolean, ByVal reserved1 As Integer, ByVal reserved2 As Integer)
        End Sub
        Public Function GetAddOnInterface() As IAlibreAddOn
            Return CType(_addOnHandle, IAlibreAddOn)
        End Function
        Public Sub ExecutePythonScript(scriptPath As String)
            Try
                Dim automationHook As Object = System.Runtime.InteropServices.Marshal.GetActiveObject("AlibreX.AutomationHook")
                Dim root As IADRoot = CType(automationHook.Root, IADRoot)
                Dim session As IADSession = Nothing
                For Each s As IADSession In root.Sessions
                    session = s
                    Exit For
                Next
                If session Is Nothing Then
                    System.Windows.MessageBox.Show("No active Alibre session found. Please open a part or assembly in Alibre Design.", "No Session")
                    Return
                End If
                If _addOnHandle IsNot Nothing Then
                    AddOnRibbon.CurrentScriptToExecute = scriptPath
                    _addOnHandle.InvokeCommand(1001, session.Identifier)
                Else
                    System.Windows.MessageBox.Show("Add-on not initialized. Please ensure the alibre-vscodium-addon is loaded.", "Add-on Error")
                End If
            Catch ex As Exception
                System.Windows.MessageBox.Show("Error executing script from external call:" & vbLf & ex.ToString(), "Script Execution Error")
            End Try
        End Sub
    End Module
    Public Class AddOnRibbon
        Implements IAlibreAddOn
        Private Const ROOT_ID As Integer = 901
        Private Const CMD As Integer = 1001
        Private Const CMD_OPEN_EDITOR As Integer = 1002
        Private Const CMD_EXECUTE_EXTERNAL As Integer = 1003
        Private ReadOnly _AlibreRoot As IADRoot
        Private ReadOnly _parentWinHandle As IntPtr
        Public Shared CurrentScriptToExecute As String = ""
        Public Sub New(alibreRoot As IADRoot, parentWinHandle As IntPtr)
            Try
                Debug.WriteLine("AddOnRibbon constructor starting...")
                _AlibreRoot = alibreRoot
                _parentWinHandle = parentWinHandle
                Debug.WriteLine("AddOnRibbon constructor completed successfully")
            Catch ex As Exception
                Debug.WriteLine("CRITICAL ERROR in AddOnRibbon constructor: " & ex.Message)
                Debug.WriteLine("Constructor exception details: " & ex.ToString())
            End Try
        End Sub
        Public ReadOnly Property RootMenuItem As Integer Implements IAlibreAddOn.RootMenuItem
            Get
                Return ROOT_ID
            End Get
        End Property
        <STAThread>
        Public Function InvokeCommand(menuId As Integer, sessionIdentifier As String) As IAlibreAddOnCommand Implements IAlibreAddOn.InvokeCommand
            Try
                Dim session As IADSession = Nothing
                If _AlibreRoot IsNot Nothing Then
                    Try
                        If Not String.IsNullOrEmpty(sessionIdentifier) Then
                            For Each s As IADSession In _AlibreRoot.Sessions
                                If String.Equals(s.Identifier, sessionIdentifier, StringComparison.OrdinalIgnoreCase) Then
                                    session = s
                                    Exit For
                                End If
                            Next
                        End If
                        If session Is Nothing Then
                            For Each s As IADSession In _AlibreRoot.Sessions
                                session = s
                                Exit For
                            Next
                        End If
                    Catch
                    End Try
                End If
                Dim runner As New ScriptRunner1(_AlibreRoot)
                Dim addOnDirectory As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                Dim scriptsPath = Path.Combine(addOnDirectory, "scripts")
                Select Case menuId
                    Case CMD
                        runner.ExecuteScript(session, "Template.py")
                    Case CMD_EXECUTE_EXTERNAL
                        If Not String.IsNullOrEmpty(CurrentScriptToExecute) Then
                            Try
                                Debug.WriteLine($"Executing external script: {CurrentScriptToExecute}")
                                runner.ExecuteScriptDirect(session, CurrentScriptToExecute)
                                Debug.WriteLine($"External script execution completed: {CurrentScriptToExecute}")
                            Catch ex As Exception
                                Debug.WriteLine($"External script execution failed: {ex.Message}")
                                System.Windows.MessageBox.Show($"Error executing external script: {ex.Message}", "Script Execution Error")
                                Throw 
                            Finally
                                CurrentScriptToExecute = ""
                            End Try
                        End If
                    Case CMD_OPEN_EDITOR
                        OpenVSCodiumEditor()
                End Select
            Catch ex As Exception
                System.Windows.MessageBox.Show("An error occurred while invoking the command:" & vbLf & ex.ToString(), "alibre-vscodium-addon")
            End Try
            Return Nothing
        End Function
        Private Sub OpenVSCodiumEditor()
            Try
                Dim addOnDirectory As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                Dim projectRoot As String = Path.GetDirectoryName(addOnDirectory) 
                Dim editorPath As String = Path.Combine(projectRoot, "editor", "VSCodium.exe")
                Dim templateScriptPath As String = Path.Combine(addOnDirectory, "scripts", "Template.py")
                Dim workspaceConfigPath As String = Path.Combine(addOnDirectory, "scripts", ".vscode")
                Debug.WriteLine(editorPath)
                If Not File.Exists(editorPath) Then
                    System.Windows.MessageBox.Show("VS Codium editor not found at: " & editorPath, "Editor Not Found")
                    Return
                End If
                If Not File.Exists(templateScriptPath) Then
                    System.Windows.MessageBox.Show("Template.py script not found at: " & templateScriptPath, "Template Not Found")
                    Return
                End If
                CreateVSCodeWorkspaceConfig(workspaceConfigPath, addOnDirectory)
                Dim processInfo As New ProcessStartInfo()
                processInfo.FileName = editorPath
                processInfo.Arguments = """" & Path.GetDirectoryName(templateScriptPath) & """ """ & templateScriptPath & """"
                processInfo.UseShellExecute = True
                processInfo.WorkingDirectory = Path.GetDirectoryName(templateScriptPath)
                Process.Start(processInfo)
            Catch ex As Exception
                System.Windows.MessageBox.Show("An error occurred while opening the editor:" & vbLf & ex.ToString(), "Editor Error")
            End Try
        End Sub
        Private Sub CreateVSCodeWorkspaceConfig(workspaceConfigPath As String, addOnDirectory As String)
            Try
                If Not Directory.Exists(workspaceConfigPath) Then
                    Directory.CreateDirectory(workspaceConfigPath)
                End If
                Dim tasksJsonPath As String = Path.Combine(workspaceConfigPath, "tasks.json")
                Dim projectRoot As String = Path.GetDirectoryName(addOnDirectory) 
                Dim cliToolPath As String = Path.Combine(projectRoot, "cli", "alibre-vscodium-cli.exe")
                Dim escapedPath As String = cliToolPath.Replace("\", "\\")
                Dim tasksJson As String = "{
    ""version"": ""2.0.0"",
    ""tasks"": [
        {
            ""label"": ""Execute in Alibre"",
            ""type"": ""shell"",
            ""command"": """ & escapedPath & """,
            ""args"": [
                ""execute"",
                ""${file}""
            ],
            ""group"": {
                ""kind"": ""build"",
                ""isDefault"": true
            },
            ""presentation"": {
                ""echo"": true,
                ""reveal"": ""always"",
                ""focus"": false,
                ""panel"": ""shared"",
                ""showReuseMessage"": true,
                ""clear"": false
            },
            ""problemMatcher"": []
        },
        {
            ""label"": ""Check Alibre Status"",
            ""type"": ""shell"",
            ""command"": """ & escapedPath & """,
            ""args"": [
                ""status""
            ],
            ""group"": ""test"",
            ""presentation"": {
                ""echo"": true,
                ""reveal"": ""always"",
                ""focus"": false,
                ""panel"": ""shared"",
                ""showReuseMessage"": true,
                ""clear"": false
            },
            ""problemMatcher"": []
        }
    ]
}"
                File.WriteAllText(tasksJsonPath, tasksJson)
                Dim keybindingsJsonPath As String = Path.Combine(workspaceConfigPath, "keybindings.json")
                Dim keybindingsJson As String = "[
    {
        ""key"": ""f5"",
        ""command"": ""workbench.action.tasks.runTask"",
        ""args"": ""Execute in Alibre"",
        ""when"": ""editorTextFocus && resourceExtname == ""
    }
]"
                File.WriteAllText(keybindingsJsonPath, keybindingsJson)
            Catch ex As Exception
            End Try
        End Sub
        Public Class ScriptRunner1
            Private ReadOnly _engine As ScriptEngine
            Private ReadOnly _alibreRoot As IADRoot
            Public Sub New(alibreRoot As IADRoot)
                _alibreRoot = alibreRoot
                _engine = Python.CreateEngine()
                Dim alibreInstallPath As String = System.Reflection.Assembly.GetAssembly(GetType(IADRoot)).Location.Replace("\Program\AlibreX.dll", "")
                Dim searchPaths = _engine.GetSearchPaths()
                searchPaths.Add(Path.Combine(alibreInstallPath, "Program"))
                searchPaths.Add(Path.Combine(alibreInstallPath, "Program", "Addons", "AlibreScript", "PythonLib"))
                searchPaths.Add(Path.Combine(alibreInstallPath, "Program", "Addons", "AlibreScript"))
                _engine.SetSearchPaths(searchPaths)
            End Sub
            Public Sub ExecuteScript(session As IADSession, mainScriptFileName As String)
                Try
                    Dim addOnDirectory As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
                    Dim ScriptsPath As String = Path.Combine(addOnDirectory, "scripts")
                    Dim setupScriptPath As String = Path.Combine(ScriptsPath, "alibre_setup.py")
                    Dim mainScriptPath As String = Path.Combine(ScriptsPath, mainScriptFileName)
                    If (Not File.Exists(setupScriptPath)) OrElse (Not File.Exists(mainScriptPath)) Then
                        System.Windows.MessageBox.Show("Error: Script not found." & vbLf & "Setup: " & setupScriptPath & vbLf & "Main: " & mainScriptPath, "Script Error")
                        Return
                    End If
                    Dim scope As ScriptScope = _engine.CreateScope()
                    scope.SetVariable("ScriptFileName", mainScriptFileName)
                    scope.SetVariable("ScriptFolder", ScriptsPath)
                    scope.SetVariable("SessionIdentifier", session.Identifier)
                    scope.SetVariable("Arguments", New List(Of String)())
                    scope.SetVariable("AlibreRoot", _alibreRoot)
                    scope.SetVariable("CurrentSession", session)
                    _engine.ExecuteFile(setupScriptPath, scope)
                    _engine.ExecuteFile(mainScriptPath, scope)
                Catch ex As Exception
                    System.Windows.MessageBox.Show("An error occurred while running the script:" & vbLf & ex.ToString(), "Python Execution Error")
                End Try
            End Sub
            Public Sub ExecuteScriptDirect(session As IADSession, scriptPath As String)
                Try
                    Debug.WriteLine($"ExecuteScriptDirect called with path: {scriptPath}")
                    If Not File.Exists(scriptPath) Then
                        Debug.WriteLine($"Script file not found: {scriptPath}")
                        System.Windows.MessageBox.Show("Error: Script not found." & vbLf & "Path: " & scriptPath, "Script Error")
                        Return
                    End If
                    Dim addOnDirectory As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
                    Dim ScriptsPath As String = Path.Combine(addOnDirectory, "scripts")
                    Dim setupScriptPath As String = Path.Combine(ScriptsPath, "alibre_setup.py")
                    Debug.WriteLine($"Setup script path: {setupScriptPath}")
                    If Not File.Exists(setupScriptPath) Then
                        Debug.WriteLine($"Setup script not found: {setupScriptPath}")
                        System.Windows.MessageBox.Show("Error: Setup script not found." & vbLf & "Setup: " & setupScriptPath, "Script Error")
                        Return
                    End If
                    Debug.WriteLine($"Creating IronPython scope and executing scripts...")
                    Dim scope As ScriptScope = _engine.CreateScope()
                    scope.SetVariable("ScriptFileName", Path.GetFileName(scriptPath))
                    scope.SetVariable("ScriptFolder", ScriptsPath)  
                    scope.SetVariable("SessionIdentifier", session.Identifier)
                    scope.SetVariable("Arguments", New List(Of String)())
                    scope.SetVariable("AlibreRoot", _alibreRoot)
                    scope.SetVariable("CurrentSession", session)
                    Debug.WriteLine($"Executing setup script: {setupScriptPath}")
                    _engine.ExecuteFile(setupScriptPath, scope)
                    Debug.WriteLine($"Executing main script: {scriptPath}")
                    _engine.ExecuteFile(scriptPath, scope)
                    Debug.WriteLine($"Script execution completed successfully")
                Catch ex As Exception
                    Debug.WriteLine($"Script execution error: {ex.Message}")
                    System.Windows.MessageBox.Show("An error occurred while running the script:" & vbLf & ex.ToString(), "Python Execution Error")
                End Try
            End Sub
        End Class
        Public Function HasSubMenus(menuId As Integer) As Boolean Implements IAlibreAddOn.HasSubMenus
            If menuId = ROOT_ID Then Return True
            Return False
        End Function
        Public Function SubMenuItems(menuId As Integer) As Array Implements IAlibreAddOn.SubMenuItems
            If menuId = ROOT_ID Then
                Return New Integer() {CMD, CMD_OPEN_EDITOR}
            End If
            Return Nothing
        End Function
        Public Function MenuItemText(menuId As Integer) As String Implements IAlibreAddOn.MenuItemText
            Select Case menuId
                Case ROOT_ID : Return "VSCodium Editor"
                Case CMD : Return "Run Template Script"
                Case CMD_OPEN_EDITOR : Return "Open Editor"
                Case Else : Return String.Empty
            End Select
        End Function
        Public Function MenuItemState(menuId As Integer, sessionIdentifier As String) As ADDONMenuStates Implements IAlibreAddOn.MenuItemState
            Return ADDONMenuStates.ADDON_MENU_ENABLED
        End Function
        Public Function MenuItemToolTip(menuId As Integer) As String Implements IAlibreAddOn.MenuItemToolTip
            Select Case menuId
                Case ROOT_ID : Return "VSCodium Editor Tools"
                Case CMD : Return "Execute the Template.py script"
                Case CMD_OPEN_EDITOR : Return "Open VSCodium editor with Template.py preloaded"
                Case Else : Return String.Empty
            End Select
        End Function
        Public Function MenuIcon(menuID As Integer) As String Implements IAlibreAddOn.MenuIcon
            Select Case menuID
                Case ROOT_ID : Return "logo.ico"
                Case CMD : Return "logo.ico"
                Case CMD_OPEN_EDITOR : Return "logo.ico"
                Case Else : Return ""
            End Select
        End Function
        Public Function PopupMenu(menuId As Integer) As Boolean Implements IAlibreAddOn.PopupMenu
            Return False
        End Function
        Public Function HasPersistentDataToSave(sessionIdentifier As String) As Boolean Implements IAlibreAddOn.HasPersistentDataToSave
            Return False
        End Function
        Public Sub SaveData(pCustomData As IStream, sessionIdentifier As String) Implements IAlibreAddOn.SaveData
        End Sub
        Public Sub LoadData(pCustomData As IStream, sessionIdentifier As String) Implements IAlibreAddOn.LoadData
        End Sub
        Public Function UseDedicatedRibbonTab() As Boolean Implements IAlibreAddOn.UseDedicatedRibbonTab
            Return False
        End Function
        Private Sub IAlibreAddOn_setIsAddOnLicensed(isLicensed As Boolean) Implements IAlibreAddOn.setIsAddOnLicensed
        End Sub
    End Class
End Namespace
