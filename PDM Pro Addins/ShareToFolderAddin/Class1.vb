Imports EPDM.Interop.epdm
Imports System.Runtime.InteropServices

<Guid("E85D48B5-18D8-4B6D-A313-B0B12813605F")>
<ComVisible(True)>
Public Class SharedToFolderTaskAddin
    Implements IEdmAddIn5

    Public Sub GetAddInInfo(
    ByRef poInfo As EdmAddInInfo,
    ByVal poVault As IEdmVault5,
    ByVal poCmdMgr As IEdmCmdMgr5) _
    Implements IEdmAddIn5.GetAddInInfo

        Try
            poInfo.mbsAddInName =
              "Share File To Folder Task Addin"
            poInfo.mbsCompany = "Safety Rail Company"
            poInfo.mbsDescription =
              "Task Addin that shares the files to a specified Folder"
            poInfo.mlAddInVersion = 2

            'Minimum SOLIDWORKS PDM Professional version
            'needed for VB.Net Task Add-Ins is 10.0
            poInfo.mlRequiredVersionMajor = 10
            poInfo.mlRequiredVersionMinor = 0

            'Register this add-in as a task add-in
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskRun)
            'Register this add-in to be called when
            'selected as a task in the Administration tool
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskSetup)
            'Register this add-in to be called when the set-up wizard is closed
            poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskSetupButton)
        Catch ex As Runtime.InteropServices.COMException
            MsgBox("HRESULT = 0x" +
              ex.ErrorCode.ToString("X") + vbCrLf +
              ex.Message)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub OnCmd(ByRef poCmd As EdmCmd, ByRef ppoData As EdmCmdData()) Implements IEdmAddIn5.OnCmd

        Try
            'PauseToAttachProcess(poCmd.meCmdType.ToString())

            Select Case poCmd.meCmdType
                Case EdmCmdType.EdmCmd_TaskRun
                    OnTaskRun(poCmd, ppoData)
                Case EdmCmdType.EdmCmd_TaskSetup
                    OnTaskSetup(poCmd, ppoData)
                Case EdmCmdType.EdmCmd_TaskSetupButton
                    OnTaskSetupButton(poCmd, ppoData)
            End Select
        Catch ex As Runtime.InteropServices.COMException
            MsgBox("HRESULT = 0x" +
              ex.ErrorCode.ToString("X") + vbCrLf +
              ex.Message)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'Store the custom set-up page here so it can be accessed from both OnTaskSetup and OnTaskSetupButton
    Dim currentSetupPage As MenuCommandPage

    Private Sub OnTaskSetupButton(ByRef poCmd As EdmCmd, ByRef ppoData As System.Array)
        Dim props As IEdmTaskProperties
        props = poCmd.mpoExtra

        'The custom set-up page in currentSetupPage
        'was created in method OnTaskSetup;
        'StoreData saves the contents of the edit
        'box in the user control to
        'IEdmTaskProperties in poCmd.mpoExtra
        If poCmd.mbsComment = "OK" And currentSetupPage IsNot Nothing Then
            currentSetupPage.StoreData(poCmd)
            'Set up the menu commands to launch this task
            Dim cmds(0) As EdmTaskMenuCmd
            cmds(0).mbsMenuString = props.GetValEx("MenuCommand")
            cmds(0).mbsStatusBarHelp = props.GetValEx("HelpText")
            cmds(0).mlCmdID = 1
            cmds(0).mlEdmMenuFlags =
              EdmMenuFlags.EdmMenu_MustHaveSelection
            If props.GetValEx("Checked") = "True" Then
                props.SetMenuCmds(cmds)
            End If
        End If
        currentSetupPage = Nothing
    End Sub

    Private Sub OnTaskRun(ByRef poCmd As EdmCmd, ByRef ppoData As System.Array)

        'Get the task instance interface
        Dim inst As IEdmTaskInstance
        inst = poCmd.mpoExtra
        Try
            'Keep the task list status up to date
            inst.SetStatus _
              (EdmTaskStatus.EdmTaskStat_Running)

            'Format a message that will be displayed
            'in the task list
            inst.SetProgressRange(10, 1, "Task is running.")

            Dim pdmVault As IEdmVault5
            pdmVault = poCmd.mpoVault

            If Not pdmVault.IsLoggedIn Then
                'Log into selected vault as the current user
                pdmVault.LoginAuto("SRCVault", poCmd.mlParentWnd)
            End If

            inst.SetProgressPos(4, "Task is sharing the files.")

            GetLocalCopies(pdmVault, ppoData, inst)

            inst.SetProgressPos(10, "Task finished sharing the files.")

            'Dim ProgresssMsg As String
            'If (Items.Count > 0) Then
            'ProgresssMsg = "Found " +
            'Items.Count.ToString() + " files."
            'Else
            'ProgresssMsg = ("No files found.")
            'End If

            'inst.SetProgressPos(10, ProgresssMsg)
            inst.SetStatus(
            EdmTaskStatus.EdmTaskStat_DoneOK, 0, "", , "")

        Catch ex As Runtime.InteropServices.COMException
            inst.SetStatus _
              (EdmTaskStatus.EdmTaskStat_DoneFailed,
              ex.ErrorCode, "The test task failed!")
        Catch ex As Exception
            inst.SetStatus _
              (EdmTaskStatus.EdmTaskStat_DoneFailed,
              0, "Non COM test task failure!")
        End Try
    End Sub

    Private Sub OnTaskSetup(ByRef poCmd As EdmCmd,
      ByRef ppoData As System.Array)

        Try
            'Get the property interface used to
            'access the framework
            Dim props As IEdmTaskProperties
            props = poCmd.mpoExtra

            'Set the property flag that says you want a
            'menu item for the user to launch the task
            'and a flag to support scheduling
            props.TaskFlags =
              EdmTaskFlag.EdmTask_SupportsInitExec +
              EdmTaskFlag.EdmTask_SupportsChangeState

            currentSetupPage = New MenuCommandPage
            currentSetupPage.CreateControl()
            currentSetupPage.LoadData(poCmd)

            Dim pages(0) As EdmTaskSetupPage
            pages(0).mbsPageName = "Menu Command"
            pages(0).mlPageHwnd = currentSetupPage.Handle.ToInt32
            pages(0).mpoPageImpl = currentSetupPage

            props.SetSetupPages(pages)


        Catch ex As Runtime.InteropServices.COMException
            MsgBox("HRESULT = 0x" +
              ex.ErrorCode.ToString("X") + vbCrLf +
              ex.Message)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub PauseToAttachProcess(
      ByVal callbackType As String)

        Try
            'If the debugger isn't already attached to a
            'process, 
            If Not Debugger.IsAttached() Then
                'Launch the debug dialog
                'Debugger.Launch()
                'or
                'use a MsgBox dialog to pause execution
                'and allow the user time to attach it
                MsgBox("Attach debugger to process """ +
                  Process.GetCurrentProcess.ProcessName() +
                  """ for callback """ + callbackType +
                  """ before clicking OK.")
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub GetLocalCopies(ByVal vault As IEdmVault5, ByVal vFilePaths As Array, ByVal inst As IEdmTaskInstance)

        If vFilePaths.Length > 0 Then

            Dim i As Integer

            Dim folderpath As String
            folderpath = vault.RootFolderPath

            Dim PLANT_DRAWINGS As String
            PLANT_DRAWINGS = folderpath + "\" + inst.GetValEx("SharedPath")

            For i = 0 To UBound(vFilePaths)

                Dim pdmFile As IEdmFile5
                Dim pdmFolder As IEdmFolder5

                Dim filename As String
                Dim filepathcurrent As String
                Dim folderpathcurrent As String
                Dim filepathshared As String
                Dim filepath As String
                Dim extension As String

                filepath = vFilePaths(i).mbsStrData1

                filename = Mid(filepath, InStrRev(filepath, "\") + 1, InStrRev(filepath, ".") - InStrRev(filepath, "\") - 1)
                'extension = Right(filepath, Len(filepath) - InStrRev(filepath, "."))
                extension = inst.GetValEx("SharedExtension")

                'filename = PLANT_DRAWINGS & "\" & filename & "." & extension
                filepathshared = PLANT_DRAWINGS & "\" & filename & "." & extension

                If inst.GetValEx("SubFolder") = "" Then
                    filepathcurrent = Left(filepath, InStrRev(filepath, ".") - 1) + "." + extension
                    folderpathcurrent = Left(filepath, InStrRev(filepath, "\") - 1)
                Else
                    folderpathcurrent = Left(filepath, InStrRev(filepath, "\")) + CStr(inst.GetValEx("SubFolder"))
                    filepathcurrent = folderpathcurrent + "\" + filename + "." + extension
                End If

                pdmFolder = vault.GetFolderFromPath(PLANT_DRAWINGS)

                'Check to see if a file with the same name exists in the vault
                If vault.GetFileFromPath(filepathshared, pdmFolder) Is Nothing Then

                    pdmFolder = vault.GetFolderFromPath(folderpathcurrent)

                    'Check to see if the file to be shared exists
                    If vault.GetFileFromPath(filepathcurrent, pdmFolder) IsNot Nothing Then

                        'Share the file
                        pdmFolder = vault.GetFolderFromPath(filepathcurrent)
                        pdmFile = vault.GetFileFromPath(filepathcurrent, pdmFolder)
                        pdmFolder = vault.GetFolderFromPath(PLANT_DRAWINGS)
                        pdmFolder.AddFileShared(pdmFile.ID, 0)

                    End If

                ElseIf vault.GetFileFromPath(filepathshared, pdmFolder).ID = vault.GetFileFromPath(filepathcurrent, pdmFolder).ID Then
                    'file already exists do nothing

                Else
                    'file exists but needs to be updated and replaced

                    'delete existing file
                    pdmFile = vault.GetFileFromPath(filepathshared, pdmFolder)
                    pdmFolder = vault.GetFolderFromPath(PLANT_DRAWINGS)
                    pdmFolder.DeleteFile(0, pdmFile.ID, True)

                    'Share the file
                    pdmFolder = vault.GetFolderFromPath(filepathcurrent)
                    pdmFile = vault.GetFileFromPath(filepathcurrent, pdmFolder)
                    pdmFolder = vault.GetFolderFromPath(PLANT_DRAWINGS)
                    pdmFolder.AddFileShared(pdmFile.ID, 0)

                End If

                inst.SetProgressPos(7, filename)

            Next

        End If

    End Sub

End Class
