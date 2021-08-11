Imports EPDM.Interop.epdm
Imports System.Runtime.InteropServices

<Guid("FC57CC4E-383F-403A-A520-D852FDA6D1F7")>
<ComVisible(True)>
Public Class GetLatestVersionTask
    Implements IEdmAddIn5

    Public Sub GetAddInInfo(
    ByRef poInfo As EdmAddInInfo,
    ByVal poVault As IEdmVault5,
    ByVal poCmdMgr As IEdmCmdMgr5) _
    Implements IEdmAddIn5.GetAddInInfo

        Try
            poInfo.mbsAddInName =
              "Get Latest Version Task Addin"
            poInfo.mbsCompany = "Safety Rail Company"
            poInfo.mbsDescription =
              "Task Addin that gets the latest version of files "
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
            End Select
        Catch ex As Runtime.InteropServices.COMException
            MsgBox("HRESULT = 0x" +
              ex.ErrorCode.ToString("X") + vbCrLf +
              ex.Message)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
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

            inst.SetProgressPos(4, "Task is getting the latest.")

            GetLocalCopies(pdmVault, ppoData, inst)

            inst.SetProgressPos(10, "Task finished getting latest versions.")

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
              EdmTaskFlag.EdmTask_SupportsScheduling +
              EdmTaskFlag.EdmTask_SupportsChangeState

            'Set up the menu commands to launch this task
            Dim cmds(0) As EdmTaskMenuCmd
            cmds(0).mbsMenuString = "Tasks\Get Latest Version task"
            cmds(0).mbsStatusBarHelp =
              "This command runs the task add-in to get the" +
              " Latest Version of the Selected Files."
            cmds(0).mlCmdID = 1
            cmds(0).mlEdmMenuFlags =
              EdmMenuFlags.EdmMenu_MustHaveSelection
            props.SetMenuCmds(cmds)
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

            Dim pdmBatchGetUtil As IEdmBatchGet
            pdmBatchGetUtil = vault.CreateUtility(EdmUtility.EdmUtil_BatchGet)

            Dim i As Integer
            Dim filename As String

            Dim pdmSelItems() As EdmSelItem
            ReDim pdmSelItems(UBound(vFilePaths))

            For i = 0 To UBound(vFilePaths)

                pdmSelItems(i).mlDocID = vFilePaths(i).mlObjectID1
                pdmSelItems(i).mlProjID = vFilePaths(i).mlObjectID2
                filename = vFilePaths(i).mbsStrData1

                inst.SetProgressPos(7, filename)

            Next

            pdmBatchGetUtil.AddSelection(vault, pdmSelItems)
            pdmBatchGetUtil.CreateTree(0, EdmGetCmdFlags.Egcf_RefreshFileListing)
            pdmBatchGetUtil.GetFiles(0)

        End If

    End Sub

End Class
