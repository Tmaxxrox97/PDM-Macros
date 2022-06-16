Imports EPDM.Interop.epdm
Imports Renci.SshNet
Imports System.Runtime.InteropServices
Imports System.IO

<Guid("F26CE2D2-CD52-435B-88E4-579C3517574C")>
<ComVisible(True)>
Public Class UploadToSFTP
    Implements IEdmAddIn5

    Public Sub GetAddInInfo(
    ByRef poInfo As EdmAddInInfo,
    ByVal poVault As IEdmVault5,
    ByVal poCmdMgr As IEdmCmdMgr5) _
    Implements IEdmAddIn5.GetAddInInfo

        Try
            poInfo.mbsAddInName =
              "Upload to SFTP"
            poInfo.mbsCompany = "Safety Rail Company"
            poInfo.mbsDescription =
              "Uploads Files to specified SFTP site"
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

            inst.SetProgressPos(4, "Task is uploading the files.")

            UploadFiles(pdmVault, ppoData, inst)

            inst.SetProgressPos(10, "Task finished uploading the files.")

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
              ex.ErrorCode, "The task failed!")
        Catch ex As Exception
            inst.SetStatus _
              (EdmTaskStatus.EdmTaskStat_DoneFailed,
              0, "Non COM task failure!")
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

    Private Sub UploadFiles(ByVal vault As IEdmVault5, ByVal vFilePaths As Array, ByVal inst As IEdmTaskInstance)

        If vFilePaths.Length > 0 Then

            Dim serverpath As String
            Dim username As String
            Dim password As String

            Dim index As Integer
            Dim vParts As Object

            serverpath = inst.GetValEx("ServerPath")
            username = inst.GetValEx("Username")
            password = inst.GetValEx("Password")

            Dim client As New SftpClient(serverpath, username, password)
            client.Connect()

            Dim i As Integer

            For i = 0 To UBound(vFilePaths)

                Dim filename As String
                Dim filepath As String
                Dim remote As String

                filepath = vFilePaths(i).mbsStrData1

                filename = Right(filepath, Len(filepath) - InStrRev(filepath, "\"))

                vParts = LoadCSV("C:\SRC Vault\Shared\Macros\UploadtoSFTP.csv")

                index = LoopThrough2DArray(vParts, filename)

                If Index <> 0 Then
                    GoTo LastLine
                End If

                remote = inst.GetValEx("Remote") + filename

                inst.SetProgressPos(7, filename)

                Using stream As Stream = File.OpenRead(filepath)
                    client.UploadFile(stream, remote)
                End Using

LastLine:

            Next

            client.Disconnect()

        End If

    End Sub

    Private Function LoadCSV(file_name As String) As Object
        Dim whole_file As String
        Dim lines As Object
        Dim one_line As Object
        Dim num_rows As Long
        Dim num_cols As Long
        Dim the_array(1, 1) As Object
        Dim R As Long
        Dim C As Long

        Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(file_name)

            whole_file = MyReader.ReadToEnd

        End Using

        lines = Split(whole_file, vbCrLf)

        ' Dimension the array.
        num_rows = UBound(lines)
        one_line = Split(lines(0), ",")
        num_cols = UBound(one_line)
        ReDim the_array(num_rows, num_cols)

        ' Copy the data into the array.
        For R = 0 To num_rows
            If Len(lines(R)) > 0 Then
                one_line = Split(lines(R), ",")
                For C = 0 To num_cols
                    the_array(R, C) = one_line(C)
                Next C
            End If
        Next R

        ' Prove we have the data loaded.
        For R = 0 To num_rows
            For C = 0 To num_cols
                Debug.Print(the_array(R, C) & "|")
            Next C
            Debug.Print("")
        Next R
        Debug.Print("=======")

        Return the_array

    End Function

    Private Function LoopThrough2DArray(varArray As Object, strFind As String) As Integer
        'declare variables for the loop
        Dim i As Long
        'loop for the first dimension
        For i = LBound(varArray, 1) To UBound(varArray, 1)
            'if we find the value, then msgbox to say that we have the value and exit the function
            If varArray(i, 0) = strFind Then
                Return i
                Exit Function
            End If
        Next i
        If i > UBound(varArray, 1) Then
            i = 0
        End If
        Return i
    End Function

End Class
