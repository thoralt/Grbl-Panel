Imports System.IO
Imports GrblPanel.GrblIF

Partial Class GrblGui

    Public Class GrblGcode

        ' A Class to handle reading, parsing, removing white space
        '   - Handles the sending to Grbl using either the simple or advanced protocols
        '   - Handles introducing canned cycles (M06, G81/2/3)

#Region "Private fields"
        Private _gui As GrblGui
        Private _queue As GrblQueue
#End Region

#Region "Public methods"
        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="gui">The main GUI instance</param>
        Public Sub New(ByRef gui As GrblGui)
            _gui = gui
            _queue = gui.grblQueue
        End Sub

        ''' <summary>
        ''' Loads a given file into the queue.
        ''' </summary>
        ''' <param name="file">Filename to load</param>
        ''' <returns>True if successful</returns>
        Public Function loadGCodeFile(ByVal file As String) As Boolean
            Dim data As String

            ' Start from clean slate
            resetGcode(True)

            Try
                ' stop the ListView from updating, increases performance while inserting many lines
                _gui.lvGcode.BeginUpdate()

                ' Load the file
                Dim inputFile As StreamReader = Nothing
                inputFile = My.Computer.FileSystem.OpenTextFileReader(file)

                ' count the lines while loading up
                Dim inputcount As Integer = 0
                Do While Not inputFile.EndOfStream
                    data = inputFile.ReadLine()    ' Issue #20, ignore '%'
                    If data <> "%" Then
                        ' enqueue one line of GCode
                        ' this also fires QueueAddedItem event which in turn adds the item to the ListView
                        _queue.Enqueue(data, False, inputcount)
                        inputcount += 1
                    End If
                Loop

                _gui.lvGcode.EnsureVisible(0)           ' show top of file for user to verify etc

                If Not IsNothing(inputFile) Then
                    inputFile.Close()
                End If        ' Issue #19
            Catch ex As Exception
                MessageBox.Show("An exception occured while trying to read the file:" & vbNewLine &
                                vbNewLine & ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            Finally
                ' enable updating of ListView to show our changes
                _gui.lvGcode.EndUpdate()
            End Try

            Return True
        End Function

        ''' <summary>
        ''' Clean up this instance.
        ''' </summary>
        Public Sub shutdown()
            ' Close up shop
            resetGcode(True)
        End Sub

        ''' <summary>
        ''' Clears the ListView, queue, progress bar and text elements
        ''' </summary>
        ''' <param name="fullstop"></param>
        Private Sub resetGcode(ByVal fullstop As Boolean)
            ' Clear out all variables etc to initial state
            _gui.lblTotalLines.Text = ""
            _gui.tbGCodeMessage.Text = ""
            _gui.tbGcodeFile.Text = ""
            _gui.TransmissionProgress.Value = 0

            ' TODO: resetGcode() is always called with fullstop=True -> should we remove fullstop?
            If fullstop Then
                ' Clear the list of gcode block sent
                _gui.gcodeview.Clear()
                _queue.Clear()
            End If
        End Sub
    End Class

    Private Sub btnCheckMode_Click(sender As Object, e As EventArgs) Handles btnCheckMode.Click
        ' Enable/disable Check mode in Grbl
        ' Just send a $C, this toggles Check state in Grbl
        'grblPort.sendData("$C")
        grblQueue.ExecuteImmediateCommand("$C")
    End Sub

    ''' <summary>
    ''' This event handler deals with the gcode file related buttons
    ''' Implements a simple state machine to keep user from clicking the wrong buttons
    ''' Uses button.tag instead of .text so the text doesn't mess up the images on the buttons
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnFileGroup_Click(sender As Object, e As EventArgs) Handles btnFileSend.Click, btnFileSelect.Click, btnFilePause.Click, btnFileStop.Click,
                                    btnFileReload.Click
        Dim args As Button = sender
        Select Case args.Tag
            Case "File"
                Dim str As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                'ofdGcodeFile.InitialDirectory = Path.Combine(Path.GetFullPath(str), "*") ' commented out because of FileNotFoundException
                If tbSettingsDefaultExt.Text <> "" Then
                    ofdGcodeFile.Filter = String.Format("Gcode |*.{0}|All Files |*.*", tbSettingsDefaultExt.Text)
                    ofdGcodeFile.DefaultExt = String.Format(".{0}", tbSettingsDefaultExt.Text)
                End If
                'ofdGcodeFile.FileName = "File" ' commented out to allow re-opening the last file
                If ofdGcodeFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    'gcode.openGCodeFile(ofdGcodeFile.FileName)
                    gcode.loadGCodeFile(ofdGcodeFile.FileName)
                    tbGcodeFile.Text = ofdGcodeFile.FileName
                    lblTotalLines.Text = grblQueue.numberOfLines().ToString()

                    btnFileSelect.Enabled = True    ' Allow changing your mind about the file
                    btnFileSend.Enabled = True
                    btnFilePause.Enabled = False
                    btnFileStop.Enabled = False
                    btnFileReload.Enabled = False
                    ' reset filter in case user changes ext on Settings tab
                    ofdGcodeFile.Filter = ""
                    ofdGcodeFile.DefaultExt = ""
                End If
            Case "Send"
                ' Send a gcode file to Grbl
                setSubPanels("GCodeStream")
                grblQueue.Resume()

                btnFileSelect.Enabled = False
                btnFileSend.Enabled = False
                btnFilePause.Enabled = True
                btnFileStop.Enabled = True
                btnFileReload.Enabled = False

            Case "Pause"
                grblQueue.Pause()

                btnFileSelect.Enabled = False
                btnFileSend.Tag = "Resume"
                btnFileSend.Enabled = True
                btnFilePause.Enabled = False
                btnFileStop.Enabled = True
                btnFileReload.Enabled = False

            Case "Stop"
                gcodeview.StopSending()

                ' Re-enable manual control
                setSubPanels("Idle")
                btnFileSelect.Enabled = True
                btnFileSend.Tag = "Send"
                btnFileSend.Enabled = False
                btnFilePause.Enabled = False
                btnFileStop.Enabled = False
                btnFileReload.Enabled = True

            Case "Resume"
                grblQueue.Resume()
                btnFileSelect.Enabled = False
                btnFileSend.Tag = "Send"
                btnFileSend.Enabled = False
                btnFilePause.Enabled = True
                btnFileStop.Enabled = True
                btnFileReload.Enabled = False

            Case "Reload"
                ' Reload the same file 
                gcode.loadGCodeFile(ofdGcodeFile.FileName)
                tbGcodeFile.Text = ofdGcodeFile.FileName
                lblTotalLines.Text = grblQueue.numberOfLines().ToString()

                btnFileSelect.Enabled = True    ' Allow changing your mind about the file
                btnFileSend.Enabled = True
                btnFilePause.Enabled = False
                btnFileStop.Enabled = False
                btnFileReload.Enabled = False

        End Select
    End Sub

#End Region

    Public Class GrblGcodeView
        ' A class to manage the Gcode list view and GrblQueue events
        ' This contains the GCode queue going to Grbl
        ' GrblGui owns the lvGcode control but this class manages its content
#Region "Private fields"
        Private _gui As GrblGui
        Private _lview As ListView
        Private _message As TextBox
        Private _progress As ProgressBar
        Private _bufferLevel As ProgressBar
        Private _alarmDescription As Label
        Private WithEvents _queue As GrblQueue

        ''' <summary>
        ''' This is the item index which will be processed after resume (gets set by QueuePaused event)
        ''' and is used to stop the QueueItemStateChanged event from marking the paused item as executing
        ''' </summary>
        Private _pausedItem As Integer = -1
#End Region

#Region "Public methods"
        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="gui">The main GUI instance</param>
        Public Sub New(ByRef gui As GrblGui)
            _gui = gui
            _lview = gui.lvGcode
            _queue = gui.grblQueue
            _message = gui.tbGCodeMessage
            _progress = gui.TransmissionProgress
            _bufferLevel = gui.GrblBufferLevel
            _alarmDescription = gui.lblAlarmDescription
        End Sub

        ''' <summary>
        ''' Clears the gcode ListView, progress bar and queue.
        ''' </summary>
        Public Sub Clear()
            _lview.Items.Clear()
            _lview.Update()
            _queue.Clear()
            _pausedItem = -1
            _progress.Value = 0
        End Sub

        ''' <summary>
        ''' Stops the queue and marks the entry after the last one being sent as stopped in the UI
        ''' </summary>
        Public Sub StopSending()
            _queue.Pause()

            Dim stoppedItem As ListViewItem = Nothing

            If _pausedItem > -1 And _pausedItem < _lview.Items.Count Then
                ' if paused, mark paused entry in ListView as stopped
                stoppedItem = _lview.Items(_pausedItem)
            Else
                ' check if there's an entry after the last buffered/last acknowledged item
                Dim lastBufferedItemIndex As Integer = -1
                Dim lastAcknowledgedItemIndex As Integer = -1
                Dim index As Integer = -1
                For Each item As ListViewItem In _lview.Items
                    If item.SubItems(0).Text = "buff" Then
                        lastBufferedItemIndex = item.Index
                    ElseIf item.SubItems(0).Text = "OK" OrElse item.SubItems(0).Text = "err" Then
                        lastAcknowledgedItemIndex = item.Index
                    End If
                Next

                If lastBufferedItemIndex > lastAcknowledgedItemIndex Then
                    index = lastBufferedItemIndex + 1
                Else
                    index = lastAcknowledgedItemIndex + 1
                End If
                If index >= 0 And index < _lview.Items.Count Then
                    stoppedItem = _lview.Items(index)
                End If
            End If

            ' mark as stopped
            If stoppedItem IsNot Nothing Then
                stoppedItem.BackColor = Color.Black
                stoppedItem.ForeColor = Color.White
                stoppedItem.SubItems(0).Text = "stop"
            End If
        End Sub

#End Region

#Region "GrblQueue event handlers"

        ''' <summary>
        ''' Gets fired when GrblQueue gets ALARM from grbl. Displays alarm description on GUI.
        ''' </summary>
        ''' <param name="description">Grbl's alarm text</param>
        Public Delegate Sub QueueAlarm_Delegate(description As String)
        Private Sub QueueAlarm(description As String) Handles _queue.QueueAlarm
            If _alarmDescription.InvokeRequired Then
                Dim _delegate As New QueueAlarm_Delegate(AddressOf QueueAlarm)
                _alarmDescription.BeginInvoke(_delegate, New Object() {description})
            Else
                _alarmDescription.Text = description
                _alarmDescription.BackColor = Color.Red
                _alarmDescription.ForeColor = Color.White
                _alarmDescription.Visible = True

                ' stop sending more commands to grbl
                _gui.btnFileGroup_Click(_gui.btnFileStop, Nothing)
            End If
        End Sub

        ''' <summary>
        ''' Gets called when GrblQueue detects an error condition in grbl's output.
        ''' Pauses the GUI and queue.
        ''' </summary>
        ''' <param name="description"></param>
        Public Delegate Sub QueueError_Delegate(description As String)
        Private Sub QueueError(description As String) Handles _queue.QueueError
            If _alarmDescription.InvokeRequired Then
                Dim _delegate As New QueueError_Delegate(AddressOf QueueError)
                _alarmDescription.BeginInvoke(_delegate, New Object() {description})
            Else
                _alarmDescription.Text = description
                _alarmDescription.BackColor = Color.Yellow
                _alarmDescription.ForeColor = Color.Black
                _alarmDescription.Visible = True

                ' pause on error
                _gui.btnFileGroup_Click(_gui.btnFilePause, Nothing)
            End If
        End Sub

        ''' <summary>
        ''' Grbl has been reset after an alarm condition and is operating normally at the moment. Hides the alarm message text.
        ''' </summary>
        Public Delegate Sub QueueResetAlarm_Delegate()
        Private Sub QueueResetAlarm() Handles _queue.QueueResetAlarm
            If _alarmDescription.InvokeRequired Then
                Dim _delegate As New QueueResetAlarm_Delegate(AddressOf QueueResetAlarm)
                _alarmDescription.BeginInvoke(_delegate)
            Else
                ' reset alarm display
                _alarmDescription.Visible = False
            End If
        End Sub

        ''' <summary>
        ''' Reports grbl's buffer level in %
        ''' sets the content of text box tbGCodeMessage
        ''' the last queued command.
        ''' </summary>
        Public Delegate Sub QueueGrblBufferLevel_Delegate(level As Integer)
        Private Sub QueueGrblBufferLevel(level As Integer) Handles _queue.QueueGrblBufferLevel
            If _bufferLevel.InvokeRequired Then
                Dim _delegate As New QueueGrblBufferLevel_Delegate(AddressOf QueueGrblBufferLevel)
                _bufferLevel.BeginInvoke(_delegate, New Object() {level})
            Else
                _bufferLevel.Value = level
            End If
        End Sub

        ''' <summary>
        ''' Gets called whenever an item has been added to the queue
        ''' -> display the item in the GUI
        ''' the last queued command.
        ''' </summary>
        Public Delegate Sub QueueAddedItem_Delegate(item As GrblQueueItem)
        Private Sub QueueAddedItem(item As GrblQueueItem) Handles _queue.QueueAddedItem
            If _lview.InvokeRequired Then
                Dim _delegate As New QueueAddedItem_Delegate(AddressOf QueueAddedItem)
                _lview.BeginInvoke(_delegate, New Object() {item})
            Else
                Dim lvi As New ListViewItem
                lvi.Text = ""                             ' This is for Status of command
                lvi.SubItems.Add(item.Index.ToString())   ' file line number
                lvi.SubItems.Add(item.Text)               ' This is the Gcode block
                _lview.Items.Add(lvi)
            End If
        End Sub

        ''' <summary>
        ''' Reports queue progress in %, sets the progress bar in the GUI
        ''' the last queued command.
        ''' </summary>
        Public Delegate Sub QueueProgress_Delegate(progressValue As Integer)
        Private Sub QueueProgress(progressValue As Integer) Handles _queue.QueueProgress
            If _progress.InvokeRequired Then
                Dim _delegate As New QueueProgress_Delegate(AddressOf QueueProgress)
                _progress.BeginInvoke(_delegate, New Object() {progressValue})
            Else
                _progress.Value = progressValue
            End If
        End Sub

        ''' <summary>
        ''' Gets called whenever the queue finds a gcode comment "(... comment text ...)",
        ''' sets the content of text box tbGCodeMessage
        ''' the last queued command.
        ''' </summary>
        Public Delegate Sub QueueComment_Delegate(comment As String)
        Private Sub QueueComment(comment As String) Handles _queue.QueueComment
            If _message.InvokeRequired Then
                Dim _delegate As New QueueComment_Delegate(AddressOf QueueComment)
                _message.BeginInvoke(_delegate, New Object() {comment})
            Else
                _message.Text = comment
            End If
        End Sub

        ''' <summary>
        ''' Gets called whenever the queue is paused, sets the state of
        ''' the current command in the GUI.
        ''' the last queued command.
        ''' </summary>
        Public Delegate Sub QueueResuming_Delegate(item As GrblQueueItem)
        Private Sub QueueResuming(item As GrblQueueItem) Handles _queue.QueueResuming
            If _lview.InvokeRequired Then
                Dim _delegate As New QueueResuming_Delegate(AddressOf QueueResuming)
                _lview.BeginInvoke(_delegate, New Object() {item})
            Else
                ' mark the active command as executing
                If item.Index > 0 AndAlso item.Index < _lview.Items.Count() Then
                    _pausedItem = -1 ' there is no paused item at the moment
                    _lview.Items(item.Index).BackColor = Color.Orange
                    _lview.Items(item.Index).SubItems(0).Text = "exec"
                    _lview.Update()
                End If
            End If
        End Sub

        ''' <summary>
        ''' Gets called whenever the queue is paused, sets the state of
        ''' the current command in the GUI.
        ''' </summary>
        Public Delegate Sub QueuePaused_Delegate(item As GrblQueueItem)
        Private Sub QueuePaused(item As GrblQueueItem) Handles _queue.QueuePaused
            If _lview.InvokeRequired Then
                Dim _delegate As New QueuePaused_Delegate(AddressOf QueuePaused)
                _lview.BeginInvoke(_delegate, New Object() {item})
            Else
                ' mark the active command as paused
                If item.Index > 0 AndAlso item.Index < _lview.Items.Count() Then
                    _pausedItem = item.Index
                    _lview.Items(_pausedItem).BackColor = Color.LightBlue
                    _lview.Items(_pausedItem).SubItems(0).Text = "pause"
                    _lview.Update()
                End If
            End If
        End Sub

        ''' <summary>
        ''' Gets called whenever the queue finished sending and grbl acknowledged
        ''' the last queued command.
        ''' </summary>
        Public Delegate Sub QueueFinished_Delegate()
        Private Sub QueueFinished() Handles _queue.QueueFinished
            If _lview.InvokeRequired Then
                Dim _delegate As New QueueFinished_Delegate(AddressOf QueueFinished)
                _lview.BeginInvoke(_delegate)
            Else
                ' queue finished -> enable GUI controls
                _gui.btnFileGroup_Click(_gui.btnFileStop, Nothing)
            End If
        End Sub

        ''' <summary>
        ''' This delegate gets called everytime an item in the queue changes its state
        ''' and updates the GUI list view colors depending on the item state.
        ''' </summary>
        ''' <param name="item"></param>
        ''' <param name="oldState"></param>
        ''' <param name="newState"></param>
        Public Delegate Sub QueueItemStatusChanged_Delegate(item As GrblQueueItem,
                                           oldState As GrblQueueItem.ItemState,
                                           newState As GrblQueueItem.ItemState)
        Private Sub QueueItemStatusChanged(item As GrblQueueItem,
                                           oldState As GrblQueueItem.ItemState,
                                           newState As GrblQueueItem.ItemState) Handles _queue.QueueItemStateChanged

            If _lview.InvokeRequired Then
                Dim _delegate As New QueueItemStatusChanged_Delegate(AddressOf QueueItemStatusChanged)
                _lview.BeginInvoke(_delegate, New Object() {item, oldState, newState})
            Else
                ' safety check
                If item.Index < 0 OrElse item.Index >= _lview.Items.Count Then Return

                If newState = GrblQueueItem.ItemState.waiting Then
                    ' an item has been added to our queue and is waiting to be sent to grbl
                    ' -> color it grey
                    _lview.Items(item.Index).BackColor = Color.LightGray
                    _lview.Update()
                End If

                If newState = GrblQueueItem.ItemState.sent Then
                    ' an item has been transferred to grbl's internal buffer
                    ' -> color it yellow
                    _lview.Items(item.Index).BackColor = Color.LightYellow
                    _lview.Items(item.Index).SubItems(0).Text = "buff"

                    Dim index As Integer = item.Index + 2
                    If index >= _lview.Items.Count() Then index = _lview.Items.Count - 1
                    _lview.EnsureVisible(index) ' scroll to latest command in grbl's buffer
                    _lview.Update()
                End If

                If newState = GrblQueueItem.ItemState.acknowledged Then
                    ' an item has been acknowledged by grbl which means the corresponding
                    ' command has been executed -> color it green
                    'Console.WriteLine("Ack " & item.Index & " " & item.Text)
                    If item.ErrorFlag Then
                        _lview.Items(item.Index).BackColor = Color.Red
                        _lview.Items(item.Index).SubItems(0).Text = "Err"
                    Else
                        _lview.Items(item.Index).BackColor = Color.LightGreen
                        _lview.Items(item.Index).SubItems(0).Text = "OK"
                    End If

                    ' if a next item does exist, color it orange to mark it as currently executing
                    If item.Index + 1 < _lview.Items.Count AndAlso item.Index + 1 <> _pausedItem Then
                        Dim index As Integer = item.Index + 1
                        _lview.Items(index).BackColor = Color.Orange
                        _lview.Items(index).SubItems(0).Text = "exec"
                    End If
                    _lview.Update()
                End If
            End If
        End Sub
#End Region

    End Class


End Class