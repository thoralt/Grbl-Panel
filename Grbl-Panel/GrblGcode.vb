Imports System.IO
Imports GrblPanel.GrblIF

Partial Class GrblGui

    Public Class GrblGcode

        ' A Class to handle reading, parsing, removing white space
        '   - Handles the sending to Grbl using either the simple or advanced protocols
        '   - Handles introducing canned cycles (M06, G81/2/3)
        Private _gui As GrblGui
        Private _wtgForAck As Boolean = False         '
        Private _runMode As Boolean = True            '
        Private _sendAnotherLine As Boolean = False   '
        Private _lineCount As Integer = 0               ' No of lines left to send
        Private _linesDone As Integer = 0

        ' Handle file read (Gcode in) and Write (Gcode save)
        ' While we are sending the file, lock out manual functions
        Private _inputfh As StreamReader
        Private _inputcount As Integer

        Public Sub New(ByRef gui As GrblGui)
            _gui = gui
        End Sub

        Public Sub enableGCode(ByVal action As Boolean)
            ' Can't use new if we need to reference _gui as it causes Form creation errors
            _gui.gbGcode.Enabled = action
            If action = True Then
                ' Enable looking at responses now, for use by manual commands
                _gui.grblPort.addRcvDelegate(AddressOf _gui.processLineEvent)
                _gui.btnFileSelect.Enabled = True
            Else
                _gui.grblPort.deleteRcvDelegate(AddressOf _gui.processLineEvent)
            End If
        End Sub

        Public Function loadGCodeFile(ByVal file As String) As Boolean
            Dim data As String

            _gui.lvGcode.BeginUpdate()

            ' Start from clean slate
            resetGcode(True)
            ' Load the file, count lines
            _inputfh = My.Computer.FileSystem.OpenTextFileReader(file)
            ' count the lines while loading up
            _inputcount = 0
            Do While Not _inputfh.EndOfStream
                data = _inputfh.ReadLine()    ' Issue #20, ignore '%'
                If data <> "%" Then
                    _gui.gcodeview.Insert(data, _inputcount)
                    _inputcount += 1
                End If
            Loop

            lineCount = _inputcount

            If Not IsNothing(_inputfh) Then
                _inputfh.Close()
            End If        ' Issue #19

            _gui.lvGcode.EndUpdate()

            Return True

        End Function
        Public Sub closeGCodeFile()
            If Not IsNothing(_inputfh) Then
                _inputfh.Close()
            End If
            _gui.tbGcodeFile.Text = ""
            _inputcount = 0

        End Sub
        Public Function readGcode() As String
            ' Read a line, if EOF then return EOF
            Dim lv As ListView = _gui.lvGcode
            If _lineCount > 0 Then
                Return lv.Items(_linesDone).SubItems(2).Text
            Else
                Return "EOF"
            End If
        End Function

        Public Sub sendGcodeFile()

            ' Workflow:
            ' Disable other panels to prevent operator error
            _gui.setSubPanels("GCodeStream")
            ' set sendAnotherLine
            ' raise processLineEvent
            lineCount = _inputcount
            linesDone = 0
            wtgForAck = False
            runMode = True
            sendAnotherLine = True
            _gui.gcodeview.fileMode = True

            '_gui.processLineEvent("")              ' Prime the pump
            _gui.grblQueue.resumeSending()


        End Sub

        Public Sub sendGCodeLine(ByVal data As String)
            ' Send a line immediately
            ' This can only happen when not sending a file, buttons are interlocked
            _runMode = False
            _gui.gcodeview.fileMode = False

            If Not (data.StartsWith("$") Or data.StartsWith("?")) Then
                ' add to display
                _gui.gcodeview.Insert(data, 0)
                gcode.lineCount += 1        ' TODO is this necessary?
                ' we are always be the last item in manual mode
                _gui.gcodeview.UpdateGcodeSent(-1)
                ' Expect a response from Grbl
                wtgForAck = True
            End If
            _gui.state.ProcessGCode(data)            ' Keep Gcode State object in the loop
            _gui.grblPort.sendData(data)

        End Sub

        Public Sub sendGCodeFilePause()
            ' Pause the file send
            _sendAnotherLine = False
            _runMode = False
            _gui.grblQueue.pauseSending()
        End Sub

        Public Sub sendGCodeFileResume()
            ' Resume sending of file
            _sendAnotherLine = True
            _runMode = True
            _gui.gcodeview.fileMode = True
            _gui.processLineEvent("")              ' Prime the pump again
            _gui.grblQueue.resumeSending()
        End Sub
        Public Sub sendGCodeFileStop()

            ' reset state variables
            If runMode Then
                wtgForAck = False
                runMode = False
                sendAnotherLine = False
                _gui.gcodeview.fileMode = False        ' allow manual mode gcode send

                ' Make the fileStop button go click, to stop the file send
                ' and set the buttons
                _gui.btnFileGroup_Click(_gui.btnFileStop, Nothing)
            End If

        End Sub

        Public Sub shutdown()
            ' Close up shop
            resetGcode(True)
        End Sub

#Region "Properties"

        Property runMode() As Boolean
            Get
                Return _runMode
            End Get
            Set(value As Boolean)
                _runMode = value
            End Set
        End Property
        Property wtgForAck() As Boolean
            Get
                Return _wtgForAck
            End Get
            Set(value As Boolean)
                _wtgForAck = value
            End Set
        End Property
        Property sendAnotherLine() As Boolean
            Get
                Return _sendAnotherLine
            End Get
            Set(value As Boolean)
                _sendAnotherLine = value
            End Set
        End Property

        Property linesDone As Int64
            Get
                Return _linesDone
            End Get
            Set(value As Int64)
                _linesDone = value
            End Set
        End Property

        Property lineCount As Int64
            Get
                Return _lineCount
            End Get
            Set(value As Int64)
                _lineCount = value
            End Set
        End Property

#End Region

        Private Sub resetGcode(ByVal fullstop As Boolean)
            ' Clear out all variables etc to initial state
            lineCount = 0
            linesDone = 0
            _gui.lblTotalLines.Text = ""
            _gui.tbGCodeMessage.Text = ""
            ' clear out the file name etc
            closeGCodeFile()
            ' reset state variables
            wtgForAck = False
            runMode = False
            sendAnotherLine = False

            If fullstop Then
                ' Clear the list of gcode block sent
                _gui.gcodeview.Clear()
            End If
        End Sub
    End Class

    Public Class GrblGcodeView
        ' A class to manage the Gcode list view
        ' This contains the GCode queue going to Grbl
        ' GrblGui owns the lvGcode control but this class manages its content

        Private _lview As ListView
        Private _message As TextBox
        Private _progress As ProgressBar
        Private _bufferLevel As ProgressBar
        Private WithEvents _queue As GrblQueue
        Private _filemode As Boolean = False ' True if in File Send mode
        Private _pausedItem As Integer = -1

        Property ignoreScreenUpdate As Boolean
            Get
                Return _ignoreScreenUpdate
            End Get
            Set(value As Boolean)
                _ignoreScreenUpdate = value
                If value = False Then _lview.Update()
            End Set
        End Property
        Private _ignoreScreenUpdate As Boolean = False

        Public Sub New(ByRef view As ListView, ByRef queue As GrblQueue, ByRef message As TextBox, progress As ProgressBar, bufferLevel As ProgressBar)
            _lview = view
            _queue = queue
            _message = message
            _progress = progress
            _bufferLevel = bufferLevel
        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        Public Sub Clear()
            _lview.Items.Clear()
            _filemode = False
            _lview.Update()
            _queue.Clear()
            _pausedItem = -1
            _progress.Value = 0
        End Sub

#Region "GrblQueue Events"

        ''' <summary>
        ''' Gets called whenever the queue finds a gcode comment "(... comment text ...)",
        ''' sets the content of text box tbGCodeMessage
        ''' the last queued command.
        ''' </summary>
        Public Delegate Sub QueueGrblBufferLevel_Delegate(level As Integer)
        Private Sub QueueGrblBufferLevel(level As Integer) Handles _queue.QueueGrblBufferLevel
            If _bufferLevel.InvokeRequired Then
                ' we need to cross thread this callback
                Dim _delegate As New QueueGrblBufferLevel_Delegate(AddressOf QueueGrblBufferLevel)
                _bufferLevel.BeginInvoke(_delegate, New Object() {level})
            Else
                _bufferLevel.Value = level
            End If
        End Sub

        ''' <summary>
        ''' Gets called whenever the queue finds a gcode comment "(... comment text ...)",
        ''' sets the content of text box tbGCodeMessage
        ''' the last queued command.
        ''' </summary>
        Public Delegate Sub QueueProgress_Delegate(progressValue As Integer)
        Private Sub QueueProgress(progressValue As Integer) Handles _queue.QueueProgress
            If _progress.InvokeRequired Then
                ' we need to cross thread this callback
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
                ' we need to cross thread this callback
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
                ' we need to cross thread this callback
                Dim _delegate As New QueueResuming_Delegate(AddressOf QueueResuming)
                _lview.BeginInvoke(_delegate, New Object() {item})
            Else
                ' mark the active command as executing
                If item.Index > 0 AndAlso item.Index < _lview.Items.Count() Then
                    _pausedItem = -1
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
                ' we need to cross thread this callback
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
                ' we need to cross thread this callback
                Dim _delegate As New QueueFinished_Delegate(AddressOf QueueFinished)
                _lview.BeginInvoke(_delegate)
            Else
                ' queue finished -> enable GUI controls
                gcode.sendGCodeFileStop()
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
                ' we need to cross thread this callback
                Dim _delegate As New QueueItemStatusChanged_Delegate(AddressOf QueueItemStatusChanged)
                _lview.BeginInvoke(_delegate, New Object() {item, oldState, newState})
            Else
                If newState = GrblQueueItem.ItemState.waiting Then
                    ' an item has been added to our queue and is waiting to be sent to grbl
                    ' -> color it grey
                    _lview.Items(item.Index).BackColor = Color.LightGray
                    If Not _ignoreScreenUpdate Then _lview.Update()
                End If

                If newState = GrblQueueItem.ItemState.sent Then
                    ' an item has been transferred to grbl's internal buffer
                    ' -> color it yellow
                    _lview.Items(item.Index).BackColor = Color.LightYellow
                    _lview.Items(item.Index).SubItems(0).Text = "buff"

                    Dim index As Integer = item.Index + 2
                    If index >= _lview.Items.Count() Then index = _lview.Items.Count - 1
                    _lview.EnsureVisible(index) ' scroll to latest command in grbl's buffer
                    If Not _ignoreScreenUpdate Then _lview.Update()
                End If

                If newState = GrblQueueItem.ItemState.acknowledged Then
                    ' an item has been acknowledged by grbl which means the corresponding
                    ' command has been executed -> color it green
                    Console.WriteLine("Ack " & item.Index & " " & item.Text)
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
                    If Not _ignoreScreenUpdate Then _lview.Update()
                End If
            End If
        End Sub
#End Region

        Public Sub Insert(ByVal data As String, ByVal lineNumber As Integer)
            ' Insert a new item into the view
            Dim lvi As New ListViewItem
            lvi.Text = ""                       ' This is for Status of command
            lvi.SubItems.Add(lineNumber.ToString)    ' file line number
            lvi.SubItems.Add(data)              ' This is the Gcode block
            _lview.Items.Add(lvi)
            _queue.enqueue(data, False, lineNumber)

            _lview.EnsureVisible(0)           ' show top of file for user to verify etc
            '_lview.Update()
        End Sub

        Public Sub UpdateGCodeStatus(ByVal stat As String, ByVal index As Integer)
            Return

            ' Set the Status column of the line item
            If _filemode Then
                _lview.Items(index).Text = stat
                _lview.EnsureVisible(index)
            Else            ' we always pick the last entry
                _lview.Items(_lview.Items.Count - 1).Text = stat
                _lview.EnsureVisible(_lview.Items.Count - 1)
            End If

            _lview.Update()
        End Sub

        Public Sub UpdateGcodeSent(ByVal index As Integer)
            Return

            '  Set background to indicate the gcode line was sent
            If _filemode Then       ' Are we running a file
                _lview.Items(index).BackColor = Color.LightBlue
                _lview.EnsureVisible(index)
            Else
                _lview.Items(_lview.Items.Count + index).BackColor = Color.LightBlue
                _lview.EnsureVisible(_lview.Items.Count + index)
            End If

            _lview.Update()

        End Sub

        ReadOnly Property count As Integer
            Get
                Return GrblGui.lvGcode.Items.Count
            End Get
        End Property

        Property fileMode As Boolean
            ' Set true if we are running a gcode file
            Get
                Return _filemode
            End Get
            Set(value As Boolean)
                _filemode = value
            End Set
        End Property
    End Class

    Public Sub processLineEvent(ByVal data As String)
        Return


        ' This event handles processing and sending GCode lines from the file as well as ok/error responses from Grbl
        ' Implements simple protocol (send block, wait for ok/error loop)
        ' TODO implement stuffing protocol
        ' It runs on the UI thread, and is raised for each line received from Grbl
        ' even when there is no file to send, e.g. due to status poll response
        ' TODO THIS WILL ALL BE REPLACED WHEN we add Gcode editing and Macro insertion (for canned cycles, tool change etc)

        ' we need this to run in the UI thread so:
        'Console.WriteLine("processLineEvent: " + data)
        If Me.lvGcode.InvokeRequired Then
            ' we need to cross thread this callback
            Dim ncb As New grblDataReceived(AddressOf Me.processLineEvent)
            Me.BeginInvoke(ncb, New Object() {data})
        Else
            ' are we waiting for Ack?
            If gcode.wtgForAck Then
                ' is recvData ok or error?
                If data.StartsWith("ok") Or data.StartsWith("error") Then
                    ' Mark gcode item as ok/error
                    gcodeview.UpdateGCodeStatus(data, gcode.linesDone - 1)
                    ' No longer waiting for Ack
                    gcode.wtgForAck = False
                    If gcode.runMode Then               ' if not paused or stopped
                        ' Mark sendAnotherLine
                        gcode.sendAnotherLine = True
                    End If
                End If
            End If
            ' Do we have another line to send?
            If gcode.runMode = True Then                    ' if not paused or stopped
                If gcode.sendAnotherLine Then
                    gcode.sendAnotherLine = False
                    ' if count > 0
                    If gcode.lineCount > 0 Then
                        Dim line As String
                        ' Read another line
                        line = gcode.readGcode()
                        If Not line.StartsWith("EOF") Then  ' We never hit this but is here just in case the file gets truncated
                            ' count - 1
                            gcode.lineCount -= 1
                            ' show as sent
                            gcodeview.UpdateGcodeSent(gcode.linesDone)                  ' Mark line as sent
                            gcode.linesDone += 1
                            state.ProcessGCode(line)
                            ' Set Message if it starts with (
                            If line.StartsWith("(") Then
                                Dim templine As String = line
                                templine = templine.Remove(0, 1)
                                templine = templine.Remove(templine.Length - 1, 1)
                                tbGCodeMessage.Text = templine
                            End If
                            ' Remove all whitespace
                            line = line.Replace(" ", "")
                            ' set wtg for Ack
                            gcode.wtgForAck = True
                            ' Ship it Dano!
                            grblPort.sendData(line)
                        End If
                    Else
                        ' We reached the EOF aka linecount=0, yippee
                        gcode.sendGCodeFileStop()
                    End If
                End If
            End If
            ' Check for status repsonses that we need to handle here
            ' Extract status
            Dim status = Split(data, ",")
            If status(0) = "<Alarm" Or status(0).StartsWith("ALARM") Then
                ' Major problem so cancel the file
                ' GrblStatus has set the Alarm indicator etc
                gcode.sendGCodeFileStop()
            End If
            If status(0).StartsWith("error") Then
                ' We pause file send to allow operator to determine proceed or not
                If cbSettingsPauseOnError.Checked Then
                    btnFilePause.PerformClick()
                End If
            End If
        End If
    End Sub

    Private Sub btnCheckMode_Click(sender As Object, e As EventArgs) Handles btnCheckMode.Click
        ' Enable/disable Check mode in Grbl
        ' Just send a $C, this toggles Check state in Grbl
        grblPort.sendData("$C")
    End Sub

    Private Sub btnFileGroup_Click(sender As Object, e As EventArgs) Handles btnFileSend.Click, btnFileSelect.Click, btnFilePause.Click, btnFileStop.Click, _
                                    btnFileReload.Click
        ' This event handler deals with the gcode file related buttons
        ' Implements a simple state machine to keep user from clicking the wrong buttons
        ' Uses button.tag instead of .text so the text doesn't mess up the images on the buttons
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
                    lblTotalLines.Text = gcode.lineCount.ToString

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
                gcode.sendGcodeFile()

                btnFileSelect.Enabled = False
                btnFileSend.Enabled = False
                btnFilePause.Enabled = True
                btnFileStop.Enabled = True
                btnFileReload.Enabled = False

            Case "Pause"
                gcode.sendGCodeFilePause()

                btnFileSelect.Enabled = False
                btnFileSend.Tag = "Resume"
                btnFileSend.Enabled = True
                btnFilePause.Enabled = False
                btnFileStop.Enabled = True
                btnFileReload.Enabled = False

            Case "Stop"
                gcode.sendGCodeFilePause()
                gcode.closeGCodeFile()
                ' Re-enable manual control
                setSubPanels("Idle")

                btnFileSelect.Enabled = True
                btnFileSend.Tag = "Send"
                btnFileSend.Enabled = False
                btnFilePause.Enabled = False
                btnFileStop.Enabled = False
                btnFileReload.Enabled = True

            Case "Resume"
                gcode.sendGCodeFileResume()

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
                lblTotalLines.Text = gcode.lineCount.ToString

                btnFileSelect.Enabled = True    ' Allow changing your mind about the file
                btnFileSend.Enabled = True
                btnFilePause.Enabled = False
                btnFileStop.Enabled = False
                btnFileReload.Enabled = False


        End Select
    End Sub

End Class