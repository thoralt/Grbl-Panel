Public Class GrblQueue

#Region "Private fields"
    Private _MAX_GRBL_QUEUE_SIZE As Integer = 127

    Private _port As GrblIF = Nothing
    Private _queue As List(Of GrblQueueItem) = Nothing
    Private _bufferCapacity As Integer
    Private _bufferCapacityLock As New Object
    Private _pause As Boolean = True
    Private _drainQueueAndResume As Boolean = False

    Private _waitingItemsLock As New Object()
    Private _waitingItems As Integer = 0
    Private _sentItemsLock As New Object()
    Private _sentItems As Integer = 0
    Private _acknowledgedItemsLock As New Object()
    Private _acknowledgedItems As Integer = 0

    Private _nextSendOrderLock As New Object()
    Private _nextSendOrder As Integer = 0

    Private _executeImmediateCommandLock As New Object()
    Private _pumpLock As New Object()

    Private _alarmState As Boolean = False

#End Region

#Region "Events"
    ''' <summary>
    ''' Grbl reports an alarm.
    ''' </summary>
    ''' <param name="AlarmText">Alarm description</param>
    Public Event QueueAlarm(AlarmText As String)

    ''' <summary>
    ''' Grbl reports an error.
    ''' </summary>
    ''' <param name="ErrorText">Error description</param>
    Public Event QueueError(ErrorText As String)

    ''' <summary>
    ''' The alarm condition has been reset.
    ''' </summary>
    Public Event QueueResetAlarm()

    ''' <summary>
    ''' The state of an item has changed.
    ''' </summary>
    ''' <param name="item">Item which has changed its state</param>
    ''' <param name="oldState">Previous state</param>
    ''' <param name="newstate">Current state</param>
    Public Event QueueItemStateChanged(item As GrblQueueItem, oldState As GrblQueueItem.ItemState, newstate As GrblQueueItem.ItemState)

    ''' <summary>
    ''' The queue has finished sending all elements. There may still be items in grbl's internal buffer
    ''' which causes additional QueueItemStateChanged events.
    ''' </summary>
    Public Event QueueFinished()

    ''' <summary>
    ''' The queue has been paused.
    ''' </summary>
    ''' <param name="item">The item which will be sent when resuming</param>
    Public Event QueuePaused(item As GrblQueueItem)

    ''' <summary>
    ''' The queue is resuming operation.
    ''' </summary>
    ''' <param name="item">The item which is the first to be transmitted</param>
    Public Event QueueResuming(item As GrblQueueItem)

    ''' <summary>
    ''' The queue encountered a comment line with parantheses "(comment)"
    ''' </summary>
    ''' <param name="comment">The content of the comment</param>
    Public Event QueueComment(comment As String)

    ''' <summary>
    ''' Reports the progress of operation
    ''' </summary>
    ''' <param name="progressValue">Queue progress in %</param>
    Public Event QueueProgress(progressValue As Integer)

    ''' <summary>
    ''' Reports grbl's buffer level
    ''' </summary>
    ''' <param name="level">Grbl's buffer level in %</param>
    Public Event QueueGrblBufferLevel(level As Integer)

    ''' <summary>
    ''' An item has been added to the queue.
    ''' </summary>
    ''' <param name="item">Item which has been added</param>
    Public Event QueueAddedItem(item As GrblQueueItem)
#End Region

#Region "Public methods"
    ''' <summary>
    ''' Creates a new GrblQueue object. There should be only one
    ''' instance of GrblQueueItem using the given GrblIF. A handler is
    ''' registered to receive grbl's output.
    ''' </summary>
    ''' <param name="port">GrblIF to use for communication with grbl</param>
    Public Sub New(port As GrblIF)
        _port = port
        _port.addRcvDelegate(AddressOf Me._receiveData)
        _queue = New List(Of GrblQueueItem)
        ' TODO: Use configurable _MAX_GRBL_QUEUE_SIZE instead of constant
        _bufferCapacity = _MAX_GRBL_QUEUE_SIZE
        _pause = True
    End Sub

    ''' <summary>
    ''' Clears the queue, resets all buffers, raises QueueFinished event. Items which 
    ''' are in grbl's internal queue are NOT cancelled, grbl will continue to execute them.
    ''' </summary>
    Public Sub Clear()
        SyncLock _queue
            _queue.Clear()
        End SyncLock
        SyncLock _bufferCapacityLock
            _bufferCapacity = _MAX_GRBL_QUEUE_SIZE
        End SyncLock
        SyncLock _waitingItemsLock
            _waitingItems = 0
        End SyncLock
        SyncLock _sentItemsLock
            _sentItems = 0
        End SyncLock
        SyncLock _acknowledgedItemsLock
            _acknowledgedItems = 0
        End SyncLock
        _drainQueueAndResume = False
        _pause = True
        RaiseEvent QueueFinished()
    End Sub

    ''' <summary>
    ''' Adds an item to the queue, starts transmission if requested.
    ''' </summary>
    ''' <param name="item">GrblQueueItem to add</param>
    ''' <param name="execute">Starts transmission if True</param>
    Public Sub Enqueue(item As GrblQueueItem, execute As Boolean)
        'Console.WriteLine("GrblQueue: Adding Item " & item.Index & " '" & item.Text & "'")
        SyncLock _queue
            _queue.Add(item)
        End SyncLock

        RaiseEvent QueueAddedItem(item)

        ' trigger queue
        If execute Then Me._receiveData("(!go)")
    End Sub

    ''' <summary>
    ''' Adds an item to the queue, does not start transmission (but does not interrupt
    ''' an ongoing transmission either).
    ''' </summary>
    ''' <param name="item">GrblQueueItem to add</param>
    Public Sub Enqueue(item As GrblQueueItem)
        Enqueue(item, False)
    End Sub

    ''' <summary>
    ''' Adds a new item to the end of the queue and puts it in waiting state,
    ''' starts transmission if requested.
    ''' </summary>
    ''' <param name="text">Content (command)</param>
    ''' <param name="execute">Starts transmission if True</param>
    ''' <param name="lineNumber">Line number or last line if omitted (no check for double entries!)</param>
    Public Sub Enqueue(text As String, execute As Boolean, Optional lineNumber As Integer = -1)
        If lineNumber = -1 Then
            lineNumber = _queue.Count
        End If
        Dim item As New GrblQueueItem(text.Trim(), GrblQueueItem.ItemState.waiting, lineNumber, Me)
        Enqueue(item, execute)
    End Sub

    ''' <summary>
    ''' Adds a new item to the end of the queue and puts it in waiting state.
    ''' Does not start transmission (but does not interrupt an ongoing transmission either).
    ''' </summary>
    ''' <param name="text">Content (command)</param>
    Public Sub Enqueue(text As String)
        Enqueue(text, False)
    End Sub

    ''' <summary>
    ''' Transmits a command to grbl without using the queue. Fails if queue is currently running or
    ''' grbl is in alarm state.
    ''' </summary>
    ''' <param name="text"></param>
    ''' <returns>False if queue is busy or grbl is in alarm state</returns>
    Public Function ExecuteImmediateCommand(text As String) As Boolean
        ' do not continue if queue is currently running
        If Not _pause Then
            Console.WriteLine("GrblQueue: ExecuteImmediateCommand() failed because queue is running.")
            Return False
        End If

        ' do not execute regular commands in alarm state
        'If _alarmState And Not text.StartsWith("$") Then
        '    Console.WriteLine("GrblQueue: ExecuteImmediateCommand() failed because grbl is in alarm state.")
        '    Return False
        'End If

        SyncLock _executeImmediateCommandLock
            ' create high priority item
            Dim item As New GrblQueueItem(text.Trim(), GrblQueueItem.ItemState.waiting, _queue.Count, Me, 100)
            Enqueue(item)
            Console.WriteLine("GrblQueue: ExecuteImmediateCommand() added item: " & item.Description())

            ' repeat _pump() until at least one item has been sent
            ' since we inserted just _one_ item with high priority here, this
            ' guarantees that this item will be transmitted even if we're in pause mode
            Dim timeout As Integer = 1000
            While Not _pump() And timeout >= 0
                System.Threading.Thread.Sleep(10)
                timeout -= 1
            End While

            If timeout <= 0 Then
                Throw New Exception("Timeout while trying to send immediate item " + item.Description() + ".")
            End If
        End SyncLock

        Return True
    End Function

    ''' <summary>
    ''' Pauses the queue. No further GrblQueueItems are transmitted but grbl will continue
    ''' to execute all items in its internal buffer. Output from grbl (acknowledgements, 
    ''' errors, alarms) will be processed even if paused. Raises QueuePaused event.
    ''' </summary>
    Public Sub Pause()
        If Not _pause Then
            _pause = True
            RaiseEvent QueuePaused(_firstWaitingItem())
        End If
    End Sub

    ''' <summary>
    ''' Resumes transmission, raises QueueResuming event.
    ''' </summary>
    Public Sub [Resume]()
        If _pause Then
            _pause = False
            RaiseEvent QueueResuming(_firstWaitingItem())

            ' trigger next queue item
            _receiveData("(!go)")
        End If
    End Sub

    Public Function isRunning() As Boolean
        Return Not _pause
    End Function

    Public Function numberOfLines() As Integer
        Return _queue.Count
    End Function
#End Region

#Region "Private methods"

    ''' <summary>
    ''' Gets an ascending index used for ordering transmissions.
    ''' </summary>
    ''' <returns></returns>
    Private Function _getNextSendOrder() As Integer
        SyncLock _nextSendOrderLock
            _nextSendOrder += 1
            Return _nextSendOrder
        End SyncLock
    End Function

    ''' <summary>
    ''' Finds the first item in the queue which is in the ItemState.waiting state.
    ''' </summary>
    ''' <returns>The first waiting GrblQueueItem</returns>
    Private Function _firstWaitingItem() As GrblQueueItem
        ' find first waiting item
        Dim queueCopy As List(Of GrblQueueItem)
        SyncLock _queue
            queueCopy = New List(Of GrblQueueItem)(_queue)
        End SyncLock
        For Each item As GrblQueueItem In queueCopy
            If item.State = GrblQueueItem.ItemState.waiting Then Return item
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' Callback for queue items, is called whenever an item changes its state, 
    ''' helps to track item state counts without constantly iterating over the
    ''' whole queue. This is only public so it can be called by any GrblQueueItem,
    ''' it still is a private method not to be called from outside.
    ''' </summary>
    ''' <param name="item">GrblQueueItem which changed its state</param>
    ''' <param name="oldState">state before change</param>
    ''' <param name="newState">state after change</param>
    Public Sub _queueItemChangedState(item As GrblQueueItem, oldState As GrblQueueItem.ItemState, newState As GrblQueueItem.ItemState)
        Select Case oldState
            Case GrblQueueItem.ItemState.waiting
                SyncLock _waitingItemsLock
                    _waitingItems -= 1
                    If _waitingItems < 0 Then _waitingItems = 0 ' can happen if Clear() was called and grbl is still executing
                End SyncLock
            Case GrblQueueItem.ItemState.sent
                SyncLock _sentItemsLock
                    _sentItems -= 1
                    If _sentItems < 0 Then _sentItems = 0 ' can happen if Clear() was called and grbl is still executing
                End SyncLock
            Case GrblQueueItem.ItemState.acknowledged
                SyncLock _acknowledgedItemsLock
                    _acknowledgedItems -= 1
                    If _acknowledgedItems < 0 Then _acknowledgedItems = 0 ' can happen if Clear() was called and grbl is still executing
                End SyncLock
        End Select

        Select Case newState
            Case GrblQueueItem.ItemState.waiting
                SyncLock _waitingItemsLock
                    _waitingItems += 1
                End SyncLock
            Case GrblQueueItem.ItemState.sent
                SyncLock _sentItemsLock
                    _sentItems += 1
                End SyncLock
            Case GrblQueueItem.ItemState.acknowledged
                SyncLock _acknowledgedItemsLock
                    _acknowledgedItems += 1
                End SyncLock
        End Select

        RaiseEvent QueueItemStateChanged(item, oldState, newState)

        'Console.WriteLine("GrblQueue: w:" & _waitingItems & " s:" & _sentItems & " a:" & _acknowledgedItems)
    End Sub

    ''' <summary>
    ''' Search for the first item in the queue that has been sent but not yet acknowledged and acknowledge it
    ''' </summary>
    Private Sub _acknowledgeNextItem(err As Boolean)

        ' find next item which has been sent based on transmission order field
        ' this ensures high priority orders to be acknowledged out of order
        Dim itemToAcknowledge As GrblQueueItem = Nothing

        SyncLock _queue
            For Each item As GrblQueueItem In _queue
                If item.State = GrblQueueItem.ItemState.sent AndAlso (itemToAcknowledge Is Nothing OrElse item.Order < itemToAcknowledge.Order) Then
                    itemToAcknowledge = item
                End If
            Next

            If itemToAcknowledge Is Nothing Then
                Console.WriteLine("Trying to acknowledge next item but did not find any candidate.")
                Return
            End If

            ' acknowledge this item
            SyncLock itemToAcknowledge
                itemToAcknowledge.State = GrblQueueItem.ItemState.acknowledged
                itemToAcknowledge.ErrorFlag = err
            End SyncLock

            ' increase free buffer capacity
            SyncLock _bufferCapacityLock
                _bufferCapacity += itemToAcknowledge.Length
                If _bufferCapacity > _MAX_GRBL_QUEUE_SIZE Then _bufferCapacity = _MAX_GRBL_QUEUE_SIZE
            End SyncLock

            'Console.WriteLine("GrblQueue: Acknowledging item " & itemToAcknowledge.Description() & ", freeing " &
            '                  itemToAcknowledge.Length & " bytes, now capacity=" & _bufferCapacity)

            RaiseEvent QueueProgress(100 * itemToAcknowledge.Index / _queue.Count)
            RaiseEvent QueueGrblBufferLevel(100 * (_MAX_GRBL_QUEUE_SIZE - _bufferCapacity) / _MAX_GRBL_QUEUE_SIZE)
        End SyncLock
    End Sub

    ''' <summary>
    ''' Gets the next item to be transmitted. Items with higher priority come first.
    ''' </summary>
    ''' <returns>Next GrblQueueItem to be transmitted</returns>
    Private Function _getNextWaitingItem() As GrblQueueItem
        Dim nextItem As GrblQueueItem = Nothing
        Dim queueCopy As List(Of GrblQueueItem)
        SyncLock _queue
            ' copy it inside SyncLock because external thread possibly
            ' changes _queue during cloning
            queueCopy = New List(Of GrblQueueItem)(_queue)
        End SyncLock

        ' iterate over all items and find first item in waiting state 
        ' with the highest priority
        For Each item As GrblQueueItem In queueCopy
            If item.State = GrblQueueItem.ItemState.waiting AndAlso (nextItem Is Nothing OrElse item.Priority > nextItem.Priority) Then
                nextItem = item
            End If
        Next

        Return nextItem
    End Function

    ''' <summary>
    ''' Send next command to grbl
    ''' </summary>
    Private Function _pump() As Boolean

        SyncLock _pumpLock

            Dim didSendItems As Boolean = False

            ' check if there are items to be sent
            If _waitingItems = 0 Then
                Return False
            End If

            If _drainQueueAndResume Then
                Console.WriteLine("_pump: _drainQueueAndResume=True, _bufferCapacity=" & _bufferCapacity)
            End If

            ' do not continue if waiting for EEPROM write result and queue is still filled:
            ' Wait until GRBL acknowledged all commands and buffer is completely empty
            If _drainQueueAndResume And _bufferCapacity < _MAX_GRBL_QUEUE_SIZE Then Return False

            ' queue is now empty, EEPROM write is completed and we can continue
            _drainQueueAndResume = False

            Dim bufferIsFull As Boolean = False
            Do
                Dim item As GrblQueueItem = _getNextWaitingItem()
                If item Is Nothing Then
                    'Throw New Exception("No item with status GrblQueueItem.ItemState.waiting available, but internal counter _waitingItems=" & _waitingItems & ".")
                    Return didSendItems
                End If

                SyncLock item ' lock this particular GrblQueueItem

                    If item.Length <= _bufferCapacity Then
                        ' command fits into buffer -> send it to grbl
                        SyncLock _bufferCapacityLock
                            _bufferCapacity -= item.Length
                        End SyncLock
                        item.Order = _getNextSendOrder()
                        _port.sendData(item.Text)
                        item.State = GrblQueueItem.ItemState.sent
                        didSendItems = True

                        ' Is this a comment? If so, send message containing comment text to GUI.
                        If item.Text.StartsWith("(") Then
                            Dim comment As String = item.Text.Trim()
                            comment = comment.Substring(1, comment.Length - 2)
                            RaiseEvent QueueComment(comment)
                        End If

                        'Console.WriteLine("GrblQueue: sending item " & item.Index & " '" & item.Text & "' " & item.Length & " bytes, capacity=" & _bufferCapacity)

                        RaiseEvent QueueGrblBufferLevel(100 * (_MAX_GRBL_QUEUE_SIZE - _bufferCapacity) / _MAX_GRBL_QUEUE_SIZE)

                        ' check for commands causing grbl to write to the internal EEPROM
                        If item.Text.Contains("G10") Or
                                item.Text.Contains("G28.1") Or
                                item.Text.Contains("G30.1") Or
                                item.Text.Contains("$") Then
                            ' if EEPROM command is found: drain the queue, do not continue to
                            ' stream data until grbl finished this command
                            ' reason: grbl disables interrupts during EEPROM write and is not
                            ' able to receive characters, this would lead to missing bytes and
                            ' corrupted commands
                            _drainQueueAndResume = True
                            Console.WriteLine("Detected EEPROM write, draining queue.")
                        End If
                    Else
                        ' next command to be sent does _not_ fit into buffer
                        ' -> stop here and try again later when more commands
                        ' have been acknowledged
                        'Console.WriteLine("GrblQueue: _pump() could not send '" & item.Text & "' " & item.Length & " bytes, buffer is full")
                        bufferIsFull = True
                    End If
                End SyncLock

                ' continue to send items until the next command does not fit into the buffer
            Loop Until bufferIsFull Or _drainQueueAndResume

            Return didSendItems
        End SyncLock
    End Function

    ''' <summary>
    ''' Callback for data sent from grbl
    ''' </summary>
    ''' <param name="data">One line of data from grbl</param>
    Private Sub _receiveData(ByVal data As String)

        data = data.ToUpper().Trim()
        If Not data.StartsWith("<") And Not data.StartsWith("$") Then
            Console.WriteLine("GrblQueue: received '" & data & "'")
        End If

        ' check grbl response
        If data = "(!GO)" Then
            ' this is only to trigger transmission in case _pump() was idle
            _pump()

        ElseIf data.StartsWith("OK") Then
            ' if the response was "OK", then acknowledge the next item in
            ' the buffer and try to send more data
            _acknowledgeNextItem(False)

            ' pump next item if not paused
            If Not _pause Then _pump()

        ElseIf data.StartsWith("ERROR")
            _acknowledgeNextItem(True)
            RaiseEvent QueueError(data.Substring(7))

        ElseIf data.StartsWith("ALARM") Then
            ' this occurs directly when an alarm happens

            _acknowledgeNextItem(False)
            _alarmState = True
            _pause = True

            ' report alarm to GUI (will disable/enable buttons and display alarm message)
            RaiseEvent QueueAlarm(data.Substring(7))
            RaiseEvent QueueProgress(0)

        ElseIf data.StartsWith("<ALARM") Then
            ' this occurs if the periodical "get status" command ("?") shows a previous
            ' alarm, grbl is in alarm state
            _alarmState = True

        ElseIf data.StartsWith("<IDLE") Or
            data.StartsWith("<RUN") Or
            data.StartsWith("<HOLD") Or
            data.StartsWith("<DOOR") Or
            data.StartsWith("<HOME") Or
            data.StartsWith("<CHECK") Then

            ' reset alarm condition because grbl seems to be doing something useful
            If _alarmState Then
                _alarmState = False
                RaiseEvent QueueResetAlarm()
            End If

        End If

        ' check if queue is empty and raise event if necessary
        If _waitingItems = 0 Then
            If Not _pause Then
                RaiseEvent QueueFinished()
                RaiseEvent QueueProgress(100)
            End If
            _pause = True
            Return
        End If

    End Sub
#End Region

End Class
