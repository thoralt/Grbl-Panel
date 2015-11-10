Public Class GrblQueue

#Region "Private variables"
    Private _MAX_GRBL_QUEUE_SIZE As Integer = 127

    Private _port As GrblIF = Nothing
    Private _queue As List(Of GrblQueueItem) = Nothing
    Private _bufferCapacity As Integer
    Private _bufferCapacityLock As New Object
    Private _pause As Boolean = True
    Private _drainQueueAndResume As Boolean = False

    Private _waitingItems As Integer = 0
    Private _sentItems As Integer = 0
    Private _acknowledgedItems As Integer = 0
    Private _waitingItemsLock As New Object()
    Private _sentItemsLock As New Object()
    Private _acknowledgedItemsLock As New Object()
#End Region

#Region "Events"
    Public Event QueueItemStateChanged(item As GrblQueueItem, oldState As GrblQueueItem.ItemState, newstate As GrblQueueItem.ItemState)
    Public Event QueueFinished()
    Public Event QueuePaused(item As GrblQueueItem)
    Public Event QueueResuming(item As GrblQueueItem)
    Public Event QueueComment(comment As String)
    Public Event QueueProgress(progressValue As Integer)
    Public Event QueueGrblBufferLevel(level As Integer)
#End Region

#Region "Public functions"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="port"></param>
    Public Sub New(port As GrblIF)
        _port = port
        _port.addRcvDelegate(AddressOf Me._receiveData)
        _queue = New List(Of GrblQueueItem)
        ' TODO: Use configurable _MAX_GRBL_QUEUE_SIZE instead of constant
        _bufferCapacity = _MAX_GRBL_QUEUE_SIZE
        _pause = True
    End Sub

    Public Sub Clear()
        SyncLock _queue
            _queue.Clear()
        End SyncLock
        SyncLock _bufferCapacityLock
            _bufferCapacity = _MAX_GRBL_QUEUE_SIZE
        End SyncLock
        _drainQueueAndResume = False
        _pause = True
        RaiseEvent QueueFinished()
    End Sub

    ''' <summary>
    ''' Callback for queue items, is called whenever an item changes its state, 
    ''' helps to track item state counts without constantly iterating over the
    ''' whole queue
    ''' </summary>
    ''' <param name="item">item which changed its state</param>
    ''' <param name="oldState">state before change</param>
    ''' <param name="newState">state after change</param>
    Public Sub queueItemChangedState(item As GrblQueueItem, oldState As GrblQueueItem.ItemState, newState As GrblQueueItem.ItemState)
        Select Case oldState
            Case GrblQueueItem.ItemState.waiting
                SyncLock _waitingItemsLock
                    _waitingItems -= 1
                End SyncLock
            Case GrblQueueItem.ItemState.sent
                SyncLock _sentItemsLock
                    _sentItems -= 1
                End SyncLock
            Case GrblQueueItem.ItemState.acknowledged
                SyncLock _acknowledgedItemsLock
                    _acknowledgedItems -= 1
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
    ''' Adds an item to the queue, starts transmission if not already transmitting
    ''' </summary>
    ''' <param name="item"></param>
    Public Sub enqueue(item As GrblQueueItem, execute As Boolean)
        'Console.WriteLine("GrblQueue: Adding Item " & item.Index & " '" & item.Text & "'")
        SyncLock _queue
            _queue.Add(item)
        End SyncLock

        ' trigger queue
        If execute Then Me._receiveData("(!go)")
    End Sub

    Public Sub enqueue(item As GrblQueueItem)
        enqueue(item, False)
    End Sub

    ''' <summary>
    ''' Adds a new item to the end of the queue and puts it in waiting state.
    ''' </summary>
    ''' <param name="text">Content (command)</param>
    ''' <param name="execute">trigger queue execution</param>
    Public Sub enqueue(text As String, execute As Boolean, Optional lineNumber As Integer = -1)
        If lineNumber = -1 Then
            lineNumber = _queue.Count
        End If
        Dim item As New GrblQueueItem(text.Trim(), GrblQueueItem.ItemState.waiting, lineNumber, Me)
        enqueue(item, execute)
    End Sub

    ''' <summary>
    ''' Adds a new item to the end of the queue and puts it in waiting state.
    ''' </summary>
    ''' <param name="text">Content (command)</param>
    Public Sub enqueue(text As String)
        enqueue(text, False)
    End Sub

    Private Function firstWaitingItem() As GrblQueueItem
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
    ''' 
    ''' </summary>
    Public Sub pauseSending()
        _pause = True

        RaiseEvent QueuePaused(firstWaitingItem())
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    Public Sub resumeSending()
        _pause = False
        RaiseEvent QueueResuming(firstWaitingItem())

        ' trigger next queue item
        _receiveData("(!go)")
    End Sub
#End Region

#Region "Private functions"
    ''' <summary>
    ''' Search for the first item in the queue that has been sent but not yet acknowledged and acknowledge it
    ''' </summary>
    Private Sub _acknowledgeNextItem(err As Boolean)

        ' find first item which has been sent, but not yet acknowledged
        SyncLock _queue
            For Each item As GrblQueueItem In _queue
                If item.State = GrblQueueItem.ItemState.sent Then
                    ' acknowledge this item
                    item.State = GrblQueueItem.ItemState.acknowledged
                    item.ErrorFlag = err

                    ' increase free buffer capacity
                    SyncLock _bufferCapacityLock
                        _bufferCapacity += item.Length
                    End SyncLock

                    Console.WriteLine("GrblQueue: Acknowledging item " & item.Index & " '" & item.Text &
                                      "', freeing " & item.Length & " bytes, now capacity=" & _bufferCapacity)
                    RaiseEvent QueueProgress(100 * item.Index / _queue.Count)
                    RaiseEvent QueueGrblBufferLevel(100 * (_MAX_GRBL_QUEUE_SIZE - _bufferCapacity) / _MAX_GRBL_QUEUE_SIZE)
                    Return
                End If
            Next
        End SyncLock
        Console.WriteLine("GrblQueue: _acknowledgeNextItem(): No item to acknowledge.")
    End Sub

    ''' <summary>
    ''' Send next command to GRBL
    ''' </summary>
    Private Sub _pump()

        ' check if there are items to be sent
        If _waitingItems = 0 Then
            Return
        End If

        ' create a copy of the queue since _pump() gets called
        ' asynchronously and other threads might enqueue new items
        ' while we are working with the list -> exception "list has changed"
        Dim queueCopy As List(Of GrblQueueItem)
        SyncLock _queue
            ' copy it inside SyncLock because external thread possibly
            ' changes _queue during cloning
            queueCopy = New List(Of GrblQueueItem)(_queue)
        End SyncLock

        ' find next item to be sent
        For Each item As GrblQueueItem In queueCopy

            If _drainQueueAndResume Then
                Console.WriteLine("_drainQueueAndResume=True, _bufferCapacity=" & _bufferCapacity)
            End If

            ' do not continue if waiting for EEPROM write result and queue is still filled:
            ' Wait until buffer is completely empty because GRBL acknowledged the last command
            If _drainQueueAndResume And _bufferCapacity < _MAX_GRBL_QUEUE_SIZE Then Exit For

            ' queue is now empty, EEPROM write is completed and we can continue
            _drainQueueAndResume = False

            SyncLock item ' lock this particular GrblQueueItem

                ' check if this is the first waiting Item
                ' TODO: Should we add a pointer to the next waiting item to eliminate looping through whole list?
                If item.State = GrblQueueItem.ItemState.waiting Then
                    If item.Length <= _bufferCapacity Then
                        ' command fits into buffer -> send it to grbl
                        SyncLock _bufferCapacityLock
                            _bufferCapacity -= item.Length
                        End SyncLock
                        _port.sendData(item.Text)
                        item.State = GrblQueueItem.ItemState.sent

                        ' Is this a comment? If so, send message containing comment text to GUI.
                        If item.Text.StartsWith("(") Then
                            Dim comment As String = item.Text.Trim()
                            comment = comment.Substring(1, comment.Length - 2)
                            RaiseEvent QueueComment(comment)
                        End If

                        Console.WriteLine("GrblQueue: sending item " & item.Index & " '" & item.Text & "' " & item.Length & " bytes, capacity=" & _bufferCapacity)

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
                        Console.WriteLine("GrblQueue: _pump() could not send '" & item.Text & "' " & item.Length & " bytes, buffer is full")
                        Exit For
                    End If
                End If
            End SyncLock
        Next
    End Sub

    ''' <summary>
    ''' Callback for data sent from grbl
    ''' </summary>
    ''' <param name="data">One line of data from grbl</param>
    Private Sub _receiveData(ByVal data As String)

        data = data.ToUpper().Trim()
        If Not data.StartsWith("<") And Not data.StartsWith("$") Then
            Console.WriteLine("GrblQueue: received '" & data & "'")
        End If

        ' check if queue is empty and raise event if necessary
        If _waitingItems = 0 Then
            If Not _pause Then RaiseEvent QueueFinished()
            _pause = True
            Return
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

            ' TODO: Pause on error?
            'Me.pauseSending()
            If Not _pause Then _pump()

        ElseIf (data.StartsWith("<ALARM") Or data.StartsWith("ALARM")) And Not _pause
            ' alarm always stops sending

            ' mark every item as acknowledged so it does
            ' not get sent again when restarting

            ' TODO: marking every item as acknowledged is not very elegant. Suggestions?
            SyncLock _queue
                For Each item As GrblQueueItem In _queue
                    item.State = GrblQueueItem.ItemState.acknowledged
                Next
            End SyncLock
            SyncLock _bufferCapacityLock
                _bufferCapacity = _MAX_GRBL_QUEUE_SIZE
            End SyncLock
            _drainQueueAndResume = False
            _pause = True
            RaiseEvent QueueFinished()
        End If
    End Sub
#End Region

End Class
