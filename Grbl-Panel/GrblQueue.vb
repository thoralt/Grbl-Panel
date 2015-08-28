Public Class GrblQueue

#Region "Private variables"
    Private _MAX_GRBL_QUEUE_SIZE As Integer = 127

    Private _port As GrblIF = Nothing
    Private _queue As List(Of GrblQueueItem) = Nothing
    Private _isSending As Boolean = False
    Private _bufferCapacity As Integer
    Private _pause As Boolean = False
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
    End Sub

    ''' <summary>
    ''' Adds an item to the queue, starts transmission if not already transmitting
    ''' </summary>
    ''' <param name="item"></param>
    Public Sub enqueue(item As GrblQueueItem)
        Console.WriteLine("GrblQueue: Adding Item '" & item.Text & "'")
        _queue.Add(item)

        ' trigger queue if currently idle
        If Not _isSending Then Me._receiveData("ok")
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="text"></param>
    Public Sub enqueue(text As String)
        Dim item As New GrblQueueItem(text, GrblQueueItem.ItemStatus.waiting)
        enqueue(item)
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    Public Sub pauseSending()
        _pause = True
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    Public Sub resumeSending()
        _pause = False

        ' trigger next queue item
        _receiveData("ok")
    End Sub
#End Region

#Region "Private functions"
    ''' <summary>
    ''' Search for the first item in the queue that has been sent but not yet acknowledged and acknowledge it
    ''' </summary>
    Private Sub _acknowledgeNextItem()

        ' find first item which has been sent, but not yet acknowledged
        For Each item As GrblQueueItem In _queue
            If item.Status = GrblQueueItem.ItemStatus.sent Then
                ' acknowledge this item
                item.Status = GrblQueueItem.ItemStatus.acknowledged

                ' increase free buffer capacity
                _bufferCapacity += item.Length

                Exit For
            End If
        Next
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    Private Sub _pump()
        ' find next item to be sent
        For Each item As GrblQueueItem In _queue
            If item.Status = GrblQueueItem.ItemStatus.waiting Then
                If item.Length <= _bufferCapacity Then
                    ' command fits into buffer -> send it to grbl
                    _bufferCapacity -= item.Length
                    _port.sendData(item.Text)
                Else
                    ' next command to be sent does _not_ fit into buffer
                    ' -> stop here and try again later when more commands
                    ' have been acknowledged
                    Exit For
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Callback for data sent from grbl
    ''' </summary>
    ''' <param name="data">One line of data from grbl</param>
    Private Sub _receiveData(ByVal data As String)

        data = data.ToLower()

        ' do not continue if the queue is paused
        If _pause Then Return

        ' check if we are done
        If _queue.Count = 0 Then
            _isSending = False
            Return
        End If

        ' check grbl response
        If data.StartsWith("ok") Then
            ' if the response was "OK", then acknowledge the next item in
            ' the buffer and try to send more data
            _acknowledgeNextItem()
            _pump()

        ElseIf data.StartsWith("error")
            ' error means pause
            _acknowledgeNextItem()
            Me.pauseSending()

        ElseIf data.StartsWith("<alarm") Or data.StartsWith("alarm")
            ' alarm always stops sending
            _isSending = False

            ' mark every item as acknowledged so it does
            ' not get sent again when restarting
            For Each item As GrblQueueItem In _queue
                item.Status = GrblQueueItem.ItemStatus.acknowledged
            Next
        End If
    End Sub
#End Region

End Class
