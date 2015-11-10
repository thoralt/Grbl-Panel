Public Class GrblQueueItem

#Region "Enums"
    Public Enum ItemState
        ''' <summary>
        ''' The state is undefined (not yet initialized).
        ''' </summary>
        undefined

        ''' <summary>
        ''' The item has not yet been sent to grbl.
        ''' </summary>
        waiting

        ''' <summary>
        ''' The item has been sent to grbl's internal buffer and is now waiting to be executed.
        ''' </summary>
        sent

        ''' <summary>
        ''' Grbl has acknowledged this item because the corresponding command has been executed.
        ''' </summary>
        acknowledged
    End Enum
#End Region

#Region "Private variables"
    Private _parent As GrblQueue
#End Region

#Region "Properties"
    ''' <summary>
    ''' The position of the queue item in the queue
    ''' </summary>
    ''' <returns></returns>
    Property Index() As Integer
        Get
            Return _index
        End Get
        Set(value As Integer)
            _index = value
        End Set
    End Property
    Private _index As Integer

    ''' <summary>
    ''' The queue item state: undefined, waiting, sent, acknowledged
    ''' </summary>
    ''' <returns></returns>
    Property State() As ItemState
        Get
            Return _state
        End Get
        Set(newState As ItemState)
            Dim oldState As ItemState = _state
            _state = newState ' modify before kicking off notification so the internal state change is not delayed
            If newState <> oldState Then
                _parent.queueItemChangedState(Me, oldState, newState)
            End If
        End Set
    End Property
    Private _state As ItemState = ItemState.undefined

    ''' <summary>
    ''' The content (command) of the queue item
    ''' </summary>
    ''' <returns></returns>
    Property Text() As String
        Get
            Return _text
        End Get
        Set(value As String)
            _text = value
            _text = _text.TrimEnd()
        End Set
    End Property
    Private _text As String

    ''' <summary>
    ''' This calculates the length of the stored text including line feed (which will be added when sending the command).
    ''' </summary>
    ''' <returns></returns>
    ReadOnly Property Length() As Integer
        Get
            Return _text.Length() + 1
        End Get
    End Property

    ''' <summary>
    ''' Error flag
    ''' </summary>
    ''' <returns></returns>
    Public Property ErrorFlag As Boolean

#End Region

#Region "Constructor"
    ''' <summary>
    ''' Creates a new instance of GrblQueueItem
    ''' </summary>
    ''' <param name="pText">Content (command)</param>
    ''' <param name="pState">usually ItemState.waiting</param>
    ''' <param name="pIndex">position in the queue</param>
    Public Sub New(pText As String, pState As ItemState, pIndex As Integer, pParent As GrblQueue)
        _parent = pParent
        Text = pText
        Index = pIndex
        State = pState
        ErrorFlag = False
    End Sub
#End Region

End Class
