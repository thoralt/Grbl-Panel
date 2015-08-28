Public Class GrblQueueItem

#Region "Enums"
    Enum ItemStatus
        waiting
        sent
        acknowledged
    End Enum
#End Region

#Region "Private variables"
    Private _status As ItemStatus
    Private _text As String
#End Region

#Region "Properties"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Property Status() As ItemStatus
        Get
            Return _status
        End Get
        Set(value As ItemStatus)
            _status = value
        End Set
    End Property

    ''' <summary>
    ''' 
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

    ''' <summary>
    ''' This calculates the length of the stored text including line feed (which will be added when sending the command).
    ''' </summary>
    ''' <returns></returns>
    ReadOnly Property Length() As Integer
        Get
            Return _text.Length() + 1
        End Get
    End Property
#End Region

#Region "Constructor"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="NewText"></param>
    ''' <param name="NewStatus"></param>
    Public Sub New(NewText As String, NewStatus As ItemStatus)
        Status = NewStatus
        Text = NewText
    End Sub
#End Region

End Class
