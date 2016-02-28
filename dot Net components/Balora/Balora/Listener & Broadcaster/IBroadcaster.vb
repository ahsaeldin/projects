Public Interface IBroadcaster
    Sub AddListener(listener As IListener)
    Sub NotifyListeners(message As Object)
    Sub RemoveListener(listener As IListener)
End Interface
