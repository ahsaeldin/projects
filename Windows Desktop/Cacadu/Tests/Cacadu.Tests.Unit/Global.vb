Public Module GlobalRoutines

    Public Sub PrepareCacaduComponents()
        Balora.Settings.CAG = Balora.Serializer.BinarySerializer.MD5SerializedObject("{38D4C0D0-8E45-40E6-B40A-67AAFFE264E1}")
        Cacadore.Settings.CAG = Balora.Serializer.BinarySerializer.MD5SerializedObject("{0A70981F-B7D2-4E25-8379-8EBAB7BC9A96}")
        Cacadore.Settings.CrudObject = New TecDAL.TypedDataSetCrud
        Balora.Settings.LogFileName = "CacaduErrorLog.txt"
    End Sub

End Module
