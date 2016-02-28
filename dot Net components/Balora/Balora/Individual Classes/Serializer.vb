Option Strict Off
Option Explicit On

Imports System.IO
Imports System.Xml.Serialization
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Windows.Forms

Namespace Serializer
    ''' <summary>
    ''' Several methods of Binary Serialization.
    ''' </summary>
    Public Class BinarySerializer
        Inherits BaloraBase

        Public Overloads Shared Function SerializeObject(ByVal obj As Object, ByVal saveToPath As String) As Boolean
            Try
                Dim bFormatter As New BinaryFormatter
                Dim fStream As New FileStream(saveToPath, FileMode.Create)
                bFormatter.Serialize(fStream, obj)
                fStream.Close()
                fStream.Dispose()
                fStream = Nothing
                bFormatter = Nothing
                Return True
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Alerter.REP("Error during serializing object.", ex, True)
                Return False
            End Try
        End Function

        Public Overloads Shared Function SerializeObject(ByVal obj As Object, ByRef saveToMemoryStream As MemoryStream) As Boolean
            Try
                Dim bFormatter As New BinaryFormatter
                bFormatter.Serialize(saveToMemoryStream, obj)
                bFormatter = Nothing
                Return True
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Alerter.REP("Error during serializing object.", ex, True)
                Return False
            End Try
        End Function

        Public Overloads Shared Function DeserializeObject(ByVal readFromPath As String, Optional ByRef IsFailed As Boolean = False) As Object
            'ex:
            'CacaduErrorReporter = CType(Util.DeserializeObject("c:\ahmed.dat"), ErrorReporter)
            Dim obj As Object = Nothing
            Dim fStream = New FileStream(readFromPath, FileMode.Open)
            Dim BinFormatter As BinaryFormatter = New BinaryFormatter
            Try
                obj = BinFormatter.Deserialize(fStream)
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Alerter.REP("Object file is corrupted, error restoring a serialized object.", ex, True)
                IsFailed = True
            End Try
            fStream.Close()
            fStream.Dispose()
            fStream = Nothing
            BinFormatter = Nothing
            Return obj
        End Function

        Public Overloads Shared Function DeserializeObject(ByVal bArray As Byte()) As Object
            Dim obj As Object = Nothing
            Dim memStream As New MemoryStream(bArray)
            Dim BinFormatter As BinaryFormatter = New BinaryFormatter
            Try
                obj = BinFormatter.Deserialize(memStream)
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Alerter.REP("Error : Restoring a serialized object.", ex, True)
                Return Nothing
            End Try
            memStream.Close()
            memStream.Dispose()
            memStream = Nothing
            BinFormatter = Nothing
            Return obj
        End Function

        Public Shared Sub SerializeListView(ByVal lstView As ListView, ByVal saveToPath As String)
            Try
                Dim BinFormatter As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
                Dim FS As New System.IO.FileStream(saveToPath, IO.FileMode.Create)
                BinFormatter.Serialize(FS, New ArrayList(lstView.Items))
                FS.Flush()
                FS.Dispose()
                FS = Nothing
                BinFormatter = Nothing
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Alerter.REP("Error during serializing listView.", ex, True)
            End Try
        End Sub

        Public Shared Sub DeserializeListView(ByVal lstView As ListView, ByVal readFromPath As String)
            Try
                Dim BinFormatter As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
                Dim FS As New System.IO.FileStream(readFromPath, IO.FileMode.Open)
                lstView.Items.AddRange(BinFormatter.Deserialize(FS).ToArray(GetType(ListViewItem)))
                If lstView.Items.Count = 0 Then
                    MsgBox("Nothing to load; the macro file is empty." _
                    , MsgBoxStyle.Critical, "Perfect Macro Recorder")
                End If
                FS.Dispose()
                FS = Nothing
                BinFormatter = Nothing
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Alerter.REP("Error : Deserializing list view.", ex, True)
            End Try
        End Sub

        Public Shared Function SerializedArrayList(ByVal arrList As ArrayList, ByVal saveToPath As String) As Boolean
            Try
                Dim bFormatter As New BinaryFormatter
                Dim fStream As New FileStream(saveToPath, FileMode.Create)
                bFormatter.Serialize(fStream, arrList)
                fStream.Close()
                fStream.Dispose()
                fStream = Nothing
                bFormatter = Nothing
                Return True
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Alerter.REP("Error during serializing ArrayList", ex, True)
                Return False
            End Try
        End Function

        Public Shared Function DeserializeArrayList(ByVal readFromPath As String, Optional ByRef IsFailed As Boolean = False) As ArrayList
            Dim RestoredEventsList As ArrayList = Nothing
            Dim fStream = New FileStream(readFromPath, FileMode.Open)
            Dim BinFormatter As BinaryFormatter = New BinaryFormatter
            Try
                RestoredEventsList = CType(BinFormatter.Deserialize(fStream), ArrayList)
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Alerter.REP("Error : Restoring a serialized ArrayList.", ex, True)
                IsFailed = True
            End Try
            fStream.Close()
            fStream.Dispose()
            DeserializeArrayList = RestoredEventsList
        End Function

        ''' <summary>
        ''' Take a serializable object, serialize it and generate md5 hash from it.
        ''' </summary>
        ''' <param name="obj">A serializable object</param>
        ''' <returns>MD5 hash of the serialized object</returns>
        ''' <remarks></remarks>
        Public Shared Function MD5SerializedObject(ByVal obj As Object) As String
            Dim md5Hasher As New Encryption.Hash(Encryption.Hash.Provider.MD5)
            Dim dataToBeEncrypted As New Encryption.Data(Util.ObjectToByteArray(obj))
            md5Hasher.Calculate(dataToBeEncrypted)
            dataToBeEncrypted = Nothing
            Dim res = md5Hasher.Value.ToHex
            md5Hasher = Nothing
            Return res
        End Function
    End Class

    ''' <summary>
    ''' A class for XML Serialization.
    ''' </summary>
    Public Class XMLSerialr
        Inherits BaloraBase
        Public Shared Function SerializeObject(ByVal obj As Object, ByVal saveToPath As String) As Boolean
            Try
                'Serialize object to a text file.
                Dim objStreamWriter As New StreamWriter(saveToPath)
                Dim x As New XmlSerializer(obj.GetType)
                x.Serialize(objStreamWriter, obj)
                objStreamWriter.Close()
                Return True
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Alerter.REP("Error during xml serializing object", ex, True)
                Return False
            End Try
        End Function
    End Class
End Namespace


