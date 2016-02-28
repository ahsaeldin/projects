'Balora version:0.1  .. 06/11/2010
' Contains a different sets to deal with resources from reading, writing and creating resources files .
Option Strict On
Option Explicit On

#Region "Imported NameSpaces"

Imports System.Resources
Imports System.Reflection

#End Region

Namespace Balora
    Partial Friend Class xResources

#Region "Methods"

        Public Shared Function WriteXMLResource(ByVal generatedResFilePath As String, ByVal resources As Hashtable) As Boolean

            ' Generate a resource file and save it to the path passed in parameters .
            ' And the resources passed in resources hashtable (great idea) .
            ' Currently it support strings, bitmaps, and custom files (Byte()) .
            ' And I didn't tested it for other types of resources yet.
            Dim ResWriter As New ResXResourceWriter(generatedResFilePath)
            Dim resourceEntry As ResXDataNode = Nothing

            Try
                For Each key As String In resources.Keys
                    Select Case resources(key).GetType().ToString()
                        Case "System.String"
                            resourceEntry = Nothing
                            resourceEntry = New ResXDataNode(key, CType(resources(key), String))
                            ResWriter.AddResource(resourceEntry)
                        Case "System.Drawing.Bitmap"
                            resourceEntry = Nothing
                            resourceEntry = New ResXDataNode(key, CType(resources(key), Bitmap))
                            ResWriter.AddResource(resourceEntry)
                        Case "System.Byte[]"
                            resourceEntry = Nothing
                            resourceEntry = New ResXDataNode(key, CType(resources(key), Byte()))
                            ResWriter.AddResource(resourceEntry)
                    End Select
                Next
                ResWriter.Generate()
                ResWriter.Close()
                Return True
            Catch ex As Exception
                Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error generating macro data.")
                Return False
            End Try

        End Function

        Public Shared Function WriteResource(ByVal generatedResFilePath As String, ByVal resources As Hashtable) As Boolean
            'You pass the resource file name or path like this project1.myResources
            'and a hashtable of the the resources you want to save.

            ' Generate a resource file and save it to the path passed in parameters .
            ' And the resources passed in resources hashtable (great idea) .
            ' Currently it support strings, bitmaps, and custom files (Byte()) .
            ' And I didn't tested it for other types of resources yet.
            ' Creates a resource writer.
            Dim ResWriter As New ResourceWriter(generatedResFilePath)
            Try
                For Each key As String In resources.Keys
                    Select Case resources(key).GetType().ToString()
                        Case "System.String"
                            ResWriter.AddResource(key, CType(resources(key), String))
                        Case "System.Drawing.Icon"
                            ResWriter.AddResource(key, CType(resources(key), Icon))
                        Case "System.Byte[]"
                            ResWriter.AddResource(key, CType(resources(key), Byte()))
                    End Select
                Next
                ResWriter.Generate()
                ResWriter.Close()
                Return True
            Catch ex As Exception
                Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error generating macro data.")
                Return False
            End Try

        End Function

        Public Shared Function ReadResourceItem(ByVal resourceLocation As String, ByVal itemName As String) As Object
            'You pass the resource file name or path like this project1.myResources
            'followed by the resource item you want to retrieve.
            'Note that you must know the object type you are retrieving in order to
            'convert to it like this for example: icon = CType(returnedObjByReadResourceItem, System.Drawing.Icon)
            Dim reader As New ResourceReader(resourceLocation)
            Dim en As IDictionaryEnumerator = reader.GetEnumerator()
            ' Goes through the enumerator, printing out the key and value pairs.
            While en.MoveNext()
                If en.Key.ToString = itemName Then
                    ReadResourceItem = en.Value
                    Exit Function
                    'This an exmple of how to check the item object type
                    Select Case en.Value.GetType().ToString()
                        Case "System.String"
                            MsgBox("String")
                        Case "System.Drawing.Icon"
                            MsgBox("Icon")
                        Case "System.Byte[]"
                            MsgBox("Byte")
                    End Select
                End If
            End While
            reader.Close()
            Return Nothing

        End Function

        Public Shared Function ReadEmbeddedResourceItemAtRuntime(ByVal resourceNameWithoutExtention As String, ByVal itemName As String, Optional ByVal fromAssemply As Assembly = Nothing) As Object

            Dim ResourceManager As Global.System.Resources.ResourceManager = Nothing

            If fromAssemply Is Nothing Then
                ResourceManager = New Global.System.Resources.ResourceManager(resourceNameWithoutExtention, Assembly.GetExecutingAssembly())
            Else
                ResourceManager = New Global.System.Resources.ResourceManager(resourceNameWithoutExtention, fromAssemply)
            End If

            Dim obj As Object = ResourceManager.GetObject(itemName)

            Return obj

        End Function

#End Region

    End Class
End Namespace

