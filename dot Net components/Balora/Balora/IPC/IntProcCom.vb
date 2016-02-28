Imports System.Runtime.InteropServices
Imports System.Windows.Forms

'//My internal process communications class
Friend Class IntProcCom
    Inherits BaloraBase

    'To use this method you need to uncomment Win32API.vb
    ''//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Friend Function SendStringToAprocess(ByVal text As String) As Integer
        '    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
        Try
            '        Dim CMHwnd As IntPtr = FindWindow(vbNullString, "Perfect Macro Recorder")
            '        'get string from textbox
            '        Dim str As String = text
            '        'string to byte array
            '        Dim B() As Byte = System.Text.Encoding.Default.GetBytes(str)
            '        'allocate memory space for byte array, and get a pointer to it
            '        Dim lpB As IntPtr = Marshal.AllocHGlobal(B.Length)
            '        'copy the byte array into memory  
            '        Marshal.Copy(B, 0, lpB, B.Length)
            '        'setup a standard structure for the WM_COPYDATA message
            '        Dim CD As COPYDATASTRUCT
            '        With CD
            '            .dwData = CType(0, IntPtr) 'can be used for custom indexing between apps
            '            .cbData = B.Length 'length of data
            '            .lpData = CType(lpB.ToInt32, IntPtr) 'pointer to the data
            '        End With
            '        'clean up array
            '        Erase B
            '        'allocate memory space for structure, and get a pointer to it
            '        Dim lpCD As IntPtr = Marshal.AllocHGlobal(CD.cbData)
            '        'copy structure to allocated memory place
            '        Marshal.StructureToPtr(CD, lpCD, False)

            '        'send message
            '        'SendMessage(CMHwnd, WM_COPYDATA, 0, lpCD.ToInt32)
            '        Dim Res As Integer
            '        Dim SendingResult As Integer
            '        Res = SendMessageTimeout(CMHwnd, WM_COPYDATA, 0, lpCD.ToInt32, SMTO_ABORTIFHUNG, 3000, SendingResult)

            '        Return Res

        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Alerter.REP("Error Sending Path to the other instance", ex, True)
            Return 0
        End Try

        Return Nothing
    End Function

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Shared Sub CopySharedDataToClipboard(ByVal format As String, ByVal [string] As String)
        '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
        'الوظيفة دي ووظيفة
        'PasteDataFromClipboard
        'عشان النسخ واللصق من وإلي الكيب بورد بتنسيقات معينة
        'مما يعنني أولا علي ارسال رسائل بين حالات البرنامج المختلفة
        'لتفادي استخدام
        'WM_COPYDATA 
        'في وظيفة
        'MyApplication_Startup
        'التي تسخدم في إرسال مسار ملف السي ماك إلي النسخة التي تعمل
        'فعلياً من البرنامج
        ' Creates a new data format.
        Dim CDataFormat As DataFormats.Format = DataFormats.GetFormat(format)
        ' Creates a new object and store it in a DataObject using myFormat 
        ' as the type of format. 
        Dim CopiedText As String = [string]
        Dim DataObject As New DataObject(CDataFormat.Name, CopiedText)
        'لازم تاني بارامتر يبقي ترو عشان تقدر تنقل محتويات الكليب بوورد بين برنامجين
        Clipboard.SetDataObject(DataObject, True)
    End Sub

    Public Shared Function GetSharedDataFromClipboard(ByVal format As String) As String
        'الوظيفة دي ووظيفة
        'PasteDataFromClipboard
        'عشان النسخ واللصق من وإلي الكيب بورد بتنسيقات معينة
        'مما يعنني أولا علي ارسال رسائل بين حالات البرنامج المختلفة
        'لتفادي استخدام
        'WM_COPYDATA 
        'في وظيفة
        'MyApplication_Startup
        'التي تسخدم في إرسال مسار ملف السي ماك إلي النسخة التي تعمل
        'فعلياً من البرنامج
        ' Creates a new data format.
        Dim CMFormat As DataFormats.Format = DataFormats.GetFormat(format)
        ' Performs some processing steps.
        ' Retrieves the data from the clipboard.
        Dim RetrievedItemsObject As IDataObject = Clipboard.GetDataObject()
        ' Converts the IDataObject type to MyNewObject type. 
        Dim CopiedItems As String = CType(RetrievedItemsObject.GetData(CMFormat.Name), String)
        GetSharedDataFromClipboard = CopiedItems
    End Function

End Class
