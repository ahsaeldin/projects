'الاحداث الموجودة في ملف الماك نوعين منفصلين تماماً
'النوع الأول هي تلك الاحداث التي يتم إلتقاطها من تحركات
'المستخدم أثناء تسجيل الماكرو وهم احداث الماوس والكي بورد
'وعنصر الإنتظار وأخيراً عنصر الباك جروند يا معلم .
'النوع الثاني من الإحداث هو العناصر الجديدة التي يتم إضافتها
'من شريط الأدوات خاصة الكيوت ماكرو والذي يضيف العناصر إلي الليست 
'ويتم إنعكاسها علي البلاي باك ولااحتاج ان يتم اضافتها علي كلاس الماكرو ايفنت الشهيرة
'وبالتالي أنت لن تضيف أي احداث جديدة إلي الماكرو ايفنت كلاس
'باستثناء عجلة الماوس و الباك جروند يا معلم

Imports System.Windows.Forms

Friend Class frmMoreItems

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
