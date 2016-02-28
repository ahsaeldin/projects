Module Themes
    Structure Skins
        Shared Sub RegisterSkins()
            'Enable title bar skinning 
            'More info:-
            'http://documentation.devexpress.com/#WindowsForms/CustomDocument4874
            'http://documentation.devexpress.com/#WindowsForms/DevExpressXtraEditorsXtraForm_HtmlTexttopic
            DevExpress.Skins.SkinManager.EnableFormSkins()
            'For more information, search DevExpress help for 'How to: Use a Skin from an External Skin Library'
            DevExpress.Skins.SkinManager.Default.RegisterAssembly(GetType(DevExpress.UserSkins.BonusSkins).Assembly)
            DevExpress.Skins.SkinManager.Default.RegisterAssembly(GetType(DevExpress.UserSkins.BonusSkins).Assembly)
        End Sub

        Friend Shared Sub RandomlyChangeSkin()
            'To prevent calling skin_changed and get settings over-written.
            'RemoveHandlerForStyleChangedEvent()
            Dim randNum As New Random
            Dim skinArr As ArrayList = Skins.GetRandomSkin
            Dim chosenSkinName As String = CStr(skinArr(randNum.Next(38)))
#If DEBUG Then
            'MainForm.StatusbarLeftContainer.Caption = "Skin Name: " & chosenSkinName
#End If
            SetSkin(chosenSkinName)
            SaveThemeSetting("Random")
        End Sub

        Shared Sub ChangeSkin(skinName As String)
            Select Case skinName
                Case "Random"
                    RandomlyChangeSkin()
                Case "Default"
                    SetSkin("Seven")
                Case Else
                    SetSkin(skinName)
            End Select
        End Sub

        Shared Sub AddHandlerForStyleChangedEvent()
            AddHandler UI.MainForm.MainFormLookAndFeel.LookAndFeel.StyleChanged, AddressOf Skins.Skin_Changed
        End Sub

        Shared Sub RemoveHandlerForStyleChangedEvent()
            RemoveHandler UI.MainForm.MainFormLookAndFeel.LookAndFeel.StyleChanged, AddressOf Skins.Skin_Changed
        End Sub

        Shared Sub ChangeThemeComboCaption(skinName As String)
            UI.MainForm.ThemeBarSubItem.Caption = skinName
        End Sub

        Shared Sub SetSkin(skinName As String)
            ChangeThemeComboCaption(skinName)
            UI.MainForm.MainFormLookAndFeel.LookAndFeel.SkinName = skinName
        End Sub

        Shared Function GetRandomSkin() As ArrayList
            Dim randomArrayList As New ArrayList
            randomArrayList.Add(Lilian) : randomArrayList.Add(DevExpress_Dark_Style) : randomArrayList.Add(iMaginary)
            randomArrayList.Add(Black) : randomArrayList.Add(Blue) : randomArrayList.Add(Coffee)
            randomArrayList.Add(Liquid_Sky) : randomArrayList.Add(London_Liquid_Sky) : randomArrayList.Add(Glass_Oceans)
            randomArrayList.Add(Stardust) : randomArrayList.Add(Xmas_2008_Blue) : randomArrayList.Add(Valentine)
            randomArrayList.Add(McSkin) : randomArrayList.Add(Summer_2008) : randomArrayList.Add(Pumpkin)
            randomArrayList.Add(Dark_Side) : randomArrayList.Add(Springtime) : randomArrayList.Add(Darkroom)
            randomArrayList.Add(Foggy) : randomArrayList.Add(High_Contrast) : randomArrayList.Add(Seven)
            randomArrayList.Add(Seven_Classic) : randomArrayList.Add(Sharp) : randomArrayList.Add(Sharp_Plus)
            randomArrayList.Add(The_Asphalt_World) : randomArrayList.Add(Blueprint) : randomArrayList.Add(Whiteprint)
            randomArrayList.Add(VS2010) : randomArrayList.Add(Office_2007_Blue) : randomArrayList.Add(Office_2007_Black)
            randomArrayList.Add(Office_2007_Silver) : randomArrayList.Add(Office_2007_Green) : randomArrayList.Add(Office_2007_Pink)
            randomArrayList.Add(Office_2010_Blue) : randomArrayList.Add(Office_2010_Black) : randomArrayList.Add(Office_2010_Silver)
            randomArrayList.Add(Caramel) : randomArrayList.Add(Money_Twins) : randomArrayList.Add(Metropolis)
            Return randomArrayList
        End Function

        Shared Sub SaveThemeSetting(themeName As String)
            My.Settings.SN = themeName
            My.Settings.Save()
        End Sub

        Shared Sub OnThemeItemClick(themeName As String)
            'To prevent calling skin_changed and get settings over-written.
            'RemoveHandlerForStyleChangedEvent()
            Select Case themeName
                Case "Random"
                    Skins.RandomlyChangeSkin()
                Case "Default"
                    Skins.ChangeSkin("Seven")
            End Select
            SaveThemeSetting(themeName)
            'AddHandlerForStyleChangedEvent()
        End Sub

        Shared Sub Skin_Changed(sender As Object, e As EventArgs)
            Dim uLookAndFeel As DevExpress.LookAndFeel.Design.UserLookAndFeelDefault = CType(sender, DevExpress.LookAndFeel.Design.UserLookAndFeelDefault)
            Dim skinName As String = uLookAndFeel.ActiveSkinName
            ChangeThemeComboCaption(skinName)
            SaveThemeSetting(skinName)
            'RemoveHandlerForStyleChangedEvent()
        End Sub

        Shared Sub FillThemesCombo()
            DevExpress.XtraBars.Helpers.SkinHelper.InitSkinPopupMenu(UI.MainForm.ThemeBarSubItem)
        End Sub
#Region "Skins Names"
        Shared Lilian As String = "Lilian" : Shared DevExpress_Dark_Style As String = "DevExpress Dark Style"
        Shared iMaginary As String = "iMaginary" : Shared Black As String = "Black"
        Shared Blue As String = "Blue" : Shared Coffee As String = "Coffee"
        Shared Liquid_Sky As String = "Liquid Sky" : Shared London_Liquid_Sky As String = "London Liquid Sky"
        Shared Glass_Oceans As String = "Glass Oceans" : Shared Stardust As String = "Stardust"
        Shared Xmas_2008_Blue As String = "Xmas 2008 Blue" : Shared Valentine As String = "Valentine"
        Shared McSkin As String = "McSkin" : Shared Summer_2008 As String = "Summer 2008"
        Shared Pumpkin As String = "Pumpkin" : Shared Dark_Side As String = "Dark Side"
        Shared Springtime As String = "Springtime" : Shared Darkroom As String = "Darkroom"
        Shared Foggy As String = "Foggy" : Shared High_Contrast As String = "High Contrast"
        Shared Seven As String = "Seven" : Shared Seven_Classic As String = "Seven Classic"
        Shared Sharp As String = "Sharp" : Shared Sharp_Plus As String = "Sharp Plus"
        Shared The_Asphalt_World As String = "The Asphalt World" : Shared Blueprint As String = "Blueprint"
        Shared Whiteprint As String = "Whiteprint" : Shared VS2010 As String = "VS2010"
        Shared Office_2007_Blue As String = "Office 2007 Blue" : Shared Office_2007_Black As String = "Office 2007 Black"
        Shared Office_2007_Silver As String = "Office 2007 Silver" : Shared Office_2007_Green As String = "Office 2007 Green"
        Shared Office_2007_Pink As String = "Office 2007 Pink" : Shared Office_2010_Blue As String = "Office 2010 Blue"
        Shared Office_2010_Black As String = "Office 2010 Black" : Shared Office_2010_Silver As String = "Office 2010 Silver"
        Shared Caramel As String = "Caramel" : Shared Money_Twins As String = "Money Twins"
        Shared Metropolis As String = "Metropolis"
#End Region
    End Structure
End Module
