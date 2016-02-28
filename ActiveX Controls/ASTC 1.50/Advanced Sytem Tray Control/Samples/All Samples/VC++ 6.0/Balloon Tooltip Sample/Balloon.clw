; CLW file contains information for the MFC ClassWizard

[General Info]
Version=1
LastClass=CBalloonDlg
LastTemplate=CDialog
NewFileInclude1=#include "stdafx.h"
NewFileInclude2=#include "Balloon.h"
ODLFile=Balloon.odl

ClassCount=4
Class1=CBalloonApp
Class2=CBalloonDlg
Class3=CAboutDlg
Class4=CBalloonDlgAutoProxy

ResourceCount=5
Resource1=IDD_ABOUTBOX
Resource2=IDR_MAINFRAME
Resource3=IDD_BALLOON_DIALOG
Resource4=IDD_ABOUTBOX (English (U.S.))
Resource5=IDD_BALLOON_DIALOG (English (U.S.))

[CLS:CBalloonApp]
Type=0
HeaderFile=Balloon.h
ImplementationFile=Balloon.cpp
Filter=N

[CLS:CBalloonDlg]
Type=0
HeaderFile=BalloonDlg.h
ImplementationFile=BalloonDlg.cpp
Filter=D
LastObject=IDC_BUTTON1
BaseClass=CDialog
VirtualFilter=dWC

[CLS:CAboutDlg]
Type=0
HeaderFile=BalloonDlg.h
ImplementationFile=BalloonDlg.cpp
Filter=D

[DLG:IDD_ABOUTBOX]
Type=1
ControlCount=4
Control1=IDC_STATIC,static,1342177283
Control2=IDC_STATIC,static,1342308352
Control3=IDC_STATIC,static,1342308352
Control4=IDOK,button,1342373889
Class=CAboutDlg

[CLS:CBalloonDlgAutoProxy]
Type=0
HeaderFile=DlgProxy.h
ImplementationFile=DlgProxy.cpp
BaseClass=CCmdTarget
Filter=N

[DLG:IDD_BALLOON_DIALOG]
Type=1
ControlCount=3
Control1=IDOK,button,1342242817
Control2=IDCANCEL,button,1342242816
Control3=IDC_STATIC,static,1342308352
Class=CBalloonDlg

[DLG:IDD_BALLOON_DIALOG (English (U.S.))]
Type=1
Class=CBalloonDlg
ControlCount=7
Control1=IDC_STATIC,button,1342177287
Control2=IDC_STATIC,static,1342308352
Control3=IDC_BUTTON1,button,1342242816
Control4=IDC_BUTTON2,button,1342242816
Control5=IDC_BUTTON3,button,1342242816
Control6=IDC_BALLOON2,{30E583AF-8E26-11DA-A1C2-807A50C10000},1342242816
Control7=IDC_BALLOON1,{30E583AF-8E26-11DA-A1C2-807A50C10000},1342242816

[DLG:IDD_ABOUTBOX (English (U.S.))]
Type=1
Class=?
ControlCount=4
Control1=IDC_STATIC,static,1342177283
Control2=IDC_STATIC,static,1342308480
Control3=IDC_STATIC,static,1342308352
Control4=IDOK,button,1342373889

