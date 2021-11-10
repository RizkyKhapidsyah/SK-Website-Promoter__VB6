VERSION 5.00
Begin VB.Form frmProfiles 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Save Profile As ..."
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdSaveProfile 
      Caption         =   "Save"
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.TextBox ProfileName 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   360
         Width           =   5295
      End
      Begin VB.ListBox Profiles 
         Height          =   1815
         Left            =   1800
         TabIndex        =   5
         Top             =   720
         Width           =   5295
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   360
         Picture         =   "frmProfiles.frx":0000
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   4
         Top             =   2160
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Save Profile As:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Available Profiles:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmProfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intFile As Integer, lngRecLength As Long, lngNextID As Long
Dim lngTotalRecords As Long, lngID As Long
Dim NumRecords As Integer
Dim intFileNum As Integer
Dim Records As Integer


' Type Profile is for savings into random _
  access file. (Dont's change that once you have saved anything)
Private Type Profile
    ProfileName As String * 20
    URL As String * 100
    Title As String * 100
    Descr As String * 264
    Key1 As String * 12
    Key2 As String * 12
    Key3 As String * 12
    Key4 As String * 12
    Key5 As String * 12
    Key6 As String * 12
    Key7 As String * 12
    Key8 As String * 12
    Key9 As String * 12
    Key10 As String * 12
    Key11 As String * 12
    Key12 As String * 12
    Key13 As String * 12
    Key14 As String * 12
    Key15 As String * 12
    Key16 As String * 12
    Key17 As String * 12
    Key18 As String * 12
    NameY As String * 25
    Compamny As String * 25
    City As String * 15
    Country As String * 20
    email As String * 30
    Address As String * 50
    Province As String * 15
    Postal As String * 9
    Phone As Integer

End Type

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
' Load up the profiles
Dim udtProfToView As Profile
'Open File
intFile = FreeFile
lngRecLength = LenB(udtProfToView)
Open App.Path & "\profiles\profiles.dat" For Random As #intFile Len = lngRecLength
'# of Rec.
If LOF(intFile) Mod lngRecLength = 0 Then
NumRecords = (LOF(intFile) \ lngRecLength)
Else
NumRecords = (LOF(intFile) \ lngRecLength) + 1
End If
lngTotalRecords = NumRecords
'View Rec if Valid
    If lngTotalRecords = 0 Then
Records = 1
    Exit Sub
    
    Close #intFile
    End If
lngID = 0
Do
    If lngID > lngTotalRecords Then
Records = lngTotalRecords + 1

Exit Sub
Close #intFile
    Else
   
lngID = lngID + 1
 If lngID > 0 And lngID <= lngTotalRecords Then
Get #intFile, lngID, udtProfToView
' Add for selection in case of load up.
Profiles.AddItem udtProfToView.ProfileName
 End If
    End If
Loop
Close #intFile
End Sub
Private Sub cmdSaveProfile_Click()
' Procedure to save the profile into random access file.


Dim udtNewProfile As Profile
'Dim intFile As Integer, lngRecLength As Long, lngNewID As Long
'Open File
intFile = FreeFile
lngRecLength = LenB(udtNewProfile)
Open App.Path & "\profiles\profiles.dat" For Random As #intFile Len = lngRecLength

'Adds New profile to file
lngNewID = Records
udtNewProfile.ProfileName = ProfileName.Text
udtNewProfile.Address = frmMain.txtAddress.Text
udtNewProfile.City = frmMain.txtCity.Text
udtNewProfile.Compamny = frmMain.txtCompany.Text
udtNewProfile.Country = frmMain.txtCountry.Text
udtNewProfile.Descr = frmMain.txtDes.Text
udtNewProfile.email = frmMain.txtMail.Text
udtNewProfile.Key1 = frmMain.k1.Text
udtNewProfile.Key2 = frmMain.k2.Text
udtNewProfile.Key3 = frmMain.k3.Text
udtNewProfile.Key4 = frmMain.k4.Text
udtNewProfile.Key5 = frmMain.k5.Text
udtNewProfile.Key6 = frmMain.k6.Text
udtNewProfile.Key7 = frmMain.k7.Text
udtNewProfile.Key8 = frmMain.k8.Text
udtNewProfile.Key9 = frmMain.k9.Text
udtNewProfile.Key10 = frmMain.k10.Text
udtNewProfile.Key11 = frmMain.k11.Text
udtNewProfile.Key12 = frmMain.k12.Text
udtNewProfile.Key13 = frmMain.k13.Text
udtNewProfile.Key14 = frmMain.k14.Text
udtNewProfile.Key15 = frmMain.k15.Text
udtNewProfile.Key16 = frmMain.k16.Text
udtNewProfile.Key17 = frmMain.k17.Text
udtNewProfile.Key18 = frmMain.k18.Text
udtNewProfile.NameY = frmMain.txtName.Text
'udtNewProfile.Phone = frmMain.txtPhone.Text
udtNewProfile.Postal = frmMain.txtPostal.Text
udtNewProfile.Province = frmMain.txtProvince.Text
udtNewProfile.Title = frmMain.txtTitle.Text
udtNewProfile.URL = frmMain.txtURL.Text
Put #intFile, lngNewID, udtNewProfile
Profiles.AddItem udtNewProfile.ProfileName
Close #intFile
Unload Me
End Sub

Private Sub cmdLoad_Click()
' We  have to dim and Trim every field that's gonna be load up, because
' we don't want for example a URL which has a name and then 30 spaces
' behind it. Duh, it wouldn't do any good. :)
Dim Address_T As String, City_T As String, Company_T As String, Country_T As String, Des_T As String, Mail_T As String
Dim k1_T As String, k2_T As String, k3_T As String, k4_T As String, k5_T As String, k6_T As String, k7_T As String, k8_T As String, k9_T As String, k10_T As String
Dim k11_T As String, k12_T As String, k13_T As String, k14_T As String, k15_T As String, k16_T As String, k17_T As String, k18_T As String
Dim Name_T As String, Postal_T As String, Province_T As String, Title_T As String, URL_T As String



' All right let's load up already saved profile.
Dim udtLoadProfile As Profile
Dim strTrimed1 As String, strTrimed2 As String

intFile = FreeFile
lngRecLength = LenB(udtLoadProfile)
Open App.Path & "\profiles\profiles.dat" For Random As #intFile Len = lngRecLength

'# of Rec.
If LOF(intFile) Mod lngRecLength = 0 Then
NumRecords = (LOF(intFile) \ lngRecLength)
Else
NumRecords = (LOF(intFile) \ lngRecLength) + 1
End If
lngTotalRecords = NumRecords


lngID = 1

Do

If lngID > lngTotalRecords Then
    MsgBox lngTotalRecords
    Exit Sub
End If

' (Royal Pain in the Ass!)
' It took me like 40 minutes to figure out how to do it.
' Even though it's very!!! simple.

Get #intFile, lngID, udtLoadProfile
  strTrimed1 = Trim(Profiles.Text)
  strTrimed2 = Trim(udtLoadProfile.ProfileName)

                    If strTrimed1 = strTrimed2 Then



Get #intFile, lngID, udtLoadProfile



Address_T = udtLoadProfile.Address
City_T = udtLoadProfile.City
Company_T = udtLoadProfile.Compamny
Country_T = udtLoadProfile.Country
Des_T = udtLoadProfile.Descr
Mail_T = udtLoadProfile.email
k1_T = udtLoadProfile.Key1
k2_T = udtLoadProfile.Key2
k3_T = udtLoadProfile.Key3
k4_T = udtLoadProfile.Key4
k5_T = udtLoadProfile.Key5
k6_T = udtLoadProfile.Key6
k7_T = udtLoadProfile.Key7
k8_T = udtLoadProfile.Key8
k9_T = udtLoadProfile.Key9
k10_T = udtLoadProfile.Key10
k11_T = udtLoadProfile.Key11
k12_T = udtLoadProfile.Key12
k13_T = udtLoadProfile.Key13
k14_T = udtLoadProfile.Key14
k15_T = udtLoadProfile.Key15
k16_T = udtLoadProfile.Key16
k17_T = udtLoadProfile.Key17
k18_T = udtLoadProfile.Key18
Name_T = udtLoadProfile.NameY
Postal_T = udtLoadProfile.Postal
Province_T = udtLoadProfile.Province
Title_T = udtLoadProfile.Title
URL_T = udtLoadProfile.URL
'
'
'Now we have to Trim(them all) and paste them onto the form.
'
'
frmMain.txtAddress.Text = Trim(Address_T)
frmMain.txtCity.Text = Trim(City_T)
frmMain.txtCompany.Text = Trim(Company_T)
frmMain.txtCountry.Text = Trim(Country_T)
frmMain.txtDes.Text = Trim(Des_T)
frmMain.txtMail.Text = Trim(Mail_T)
frmMain.k1.Text = Trim(k1_T)
frmMain.k2.Text = Trim(k2_T)
frmMain.k3.Text = Trim(k3_T)
frmMain.k4.Text = Trim(k4_T)
frmMain.k5.Text = Trim(k5_T)
frmMain.k6.Text = Trim(k6_T)
frmMain.k7.Text = Trim(k7_T)
frmMain.k8.Text = Trim(k8_T)
frmMain.k9.Text = Trim(k9_T)
frmMain.k10.Text = Trim(k10_T)
frmMain.k11.Text = Trim(k11_T)
frmMain.k12.Text = Trim(k12_T)
frmMain.k13.Text = Trim(k13_T)
frmMain.k14.Text = Trim(k14_T)
frmMain.k15.Text = Trim(k15_T)
frmMain.k16.Text = Trim(k16_T)
frmMain.k17.Text = Trim(k17_T)
frmMain.k18.Text = Trim(k18_T)
frmMain.txtName.Text = Trim(Name_T)
frmMain.txtPostal.Text = Trim(Postal_T)
frmMain.txtProvince.Text = Trim(Province_T)
frmMain.txtTitle.Text = Trim(Title_T)
frmMain.txtURL.Text = Trim(URL_T)
Unload Me
Exit Sub
End If
           
           lngID = lngID + 1
        Loop


End Sub

Private Sub Profiles_DblClick()
' Double click as well
    cmdLoad_Click
End Sub
