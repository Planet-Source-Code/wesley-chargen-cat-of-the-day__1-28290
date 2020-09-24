VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Cat of the Day"
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
   ForeColor       =   &H000000FF&
   Icon            =   "catofday.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PBoxCat 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Loading..."
      Height          =   195
      Left            =   3308
      TabIndex        =   1
      Top             =   2460
      Width           =   705
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type TGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function OleLoadPicturePath Lib "oleaut32.dll" (ByVal szURLorPath As Long, ByVal punkCaller As Long, ByVal dwReserved As Long, ByVal clrReserved As OLE_COLOR, ByRef riid As TGUID, ByRef ppvRet As IPicture) As Long

Dim strYear, strMonth, strDay, strFilename As String

Public Function LoadPicture(ByVal strFilename As String) As Picture
Dim myTGUID As TGUID
myTGUID.Data1 = &H7BF80980
myTGUID.Data2 = &HBF32
myTGUID.Data3 = &H101A
myTGUID.Data4(0) = &H8B
myTGUID.Data4(1) = &HBB
myTGUID.Data4(2) = &H0
myTGUID.Data4(3) = &HAA
myTGUID.Data4(4) = &H0
myTGUID.Data4(5) = &H30
myTGUID.Data4(6) = &HC
myTGUID.Data4(7) = &HAB

On Error GoTo LblError
OleLoadPicturePath StrPtr(strFilename), 0, 0, 0, myTGUID, LoadPicture
Exit Function

LblError:
Set LoadPicture = VB.LoadPicture(strFilename)

End Function

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdMini_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Form_Load()
frmMain.Show
strYear = Format(Now, "yyyy")
strMonth = Format(Now, "MMMM")
strDay = Format(Now, "dd")
strFilename = "http://www.catoftheday.com/archive/" + strYear + "/" + strMonth + "/" + strDay + ".jpg"
PBoxCat.Picture = LoadPicture(strFilename)
frmMain.Height = PBoxCat.Height
frmMain.Width = PBoxCat.Width
End Sub

Private Sub PBoxCat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
PopupMenu frmMnu.mnuAll
End If
End Sub
