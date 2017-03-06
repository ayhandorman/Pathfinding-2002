VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pathfinding by Ayhan Dorman"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   562
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   511
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7050
      Top             =   7110
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DUR"
      Height          =   330
      Left            =   6990
      TabIndex        =   6
      Top             =   7680
      Width           =   615
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Hedef Nokta"
      Height          =   225
      Left            =   105
      TabIndex        =   5
      Top             =   7920
      Width           =   1635
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Engel"
      Height          =   225
      Left            =   105
      TabIndex        =   4
      Top             =   8160
      Width           =   795
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Baþlangýç Noktasý"
      Height          =   225
      Left            =   105
      TabIndex        =   3
      Top             =   7680
      Value           =   -1  'True
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   7560
      Left            =   45
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   2
      Top             =   60
      Width           =   7560
      Begin VB.Shape ben 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   150
         Left            =   3000
         Top             =   2100
         Width           =   150
      End
      Begin VB.Shape sen 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   150
         Left            =   750
         Top             =   2100
         Width           =   150
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TEMÝZLE"
      Height          =   330
      Left            =   6330
      TabIndex        =   1
      Top             =   8040
      Width           =   1260
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BUL"
      Default         =   -1  'True
      Height          =   330
      Left            =   6330
      TabIndex        =   0
      Top             =   7680
      Width           =   615
   End
   Begin VB.Image bosluk 
      Height          =   135
      Left            =   2160
      Picture         =   "path.frx":0000
      Top             =   7920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image engel 
      Height          =   135
      Left            =   2160
      Picture         =   "path.frx":013E
      Top             =   7740
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'noktanýn yönü ve içeriði
Dim nokta(-1 To 50, -1 To 50) As Byte

Dim X As Integer

Dim Y As Integer

Dim bulundu As Boolean

Dim adimx(1 To 5000) As Byte

Dim adimy(1 To 5000) As Byte

Dim adim As Integer

Dim sayac As Integer

Dim git As Integer

Dim toplam_bakilan As Integer

Dim toplam_bakilan_eski As Integer


Private Sub Command3_Click()

	sen.Left = sen.Left - sen.Left Mod 10
	
	sen.Top = sen.Top - sen.Top Mod 10
	
	Timer2.Enabled = False
		
	adim = 0
	
	git = 0
	
	Command1.Enabled = True
	
End Sub


Private Sub Form_Load()

	cizim
	
	For X = -1 To 50
	
		nokta(X, -1) = 9
		
		nokta(-1, X) = 9
		
		nokta(X, 50) = 9
		
		nokta(50, X) = 9
		
	Next
	
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

	On Local Error Resume Next
	
	If Button = 1 Then
	
		If Option1.Value Then sen.Left = X - X Mod 10: sen.Top = Y - Y Mod 10
		
		If Option2.Value Then ben.Left = X - X Mod 10: ben.Top = Y - Y Mod 10
		
		If Option3.Value Then
		
			nokta((X - X Mod 10) \ 10, (Y - Y Mod 10) \ 10) = 9
			
			Picture1.PaintPicture engel.Picture, X - X Mod 10 + 1, Y - Y Mod 10 + 1
			
		End If
		
	End If
	
	If Button = 2 Then
	
		nokta((X - X Mod 10) \ 10, (Y - Y Mod 10) \ 10) = 0
		
		Picture1.PaintPicture bosluk.Picture, X - X Mod 10 + 1, Y - Y Mod 10 + 1
		
	End If
	
	If sen.Left < 0 Then sen.Left = 0
	
	If sen.Left > 490 Then sen.Left = 490
	
	If sen.Top < 0 Then sen.Top = 0
	
	If sen.Top > 490 Then sen.Top = 490
	
	If ben.Left < 0 Then ben.Left = 0
	
	If ben.Left > 490 Then ben.Left = 490
	
	If ben.Top < 0 Then ben.Top = 0
	
	If ben.Top > 490 Then ben.Top = 490
	
	
End Sub


Sub cizim()

	For X = 0 To 49
	
		For Y = 0 To 49
		
			Picture1.Line (X * 10, Y * 10)-((X + 1) * 10, Y * 10)
			
			Picture1.Line (X * 10, Y * 10)-(X * 10, (Y + 1) * 10)
			
		Next
		
	Next
	
End Sub


Private Sub Command2_Click()

	Picture1.Cls
	
	For X = 0 To 49
	
		For Y = 0 To 49
		
			If nokta(X, Y) > 0 And nokta(X, Y) < 9 Then Picture1.PaintPicture bosluk.Picture, X * 10 + 1, Y * 10 + 1: nokta(X, Y) = 0
			
			If nokta(X, Y) = 9 Then Picture1.PaintPicture engel.Picture, X * 10 + 1, Y * 10 + 1
			
		Next
		
	Next
	
	cizim
	
End Sub


Private Sub Command1_Click()

	toplam_bakilan_eski = 0
	
	nokta(sen.Left \ 10, sen.Top \ 10) = 1
	
	bulundu = False
	
	Command1.Enabled = False
	
	On Local Error Resume Next
	
	'''''''arama''''''''
	
	Do
	
		For X = 0 To 49
		
			For Y = 0 To 49
			
				If bulundu Then Exit Sub
				If nokta(X - 1, Y) <> 0 And nokta(X - 1, Y - 1) <> 0 And nokta(X, Y - 1) <> 0 And nokta(X + 1, Y - 1) <> 0 And nokta(X + 1, Y) <> 0 And nokta(X + 1, Y + 1) <> 0 And nokta(X, Y + 1) <> 0 And nokta(X - 1, Y + 1) <> 0 Then GoTo devam

				If nokta(X, Y) > 0 And nokta(X, Y) < 9 Then
					If nokta(X - 1, Y) = 0 Then nokta(X - 1, Y) = 1
					If nokta(X - 1, Y - 1) = 0 Then nokta(X - 1, Y - 1) = 2
					If nokta(X, Y - 1) = 0 Then nokta(X, Y - 1) = 3
					If nokta(X + 1, Y - 1) = 0 Then nokta(X + 1, Y - 1) = 4
					If nokta(X + 1, Y) = 0 Then nokta(X + 1, Y) = 5
					If nokta(X + 1, Y + 1) = 0 Then nokta(X + 1, Y + 1) = 6
					If nokta(X, Y + 1) = 0 Then nokta(X, Y + 1) = 7
					If nokta(X - 1, Y + 1) = 0 Then nokta(X - 1, Y + 1) = 8

					If nokta(X - 1, Y) > 0 And nokta(X - 1, Y) < 9 And ben.Left = (X - 1) * 10 And ben.Top = Y * 10 Then bulundu = True
					If nokta(X - 1, Y - 1) > 0 And nokta(X - 1, Y - 1) < 9 And ben.Left = (X - 1) * 10 And ben.Top = (Y - 1) * 10 Then bulundu = True
					If nokta(X, Y - 1) > 0 And nokta(X, Y - 1) < 9 And ben.Left = X * 10 And ben.Top = (Y - 1) * 10 Then bulundu = True
					If nokta(X + 1, Y - 1) > 0 And nokta(X + 1, Y - 1) < 9 And ben.Left = (X + 1) * 10 And ben.Top = (Y - 1) * 10 Then bulundu = True
					If nokta(X + 1, Y) > 0 And nokta(X + 1, Y) < 9 And ben.Left = (X + 1) * 10 And ben.Top = Y * 10 Then bulundu = True
					If nokta(X + 1, Y + 1) > 0 And nokta(X + 1, Y + 1) < 9 And ben.Left = (X + 1) * 10 And ben.Top = (Y + 1) * 10 Then bulundu = True
					If nokta(X, Y + 1) > 0 And nokta(X, Y + 1) < 9 And ben.Left = X * 10 And ben.Top = (Y + 1) * 10 Then bulundu = True
					If nokta(X - 1, Y + 1) > 0 And nokta(X - 1, Y + 1) < 9 And ben.Left = (X - 1) * 10 And ben.Top = (Y + 1) * 10 Then bulundu = True
					'''''''kestirme yola indirgeme
					If bulundu Then
						adim = 1
						adimx(adim) = X
						adimy(adim) = Y
						Do
							Select Case nokta(X, Y)
								Case 1: X = X + 1
								Case 2: X = X + 1: Y = Y + 1
								Case 3: Y = Y + 1
								Case 4: X = X - 1: Y = Y + 1
								Case 5: X = X - 1
								Case 6: X = X - 1: Y = Y - 1
								Case 7: Y = Y - 1
								Case 8: X = X + 1: Y = Y - 1
							End Select
							adim = adim + 1
							adimx(adim) = X
							adimy(adim) = Y
						Loop Until sen.Left = X * 10 And sen.Top = Y * 10
						git = adim - 1
						Picture1.ForeColor = 33023
						For sayac = 1 To adim - 1
							Picture1.Line (adimx(sayac) * 10 + 5, adimy(sayac) * 10 + 5)-(adimx(sayac + 1) * 10 + 5, adimy(sayac + 1) * 10 + 5)
						Next
						Picture1.ForeColor = 0
						Timer2.Enabled = True
					End If
				End If
				devam:
			Next
		Next
		
		'yol bulunamadýysa
		toplam_bakilan = 0
		
		For X = 0 To 49
		
			For Y = 0 To 49
			
				If nokta(X, Y) <> 0 And nokta(X, Y) <> 9 Then toplam_bakilan = toplam_bakilan + 1
				
			Next
			
		Next
	
		If toplam_bakilan = toplam_bakilan_eski Then
		
			MsgBox "Yol bulunamadý.", vbOKOnly, "Bulunamadý!"
			
			toplam_bakilan_eski = 0
			
			sen.Left = sen.Left - sen.Left Mod 10
			
			sen.Top = sen.Top - sen.Top Mod 10
			
			Timer2.Enabled = False
			
			adim = 0
			
			git = 0
			
			Exit Sub
			
		End If
		
		toplam_bakilan_eski = toplam_bakilan
		
	Loop Until bulundu = True
	
End Sub


Private Sub Timer2_Timer()

	On Local Error Resume Next
	
	If git = 0 Then Timer2.Enabled = False: Exit Sub
	
	If sen.Left < adimx(git) * 10 Then sen.Left = sen.Left + 1
	
	If sen.Left > adimx(git) * 10 Then sen.Left = sen.Left - 1
	
	If sen.Top < adimy(git) * 10 Then sen.Top = sen.Top + 1
	
	If sen.Top > adimy(git) * 10 Then sen.Top = sen.Top - 1
	
	If adimx(git) * 10 = sen.Left And adimy(git) * 10 = sen.Top Then git = git - 1
	
End Sub
