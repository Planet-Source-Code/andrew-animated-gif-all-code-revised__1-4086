VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Animated GIF"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txFile 
      Height          =   330
      Left            =   165
      TabIndex        =   1
      Top             =   150
      Width           =   3780
   End
   Begin VB.Timer AnimationTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   195
      Top             =   555
   End
   Begin VB.CommandButton btnPlay 
      Caption         =   "PLAY"
      Height          =   465
      Left            =   4200
      TabIndex        =   0
      Top             =   105
      Width           =   1380
   End
   Begin VB.Image AnimatedGIF 
      Appearance      =   0  'Flat
      Height          =   900
      Index           =   0
      Left            =   855
      Top             =   615
      Width           =   1500
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim RepeatTimes&
Dim RepeatCount&
Dim FrameCount&
Dim TotalFrames&


Private Sub btnPlay_Click()
       
    Call LoadAniGif(txFile.Text, AnimatedGIF)

End Sub


Sub LoadAniGif(xFile As String, xImgArray)

    If Not IIf(Dir$(xFile) = "", False, True) Or xFile = "" Then
        MsgBox "File not found.", vbExclamation, "File Error"
        Exit Sub
    End If
        
    Dim F1, F2
    Dim AnimatedGIFs() As String
    Dim imgHeader As String
    Static buf$, picbuf$
    Dim fileHeader As String
    Dim imgCount
    Dim i&, j&, xOff&, yOff&, TimeWait&
    Dim GifEnd
    GifEnd = Chr(0) & "!ù"
    
    AnimationTimer.Enabled = False
    For i = 1 To xImgArray.Count - 1
        Unload xImgArray(i)
    Next i
    
    F1 = FreeFile
On Error GoTo badFile:
    Open xFile For Binary Access Read As F1
        buf = String(LOF(F1), Chr(0))
        Get #F1, , buf
    Close F1
    
    i = 1
    imgCount = 0
    
    j = (InStr(1, buf, GifEnd) + Len(GifEnd)) - 2
    fileHeader = Left(buf, j)
    i = j + 2
    
    If Len(fileHeader) >= 127 Then
        RepeatTimes& = Asc(Mid(fileHeader, 126, 1)) + (Asc(Mid(fileHeader, 127, 1)) * CLng(256))
    Else
        RepeatTimes = 0
    End If


    Do
        imgCount = imgCount + 1
        j = InStr(i, buf, GifEnd) + Len(GifEnd)
        If j > Len(GifEnd) Then
            F2 = FreeFile
            Open "tmp.gif" For Binary As F2
                picbuf = String(Len(fileHeader) + j - i, Chr(0))
                picbuf = fileHeader & Mid(buf, i - 1, j - i)
                Put #F2, 1, picbuf
                imgHeader = Left(Mid(buf, i - 1, j - i), 16)
            Close F2
            
            TimeWait = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * CLng(256))) * CLng(10)
            If imgCount > 1 Then
                xOff = Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * CLng(256))
                yOff = Asc(Mid(imgHeader, 11, 1)) + (Asc(Mid(imgHeader, 12, 1)) * CLng(256))
                Load xImgArray(imgCount - 1)
                xImgArray(imgCount - 1).ZOrder 0
                xImgArray(imgCount - 1).Left = xImgArray(0).Left + (xOff * CLng(15))
                xImgArray(imgCount - 1).Top = xImgArray(0).Top + (yOff * CLng(15))
            End If
            xImgArray(imgCount - 1).Tag = TimeWait
            xImgArray(imgCount - 1).Picture = LoadPicture("tmp.gif")
            Kill ("tmp.gif")
            
            i = j '+ 1
        End If
        DoEvents
    Loop Until j = Len(GifEnd)
    
    If i < Len(buf) Then
        F2 = FreeFile
        Open "tmp.gif" For Binary As F2
            picbuf = String(Len(fileHeader) + Len(buf) - i, Chr(0))
            picbuf = fileHeader & Mid(buf, i - 1, Len(buf) - i)
            Put #F2, 1, picbuf
            imgHeader = Left(Mid(buf, i - 1, Len(buf) - i), 16)
        Close F2

        TimeWait = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * CLng(256))) * CLng(10)
        If imgCount > 1 Then
            xOff = Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * CLng(256))
            yOff = Asc(Mid(imgHeader, 11, 1)) + (Asc(Mid(imgHeader, 12, 1)) * CLng(256))
            Load xImgArray(imgCount - 1)
            xImgArray(imgCount - 1).ZOrder 0
            xImgArray(imgCount - 1).Left = xImgArray(0).Left + (xOff * CLng(15))
            xImgArray(imgCount - 1).Top = xImgArray(0).Top + (yOff * CLng(15))
        End If
        xImgArray(imgCount - 1).Tag = TimeWait
        xImgArray(imgCount - 1).Picture = LoadPicture("tmp.gif")
        Kill ("tmp.gif")
    End If
    
    FrameCount = 0
    TotalFrames = xImgArray.Count - 1
    
On Error GoTo badTime
    AnimationTimer.Interval = CInt(xImgArray(0).Tag)
badTime:
    AnimationTimer.Enabled = True
Exit Sub
badFile:
    MsgBox "File not found.", vbExclamation, "File Error"

End Sub

Private Sub AnimationTimer_Timer()


    If FrameCount < TotalFrames Then
        FrameCount = FrameCount + 1
        AnimatedGIF(FrameCount).Visible = True
        AnimationTimer.Interval = CLng(AnimatedGIF(FrameCount).Tag)
    Else
        FrameCount = 0
        For i = 1 To AnimatedGIF.Count - 1
            AnimatedGIF(i).Visible = False
        Next i
        AnimationTimer.Interval = CLng(AnimatedGIF(FrameCount).Tag)
    End If
    
'    For i = 0 To AnimatedGIF.Count
'        If i = AnimatedGIF.Count Then
'            If RepeatTimes > 0 Then
'                RepeatCount = RepeatCount + 1
'                If RepeatCount > RepeatTimes Then
'                    AnimationTimer.Enabled = False
'                    Exit Sub
'                End If
'            End If
'            For j = 1 To AnimatedGIF.Count - 1
'                AnimatedGIF(j).Visible = False
'            Next j
'On Error GoTo badTime
'            AnimationTimer.Interval = CLng(AnimatedGIF(0).Tag)
'badTime:
'            Exit For
'        End If
'        If AnimatedGIF(i).Visible = False Then
'On Error GoTo badTime2
'            AnimationTimer.Interval = CLng(AnimatedGIF(i).Tag)
'badTime2:
'            AnimatedGIF(i).Visible = True
'            Exit For
'        End If
'    Next i

End Sub

