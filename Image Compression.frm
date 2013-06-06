VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "JPEG ENCODER"
   ClientHeight    =   11055
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   11055
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame BinaryShiftCoding 
      Caption         =   "Binary Shift Codung"
      Height          =   975
      Left            =   8400
      TabIndex        =   36
      Top             =   7440
      Width           =   2295
      Begin VB.CommandButton BinaryShift 
         Caption         =   "Execute"
         Height          =   375
         Left            =   600
         TabIndex        =   37
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8520
      Top             =   9960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Caption         =   "RLE"
      Height          =   975
      Left            =   8400
      TabIndex        =   34
      Top             =   6240
      Width           =   2295
      Begin VB.CommandButton RLE 
         Caption         =   "Execute"
         Height          =   375
         Left            =   600
         TabIndex        =   35
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "DPCM"
      Height          =   975
      Left            =   8400
      TabIndex        =   32
      Top             =   5040
      Width           =   2295
      Begin VB.CommandButton DPCM 
         Caption         =   "Execute"
         Height          =   375
         Left            =   600
         TabIndex        =   33
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1935
      Left            =   3000
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   13
      Top             =   2880
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1935
      Left            =   240
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   11
      Top             =   2880
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1935
      Left            =   3000
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   10
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton zigzagscan 
      Caption         =   "Scan"
      Height          =   372
      Left            =   9000
      TabIndex        =   9
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame zigzagframe 
      Caption         =   "ZigZag Scan"
      Height          =   1095
      Left            =   8400
      TabIndex        =   8
      Top             =   3720
      Width           =   2292
   End
   Begin VB.CommandButton quantization 
      Caption         =   "Quantize"
      Height          =   372
      Left            =   9000
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Quantization"
      Height          =   975
      Left            =   8400
      TabIndex        =   6
      Top             =   2520
      Width           =   2292
   End
   Begin VB.CommandButton DCT_Tran 
      Caption         =   "Transform"
      Height          =   372
      Left            =   9000
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton ColorTransform 
      BackColor       =   &H8000000C&
      Caption         =   "Transform"
      Height          =   372
      Left            =   9000
      MaskColor       =   &H00000000&
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Color Transform"
      Height          =   975
      Left            =   8400
      TabIndex        =   2
      Top             =   120
      Width           =   2292
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1935
      Left            =   240
      Negotiate       =   -1  'True
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H8000000C&
      Height          =   10815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8055
      Begin VB.PictureBox Picture11 
         Height          =   1935
         Left            =   5520
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   125
         TabIndex        =   28
         Top             =   8040
         Width           =   1935
      End
      Begin VB.PictureBox Picture10 
         Height          =   1935
         Left            =   2880
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   125
         TabIndex        =   27
         Top             =   8040
         Width           =   1935
      End
      Begin VB.PictureBox Picture9 
         Height          =   1935
         Left            =   120
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   125
         TabIndex        =   26
         Top             =   8040
         Width           =   1935
      End
      Begin VB.PictureBox Picture8 
         Height          =   1935
         Left            =   5520
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   125
         TabIndex        =   22
         Top             =   5280
         Width           =   1935
      End
      Begin VB.PictureBox Picture7 
         Height          =   1935
         Left            =   2880
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   125
         TabIndex        =   21
         Top             =   5280
         Width           =   1935
      End
      Begin VB.PictureBox Picture6 
         Height          =   1935
         Left            =   120
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   125
         TabIndex        =   17
         Top             =   5280
         Width           =   1935
      End
      Begin VB.PictureBox Picture5 
         AutoRedraw      =   -1  'True
         Height          =   1935
         Left            =   5520
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   125
         TabIndex        =   12
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ZIGZAG OF [V] COMPONENT"
         Height          =   255
         Left            =   5520
         TabIndex        =   31
         Top             =   10080
         Width           =   2295
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ZIGZAG OF [U] COMPONENT"
         Height          =   255
         Left            =   2880
         TabIndex        =   30
         Top             =   10080
         Width           =   2295
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ZIGZAG OF [Y] COMPONENT"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   10080
         Width           =   2295
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "QUANTIZATION OF [V]                COMPONENT"
         Height          =   375
         Left            =   5640
         TabIndex        =   25
         Top             =   7320
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "QUANTIZATION  OF [U]           COMPONENT"
         Height          =   375
         Left            =   3000
         TabIndex        =   24
         Top             =   7320
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "QUANTIZATION  OF [Y]          COMPONENT"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   7320
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "DCT OF [V] COMPONENT"
         Height          =   255
         Left            =   5520
         TabIndex        =   20
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "DCT OF [U] COMPONENT"
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "DCT OF [Y] COMPONENT"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "RGB TO YUV CONVERTER"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ORIGINAL IMAGE"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   2280
         Width           =   2055
      End
   End
   Begin VB.Frame DCT 
      Caption         =   "DCT Tranform"
      Height          =   975
      Left            =   8400
      TabIndex        =   3
      Top             =   1320
      Width           =   2292
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Top             =   4680
      Width           =   975
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu open 
         Caption         =   "Open"
      End
      Begin VB.Menu save 
         Caption         =   "Save"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Structs for JPEG header
Private Type APP0infotype
    marker As Long
    length As Long
    JFIFsignature(4) As Byte
    versionhi As Byte
    versionlo As Byte
    xyunits As Byte
    xdensity As Long
    ydensity As Long
    thumbnwidth As Byte
    thumbnheight As Byte
End Type
 
Private Type SOF0infotype
    marker As Long
    length As Long
    precision As Byte
    height As Long
    width As Long
    nrofcomponents As Byte
    IdY As Byte
    HVY As Byte
    QTY As Byte
    IdCb As Byte
    HVCb As Byte
    QTCb As Byte
    IdCr As Byte
    HVCr As Byte
    QTCr As Byte
End Type
    
Private Type DQTinfotype
     marker As Long
     length As Long
     QTYinfo As Byte
     Ytable(63) As Byte
     QTCbinfo As Byte
     Cbtable(63) As Byte
End Type
' Ytable from DQTinfo should be equal to a scaled and zizag reordered version
' of the table which can be found in "tables.h": std_luminance_qt
' Cbtable , similar = std_chrominance_qt
' We'll init them in the program using set_DQTinfo function
Private Type DHTinfotype
     marker As Long
     length As Long
     HTYDCinfo As Byte
     YDC_nrcodes(15) As Byte
     YDC_values(11) As Byte
     HTYACinfo As Byte
     YAC_nrcodes(15) As Byte
     YAC_values(161) As Byte
     HTCbDCinfo As Byte
     CbDC_nrcodes(15) As Byte
     CbDC_values(11) As Byte
     HTCbACinfo As Byte
     CbAC_nrcodes(15) As Byte
     CbAC_values(161) As Byte
End Type

Private Type SOSinfotype
     marker As Long
     length As Long
     nrofcomponents As Byte
     IdY As Byte
     HTY As Byte
     IdCb As Byte
     HTCb As Byte
     IdCr As Byte
     HTCr As Byte
     Ss As Byte
     Se As Byte
     Bf As Byte
End Type

Private Type bitstring
    length As Byte
    value As Long
End Type

Private Type colorRGB
    B As Byte
    G As Byte
    R As Byte
End Type

Private Type RLE_datatype
    length As Integer
    size As Integer
End Type

Private Type rle_probability
    value As RLE_datatype
    number_appearance  As Integer
End Type

Dim wid, Hgt
Dim bytenew As Byte
Dim bytepos As Integer
Dim mask() As Variant
Dim YDC_HT(11) As bitstring
Dim UDC_HT(11) As bitstring
Dim YAC_HT(255) As bitstring
Dim UAC_HT(255) As bitstring
Dim category(65534) As Byte
Dim bitcode(65534) As bitstring
Dim YRtab(255), YGtab(255), YBtab(255) As Long
Dim URtab(255), UGtab(255), UBtab(255) As Long
Dim VRtab(255), VGtab(255), VBtab(255) As Long
Dim fdtbl_Y(63) As Single
Dim fdtbl_U(63) As Single
Dim RGB_buffer() As colorRGB
Dim Y_DU() As Integer
Dim U_DU() As Integer
Dim V_DU() As Integer
 
Dim zigzag() As Variant
Dim std_luminance_qt() As Variant
Dim std_chrominance_qt() As Variant
Dim std_dc_luminance_nrcodes() As Variant
Dim std_dc_luminance_values() As Variant
Private std_dc_chrominance_nrcodes() As Variant
Dim std_dc_chrominance_values() As Variant
Dim std_ac_luminance_nrcodes() As Variant
Private std_ac_luminance_values() As Variant
Dim std_ac_chrominance_nrcodes() As Variant
Private std_ac_chrominance_values() As Variant

Dim SOSinfo As SOSinfotype
Dim SOF0info As SOF0infotype
Dim APP0info As APP0infotype
Dim DQTinfo As DQTinfotype
Dim DHTinfo As DHTinfotype

Dim BMP_filename As String
Dim JPG_filename As String
Dim DCT_Result() As Single
Dim Quantization_Result() As Integer
Dim ZigZag_Result() As Integer
Dim Diff_Result() As Long
Dim DC_Result() As Integer
Dim RLE_Output() As RLE_datatype
Dim Huffman_Output As String
Dim M16zeroes() As Integer
Dim huffman_data() As Long
Dim rle_prob_output(63) As rle_probability

'Global index
Dim index_1 As Long
Dim index_2 As Long
Dim index_3 As Long
Dim index_4 As Long
Dim index_5 As Long
Dim index_6 As Long
Dim index_7 As Long
Dim size_huffman_output As Long
Dim data_jpeg() As Byte


Private Sub open_Click()

    CommonDialog1.Filter = "*.bmp,*.jpg"
    CommonDialog1.ShowOpen
    BMP_filename = CommonDialog1.FileName
    'BMP_filename = "C:\Users\Administrator\Desktop\BaboonRGB_truebmp_.bmp"
    
    If BMP_filename = "" Then
        MsgBox "Please choose your bitmap image", vbCritical
        
        Else
            Call main
    End If
    If wid <= 128 And Hgt <= 128 Then
        Picture1.Picture = LoadPicture(BMP_filename)
        Else
        MsgBox "Image size too large so this will not displayed!", vbExclamation, "Load image"
    End If
    
End Sub
'Because VB6 doesn't have shift oprator
Private Function shift2Right(ByVal x As Double) As Integer

    If x < 0 Then
        If Abs(x) - Abs(Fix(x)) > 0 Then
            x = Fix(x) - 1
        Else
            x = Fix(x)
        End If
            
    Else
        x = Fix(x)
    End If
    
    shift2Right = x
    
End Function
Private Sub ColorTransform_Click()

    Dim ypos, xpos, i, j As Long
    Dim R, G, B As Byte
    Dim location As Currency
    Dim x1, x2, x3 As Double
    Dim pos As Long
    
    'Y, U, V array
    ReDim Y_DU(wid * Hgt * 3 - 1) As Integer
    ReDim U_DU(wid * Hgt * 3 - 1) As Integer
    ReDim V_DU(wid * Hgt * 3 - 1) As Integer
    
    ' Scan 8x8 blocks
    For ypos = 0 To Hgt - 1 Step 8
        For xpos = 0 To wid - 1 Step 8
            'Scan all elements in each 8x8 block
            location = wid * ypos + xpos
            For j = 0 To 7
                For i = 0 To 7
                    'Load R,G, B elements
                    R = RGB_buffer(location).R
                    G = RGB_buffer(location).G
                    B = RGB_buffer(location).B
                    
                    'Convert R,G,B to YUV
                    x1 = ((YRtab(R) + YGtab(G) + YBtab(B)) / (2 ^ 16)) - 128
                    x2 = (URtab(R) + UGtab(G) + UBtab(B)) / (2 ^ 16)
                    x3 = (VRtab(R) + VGtab(G) + VBtab(B)) / (2 ^ 16)
                    
                    Y_DU(pos) = shift2Right(x1)
                    U_DU(pos) = shift2Right(x2)
                    V_DU(pos) = shift2Right(x3)
                
                    'Set Picture2Box as YUV
                    If wid <= 128 And Hgt <= 128 Then
                        Picture2.PSet (i + xpos, j + ypos), RGB(Y_DU(pos) + 128, U_DU(pos) + 128, V_DU(pos) + 128)
                    End If
                    
                    location = location + 1
                    pos = pos + 1
                Next i
                location = location + wid - 8
            Next j
            
        Next xpos
    Next ypos
    SavePicture Picture2.Image, "C:\image.bmp"
    MsgBox "Transform RGB to YUV done!", vbOKOnly, "RGB to YUV converter"
End Sub
Private Sub DCT_transform(ByRef data() As Integer, ByRef fdtbl() As Single)
    Dim tmp0 As Single
    Dim tmp1 As Single
    Dim tmp2 As Single
    Dim tmp3 As Single
    Dim tmp4 As Single
    Dim tmp5 As Single
    Dim tmp6 As Single
    Dim tmp7 As Single
    Dim tmp10 As Single
    Dim tmp11 As Single
    Dim tmp12 As Single
    Dim tmp13 As Single
    Dim z1 As Single
    Dim z2 As Single
    Dim z3 As Single
    Dim z4 As Single
    Dim z5 As Single
    Dim z11 As Single
    Dim z13 As Single
    Dim datafloat(63) As Single
    Dim temp As Single
    Dim ctr As Integer
    Dim i, j As Byte
    Dim k As Integer
    Dim m As Integer
    
    ' DCT for each 8x8 block
    For i = 0 To 63
        datafloat(i) = data(i)
    Next i

    ' Pass 1: process rows.
    For ctr = 7 To 0 Step -1
        tmp0 = datafloat(k + 0) + datafloat(k + 7)
        tmp7 = datafloat(k + 0) - datafloat(k + 7)
        tmp1 = datafloat(k + 1) + datafloat(k + 6)
        tmp6 = datafloat(k + 1) - datafloat(k + 6)
        tmp2 = datafloat(k + 2) + datafloat(k + 5)
        tmp5 = datafloat(k + 2) - datafloat(k + 5)
        tmp3 = datafloat(k + 3) + datafloat(k + 4)
        tmp4 = datafloat(k + 3) - datafloat(k + 4)

        ' Even part

        tmp10 = tmp0 + tmp3 ' phase 2
        tmp13 = tmp0 - tmp3
        tmp11 = tmp1 + tmp2
        tmp12 = tmp1 - tmp2

        datafloat(k + 0) = tmp10 + tmp11 ' phase 3
        datafloat(k + 4) = tmp10 - tmp11

        z1 = (tmp12 + tmp13) * (CSng(0.707106781)) ' c4
        datafloat(k + 2) = tmp13 + z1 ' phase 5
        datafloat(k + 6) = tmp13 - z1

        ' Odd part

        tmp10 = tmp4 + tmp5 ' phase 2
        tmp11 = tmp5 + tmp6
        tmp12 = tmp6 + tmp7

        ' The rotator is modified from fig 4-8 to avoid extra negations.
        z5 = (tmp10 - tmp12) * (CSng(0.382683433)) ' c6
        z2 = (CSng(0.5411961)) * tmp10 + z5   ' c2-c6
        z4 = (CSng(1.306562965)) * tmp12 + z5 ' c2+c6
        z3 = tmp11 * (CSng(0.707106781)) ' c4

        z11 = tmp7 + z3 ' phase 5
        z13 = tmp7 - z3

        datafloat(k + 5) = z13 + z2 ' phase 6
        datafloat(k + 3) = z13 - z2
        datafloat(k + 1) = z11 + z4
        datafloat(k + 7) = z11 - z4

        k = k + 8 ' advance pointer to next row
    Next ctr

  ' Pass 2: process columns.

    For ctr = 7 To 0 Step -1
        tmp0 = datafloat(m + 0) + datafloat(m + 56)
        tmp7 = datafloat(m + 0) - datafloat(m + 56)
        tmp1 = datafloat(m + 8) + datafloat(m + 48)
        tmp6 = datafloat(m + 8) - datafloat(m + 48)
        tmp2 = datafloat(m + 16) + datafloat(m + 40)
        tmp5 = datafloat(m + 16) - datafloat(m + 40)
        tmp3 = datafloat(m + 24) + datafloat(m + 32)
        tmp4 = datafloat(m + 24) - datafloat(m + 32)

        ' Even part

        tmp10 = tmp0 + tmp3 ' phase 2
        tmp13 = tmp0 - tmp3
        tmp11 = tmp1 + tmp2
        tmp12 = tmp1 - tmp2

        datafloat(m + 0) = tmp10 + tmp11 ' phase 3
        datafloat(m + 32) = tmp10 - tmp11

        z1 = (tmp12 + tmp13) * (CSng(0.707106781)) ' c4
        datafloat(m + 16) = tmp13 + z1 ' phase 5
        datafloat(m + 48) = tmp13 - z1

        ' Odd part

        tmp10 = tmp4 + tmp5 ' phase 2
        tmp11 = tmp5 + tmp6
        tmp12 = tmp6 + tmp7

        ' The rotator is modified from fig 4-8 to avoid extra negations.
        z5 = (tmp10 - tmp12) * (CSng(0.382683433)) ' c6
        z2 = (CSng(0.5411961)) * tmp10 + z5   ' c2-c6
        z4 = (CSng(1.306562965)) * tmp12 + z5 ' c2+c6
        z3 = tmp11 * (CSng(0.707106781)) ' c4

        z11 = tmp7 + z3 ' phase 5
        z13 = tmp7 - z3

        datafloat(m + 40) = z13 + z2 ' phase 6
        datafloat(m + 24) = z13 - z2
        datafloat(m + 8) = z11 + z4
        datafloat(m + 56) = z11 - z4

        m = m + 1 ' advance pointer to next column
    Next ctr
    
    ' Save DCT result
    For k = 0 To 63
        DCT_Result(index_1) = datafloat(k)
        index_1 = index_1 + 1
    Next k
        
End Sub
Private Sub DCT_Tran_Click()

    Dim xpos, ypos, i, j, k, pos, location As Long
    Dim Y_DU64(63) As Integer
    Dim U_DU64(63) As Integer
    Dim V_DU64(63) As Integer
    ReDim DCT_Result(wid * Hgt * 3 - 1) As Single
    
    'DCT transform
    For ypos = 0 To Hgt - 1 Step 8
        For xpos = 0 To wid - 1 Step 8
            
            For i = 0 To 63
                Y_DU64(i) = Y_DU(pos)
                U_DU64(i) = U_DU(pos)
                V_DU64(i) = V_DU(pos)
                pos = pos + 1
            Next i
                
            Call DCT_transform(Y_DU64, fdtbl_Y)     'Y element
            Call DCT_transform(U_DU64, fdtbl_U)     'U element
            Call DCT_transform(V_DU64, fdtbl_U)     'V element
            
        Next xpos
    Next ypos
    
    'Display DCT result
    For ypos = 0 To Hgt - 1 Step 8
        For xpos = 0 To wid - 1 Step 8
            location = wid * ypos + xpos
            For j = 0 To 7
                For i = 0 To 7
                    If wid <= 128 And Hgt <= 128 Then
                        Picture3.PSet (i + xpos, j + ypos), DCT_Result(k)       ' Y
                        Picture4.PSet (i + xpos, j + ypos), DCT_Result(k + 1)   ' U
                        Picture5.PSet (i + xpos, j + ypos), DCT_Result(k + 2)   ' V
                    End If
                    k = k + 3
                    location = location + 1
                Next i
                location = location + wid - 8
            Next j
        Next xpos
    Next ypos
    MsgBox "DCT transform done!", vbOKOnly, "DCT transform"
    
End Sub
Private Sub Quanti(ByRef fdtbl() As Single)
    
    Dim temp As Single
    Dim DCT64(63) As Integer
    
    For i = 0 To 63
        DCT64(i) = DCT_Result(index_2)
        index_2 = index_2 + 1
    Next i
    
    For i = 0 To 63
        temp = DCT64(i) * fdtbl(i)
        Quantization_Result(index_3) = CInt(Fix(CInt(Fix(temp + 16384.5)) - 16384))
        index_3 = index_3 + 1
    Next i

End Sub

Private Sub quantization_Click()

    Dim xpos, ypos, i, j, k As Long
    ReDim Quantization_Result(wid * Hgt * 3 - 1) As Integer
    
    For ypos = 0 To Hgt - 1 Step 8
        For xpos = 0 To wid - 1 Step 8
        
            Call Quanti(fdtbl_Y)
            Call Quanti(fdtbl_U)
            Call Quanti(fdtbl_U)
                 
        Next xpos
    Next ypos
    
    'Display Quantization result
    For ypos = 0 To Hgt - 1 Step 8
        For xpos = 0 To wid - 1 Step 8
            location = wid * ypos + xpos
            For j = 0 To 7
                For i = 0 To 7
                    If wid <= 128 And Hgt <= 128 Then
                        Picture6.PSet (i + xpos, j + ypos), Quantization_Result(k)       ' Y
                        Picture7.PSet (i + xpos, j + ypos), Quantization_Result(k + 1)   ' U
                        Picture8.PSet (i + xpos, j + ypos), Quantization_Result(k + 2)   ' V
                    End If
                    k = k + 3
                    location = location + 1
                Next i
                location = location + wid - 8
            Next j
        Next xpos
    Next ypos
    
    MsgBox "Quantization done!", vbOKOnly, "Quantization"
    
End Sub


Private Sub ShiftCoding()
    Dim i, j As Long
    Dim nrmarker As Byte
    'Dim ShiftCode_Y_Output As String
    'Dim ShiftCode_U_Output As String
    'Dim ShiftCode_V_Output As String
    ReDim M16zeroes(wid * Hgt * 3 - 1) As Integer
    Dim f10, f11, f12 As Long
        
    'input is length of RLE output
    'ShiftCode_U_Output = "ShiftCode_U.txt"
    'ShiftCode_V_Output = "ShiftCode_V.txt"
    'ShiftCode_Y_Output = "ShiftCode_Y.txt"
    
    'f10 = FreeFile()
    'f11 = FreeFile() + 1
    'f12 = FreeFile() + 2
    
    'Open ShiftCode_Y_Output For Output Access Write As #f10
    'Open ShiftCode_U_Output For Output Access Write As #f11
    'Open ShiftCode_V_Output For Output Access Write As #f12
    
    'Print #f10, "Shift Coding input of [Y] component:"
    'Print #f10, "===================================="
    'For i = 0 To index_6 - 1 Step 3
    '    Print #f10, RLE_Output(i).length;
    'Next i
    
    'Print #f11, "Shift Coding input of [U] component:"
    'Print #f11, "==================================="
    'For i = 1 To index_6 - 1 Step 3
    '    Print #f11, RLE_Output(i).length;
    'Next i
    
    'Print #f12, "Shift Coding input [V] component:"
    'Print #f12, "================================"
    'For i = 2 To index_6 - 1 Step 3
    '    Print #f12, RLE_Output(i).length;
    'Next i
    
    For i = 0 To UBound(RLE_Output) - 1
        If RLE_Output(i).length >= 16 Then
            For nrmarker = 1 To RLE_Output(i).length \ 16
                j = j + 1
            Next nrmarker
            RLE_Output(i).length = RLE_Output(i).length Mod 16
            M16zeroes(i) = j
            j = 0
        End If
    Next i
    
    'Write output
    'Print #f10, ""
    'Print #f10, ""
    'Print #f10, "Shift Coding output of [Y] component:"
    'Print #f10, "===================================="
    'For i = 0 To index_6 - 1 Step 3
    '    If M16zeroes(i) > 0 Then
    '        Print #f10, "ZLE";
    '    End If
    '    Print #f10, RLE_Output(i).length;
    '    Print #f10, RLE_Output(i).size
    'Next i
    
    'Print #f11, ""
    'Print #f11, ""
    'Print #f11, "Shift Coding output of [U] component:"
    'Print #f11, "===================================="
    'For i = 1 To index_6 - 1 Step 3
    '    If M16zeroes(i) > 0 Then
    '        Print #f11, "ZLE";
    '    End If
    '    Print #f11, RLE_Output(i).length;
    '    Print #f11, RLE_Output(i).size
    'Next i
    
    'Print #f12, ""
    'Print #f12, ""
    'Print #f12, "Shift Coding output of [V] component:"
    'Print #f12, "===================================="
    'For i = 2 To index_6 - 1 Step 3
    '    If M16zeroes(i) > 0 Then
    '        Print #f12, "ZLE";
    '    End If
        
    '    Print #f12, RLE_Output(i).length;
    '    Print #f12, RLE_Output(i).size
    'Next i
    
    'Close #f10, #f11, #f12
    
    'MsgBox "Shift Coding Done!", vbOKOnly, "Shift Coding"
    
End Sub

Private Sub zigzagscan_Click()
    
    Dim i, j, pos, pos1, k As Long
    Dim zigzag64(63) As Integer
    Dim Quanti64(63) As Integer
    ReDim ZigZag_Result(wid * Hgt * 3 - 1) As Integer
    
    For ypos = 0 To Hgt - 1 Step 8
        For xpos = 0 To wid - 1 Step 8
            
            For j = 0 To 2
                
                For i = 0 To 63
                      Quanti64(i) = Quantization_Result(pos)
                      pos = pos + 1
                Next i
            
                For i = 0 To 63
                    zigzag64(zigzag(i)) = Quanti64(i)
                Next i
                
                For i = 0 To 63
                    ZigZag_Result(pos1) = zigzag64(i)
                    pos1 = pos1 + 1
                Next i
                
            Next j
            
        Next xpos
    Next ypos
    
    'Display zigzag result
    For ypos = 0 To Hgt - 1 Step 8
        For xpos = 0 To wid - 1 Step 8
            location = wid * ypos + xpos
            For j = 0 To 7
                For i = 0 To 7
                    If wid <= 128 And Hgt <= 128 Then
                        Picture9.PSet (i + xpos, j + ypos), ZigZag_Result(k)        ' Y
                        Picture10.PSet (i + xpos, j + ypos), ZigZag_Result(k + 1)   ' U
                        Picture11.PSet (i + xpos, j + ypos), ZigZag_Result(k + 2)   ' V
                    End If
                    k = k + 3
                    location = location + 1
                Next i
                location = location + wid - 8
            Next j
        Next xpos
    Next ypos
    
     MsgBox "zigzag done!", vbOKOnly, "Zigzag"
    
End Sub
Private Sub DPCM_cal(ByRef DC As Integer)
    Dim zigzag64(63) As Integer
    Dim i As Long
    
    For i = 0 To 63
        zigzag64(i) = ZigZag_Result(index_4)
        index_4 = index_4 + 1
    Next i
    
    Diff_Result(index_5) = zigzag64(0) - DC
    DC = zigzag64(0)
    index_5 = index_5 + 1
    
End Sub
Private Sub DPCM_Click()
    ReDim Diff_Result(wid * Hgt * 3 \ 64 - 1) As Long
    ReDim DC_Result(wid * Hgt \ 64 - 1) As Integer
    Dim xpos, ypos As Long
    Dim DCY As Integer
    Dim DCU As Integer
    Dim DCV As Integer
    Dim DPCM_Y_File As String
    Dim DPCM_U_File As String
    Dim DPCM_V_File As String
    Dim pos As Long
    Dim f1, f2, f3 As Long
    
    DPCM_U_File = "DPCM_U.txt"
    DPCM_V_File = "DPCM_V.txt"
    DPCM_Y_File = "DPCM_Y.txt"
    
    f1 = FreeFile()
    f2 = FreeFile() + 1
    f3 = FreeFile() + 2
    
    Open DPCM_Y_File For Output Access Write As #f1
    Open DPCM_U_File For Output Access Write As #f2
    Open DPCM_V_File For Output Access Write As #f3
     
    ' write input DPCM
    Print #f1, "DPCM input of [Y] component:"
    Print #f1, "==========================="
    Print #f2, "DPCM input of [U] component:"
    Print #f2, "==========================="
    Print #f3, "DPCM input of [V] component:"
    Print #f3, "==========================="
    
    For ypos = 0 To Hgt - 1 Step 8
        For xpos = 0 To wid - 1 Step 8
        
                For j = 0 To 63
                    Print #f1, ZigZag_Result(pos);
                    pos = pos + 1
                Next j
                
                For j = 0 To 63
                    Print #f2, ZigZag_Result(pos);
                    pos = pos + 1
                Next j
                
                For j = 0 To 63
                    Print #f3, ZigZag_Result(pos);
                    pos = pos + 1
                Next j

        Next xpos
    Next ypos
    
    'Calculate DPCM
    For ypos = 0 To Hgt - 1 Step 8
        For xpos = 0 To wid - 1 Step 8
            
            Call DPCM_cal(DCY)
            Call DPCM_cal(DCU)
            Call DPCM_cal(DCV)
            
        Next xpos
    Next ypos
    
    'Write Output DPCM
    Print #f1, ""
    Print #f1, ""
    Print #f1, "DPCM output of [Y] component:"
    Print #f1, "============================"
    For i = 0 To UBound(Diff_Result) - 1 Step 3
        Print #f1, Diff_Result(i);
    Next i
    
    Print #f2, ""
    Print #f2, ""
    Print #f2, "DPCM output of [U] component:"
    Print #f2, "============================"
    For i = 1 To UBound(Diff_Result) - 1 Step 3
        Print #f2, Diff_Result(i);
    Next i
    
    Print #f3, ""
    Print #f3, ""
    Print #f3, "DPCM output of [V] component:"
    Print #f3, "============================"
    For i = 2 To UBound(Diff_Result) - 1 Step 3
        Print #f3, Diff_Result(i);
    Next i
    
    Close #f1
    Close #f2
    Close #f3
    
    MsgBox "DPCM Done!", vbOKOnly, "DPCM"
    
End Sub
Private Function numOfBit(ByVal a As Integer)
    Dim num As Integer
    
    If (a < 0) Then
        a = Abs(a)
        num = 1
    End If
    
    Do While a > 0
        num = num + 1
        a = a \ 2
    Loop
    
    numOfBit = num
    
End Function

Private Sub RLE_Click()
    
    Dim xpos, ypos, i, j, k, pos1, pos2, pos4 As Long
    Dim startpos As Byte
    Dim end0pos As Byte
    Dim nrzeroes As Byte
    Dim zigzag64(63) As Integer
    ReDim RLE_Output(wid * Hgt * 3 - 1) As RLE_datatype
    Dim RLE_Y_File As String
    Dim RLE_U_File As String
    Dim RLE_V_File As String
    Dim temp1 As Integer
    Dim f2, f3, f4 As Long
    
    RLE_U_File = "RLE_U.txt"
    RLE_V_File = "RLE_V.txt"
    RLE_Y_File = "RLE_Y.txt"
    
    f2 = FreeFile()
    f3 = FreeFile() + 1
    f4 = FreeFile() + 2
    
    Open RLE_Y_File For Output Access Write As #f2
    Open RLE_U_File For Output Access Write As #f3
    Open RLE_V_File For Output Access Write As #f4
     
    'Write input RLE
    Print #f2, "RLE input of component [Y]:"
    Print #f2, "=========================="
    Print #f3, "RLE input of component [U]:"
    Print #f3, "=========================="
    Print #f4, "RLE input of component [V]:"
    Print #f4, "=========================="
    
    For ypos = 0 To Hgt - 1 Step 8
        For xpos = 0 To wid - 1 Step 8
        
                For j = 0 To 63
                    Print #f2, ZigZag_Result(pos1);
                    pos1 = pos1 + 1
                Next j
                
                For j = 0 To 63
                    Print #f3, ZigZag_Result(pos1);
                    pos1 = pos1 + 1
                Next j
                
                For j = 0 To 63
                    Print #f4, ZigZag_Result(pos1);
                    pos1 = pos1 + 1
                Next j

        Next xpos
    Next ypos
    
    'Write ouput
    Print #f2, " "
    Print #f2, " "
    Print #f2, "RLE output of [Y] component:"
    Print #f2, "==========================="
    Print #f3, " "
    Print #f3, " "
    Print #f3, "RLE output of [U] cmponent:"
    Print #f3, "==========================="
    Print #f4, " "
    Print #f4, " "
    Print #f4, "RLE output of [V] component:"
    Print #f4, "==========================="
    
    For ypos = 0 To Hgt - 1 Step 8
        For xpos = 0 To wid - 1 Step 8
                'Y
                For k = 0 To 63
                    zigzag64(k) = ZigZag_Result(pos2)
                    pos2 = pos2 + 1
                Next k
                
                end0pos = 63
                Do While (end0pos > 0) And (zigzag64(end0pos) = 0)
                    end0pos = end0pos - 1
                Loop
                
                i = 1
                Do While i <= end0pos
                    startpos = i
                    Do While (zigzag64(i) = 0) And (i <= end0pos)
                        i = i + 1
                    Loop
                    nrzeroes = i - startpos
                    RLE_Output(index_6).length = nrzeroes
                    RLE_Output(index_6).size = numOfBit(zigzag64(i))
                    index_6 = index_6 + 1
                    i = i + 1
                Loop
                    ' Mask end of block
                    RLE_Output(index_6).length = -1
                    RLE_Output(index_6).size = -1
                    index_6 = index_6 + 1
                 
                For i = pos4 To index_6 - 1
                    Print #f2, RLE_Output(i).length;
                    Print #f2, RLE_Output(i).size
                    pos4 = pos4 + 1
                Next i
                
                'U
                For k = 0 To 63
                    zigzag64(k) = ZigZag_Result(pos2)
                    pos2 = pos2 + 1
                Next k
                
                end0pos = 63
                Do While (end0pos > 0) And (zigzag64(end0pos) = 0)
                    end0pos = end0pos - 1
                Loop
                
                i = 1
                Do While i <= end0pos
                    startpos = i
                    Do While (zigzag64(i) = 0) And (i <= end0pos)
                        i = i + 1
                    Loop
                    nrzeroes = i - startpos
                    RLE_Output(index_6).length = nrzeroes
                    RLE_Output(index_6).size = numOfBit(zigzag64(i))
                    index_6 = index_6 + 1
                    i = i + 1
                Loop
                
                    ' Mask end of block
                    RLE_Output(index_6).length = -1
                    RLE_Output(index_6).size = -1
                    index_6 = index_6 + 1
                    
                For i = pos4 To index_6 - 1
                    Print #f3, RLE_Output(i).length;
                    Print #f3, RLE_Output(i).size
                    pos4 = pos4 + 1
                Next i
                            
                'V
                For k = 0 To 63
                    zigzag64(k) = ZigZag_Result(pos2)
                    pos2 = pos2 + 1
                Next k
                
                end0pos = 63
                Do While (end0pos > 0) And (zigzag64(end0pos) = 0)
                    end0pos = end0pos - 1
                Loop
                
                i = 1
                Do While i <= end0pos
                    startpos = i
                    Do While (zigzag64(i) = 0) And (i <= end0pos)
                        i = i + 1
                    Loop
                    nrzeroes = i - startpos
                    RLE_Output(index_6).length = nrzeroes
                    RLE_Output(index_6).size = numOfBit(zigzag64(i))
                    index_6 = index_6 + 1
                    i = i + 1
                Loop
                    ' Mask end of block
                    RLE_Output(index_6).length = -1
                    RLE_Output(index_6).size = -1
                    index_6 = index_6 + 1
                
                For i = pos4 To index_6 - 1
                    Print #f4, RLE_Output(i).length;
                    Print #f4, RLE_Output(i).size
                    pos4 = pos4 + 1
                Next i
                           
        Next xpos
    Next ypos
       
    Close #f2
    Close #f3
    Close #f4
    
    MsgBox "RLE Done!", vbOKOnly, "RLE"
    
End Sub

Private Sub init_variable()
    bytenew = 0
    bytepos = 7
    mask = Array(1, 2, 4, 8, 16, 32, 64, 128, 256, 512, 1024, 2048, 4096, 8192, 16384, 32768)

    zigzag = Array(0, 1, 5, 6, 14, 15, 27, 28, 2, 4, 7, 13, 16, 26, 29, 42, 3, 8, 12, 17, 25, 30, 41, 43, 9, 11, 18, 24, 31, 40, 44, 53, 10, 19, 23, 32, 39, 45, 52, 54, 20, 22, 33, 38, 46, 51, 55, 60, 21, 34, 37, 47, 50, 56, 59, 61, 35, 36, 48, 49, 57, 58, 62, 63)
    
    std_luminance_qt = Array(16, 11, 10, 16, 24, 40, 51, 61, 12, 12, 14, 19, 26, 58, 60, 55, 14, 13, 16, 24, 40, 57, 69, 56, 14, 17, 22, 29, 51, 87, 80, 62, 18, 22, 37, 56, 68, 109, 103, 77, 24, 35, 55, 64, 81, 104, 113, 92, 49, 64, 78, 87, 103, 121, 120, 101, 72, 92, 95, 98, 112, 100, 103, 99)
    std_chrominance_qt = Array(17, 18, 24, 47, 99, 99, 99, 99, 18, 21, 26, 66, 99, 99, 99, 99, 24, 26, 56, 99, 99, 99, 99, 99, 47, 66, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99, 99)
    
    std_dc_luminance_nrcodes = Array(0, 0, 1, 5, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0)
    std_dc_luminance_values = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
    
    std_dc_chrominance_nrcodes = Array(0, 0, 3, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0)
    std_dc_chrominance_values = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
    
    std_ac_luminance_nrcodes = Array(0, 0, 2, 1, 3, 3, 2, 4, 3, 5, 5, 4, 4, 0, 0, 1, &H7D)
    std_ac_luminance_values = Array(&H1, &H2, &H3, &H0, &H4, &H11, &H5, &H12, &H21, &H31, &H41, &H6, &H13, &H51, &H61, &H7, &H22, &H71, &H14, &H32, &H81, &H91, &HA1, &H8, &H23, &H42, &HB1, &HC1, &H15, &H52, &HD1, &HF0, &H24, &H33, &H62, &H72, &H82, &H9, &HA, &H16, &H17, &H18, &H19, &H1A, &H25, &H26, &H27, &H28, &H29, &H2A, &H34, &H35, &H36, &H37, &H38, &H39, &H3A, &H43, &H44, &H45, &H46, &H47, &H48, &H49, &H4A, &H53, &H54, &H55, &H56, &H57, &H58, &H59, &H5A, &H63, &H64, &H65, &H66, &H67, &H68, &H69, &H6A, &H73, &H74, &H75, &H76, &H77, &H78, &H79, &H7A, &H83, &H84, &H85, &H86, &H87, &H88, &H89, &H8A, &H92, &H93, &H94, &H95, &H96, &H97, &H98, &H99, &H9A, &HA2, &HA3, &HA4, &HA5, &HA6, &HA7, &HA8, &HA9, &HAA, &HB2, &HB3, &HB4, &HB5, &HB6, &HB7, &HB8, &HB9, &HBA, &HC2, &HC3, &HC4, &HC5, &HC6, &HC7, &HC8, &HC9, &HCA, &HD2, &HD3, &HD4, &HD5, &HD6, &HD7, &HD8, &HD9, &HDA, &HE1, &HE2, &HE3, &HE4, &HE5, &HE6, &HE7, &HE8, &HE9, &HEA, &HF1, &HF2, &HF3, &HF4, &HF5, &HF6, &HF7, &HF8, &HF9, &HFA)
    
    std_ac_chrominance_nrcodes = Array(0, 0, 2, 1, 2, 4, 4, 3, 4, 7, 5, 4, 4, 0, 1, 2, &H77)
    std_ac_chrominance_values = Array(&H0, &H1, &H2, &H3, &H11, &H4, &H5, &H21, &H31, &H6, &H12, &H41, &H51, &H7, &H61, &H71, &H13, &H22, &H32, &H81, &H8, &H14, &H42, &H91, &HA1, &HB1, &HC1, &H9, &H23, &H33, &H52, &HF0, &H15, &H62, &H72, &HD1, &HA, &H16, &H24, &H34, &HE1, &H25, &HF1, &H17, &H18, &H19, &H1A, &H26, &H27, &H28, &H29, &H2A, &H35, &H36, &H37, &H38, &H39, &H3A, &H43, &H44, &H45, &H46, &H47, &H48, &H49, &H4A, &H53, &H54, &H55, &H56, &H57, &H58, &H59, &H5A, &H63, &H64, &H65, &H66, &H67, &H68, &H69, &H6A, &H73, &H74, &H75, &H76, &H77, &H78, &H79, &H7A, &H82, &H83, &H84, &H85, &H86, &H87, &H88, &H89, &H8A, &H92, &H93, &H94, &H95, &H96, &H97, &H98, &H99, &H9A, &HA2, &HA3, &HA4, &HA5, &HA6, &HA7, &HA8, &HA9, &HAA, &HB2, &HB3, &HB4, &HB5, &HB6, &HB7, &HB8, &HB9, &HBA, &HC2, &HC3, &HC4, &HC5, &HC6, &HC7, &HC8, &HC9, &HCA, &HD2, &HD3, &HD4, &HD5, &HD6, &HD7, &HD8, &HD9, &HDA, &HE2, &HE3, &HE4, &HE5, &HE6, &HE7, &HE8, &HE9, &HEA, &HF2, &HF3, &HF4, &HF5, &HF6, &HF7, &HF8, &HF9, &HFA)
    
    APP0info.marker = &HFFE0&
    APP0info.length = 16
    APP0info.JFIFsignature(0) = 74 'J
    APP0info.JFIFsignature(1) = 70 'F
    APP0info.JFIFsignature(2) = 73 'I
    APP0info.JFIFsignature(3) = 70 'F
    APP0info.JFIFsignature(4) = 0 'F
    APP0info.versionhi = 1
    APP0info.versionlo = 1
    APP0info.xyunits = 0
    APP0info.xdensity = 1
    APP0info.ydensity = 1
    APP0info.thumbnheight = 0
    APP0info.thumbnwidth = 0
    
    SOF0info.marker = &HFFC0&
    SOF0info.length = 17
    SOF0info.precision = 8
    SOF0info.height = 0
    SOF0info.width = 0
    SOF0info.nrofcomponents = 3
    SOF0info.IdY = 1
    SOF0info.HVY = &H11
    SOF0info.QTY = 0
    SOF0info.IdCb = 2
    SOF0info.HVCb = &H11
    SOF0info.QTCb = 1
    SOF0info.IdCr = 3
    SOF0info.HVCr = &H11
    SOF0info.QTCr = 1
        
    SOSinfo.marker = &HFFDA&
    SOSinfo.length = 12
    SOSinfo.nrofcomponents = 3
    SOSinfo.IdY = 1
    SOSinfo.HTY = 0
    SOSinfo.IdCb = 2
    SOSinfo.HTCb = &H11
    SOSinfo.IdCr = 3
    SOSinfo.HTCr = &H11
    SOSinfo.Ss = 0
    SOSinfo.Se = &H3F
    SOSinfo.Bf = 0
        
End Sub
Private Sub write_APP0info()
    Put #1, , CByte(APP0info.marker \ 256)
    Put #1, , CByte(APP0info.marker Mod 256)
    Put #1, , CByte(APP0info.length \ 256)
    Put #1, , CByte(APP0info.length Mod 256)
    Put #1, , CByte(74) 'J
    Put #1, , CByte(70) 'F
    Put #1, , CByte(73) 'I
    Put #1, , CByte(70) 'F
    Put #1, , CByte(0)
    Put #1, , CByte(APP0info.versionhi)
    Put #1, , CByte(APP0info.versionlo)
    Put #1, , CByte(APP0info.xyunits)
    Put #1, , CByte(APP0info.xdensity \ 256)
    Put #1, , CByte(APP0info.xdensity Mod 256)
    Put #1, , CByte(APP0info.ydensity \ 256)
    Put #1, , CByte(APP0info.ydensity Mod 256)
    Put #1, , CByte(APP0info.thumbnwidth)
    Put #1, , CByte(APP0info.thumbnheight)
End Sub

Private Sub write_SOF0info()
    ' We should overwrite width and height
     Put #1, , CByte(SOF0info.marker \ 256)
     Put #1, , CByte(SOF0info.marker Mod 256)
     Put #1, , CByte(SOF0info.length \ 256)
     Put #1, , CByte(SOF0info.length Mod 256)
     Put #1, , CByte(SOF0info.precision)
     Put #1, , CByte(SOF0info.height \ 256)
     Put #1, , CByte(SOF0info.height Mod 256)
     Put #1, , CByte(SOF0info.width \ 256)
     Put #1, , CByte(SOF0info.width Mod 256)
     Put #1, , CByte(SOF0info.nrofcomponents)
     Put #1, , CByte(SOF0info.IdY)
     Put #1, , CByte(SOF0info.HVY)
     Put #1, , CByte(SOF0info.QTY)
     Put #1, , CByte(SOF0info.IdCb)
     Put #1, , CByte(SOF0info.HVCb)
     Put #1, , CByte(SOF0info.QTCb)
     Put #1, , CByte(SOF0info.IdCr)
     Put #1, , CByte(SOF0info.HVCr)
     Put #1, , CByte(SOF0info.QTCr)
 End Sub

Private Sub write_DQTinfo()
    Dim i As Byte

    Put #1, , CByte(DQTinfo.marker \ 256)
    Put #1, , CByte(DQTinfo.marker Mod 256)
    Put #1, , CByte(DQTinfo.length \ 256)
    Put #1, , CByte(DQTinfo.length Mod 256)
    Put #1, , CByte(DQTinfo.QTYinfo)
    For i = 0 To 63
        Put #1, , CByte(DQTinfo.Ytable(i))
    Next i
    
    Put #1, , CByte(DQTinfo.QTCbinfo)
    
    For i = 0 To 63
       Put #1, , CByte(DQTinfo.Cbtable(i))
    Next i
End Sub
Private Sub write_DHTinfo()
    Dim i As Byte
    
    Put #1, , CByte(DHTinfo.marker \ 256)
    Put #1, , CByte(DHTinfo.marker Mod 256)
    Put #1, , CByte(DHTinfo.length \ 256)
    Put #1, , CByte(DHTinfo.length Mod 256)
    Put #1, , CByte(DHTinfo.HTYDCinfo)
    For i = 0 To 15
        Put #1, , CByte(DHTinfo.YDC_nrcodes(i))
    Next i
    For i = 0 To 11
        Put #1, , CByte(DHTinfo.YDC_values(i))
    Next i
    
    Put #1, , CByte(DHTinfo.HTYACinfo)
    
    For i = 0 To 15
        Put #1, , CByte(DHTinfo.YAC_nrcodes(i))
    Next i
    For i = 0 To 161
        Put #1, , CByte(DHTinfo.YAC_values(i))
    Next i
    Put #1, , CByte(DHTinfo.HTCbDCinfo)
    For i = 0 To 15
        Put #1, , CByte(DHTinfo.CbDC_nrcodes(i))
    Next i
    For i = 0 To 11
        Put #1, , CByte(DHTinfo.CbDC_values(i))
    Next i
    Put #1, , DHTinfo.HTCbACinfo
    For i = 0 To 15
        Put #1, , CByte(DHTinfo.CbAC_nrcodes(i))
    Next i
    For i = 0 To 161
        Put #1, , CByte(DHTinfo.CbAC_values(i))
    Next i
End Sub
Private Sub write_SOSinfo()
'Nothing to overwrite for SOSinfo
    Put #1, , CByte((SOSinfo.marker) \ 256)
    Put #1, , CByte((SOSinfo.marker) Mod 256)
    Put #1, , CByte((SOSinfo.length) \ 256)
    Put #1, , CByte((SOSinfo.length) Mod 256)
    Put #1, , CByte(SOSinfo.nrofcomponents)
    Put #1, , CByte(SOSinfo.IdY)
    Put #1, , CByte(SOSinfo.HTY)
    Put #1, , CByte(SOSinfo.IdCb)
    Put #1, , CByte(SOSinfo.HTCb)
    Put #1, , CByte(SOSinfo.IdCr)
    Put #1, , CByte(SOSinfo.HTCr)
    Put #1, , CByte(SOSinfo.Ss)
    Put #1, , CByte(SOSinfo.Se)
    Put #1, , CByte(SOSinfo.Bf)
End Sub
Private Function writebits(ByRef bs As bitstring)
    Dim value As Long
    Dim posval As Long
 
    value = bs.value
    posval = bs.length - 1
    
    Do While (posval >= 0)
        If (value And mask(posval)) Then
            bytenew = bytenew Or mask(bytepos)
        End If
        
        posval = posval - 1
        bytepos = bytepos - 1
        
        
        If (bytepos < 0) Then
            ' write it
            If (bytenew = 255) Then
               ' special case
                Put #1, , CByte(255) 'define more
                Put #1, , CByte(0)  ' define more
            Else
                Put #1, , CByte(bytenew) 'define more
            End If
            'reinit
            bytepos = 7
            bytenew = 0
        End If
    Loop
    
End Function
Private Sub compute_Huffman_table(ByRef nrcodes() As Variant, ByRef std_table() As Variant, ByRef HT() As bitstring)
    Dim k As Byte
    Dim j As Byte
    Dim index_in_table As Byte
    Dim codevalue As Long

    codevalue = 0
    index_in_table = 0
    
    For k = 1 To 16
        For j = 1 To nrcodes(k)
            HT(std_table(index_in_table)).value = codevalue
            HT(std_table(index_in_table)).length = k
            index_in_table = index_in_table + 1
            codevalue = codevalue + 1
        Next j
        codevalue = codevalue * 2
    Next k
End Sub
Private Sub set_quant_table(ByRef basic_table() As Variant, ByVal scale_factor As Byte, ByRef newtable() As Byte)
      Dim i As Byte
      Dim temp As Long

      For i = 0 To 63
          temp = (CLng(Fix(basic_table(i))) * scale_factor + CLng(50)) \ CLng(100)
          If temp <= CLng(0) Then
              temp = CLng(1)
          End If
          If temp > CLng(255) Then
              temp = CLng(255)
          End If
          newtable(zigzag(i)) = CByte(temp)
      Next i
End Sub
Private Sub prepare_quant_tables()
    Dim aanscalefactor() As Variant
    aanscalefactor = Array(1#, 1.387039845, 1.306562965, 1.175875602, 1#, 0.785694958, 0.5411961, 0.275899379)
    Dim row As Byte
    Dim col As Byte
    Dim i As Byte
    i = 0
    
    For row = 0 To 7
        For col = 0 To 7
            fdtbl_Y(i) = CSng(1# / (CDbl(DQTinfo.Ytable(zigzag(i))) * aanscalefactor(row) * aanscalefactor(col) * 8#))
            fdtbl_U(i) = CSng(1# / (CDbl(DQTinfo.Cbtable(zigzag(i))) * aanscalefactor(row) * aanscalefactor(col) * 8#))
            i = i + 1
        Next col
    Next row
End Sub
Private Sub precalculate_YUV_tables()
    Dim R As Long
    Dim G As Long
    Dim B As Long

    For R = 0 To 255
        YRtab(R) = CLng(Fix(65536 * 0.299 + 0.5)) * R
        URtab(R) = CLng(Fix(65536 * -0.16874 + 0.5)) * R
        VRtab(R) = CLng(Fix(32768)) * R
    Next R
    For G = 0 To 255
        YGtab(G) = CLng(Fix(65536 * 0.587 + 0.5)) * G
        UGtab(G) = CLng(Fix(65536 * -0.33126 + 0.5)) * G
        VGtab(G) = CLng(Fix(65536 * -0.41869 + 0.5)) * G
    Next G
    For B = 0 To 255
        YBtab(B) = CLng(Fix(65536 * 0.114 + 0.5)) * B
        UBtab(B) = CLng(Fix(32768)) * B
        VBtab(B) = CLng(Fix(65536 * -0.08131 + 0.5)) * B
    Next B
End Sub
Private Sub set_numbers_category_and_bitcode()
    Dim nr As Long
    Dim nrlower As Long
    Dim nrupper As Long
    Dim cat As Byte

    nrlower = 1
    nrupper = 2
    
    For cat = 1 To 15
        'Positive numbers
        For nr = nrlower + 32767 To nrupper - 1 + 32767
            category(nr) = cat
            bitcode(nr).length = cat
            bitcode(nr).value = CLng(nr - 32767)
        Next nr
        
        'Negative numbers
        For nr = -(nrupper - 1) + 32767 To -nrlower + 32767
            category(nr) = cat
            bitcode(nr).length = cat
            bitcode(nr).value = CLng(nrupper - 1 + nr - 32767)
            
        Next nr
        nrlower = nrlower * 2
        nrupper = nrupper * 2
    Next cat
End Sub
Private Sub init_Huffman_tables()
    Call compute_Huffman_table(std_dc_luminance_nrcodes, std_dc_luminance_values, YDC_HT)
    Call compute_Huffman_table(std_ac_luminance_nrcodes, std_ac_luminance_values, YAC_HT)
    Call compute_Huffman_table(std_dc_chrominance_nrcodes, std_dc_chrominance_values, UDC_HT)
    Call compute_Huffman_table(std_ac_chrominance_nrcodes, std_ac_chrominance_values, UAC_HT)
End Sub
Private Sub set_DHTinfo()
    Dim i As Byte
    DHTinfo.marker = &HFFC4&
    DHTinfo.length = &H1A2&
    DHTinfo.HTYDCinfo = 0
    
    For i = 0 To 15
        DHTinfo.YDC_nrcodes(i) = std_dc_luminance_nrcodes(i + 1)
    Next i
    For i = 0 To 11
        DHTinfo.YDC_values(i) = std_dc_luminance_values(i)
    Next i

    DHTinfo.HTYACinfo = &H10
    For i = 0 To 15
        DHTinfo.YAC_nrcodes(i) = std_ac_luminance_nrcodes(i + 1)
    Next i
    For i = 0 To 161
        DHTinfo.YAC_values(i) = std_ac_luminance_values(i)
    Next i

    DHTinfo.HTCbDCinfo = 1
    For i = 0 To 15
        DHTinfo.CbDC_nrcodes(i) = std_dc_chrominance_nrcodes(i + 1)
    Next i
    For i = 0 To 11
        DHTinfo.CbDC_values(i) = std_dc_chrominance_values(i)
    Next i

    DHTinfo.HTCbACinfo = &H11
    For i = 0 To 15
        DHTinfo.CbAC_nrcodes(i) = std_ac_chrominance_nrcodes(i + 1)
    Next i
    For i = 0 To 161
        DHTinfo.CbAC_values(i) = std_ac_chrominance_values(i)
    Next i
End Sub
Private Sub set_DQTinfo()
    Dim scalefactor As Byte
    
    scalefactor = 50
    DQTinfo.marker = &HFFDB&
    DQTinfo.length = 132
    DQTinfo.QTYinfo = 0
    DQTinfo.QTCbinfo = 1
    Call set_quant_table(std_luminance_qt, scalefactor, DQTinfo.Ytable)
    Call set_quant_table(std_chrominance_qt, scalefactor, DQTinfo.Cbtable)
End Sub
Private Sub load_data_units_from_RGB_buffer(ByVal xpos As Long, ByVal ypos As Long)
    Dim x As Byte
    Dim y As Byte
    Dim pos As Byte
    pos = 0
    Dim location As Currency
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    Dim x1 As Currency

    location = ypos * wid + xpos
    For y = 0 To 7
        For x = 0 To 7
            R = RGB_buffer(location).R
            G = RGB_buffer(location).G
            B = RGB_buffer(location).B
            
            x1 = ((YRtab(R) + YGtab(G) + YBtab(B)) / (2 ^ 16)) - 128
            If x1 < 0 Then
              x1 = Fix(x1) - 1
            Else
                x1 = Fix(x1)
            End If
            Y_DU(pos) = x1
            
            x1 = (URtab(R) + UGtab(G) + UBtab(B)) / (2 ^ 16)
            If x1 < 0 Then
                 x1 = Fix(x1) - 1
            Else
                x1 = Fix(x1)
            End If
            
            U_DU(pos) = x1
            
            x1 = (VRtab(R) + VGtab(G) + VBtab(B)) / (2 ^ 16)
            If x1 < 0 Then
                 x1 = Fix(x1) - 1
            Else
                x1 = Fix(x1)
            End If
            
            V_DU(pos) = x1
            
            
            Y_color(posi_2) = Y_DU(pos)
            U_color(posi_2) = U_DU(pos)
            V_color(posi_2) = V_DU(pos)
            
            location = location + 1
            pos = pos + 1
            posi_2 = posi_2 + 1
            
        Next x
        location = location + wid - 8
    Next y
End Sub
Private Sub load_bitmap(ByVal bitmap_name As String, ByRef width_original As Long, ByRef height_original As Long)
        
    Dim widthDiv8 As Long
    Dim heightDiv8 As Long
    Dim nr_fillingbytes As Byte
    Dim lastcolor As colorRGB
    Dim column As Long
    Dim TMPBUF(253) As Byte
    Dim tmp() As Byte
    Dim nrline_up As Long
    Dim nrline_dn As Long
    Dim nrline As Long
    Dim dimline As Long
    Dim tmpline() As colorRGB 'pointer
    Dim i, j, m As Integer
    Dim k As Long
    k = 0

    ReDim tmp(53) As Byte
    Open bitmap_name For Binary Access Read As #2
    Get #2, , tmp
    
    For m = 0 To 53
        TMPBUF(m) = tmp(m)
    Next m
    
    If (TMPBUF(0) <> 66) Or (TMPBUF(1) <> 77) Or (TMPBUF(28) <> 24) Then
        MsgBox "Need a truecolor BMP to encode", vbCritical
        End
    End If
    
    wid = (TMPBUF(19)) * 256 + TMPBUF(18)
    Hgt = (TMPBUF(23)) * 256 + TMPBUF(22)

    ' Keep the old dimensions of the image
    width_original = wid
    height_original = Hgt

    If wid Mod 8 <> 0 Then
        widthDiv8 = (wid \ 8) * 8 + 8
    Else
        widthDiv8 = wid
    End If

    If Hgt Mod 8 <> 0 Then
       heightDiv8 = (Hgt \ 8) * 8 + 8
    Else
      heightDiv8 = Hgt
    End If

    ReDim RGB_buffer(widthDiv8 * heightDiv8 - 1) As colorRGB

    If (wid * 3) Mod 4 <> 0 Then
        nr_fillingbytes = 4 - ((wid * 3) Mod 4)
    Else
       nr_fillingbytes = 0
    End If
     
    For nrline = 0 To Hgt - 1
    
        ReDim tmp(wid * 3 - 1) As Byte
        Get #2, , tmp
        
        For m = 0 To wid * 3 - 1 Step 3
            RGB_buffer((nrline * widthDiv8 + m \ 3)).B = tmp(m)
            RGB_buffer((nrline * widthDiv8 + m \ 3)).G = tmp(m + 1)
            RGB_buffer((nrline * widthDiv8 + m \ 3)).R = tmp(m + 2)
        Next m
    
        If nr_fillingbytes > 0 Then
            ReDim tmp(nr_fillingbytes) As Byte
            Get #2, , tmp
        End If
           
        lastcolor.B = RGB_buffer(nrline * widthDiv8 + wid - 1).B
        lastcolor.G = RGB_buffer(nrline * widthDiv8 + wid - 1).G
        lastcolor.R = RGB_buffer(nrline * widthDiv8 + wid - 1).R
        
        For column = wid To widthDiv8 - 1
            RGB_buffer(nrline * widthDiv8 + column).B = lastcolor.B
            RGB_buffer(nrline * widthDiv8 + column).G = lastcolor.G
            RGB_buffer(nrline * widthDiv8 + column).R = lastcolor.R
        Next column
    Next nrline

    wid = widthDiv8
    dimline = wid
    ReDim tmpline(dimline - 1) As colorRGB
    
    nrline_up = Hgt - 1
    nrline_dn = 0
    
    Do While nrline_up > nrline_dn
    
        For m = 0 To dimline - 1
            tmpline(m).B = RGB_buffer(nrline_up * wid + m).B
            tmpline(m).G = RGB_buffer(nrline_up * wid + m).G
            tmpline(m).R = RGB_buffer(nrline_up * wid + m).R
        Next m
        
        For m = 0 To dimline - 1
            RGB_buffer(nrline_up * wid + m).B = RGB_buffer(nrline_dn * wid + m).B
            RGB_buffer(nrline_up * wid + m).G = RGB_buffer(nrline_dn * wid + m).G
            RGB_buffer(nrline_up * wid + m).R = RGB_buffer(nrline_dn * wid + m).R
        Next m
        
        For m = 0 To dimline - 1
            RGB_buffer(nrline_dn * wid + m).B = tmpline(m).B
            RGB_buffer(nrline_dn * wid + m).G = tmpline(m).G
            RGB_buffer(nrline_dn * wid + m).R = tmpline(m).R
        Next m
        
        nrline_up = nrline_up - 1
        nrline_dn = nrline_dn + 1

    Loop
    
    For m = 0 To dimline - 1
           tmpline(m).B = RGB_buffer((Hgt - 1) * wid + m).B
           tmpline(m).G = RGB_buffer((Hgt - 1) * wid + m).G
           tmpline(m).R = RGB_buffer((Hgt - 1) * wid + m).R
    Next m
    
    For nrline = Hgt To heightDiv8 - 1
    
        For m = 0 To dimline - 1
           RGB_buffer(nrline * wid + m).B = tmpline(m).B
           RGB_buffer(nrline * wid + m).G = tmpline(m).G
           RGB_buffer(nrline * wid + m).R = tmpline(m).R
        Next m
        
    Next nrline
    
    Hgt = heightDiv8
    
    Close #2
        
End Sub
Private Sub init_all()
    Call init_variable
    Call set_DQTinfo
    Call set_DHTinfo
    Call init_Huffman_tables
    Call set_numbers_category_and_bitcode
    Call precalculate_YUV_tables
    Call prepare_quant_tables
End Sub
Private Sub main()

    Dim width_original As Long
    Dim height_original As Long
    Dim fillbits As bitstring
   
       
    Call load_bitmap(BMP_filename, width_original, height_original)
    Call init_all
       
    SOF0info.width = width_original
    SOF0info.height = height_original

    bytenew = 0
    bytepos = 7
    
End Sub
Private Sub Huffman_Encoder(ByRef HTDC() As bitstring, ByRef HTAC() As bitstring)
    Dim tmp, i As Integer
    Dim nrzeroes As Integer
    Dim end0pos As Integer
    Dim startpos As Integer
    Dim zigzag64(63) As Long
    Dim EOB As bitstring
    EOB = HTAC(0)
    
    'DC
    If Diff_Result(index_1) = 0 Then
        Call writebits(HTDC(0))

    Else
        Call writebits(HTDC(category(Diff_Result(index_1) + 32767)))
        Call writebits(bitcode(Diff_Result(index_1) + 32767))
    End If
    
   
    
    For k = 0 To 63
        zigzag64(k) = ZigZag_Result(index_3)
        index_3 = index_3 + 1
    Next k
    
    'AC
    end0pos = 63
    Do While (end0pos > 0) And (zigzag64(end0pos) = 0)
        end0pos = end0pos - 1
    Loop
                
    i = 1
    Do While i <= end0pos
        startpos = i
        Do While (zigzag64(i) = 0) And (i <= end0pos)
                i = i + 1
        Loop
        
       'Shift coding
        tmp = M16zeroes(index_4)
        Do While tmp > 0
            Call writebits(HTAC(240))
        
            tmp = tmp - 1
        Loop
        index_4 = index_4 + 1
      
        Call writebits(HTAC(RLE_Output(index_2).length * 16 + category(zigzag64(i) + 32767)))
        Call writebits(bitcode(zigzag64(i) + 32767))

        i = i + 1
        index_2 = index_2 + 1
    Loop
    
    If end0pos <> 63 Then
        Call writebits(EOB)
        
    End If
    
    index_1 = index_1 + 1
    
End Sub
Private Sub Hufman_Click()

    Call ShiftCoding
        
    Dim Huffman_Input As String
    Dim pos As Long
    
    Huffman_Input = "Huffman_Input.txt"
    Huffman_Output = "Huffman_Output.txt"
    
     
    Open Huffman_Output For Binary Access Write As #1
    Open Huffman_Input For Output As #2
    
    index_1 = 0
    index_2 = 0
    index_3 = 0
    index_4 = 0
    
    'Write input: DPCM,RLE,Shift coding
    Print #2, "DPCM component:"
    Print #2, "==============="
    For i = 0 To UBound(Diff_Result) - 1
        Print #2, Diff_Result(pos);
        pos = pos + 1
    Next i
    
    Print #2, , ""
    Print #2, "Shift Coding component:"
    Print #2, "======================="
    For i = 0 To index_6 - 1
        If M16zeroes(i) > 0 Then
            Print #2, "ZLE";
        End If
        Print #2, RLE_Output(i).length;
        Print #2, RLE_Output(i).size
    Next i
    
    Huffman encoding
    For ypos = 0 To Hgt - 1 Step 8
        For xpos = 0 To wid - 1 Step 8
            Call Huffman_Encoder(YDC_HT, YAC_HT)
            Call Huffman_Encoder(UDC_HT, UAC_HT)
            Call Huffman_Encoder(UDC_HT, UAC_HT)
        Next xpos
    Next ypos
    
    size_huffman_output = Seek(1)
    
    Close #1
    Close #2
    
    MsgBox "Huffman Shift Coding Done!", vbOKOnly, "Huffman Shift Coding"

End Sub
Private Sub save_Click()
    ReDim data_jpeg(size_huffman_output) As Byte
    Dim xpos, ypos As Long
    JPG_filename = BMP_filename
    JPG_filename = Replace(JPG_filename, ".bmp", ".jpg")
    
    Open JPG_filename For Binary Access Write As #1
    Open Huffman_Output For Binary Access Read As #2
    
    'Write header
    Put #1, , CByte(&HFFD8& \ 256)
    Put #1, , CByte(&HFFD8& Mod 256)
    Call write_APP0info
    Call write_DQTinfo
    Call write_SOF0info
    Call write_DHTinfo
    Call write_SOSinfo
    
    'Write huffman output
    Get #2, , data_jpeg
    Put #1, , data_jpeg
    
    'Write footer
    Dim fillbits As bitstring
    fillbits.length = 7
    fillbits.value = 2 ^ 7 - 1
    Call writebits(fillbits)
    
    Put #1, , CByte(&HFFD9& \ 256)
    Put #1, , CByte(&HFFD9& Mod 256)
    
    Close #1, #2

    MsgBox "Save Done!", vbOKOnly, "Save"
    
End Sub
' Sort length decreasing
Private Sub sort_length(ByRef data() As RLE_datatype)
     Dim i, j, num As Long
     Dim tmp As RLE_datatype
     
     Do While Not (data(i).length = 0 And data(i).size = 0)
        num = num + 1
        i = i + 1
     Loop
     
    For i = 1 To num - 1
        j = i
        Do While j > 0
         
         If data(j - 1).length > data(j).length Then
              tmp = data(j)
              data(j) = data(j - 1)
              data(j - 1) = tmp
         End If
         j = j - 1
        Loop
    Next i

End Sub
' Sort size decreasing in the same length
Private Sub sort_size(ByRef data() As RLE_datatype)
    Dim x, y, i, j, min, temp1, temp2, index, index1, num, num1, head As Long
    head = 0
    
     Do While Not (data(i).length = 0 And data(i).size = 0)
        num1 = num1 + 1
        i = i + 1
     Loop
    
    For x = 1 To num1 - 1
        
        If (data(x).length = data(x - 1).length) Then
            If x = num1 - 1 Then
                For i = head To x - 1
                      min = i
                      
                      For j = i + 1 To x
                            If (data(j).size < data(min).size) Then
                                min = j
                            End If
                      Next j
                    
                      temp1 = data(i).length
                      temp2 = data(i).size
                      
                      data(i).length = data(min).length
                      data(i).size = data(min).size
                      
                      data(min).length = temp1
                      data(min).size = temp2
            
                Next i
            End If
            Else
                For i = head To x - 2
                      min = i
                      
                      For j = i + 1 To x - 1
                            If (data(j).size < data(min).size) Then
                                min = j
                            End If
                      Next j
                    
                      temp1 = data(i).length
                      temp2 = data(i).size
                      
                      data(i).length = data(min).length
                      data(i).size = data(min).size
                      
                      data(min).length = temp1
                      data(min).size = temp2
            
                Next i
                    head = x
            End If
    Next x
End Sub
' Reduce value & add number appear
Private Sub count_probability(ByRef data() As RLE_datatype)
    
    Dim temp As RLE_datatype
    Dim i, num, pos1 As Long
    num = 1
    temp = data(0)
    
    For i = 0 To 63
        rle_prob_output(i).number_appearance = 0
        rle_prob_output(i).value.length = 0
        rle_prob_output(i).value.size = 0
    Next i
    
    For i = 1 To UBound(data)
        If (data(i).length = data(i - 1).length) And (data(i).size = data(i - 1).size) Then
            num = num + 1
        Else
            rle_prob_output(pos1).value = temp
            rle_prob_output(pos1).number_appearance = num
            pos1 = pos1 + 1
            num = 1
            temp = data(i)
        End If
        
        If (i = UBound(data)) Then
                rle_prob_output(pos1).value = temp
                rle_prob_output(pos1).number_appearance = num
                pos1 = pos1 + 1
                temp = data(i)
            End If
    Next i
    
End Sub
' Sort probability decreasing
Private Sub sort_probability_deacreasing()
     Dim i, j, total As Long
     Dim tmp As rle_probability
     
     Do While Not (rle_prob_output(i).value.length = 0 And rle_prob_output(i).value.size = 0)
        total = total + 1
        i = i + 1
     Loop
    
     For i = 1 To total - 1
        j = i
        Do While j > 0
         
         If rle_prob_output(j - 1).number_appearance < rle_prob_output(j).number_appearance Then
              tmp = rle_prob_output(j)
              rle_prob_output(j) = rle_prob_output(j - 1)
              rle_prob_output(j - 1) = tmp
         End If
         j = j - 1
        Loop
    Next i
End Sub
' Divide into equal blocks then and each block the unique symbol then write into file
Private Sub divide_into_equal_block_and_add_unique_symbol()
    Dim i, j, total, block_divided, distance As Long
    Dim binary_code() As Long
    
    ' Divide into equal block
    Do While Not (rle_prob_output(i).value.length = 0 And rle_prob_output(i).value.size = 0)
        total = total + 1
        i = i + 1
    Loop
    
    For i = 1 To total
        If total Mod i = 0 And i > 1 Then
            block_divided = i
            Exit For
        End If
    Next
    
    If total = 1 Then
        block_divided = 1
    End If
    
    distance = total / block_divided
    
    ReDim binary_code(total - 1) As Long
    
    For i = 0 To total - 1
        binary_code(i) = j
        j = j + 1
        
        If j = distance Then
            j = 0
        End If
    Next i
    
    ' Add symbol
    For i = 1 To block_divided
    Next i
    
    'Write into file
    
End Sub
Private Sub cal_probability(ByRef data() As RLE_datatype)
    Call sort_length(data)
    Call sort_size(data)
    Call count_probability(data)
    Call sort_probability_deacreasing
End Sub
Private Sub Binary_Shift_Coding(ByRef data() As RLE_datatype)
    Call cal_probability(data)
    Call divide_into_equal_block_and_add_unique_symbol
End Sub
Private Sub BinaryShift_Click()
    
    Dim Binary_Shift As String
    Dim pos1, pos2 As Long
    Dim tmp(63) As RLE_datatype
    Dim i, j As Long
    
    Binary_Shift = "Binary_Shift_Coding.txt"
    file = FreeFile()
    
    Open Binary_Shift For Binary Access Write As #100
    
    For ypos = 0 To Hgt - 1 Step 8
        For xpos = 0 To wid - 1 Step 8
            
            ' ********* Y ********
            
            ' Write DC
            Put #100, , Diff_Result(pos1)
            pos1 = pos1 + 1
            
            Do While (RLE_Output(pos2).length <> -1)
                tmp(i) = RLE_Output(pos2)
                i = i + 1
                pos2 = pos2 + 1
            Loop
            
            pos2 = pos2 + 1
            For j = i To 63
                tmp(j).length = 0
                tmp(j).size = 0
            Next j
            
            ' Binary Shift Coding
            If i > 0 Then
                Call Binary_Shift_Coding(tmp)
            End If
            i = 0

            ' ********* U ********
            ' Write DC
            Put #100, , Diff_Result(pos1)
            pos1 = pos1 + 1
            
            Do While (RLE_Output(pos2).length <> -1)
                tmp(i) = RLE_Output(pos2)
                i = i + 1
                pos2 = pos2 + 1
            Loop
            
            pos2 = pos2 + 1
            For j = i To 63
                tmp(j).length = 0
                tmp(j).size = 0
            Next j
            
            ' Binary Shift Coding
            If i > 0 Then
                Call Binary_Shift_Coding(tmp)
            End If
            i = 0
            
            ' ********** V ********
            Put #100, , Diff_Result(pos1)
            pos1 = pos1 + 1
            
            Do While (RLE_Output(pos2).length <> -1)
                tmp(i) = RLE_Output(pos2)
                i = i + 1
                pos2 = pos2 + 1
            Loop
            
            pos2 = pos2 + 1
            For j = i To 63
                tmp(j).length = 0
                tmp(j).size = 0
            Next j
            
            ' Binary Shift Coding
            If i > 0 Then
                Call Binary_Shift_Coding(tmp)
            End If
            i = 0
        
        Next xpos
    Next ypos
    
    MsgBox "Binary Shift Coding Done!", vbOKOnly, "Binary Shift Coding"
    Close #100
    
End Sub
Private Sub exit_Click()
    Dim response As Integer
    response = MsgBox("Are you want to exit?", vbYesNo, "Exit")
    If response = vbYes Then
        End
    End If
End Sub

