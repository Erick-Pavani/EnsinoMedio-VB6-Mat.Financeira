VERSION 5.00
Begin VB.Form frmCalcular 
   BackColor       =   &H00FF0000&
   Caption         =   "Calcular Juros Compostos"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   12210
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbData 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      ItemData        =   "frmCalcular.frx":0000
      Left            =   8910
      List            =   "frmCalcular.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   4920
      Width           =   2505
   End
   Begin VB.ComboBox cmbTaxa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      ItemData        =   "frmCalcular.frx":006C
      Left            =   8910
      List            =   "frmCalcular.frx":0082
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   4170
      Width           =   2505
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   6210
      TabIndex        =   14
      Top             =   6450
      Width           =   2685
   End
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "Calcular"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3180
      TabIndex        =   13
      Top             =   6480
      Width           =   2565
   End
   Begin VB.TextBox txtN 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5790
      TabIndex        =   8
      Top             =   4860
      Width           =   2145
   End
   Begin VB.TextBox txtI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5790
      TabIndex        =   6
      Top             =   4110
      Width           =   2115
   End
   Begin VB.TextBox txtFv 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5790
      TabIndex        =   4
      Top             =   3330
      Width           =   2085
   End
   Begin VB.TextBox txtPv 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5790
      TabIndex        =   1
      Top             =   2490
      Width           =   2085
   End
   Begin VB.Label lbl5 
      BackColor       =   &H00FF0000&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8100
      TabIndex        =   17
      Top             =   4140
      Width           =   525
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF0000&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5070
      TabIndex        =   12
      Top             =   3300
      Width           =   465
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5070
      TabIndex        =   11
      Top             =   4080
      Width           =   465
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5070
      TabIndex        =   10
      Top             =   4830
      Width           =   465
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00FF0000&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5070
      TabIndex        =   9
      Top             =   2460
      Width           =   465
   End
   Begin VB.Label lblN 
      BackColor       =   &H00FF0000&
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4170
      TabIndex        =   7
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label lblI 
      BackColor       =   &H00FF0000&
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4290
      TabIndex        =   5
      Top             =   4020
      Width           =   285
   End
   Begin VB.Label lblFv 
      BackColor       =   &H00FF0000&
      Caption         =   "FV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4050
      TabIndex        =   3
      Top             =   3270
      Width           =   795
   End
   Begin VB.Label lblPv 
      BackColor       =   &H00FF0000&
      Caption         =   "PV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4050
      TabIndex        =   2
      Top             =   2460
      Width           =   795
   End
   Begin VB.Label lblJuros 
      BackColor       =   &H00FF0000&
      Caption         =   "Juros Compostos"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4410
      TabIndex        =   0
      Top             =   390
      Width           =   3765
   End
End
Attribute VB_Name = "frmCalcular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalcular_Click()
Vazio = 0
If txtPv.Text = "" Then
    Vazio = Vazio + 1
End If
If txtFv.Text = "" Then
    Vazio = Vazio + 1
End If
If txtI.Text = "" Then
    Vazio = Vazio + 1
End If
If txtN.Text = "" Then
    Vazio = Vazio + 1
End If
If Vazio <> 1 Or (txtI.Text <> "" And txtN.Text <> "" And txtFv.Text <> "" And txtPv.Text <> "") Then
    Call MsgBox("Deixe um campo apenas vazio e digite apenas números!")
Else
    If cmbTaxa.Text <> cmbData.Text Then
        If cmbTaxa.Text = "Dia - a.d" And cmbData.Text = "Mês - a.m" Then
            txtN.Text = txtN.Text * 30
        ElseIf cmbTaxa = "Dia - a.d" And cmbData.Text = "Ano - a.a" Then
            txtN.Text = txtN.Text * 360
        ElseIf cmbTaxa.Text = "Dia - a.d" And cmbData.Text = "Bimestre - a.b" Then
            txtN.Text = txtN.Text * 60
        ElseIf cmbTaxa.Text = "Dia - a.d" And cmbData.Text = "Trimestre - a.t" Then
            txtN.Text = txtN.Text * 90
        ElseIf cmbTaxa.Text = "Dia - a.d" And cmbData.Text = "Semestre - a.s" Then
            txtN.Text = txtN.Text * 180
        ElseIf cmbTaxa.Text = "Mês - a.m" And cmbData.Text = "Dia - a.d" Then
            txtN.Text = txtN.Text / 30
        ElseIf cmbTaxa.Text = "Mês - a.m" And cmbData.Text = "Ano - a.a" Then
            txtN.Text = txtN.Text * 12
        ElseIf cmbTaxa.Text = "Mês - a.m" And cmbData.Text = "Bimestre - a.b" Then
            txtN.Text = txtN.Text * 2
        ElseIf cmbTaxa.Text = "Mês - a.m" And cmbData.Text = "Trimestre - a.t" Then
            txtN.Text = txtN.Text * 3
        ElseIf cmbTaxa.Text = "Mês - a.m" And cmbData.Text = "Semestre - a.s" Then
            txtN.Text = txtN.Text * 6
        ElseIf cmbTaxa.Text = "Ano - a.a" And cmbData.Text = "Dia - a.d" Then
            txtN.Text = txtN.Text / 360
        ElseIf cmbTaxa.Text = "Ano - a.a" And cmbData.Text = "Mês - a.m" Then
            txtN.Text = txtN.Text / 12
        End If
    End If
    If txtPv.Text = "" Then
        If Not IsNumeric(txtFv.Text) Or Not IsNumeric(txtI.Text) Or Not IsNumeric(txtN.Text) Then
            Call MsgBox(" Digite apenas números!")
            Cancel = True
            Unload Me
            frmCalcular.Show
        Else
            If txtI.Text > 100 Then
                Call MsgBox("A taxa (I) deve ser menor do que 100!")
                Cancel = True
                Unload Me
                frmCalcular.Show
            Else
                txtPv.Text = (txtFv.Text) / ((1 + (txtI.Text / 100)) ^ txtN.Text)
                txtPv.Text = Round(txtPv.Text, 1)
            End If
        End If
    ElseIf txtFv.Text = "" Then
        If Not IsNumeric(txtPv.Text) Or Not IsNumeric(txtI.Text) Or Not IsNumeric(txtN.Text) Then
            Call MsgBox(" Digite apenas números!")
            Cancel = True
            Unload Me
            frmCalcular.Show
        Else
            If txtI.Text > 100 Then
                Call MsgBox("A taxa (I) deve ser menor do que 100!")
                Cancel = True
                Unload Me
                frmCalcular.Show
            Else
                txtFv.Text = (txtPv.Text) * ((1 + (txtI.Text / 100)) ^ txtN.Text)
                txtFv.Text = Round(txtFv.Text, 1)
            End If
        End If
    ElseIf txtI.Text = "" Then
        If Not IsNumeric(txtFv.Text) Or Not IsNumeric(txtPv.Text) Or Not IsNumeric(txtN.Text) Then
            Call MsgBox(" Digite apenas números!")
            Cancel = True
            Unload Me
            frmCalcular.Show
        Else
            A = (txtFv.Text / txtPv.Text)
            B = A ^ (1 / txtN.Text)
            C = (B - 1)
            txtI.Text = C * 100
            txtI.Text = Round(txtI.Text, 1)
        End If
    ElseIf txtN.Text = "" Then
        If Not IsNumeric(txtFv.Text) Or Not IsNumeric(txtPv.Text) Or Not IsNumeric(txtI.Text) Then
            Call MsgBox(" Digite apenas números!")
            Cancel = True
            Unload Me
            frmCalcular.Show
        Else
            If txtI.Text > 100 Then
                Call MsgBox("A taxa (I) deve ser menor do que 100!")
                Cancel = True
                Unload Me
                frmCalcular.Show
            Else
                txtN.Text = (Log(txtFv.Text) - Log(txtPv.Text)) / Log(1 + (txtI.Text / 100))
                txtN.Text = Round(txtN.Text, 1)
            End If
        End If
    End If
End If
End Sub
Private Sub cmdLimpar_Click()
Unload Me
frmCalcular.Show
End Sub
