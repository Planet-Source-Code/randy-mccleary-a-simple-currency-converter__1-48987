VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Currency Converter:"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2520
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "C&lear"
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cboCurrency 
      Height          =   315
      Left            =   2400
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "&Convert"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtAmount 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Currency Rates provided by: OANDA.com.                Date Checked: 10/4/03"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   6735
   End
   Begin VB.Label Label4 
      Caption         =   "US Dollar Equivalent"
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Currency Type"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Amount to convert:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblConverted 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*********************************************************************
' Currency rates were collect from the FXConverter™
' Link: http://www.oanda.com/convert/classic
' Date Checked: 10/4/2003
'*********************************************************************

Dim mdblRate(10) As Double
Dim mstrName(10) As String

Private Sub cmdConvert_Click()
   Dim dblAmount As Double
   Dim dblConverted As Double
   
   '## Get the numeric value from text box and convert to double
   dblAmount = CDbl(Val(txtAmount.Text))
   
   '## Calculate the dollar amount for the selected currency
   dblConverted = dblAmount * mdblRate(cboCurrency.ListIndex)
   
   '## Display the converted amount to 6 decimal places
   lblConverted.Caption = FormatNumber(dblConverted, 6)
End Sub

Private Sub cmdExit_Click()
   End
End Sub

Private Sub Command1_Click()
   txtAmount.Text = ""
   lblConverted.Caption = ""
   cboCurrency.ListIndex = 0
End Sub

Private Sub Form_Load()
   Dim i As Integer
   
   '## Load the Arrays to store the Rate and Currency Name
   mdblRate(0) = 0.02337
   mstrName(0) = "Afghanistan Afghani"
   mdblRate(1) = 0.008263
   mstrName(1) = "Albanian Lek"
   mdblRate(2) = 1.6612
   mstrName(2) = "British Pound"
   mdblRate(3) = 1.1569
   mstrName(3) = "Euro"
   mdblRate(4) = 0.1297
   mstrName(4) = "Hong Kong Dollar"
   mdblRate(5) = 0.009016
   mstrName(5) = "Japanse Yen"
   mdblRate(6) = 0.08879
   mstrName(6) = "Mexican Peso"
   mdblRate(7) = 0.45455
   mstrName(7) = "North Korean Won"
   mdblRate(8) = 0.03283
   mstrName(8) = "Russian Rouble"
   mdblRate(9) = 0.74822
   mstrName(9) = "Swiss Franc"
      
   '## Load the DropDown List with Name
   With cboCurrency
      For i = 0 To UBound(mstrName) - 1
         .AddItem mstrName(i)
      Next
      .ListIndex = 0
   End With
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
   '## Only allow Numeric values and the decimal point
   If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46 Then
      KeyAscii = KeyAscii
   Else
      'Return empty value
      KeyAscii = 0
   End If
End Sub
