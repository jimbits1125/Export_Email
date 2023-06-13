VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Date_Prompt 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "Date_Prompt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Date_Prompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    ' Validate the start date
    If Not IsDate(TextBox1.Value) Then
        MsgBox "Please enter a valid start date.", vbExclamation
        Exit Sub
    End If
    
    ' Validate the end date (if provided)
    If TextBox2.Value <> "" Then
        If Not IsDate(TextBox2.Value) Then
            MsgBox "Please enter a valid end date.", vbExclamation
            Exit Sub
        End If
        
        ' Compare start and end dates
        If CDate(TextBox1.Value) > CDate(TextBox2.Value) Then
            MsgBox "The end date cannot be earlier than the start date.", vbExclamation
            Exit Sub
        End If
    End If
    
    ' Process the dates
    Dim startDate As Date
    Dim endDate As Date
    
    startDate = CDate(TextBox1.Value)
    
    If TextBox2.Value = "" Then
        endDate = Date
    Else
        endDate = CDate(TextBox2.Value)
    End If
    
    ' Perform further processing with the validated dates
    ' ...
End Sub

