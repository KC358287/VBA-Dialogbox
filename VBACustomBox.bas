#Always setting variable types

Option Explicit
    Public btnDlg As DialogSheet
    Public ButtonDialog As String

Sub MakeDialogbox()

    'Create object
    ButtonDialog = "CustomButtons"
    Dim oSHL As Object: Set oSHL = CreateObject("WScript.Shell")
    Application.ScreenUpdating = False

    'check If dialogsheets exist, if yes = delete, else ok( clear errors)
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.DialogSheets(ButtonDialog).Delete
    Err.Clear
    Application.DisplayAlerts = True
    
    
    'Add custom dialogbox
    Set btnDlg = ActiveWorkbook.DialogSheets.Add
    
    'Custom modification
    With btnDlg
        .Name = ButtonDialog
        .Visible = xlSheetHidden
 
        With .DialogFrame
            .Height = 70
            .Width = 300
            .Caption = "Title"
        End With
 
        .Buttons("Button 2").Visible = False    'creating button1
        .Buttons("Button 3").Visible = False    'creating button 2
        .Labels.Add 100, 50, 100, 100
        .Labels(1).Caption = "Any questions?"
        
        'Custom 1 button, you can add more or delete all
        .Buttons.Add 220, 44, 130, 18 'Custom Button #1,index #3
        With .Buttons(3)
            .Caption = "Name1"
.OnAction = "'Sub1 """ & ButtonDialog & """'" 
        End With
        
        'Custom 2 button
        .Buttons.Add 220, 64, 130, 18 'Custom Button #2,index #4
        With .Buttons(4)
            .Caption = "Name2"
            .OnAction = "'Sub2 """ & ButtonDialog & """'"
       End With

       'If user will press X
       If .Show = False Then
            oSHL.PopUP "Canceled", 1, "Title", vbInformation
       End If

    End With
    

        'Ending with dialogsheets
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        DialogSheets("CustomButtons").Delete
        Err.Clear
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        
        btnDlg.Visible = xlSheetVeryHidden
        
       
        
End Sub

Sub Sub1(Optional Name As String = "") 'Optionnal Name required if you are doing reference to external sub 

    'Delete dialogbox
    If Len(Name) > 0 Then ActiveWorkbook.DialogSheets(Name).Hide

End Sub
