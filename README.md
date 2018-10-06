# VBA-Dialogbox
**Creating custom dialogbox without using UserForms**

If you are using macros and do not want to implement userforms in every single pc (in your company for example), you can use Dialogsheets!

## How it works?
- Create custom dialogbox by using dialogsheets.
- You can modify yours buttons and add references to external sub
  - Remember to add *Optional Name As String = ""* in external sub
  - At the end of external sub, you have to add *If Len(Name) > 0 Then ActiveWorkbook.DialogSheets(Name).Hide* to delete dialogsheet

## Custom dialogbox:
![przechwytywanie](https://user-images.githubusercontent.com/43881785/46539977-56c3a480-c8b8-11e8-9cff-51895b4e6558.PNG)
