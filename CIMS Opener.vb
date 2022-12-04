Option Explicit

Public Const SignOffPage As String = "https://cims.ventia.com.au/app/Signoff_Page.aspx?PreviousPage=9&CurrentPage=32&EditStatus=0&TitleText=Sign%20Off&CallingPage=ProjectDetails&ProjectId="
Public Const DetailsPage As String = "https://cims.ventia.com.au/app/ProjectMainPage_New.aspx?PreviousPage=32&CurrentPage=0&EditStatus=0&TitleText=Details&CallingPage=ProjectDetails&ProjectId="
Public Const MaterialPage As String = "https://cims.ventia.com.au/app/BOQMat_matnew.aspx?PreviousPage=32&CurrentPage=8&EditStatus=0&TitleText=Material&CallingPage=ProjectDetails&ProjectId="
Public Const BOQPage As String = "https://cims.ventia.com.au/app/BOQInstall.aspx?PreviousPage=8&CurrentPage=9&EditStatus=0&TitleText=BOQ%20Install&CallingPage=ProjectDetails&ProjectId="
Public Const MISCPage As String = "https://cims.ventia.com.au/app/Npsa.aspx?PreviousPage=8&CurrentPage=10&EditStatus=0&TitleText=Miscellaneous&CallingPage=ProjectDetails&ProjectId="
Public Const ScopePage As String = "https://cims.ventia.com.au/app/Scope.aspx?PreviousPage=8&CurrentPage=1&EditStatus=0&TitleText=Scope&CallingPage=ProjectDetails&ProjectId="

Private Declare PtrSafe Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hWnd As Long, _
  ByVal Operation As String, _
  ByVal Filename As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long

Private Sub OpenUrl(Url As String)

    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", Url)

End Sub

Public Sub OpenCIMSProjects()

Dim GeneratedURL As String
Dim y As Integer

'Go through each row, and determine what pages need to be opened
For y = 2 To (ActiveSheet.UsedRange.Rows.Count)
    'Details
    If Not IsEmpty(ActiveSheet.Cells(y, 2).Value) Then
        OpenUrl (DetailsPage & ActiveSheet.Cells(y, 1).Value)
    End If
    
    'Sign Off
    If Not IsEmpty(ActiveSheet.Cells(y, 3).Value) Then
        OpenUrl (SignOffPage & ActiveSheet.Cells(y, 1).Value)
    End If
        
    'Materials
    If Not IsEmpty(ActiveSheet.Cells(y, 4).Value) Then
        OpenUrl (MaterialPage & ActiveSheet.Cells(y, 1).Value)
    End If
    
    'BOQ
    If Not IsEmpty(ActiveSheet.Cells(y, 5).Value) Then
        OpenUrl (BOQPage & ActiveSheet.Cells(y, 1).Value)
    End If
    
    'MISC
    If Not IsEmpty(ActiveSheet.Cells(y, 6).Value) Then
        OpenUrl (MISCPage & ActiveSheet.Cells(y, 1).Value)
    End If
    
    'Scope
    If Not IsEmpty(ActiveSheet.Cells(y, 7).Value) Then
        OpenUrl (ScopePage & ActiveSheet.Cells(y, 1).Value)
    End If
    
Next



End Sub
