Attribute VB_Name = "Module1"
Public db As Database
Public cmp As Recordset
Public brh As Recordset
Public dpt As Recordset
Public dsg As Recordset
Public grd As Recordset
Public emp As Recordset
Public dept_dsg As Recordset
Public dsg_grd As Recordset
Public levdetails As Recordset
Public levavailed As Recordset
Public leave_query As Recordset
Public ln As Recordset
Public lm As Recordset
Public loan_query As Recordset
Public sal_mast As Recordset
Public sal_det As Recordset
Public empsal As Recordset                  'to get salary details from sal_mast
Public info As Recordset                    'to get emp details from emp_master for salary details
Public sm As Recordset
Public loan_sal As Recordset                'get employee loan details while issuing salary. salary_details table.
Public loan_rate As Recordset   'obtain rate of interest for loan selected - loan_details
Public emp_rpt As Recordset 'to obtain employee code and name for report generation
Public leav_rpt As Recordset
Public monleave As Boolean  'to check whether monthly-leave command - recordset open or not
Public emprpt As Boolean    'to check whether employee command - recordset open or not
Public sumleave As Boolean  'to check whether leave command - recordset open or not
Public salmonth As Boolean  'to check whether salary  command - recordset open or not
Public RepeatTimes As Long
Public TotalFrames As Long
Public cnn As ADODB.Connection
'Determine the total numbre of records in the table passed
'return the total record number to calling form
Public Function chk_rscount(rs_name As String) As Integer
Dim rs As Recordset
Set rs = db.OpenRecordset(rs_name, dbOpenDynaset)
chk_rscount = rs.RecordCount
rs.Close
End Function
'--------if master table contains no records then restrict-------
'--------data entry into the details table ----------
Public Sub disableall(frmname As Form)
Dim txt As Control, combo As Control, cmd As Control
For Each txt In frmname.Controls
    If TypeOf txt Is TextBox Then
        txt.Enabled = False
    End If
Next
For Each cmd In frmname.Controls
    If TypeOf cmd Is CommandButton Then
        cmd.Enabled = False
    End If
Next
For Each combo In frmname.Controls
    If TypeOf combo Is ComboBox Then
        combo.Enabled = False
    End If
Next
End Sub
'disable / Enable menu bar : Activate on form load / unload
Public Sub menu_disable()
MDIForm1.mnComp.Enabled = Not MDIForm1.mnComp.Enabled
MDIForm1.mnEmployee.Enabled = Not MDIForm1.mnEmployee.Enabled
MDIForm1.mnLeave.Enabled = Not MDIForm1.mnLeave.Enabled
MDIForm1.mnLoan.Enabled = Not MDIForm1.mnLoan.Enabled
MDIForm1.mnSalary.Enabled = Not MDIForm1.mnSalary.Enabled
MDIForm1.mn_reports.Enabled = Not MDIForm1.mn_reports.Enabled
MDIForm1.mnHelp.Enabled = Not MDIForm1.mnHelp.Enabled
MDIForm1.mn_exit.Enabled = Not MDIForm1.mn_exit.Enabled
End Sub
'--------To clear the text / combo control fields -------
'--------used in delete / add in most of the forms ----------
Public Sub clear_all(frmname As Form)
Dim txt As Control, combo As Control
For Each txt In frmname.Controls
    If TypeOf txt Is TextBox Then
        txt.Text = ""
    End If
Next
For Each combo In frmname.Controls
    If TypeOf combo Is ComboBox Then
        combo.Text = ""
    End If
Next
End Sub
'to enable / disable text box / combo box controls depending upon action
'performed in the form
Public Sub txtcmb_disable(frmname As Form)
Dim txt As Control, combo As Control, cmd As Control
For Each txt In frmname.Controls
    If TypeOf txt Is TextBox Then
        txt.Enabled = Not txt.Enabled
    End If
Next
For Each combo In frmname.Controls
    If TypeOf combo Is ComboBox Then
        combo.Enabled = Not combo.Enabled
    End If
Next
End Sub

Public Sub chk_nullvalue(ctrltype As Control)
If IsNull(ctrltype) = False Then
    MsgBox "Field " & ctrltype & " cannot contain a zero value", vbCritical, "Payroll : Data entry error"
    ctrltype.SetFocus
End If
End Sub

Public Sub txt_disable(frmname As Form)
Dim txt As Control
For Each txt In frmname.Controls
    If TypeOf txt Is TextBox Then
        txt.Enabled = Not txt.Enabled
    End If
Next
End Sub

Public Function LoadGif(sFile As String, aImg As Variant) As Boolean
    LoadGif = False
    If Dir$(sFile) = "" Or sFile = "" Then
       MsgBox "File " & sFile & " not found", vbCritical
       Exit Function
    End If
    On Error GoTo ErrHandler
    Dim fNum As Integer
    Dim imgHeader As String, fileHeader As String
    Dim buf$, picbuf$
    Dim imgCount As Integer
    Dim i&, j&, xOff&, yOff&, TimeWait&
    Dim GifEnd As String
    GifEnd = Chr(0) & Chr(33) & Chr(249)
    For i = 1 To aImg.Count - 1
        Unload aImg(i)
    Next i
    fNum = FreeFile
    Open sFile For Binary Access Read As fNum
        buf = String(LOF(fNum), Chr(0))
        Get #fNum, , buf 'Get GIF File into buffer
    Close fNum
    
    i = 1
    imgCount = 0
    j = InStr(1, buf, GifEnd) + 1
    fileHeader = Left(buf, j)
    If Left$(fileHeader, 3) <> "GIF" Then
       MsgBox "This file is not a *.gif file", vbCritical
       Exit Function
    End If
    LoadGif = True
    i = j + 2
    If Len(fileHeader) >= 127 Then
        RepeatTimes& = Asc(Mid(fileHeader, 126, 1)) + (Asc(Mid(fileHeader, 127, 1)) * 256&)
    Else
        RepeatTimes = 0
    End If

    Do ' Split GIF Files at separate pictures
       ' and load them into Image Array
        imgCount = imgCount + 1
        j = InStr(i, buf, GifEnd) + 3
        If j > Len(GifEnd) Then
            fNum = FreeFile
            Open "temp.gif" For Binary As fNum
                picbuf = String(Len(fileHeader) + j - i, Chr(0))
                picbuf = fileHeader & Mid(buf, i - 1, j - i)
                Put #fNum, 1, picbuf
                imgHeader = Left(Mid(buf, i - 1, j - i), 16)
            Close fNum
            TimeWait = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * 256&)) * 10&
            If imgCount > 1 Then
                xOff = Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * 256&)
                yOff = Asc(Mid(imgHeader, 11, 1)) + (Asc(Mid(imgHeader, 12, 1)) * 256&)
                Load aImg(imgCount - 1)
                aImg(imgCount - 1).Left = aImg(0).Left + (xOff * Screen.TwipsPerPixelX)
                aImg(imgCount - 1).Top = aImg(0).Top + (yOff * Screen.TwipsPerPixelY)
            End If
            ' Use .Tag Property to save TimeWait interval for separate Image
            aImg(imgCount - 1).Tag = TimeWait
            aImg(imgCount - 1).Picture = LoadPicture("temp.gif")
            Kill ("temp.gif")
            i = j
        End If
        DoEvents
    Loop Until j = 3
' If there are one more Image - Load it
    If i < Len(buf) Then
        fNum = FreeFile
        Open "temp.gif" For Binary As fNum
            picbuf = String(Len(fileHeader) + Len(buf) - i, Chr(0))
            picbuf = fileHeader & Mid(buf, i - 1, Len(buf) - i)
            Put #fNum, 1, picbuf
            imgHeader = Left(Mid(buf, i - 1, Len(buf) - i), 16)
        Close fNum
        TimeWait = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * 256)) * 10
        If imgCount > 1 Then
            xOff = Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * 256)
            yOff = Asc(Mid(imgHeader, 11, 1)) + (Asc(Mid(imgHeader, 12, 1)) * 256)
            Load aImg(imgCount - 1)
            aImg(imgCount - 1).Left = aImg(0).Left + (xOff * Screen.TwipsPerPixelX)
            aImg(imgCount - 1).Top = aImg(0).Top + (yOff * Screen.TwipsPerPixelY)
        End If
        aImg(imgCount - 1).Tag = TimeWait
        aImg(imgCount - 1).Picture = LoadPicture("temp.gif")
        Kill ("temp.gif")
    End If
    TotalFrames = aImg.Count - 1
    Exit Function
ErrHandler:
    MsgBox "Error No. " & Err.Number & " when reading file", vbCritical
    LoadGif = False
    On Error GoTo 0
End Function

Public Sub getconnected()
Set cnn = New ADODB.Connection
cnn.CursorLocation = adUseClient
cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & ".\Database1.mdb" & "; Persist Security Info=False;"
cnn.Open
End Sub

