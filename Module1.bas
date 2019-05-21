Attribute VB_Name = "Module1"
Option Explicit

Global Mydb As New ADODB.ConnecTion
Global Repopath As String

Global ModNo As Double

Global MySql        As String
Global MUser        As String
Global MPass        As String
Global MGrpNo       As String
Global MUserName    As String

Global MUGrpName    As String

Global Mlocation    As String     'Location as a global parameter
Global conString    As String      'Connection string a global parameter

Global MyReccount   As Double
Global MyrecMode    As Boolean
Global MyRSMode     As Boolean
Global MyRSCOUNT    As Double
Global MyRecordset  As New ADODB.Recordset
Global MyRS         As New ADODB.Recordset
Global Para         As New ADODB.Recordset
Global Ref          As New ADODB.Recordset
Global Mdate        As Date
Global myinvno      As String
Global MyName       As String
Global NumFormat    As String
Global MyFile       As String
Global MyCode       As String

Public gstrsql      As String
Public gdblRecCount As Double
Public gstrRecMode      As String
Public gstrPrintFile    As String
Public gstrPrintFile1   As String  ' The file to be printed
Public gblnFoundCde     As Boolean
Public gstrSelectedCode As String
Public gstrSelectedFirst    As String
Public gstrSelectedSecond   As String
Public gstrSelectedThird    As String
Public gstrSelectedForth    As String
Public gstrSelectedFifth    As String
Public strCompText          As String
Public MTrnId               As Integer
Global GDomainUser          As String

Public gstrCodeHeading      As String
Public gstrNameHeading      As String
Public gstrNameHeading2     As String
Public gstrNameHeading3     As String
Public gstrNameHeading4     As String

Public gstrSearchCode       As String
Public gstrEnteredCode      As String
Public MRepoPath            As String
Public i, k, N              As Integer

Public Strpassword          As String
Public Strnewpassword       As String
Public StrnewDpassword      As String
Public MusrNameEnv          As String

Public MCompID              As String
Public gstrFormName         As String
Public Muserno              As Integer
Public MuGRPno              As Integer

Public DecodedPassword      As String
Public EncodedPassword      As String

Public grsRecordSet         As New ADODB.Recordset
Public Mpara                As New ADODB.Recordset
Public Mb4Qty, MAfterQty    As Double
Public MQtyIn, MQtyOut      As Double


Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Global PCName As String
Global DBName As String
Global DataSouce As String

Public MRptName As String

'To open a PDF file
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                    ByVal hwnd As Long, _
                    ByVal lpOperation As String, _
                    ByVal lpFile As String, _
                    ByVal lpParameters As String, _
                    ByVal lpDirectory As String, _
                    ByVal nShowCmd As Long) As Long

Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWMINIMIZED As Long = 2
' End to open a PDF file


Public Sub ConnecTion()
    
    On Error GoTo ConnErr
    If Mydb.State = 1 Then
        Set Mydb = Nothing
       ' MyDB.Close
    End If

    Dim MStrConn, MLogFile As String
    
    MLogFile = App.Path & "\Conn\" & "ConnStr.txt"

    Open MLogFile For Input As #1
    'While Not EOF(1) And FindODBC = False
        Input #1, conString
       
        Debug.Print conString
'        If Left(MLine, 6) = "[AS400]" Then
'            FindODBC = True
'        End If
    'Wend
    Close #1
    'DB Name
'    Dim N1 As Integer
'    N = InStr(1, conString, "Catalog=")
'    N1 = InStr(N, conString, ";")
    'DBName = Mid(conString, (N + 8), (N1 - N - 8))
    
    'DBName = "TEXPROSQL2008"
'***********************************
    'Data Source
'    N = InStr(1, conString, "Source=")
'    N1 = Len(conString)
'    DataSouce = Right(conString, (N1 - (N + 6)))
   ' DataSouce = "TEXPROSERVER\TEXPROSQLEXPRESS"
'**********************
    
    If Mydb.State = 1 Then
        Mydb.Close
    End If
   'conString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TEXPRO_COPM2;Data Source=DC1\MYSQL2008    "
   Mydb.ConnectionString = conString
   Mydb.CursorLocation = adUseClient
   Mydb.Open conString
    Dim P As Long
    P = NameOfPC(PCName)
    DBName = Mydb.DefaultDatabase

    '*******************
    'If PCName = "DC1" Then
    '   MyDB.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TEXPRO_COPM2;Data Source=DC1\MYSQL2008 "
       'MyDB.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TEXPRO_COPM;Data Source=DC1\MYSQL2008"
    'Else
   '     MyDB.Open conString
        'mydb.Open "MYERP"
    'End If
'    ElseIf PCName = "PC1" Then
'        mydb.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MYERP;Data Source=DC1\MYSQL2000"
'
'    ElseIf PCName = "PC2" Then
'        mydb.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MYERP;Data Source=DC1\MYSQL2000"
'    Else
'        mydb.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MYERP;Data Source=HPSERVER\HP_SQLSERVER"
'    End If
    
    ModNo = 1
    MCompID = 1
    If Para.State = 1 Then
        Para.Close
    End If

    Para.Open "select repopath from m_para", Mydb, adOpenKeyset, adLockOptimistic
    MRepoPath = Para!Repopath
    Para.Close
    
'    MRepoPath = "\\TEXPROSERVER\COPM\REPORTS\"
''    If PCName = "TEXPROSERVER" Then
'''        MRepoPath = "D:\COPM\Reports\"
''         MRepoPath = App.Path & "\" & "Reports\"
''    Else
''        MRepoPath = App.Path & "\" & "Reports\"
''    End If
'
'    MRepoPath = App.Path & "\" & "Reports\"
    'MRepoPath = Repopath & "\Reports"
    NumFormat = "###,###.#0"
    
    
ConnErr:
    If Err.Number <> 0 Then
        MsgBox "connection Error - try again", vbInformation, "Please check Connection string"
        End
    End If
    
    
End Sub

Public Sub InsertUpdateSQL()
    Mydb.Execute gstrsql
End Sub

Public Function SetRecordSet() As ADODB.Recordset

        Dim rstRs As New ADODB.Recordset
        On Error GoTo Abc
        Set rstRs = SetRecordSet
        If rstRs.State = 1 Then
            rstRs.Close
        End If
        If Mydb.State = 0 Then
            ConnecTion
        End If
        rstRs.Open gstrsql, Mydb, adOpenStatic
        MyRSCOUNT = rstRs.RecordCount
        MyReccount = rstRs.RecordCount
        If rstRs.EOF = False Then
            'If IsNull(rstRs.Fields(0)) = False Then
            If rstRs.RecordCount > 0 Then
                MyRSMode = True
                MyrecMode = True
            Else
                MyRSMode = False
                MyrecMode = False
            End If
        Else
            MyRSMode = False
        End If
        Set SetRecordSet = rstRs
        Exit Function
Abc:
        MyRSMode = False
        MsgBox Err.Number & "-" & Err.Description, , ""

End Function

Public Function numbersonly(ByVal Key As Integer) As Boolean
    If Key >= 47 And Key <= 58 Or Key = 8 Or Key = 46 Or Key = 13 Or Key = 45 Then
         numbersonly = True
    End If
End Function

Public Sub CallRecordSet()
        If MyRecordset.State = 1 Then
            MyRecordset.Close
        End If
        On Error GoTo Abc
        If Mydb.State = 0 Then
            ConnecTion
        End If
        MyRecordset.CursorLocation = adUseClient
        MyRecordset.Open MySql, Mydb, adOpenKeyset, adLockOptimistic
        
        If MyRecordset.EOF = False Then
'            If IsNull(MyRecordSet.Fields(0)) = False Then
            If MyRecordset.RecordCount > 0 Then
                MyrecMode = True
                MyRecordset.MoveLast
                MyReccount = MyRecordset.RecordCount
                MyRecordset.MoveLast
                MyRecordset.MoveFirst
            Else
                MyReccount = 0
                MyrecMode = False
            End If
            
        Else
        
Abc:

            MyReccount = 0
            MyrecMode = False
        End If
        

End Sub

Public Sub CallRetRecordSet()
        If grsRecordSet.State = 1 Then
            grsRecordSet.Close
        End If
        On Error GoTo Abc
        grsRecordSet.Open gstrsql, Mydb, adOpenKeyset, adLockOptimistic
        If grsRecordSet.EOF = False Then
'            If IsNull(MyRecordSet.Fields(0)) = False Then
            If grsRecordSet.RecordCount > 0 Then
                gstrRecMode = "MoreRecs"
                grsRecordSet.MoveLast
                gdblRecCount = grsRecordSet.RecordCount
                grsRecordSet.MoveLast
                grsRecordSet.MoveFirst
            Else
                gdblRecCount = 0
                gstrRecMode = "NoRecs"
            End If
            
        Else
        
Abc:

            gstrRecMode = "NoRecs"
            gdblRecCount = 0
        End If
        
End Sub

Public Sub FillComboData(ByVal cmbCombo As ComboBox)

    cmbCombo.Clear
    
    If cmbCombo.Name = "cmbUser" Then
        gstrsql = " Select * from m_users order by user_name "
    ElseIf cmbCombo.Name = "cmbGroup" Then
        gstrsql = " Select * From M_usergroups oRDER bY Ugrp_name"
        
    ElseIf cmbCombo.Name = "cmbloca" Then
        gstrsql = " Select loc_name From m_Locmf "
        
    End If
            
    Call RetrieveRecordSet
    
    If gstrRecMode = "MoreRecs" Then
    
        While grsRecordSet.EOF = False
        
            If cmbCombo.Name = "cmbUser" Then
                cmbCombo.AddItem grsRecordSet.Fields("user_name")
            ElseIf cmbCombo.Name = "cmbGroup" Then
                cmbCombo.AddItem grsRecordSet.Fields("ugrp_name")
            ElseIf cmbCombo.Name = "cmbloca" Then
                cmbCombo.AddItem grsRecordSet.Fields("Loc_name")
            End If
            grsRecordSet.MoveNext
        Wend
        
    End If
        
End Sub


Public Sub RetrieveRecordSet()
        If grsRecordSet.State = 1 Then
            grsRecordSet.Close
        End If
        On Error GoTo Abc
        If Mydb.State = 0 Then
            ConnecTion
        End If
        grsRecordSet.Open gstrsql, Mydb, adOpenKeyset, adLockPessimistic
        If grsRecordSet.EOF = False Then
            'If IsNull(grsRecordSet.Fields(0)) = False Then
            If grsRecordSet.RecordCount > 0 Then
                gstrRecMode = "MoreRecs"
                grsRecordSet.MoveLast
                gdblRecCount = grsRecordSet.RecordCount
                grsRecordSet.MoveLast
                grsRecordSet.MoveFirst
            Else
                gdblRecCount = 0
                gstrRecMode = "NoRecs"
            End If
        Else
Abc:
            gdblRecCount = 0
            gstrRecMode = "NoRecs"
        End If
End Sub

Public Function EncodePassword(ByVal Strpassword As String) As String

        Dim intLength As Integer
        Dim i As Integer
        Dim x As Integer
        
        Strnewpassword = ""
        intLength = Len(Strpassword)
        
        For i = 1 To intLength
                                
            x = Asc(Mid(UCase(Strpassword), i, 1)) + 10
            
            Strnewpassword = Strnewpassword & x
        
        Next
        
        EncodedPassword = Strnewpassword
        
End Function

Public Function DecodePassword(ByVal Strnewpassword As String) As String

        Dim intLength As Integer
        Dim i As Integer
        StrnewDpassword = ""
        'Dim X As String
        Dim X1 As String
        Dim x, k As Integer
        Dim Mchkno As String
        Mchkno = ""
        intLength = Len(Strnewpassword)
        k = 1
        For i = 1 To intLength      '       Step 3
                
                Mchkno = Val(Mid(Strnewpassword, k, 2))
                    
                If Mchkno > 0 Then
                                        
                    If Mchkno >= 58 Or Mchkno <= 67 Then
                    
                        x = Mid(Strnewpassword, k, 2) - 10
                        k = k + 2
                    Else
                                        
                        x = Mid(Strnewpassword, k, 3) - 10
                        k = k + 2
                    End If
                    
                
                
                 X1 = Chr(x)
             
                StrnewDpassword = StrnewDpassword & X1
                
            End If
            
        Next
        
        DecodedPassword = StrnewDpassword
        
End Function

'
'Public Sub ChkRights()
'
'    'If MuGRPno = "5" Then
'    '    MDI_IMDb.mnusyssecu.Enabled = True
'
'    '    MDI_IMDb.mnusyssecu.Visible = True
'    'End If
'
'
'    MDI_IM.mnuprodmf.Enabled = False
'        MDI_IM.mnuprodcat.Enabled = False
'        MDI_IM.mnulocationmf.Enabled = False
'        MDI_IM.mnuprodgrp.Enabled = False
'        MDI_IM.mnusuppmf.Enabled = False
'        MDI_IM.mnucustmf.Enabled = False
'        MDI_IM.mnusalesrep.Enabled = False
'        MDI_IM.MnuPriceBK.Enabled = False
'
'
'
'        MDI_IM.mnugrnadd.Enabled = False
'        MDI_IM.mnutfr.Enabled = False
'        MDI_IM.mnuinvoice.Enabled = False
'        MDI_IM.mnuopbentry.Enabled = False
'        MDI_IM.Mnusupprtn.Enabled = False
'        MDI_IM.mnucustrtns.Enabled = False
'        MDI_IM.mnuReceipts.Enabled = False
'        MDI_IM.mnuexshEntry.Enabled = False
'
'        MDI_IM.mnugrnlisting.Enabled = False
'        MDI_IM.mnusalesrpt.Enabled = False
'        MDI_IM.mnusalrprtdt.Enabled = False
'        MDI_IM.mnucostsaleRpt.Enabled = False
'        MDI_IM.mnustocksheet.Enabled = False
'        MDI_IM.mnubincard.Enabled = False
'        MDI_IM.MnuLoading.Enabled = False
'        MDI_IM.MnuIssuesRpt.Enabled = False
'        MDI_IM.mnuprintInv.Enabled = False
'        MDI_IM.mnutclist.Enabled = False
'        MDI_IM.MnuCollection.Enabled = False
'
'        MDI_IM.Mnuutility.Enabled = False
'        MDI_IM.Mnuinquiry.Enabled = False
'        MDI_IM.mnusecurity.Enabled = False
'
'
'        MDI_IM.Toolbar1.Buttons(1).Enabled = False
'        MDI_IM.Toolbar1.Buttons(2).Enabled = False
'        MDI_IM.Toolbar1.Buttons(3).Enabled = False
'        MDI_IM.Toolbar1.Buttons(4).Enabled = False
'        MDI_IM.Toolbar1.Buttons(5).Enabled = False
'        MDI_IM.Toolbar1.Buttons(6).Enabled = False
'        MDI_IM.Toolbar1.Buttons(7).Enabled = False
'        MDI_IM.Toolbar1.Buttons(8).Enabled = False
'
'    Dim MyAuthoFile  As New ADODB.Recordset
'    gstrsql = "select * from user_authority where mod_no = '" & ModNo & "' and user_id = '" & MUser & "'"
'    Set MyAuthoFile = SetRecordSet
'
'    If MyAuthoFile.RecordCount > 0 Then
'        i = 2
'        Do While Not MyAuthoFile.EOF
'
'            If Trim(MyAuthoFile("mnucode")) = "mnumaster" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnumaster.Enabled = True
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnulocationmf" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnulocationmf.Enabled = True
'                    MDI_IM.Toolbar1.Buttons(1).Enabled = True
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnuprodmf" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnuprodmf.Enabled = True
'                    MDI_IM.Toolbar1.Buttons(2).Enabled = True
'
'            ElseIf Trim(Trim(MyAuthoFile("mnucode"))) = "mnuprodcat" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnuprodcat.Enabled = True
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnuprodgrp" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnuprodgrp.Enabled = True
'            ElseIf Trim(MyAuthoFile("mnucode")) = "MnuPriceBK" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.MnuPriceBK.Enabled = True
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnusuppmf" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnusuppmf.Enabled = True
'                    MDI_IM.Toolbar1.Buttons(4).Enabled = True
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnusalesrep" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnusalesrep.Enabled = True
'                    'MDI_IM.Toolbar1.Buttons(4).Enabled = False
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnucustmf" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnucustmf.Enabled = True
'                    MDI_IM.Toolbar1.Buttons(3).Enabled = True
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnucustmf" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnucustmf.Enabled = True
'                    MDI_IM.Toolbar1.Buttons(3).Enabled = False
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnudata" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnudata.Enabled = True
''            ElseIf Trim(MyAuthoFile("mnucode")) = "mnudata" And MyAuthoFile("autho_status") <> "Y" Then
''                    MDI_IM.mnudata.Enabled = False
'
''            ElseIf Trim(MyAuthoFile("mnucode")) = "mnugrnadd" And MyAuthoFile("autho_status") = "Y" Then
''                    MDI_IM.mnugrnadd.Enabled = True
''
''
''                   ' If MuGRPno = 1 And MDI_IM.mnugrnadd.Enabled = True Then 'H.O users
''                        MDI_IM.Toolbar1.Buttons(5).Enabled = True
''                   ' Else
''                   '     MDI_IM.Toolbar1.Buttons(5).Enabled = False
''                   '     MDI_IM.mnugrnadd.Enabled = False
''                   ' End If
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnugrnadd" And MyAuthoFile("autho_status") = "Y" Then
'
'                    MDI_IM.mnugrnadd.Enabled = True
'                    MDI_IM.Toolbar1.Buttons(5).Enabled = True
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnutfr" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnutfr.Enabled = True
'                    MDI_IM.Toolbar1.Buttons(6).Enabled = True
'
''            ElseIf Trim(MyAuthoFile("mnucode")) = "mnutfr" And MyAuthoFile("autho_status") = "Y" Then
''                    MDI_IM.mnutfr.Enabled = True
''                    MDI_IM.Toolbar1.Buttons(6).Enabled = False
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnuinvoice" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnuinvoice.Enabled = True
'                    MDI_IM.Toolbar1.Buttons(7).Enabled = True
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnuopbentry" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnuopbentry.Enabled = True
'                    MDI_IM.Toolbar1.Buttons(8).Enabled = True
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "Mnusupprtn" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.Mnusupprtn.Enabled = True
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnucustrtns" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnucustrtns.Enabled = True
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnuReceipts" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnuReceipts.Enabled = True
'                    MDI_IM.Toolbar1.Buttons(9).Enabled = True
'
'
''            ElseIf Trim(MyAuthoFile("mnucode")) = "mnuupdate" And MyAuthoFile("autho_status") <> "Y" Then
''                    MDI_IM.mnuupdate.Enabled = False
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnuexshEntry" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnuexshEntry.Enabled = True
'                    MDI_IM.Toolbar1.Buttons(10).Enabled = True
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnuCRNote" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnuCRNote.Enabled = True
'                   'MDI_IM.Toolbar1.Buttons(11).Enabled = True
'
'
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnureports" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnureports.Enabled = True
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnureports" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnureports.Enabled = True
'
'           ElseIf Trim(MyAuthoFile("mnucode")) = "mnugrnlisting" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnugrnlisting.Enabled = True
'           ElseIf Trim(MyAuthoFile("mnucode")) = "mnusalesrpt" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnusalesrpt.Enabled = True
'
'           ElseIf Trim(MyAuthoFile("mnucode")) = "mnucostsaleRpt" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnucostsaleRpt.Enabled = True
'           ElseIf Trim(MyAuthoFile("mnucode")) = "mnucostsaleRpt" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnucostsaleRpt.Enabled = True
'
'           ElseIf Trim(MyAuthoFile("mnucode")) = "mnusalrprtdt" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnusalrprtdt.Enabled = True
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnusalestrendreport" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnusalestrendreport.Enabled = True
'                    '
'           ElseIf Trim(MyAuthoFile("mnucode")) = "mnustocksheet" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnustocksheet.Enabled = True
'           ElseIf Trim(MyAuthoFile("mnucode")) = "mnubincard" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnubincard.Enabled = True
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "MnuIssuesRpt" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.MnuIssuesRpt.Enabled = True
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnuprintInv" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnuprintInv.Enabled = True
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnutclist" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnutclist.Enabled = True
'            ElseIf Trim(MyAuthoFile("mnucode")) = "MnuCollection" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.MnuCollection.Enabled = True
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "Mnuutility" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.Mnuutility.Enabled = True
'            ElseIf Trim(MyAuthoFile("mnucode")) = "MnuLoading" And MyAuthoFile("autho_status") = "Y" Then
'
'                    MDI_IM.MnuLoading.Enabled = True
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnusummaryRpt" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnusummaryRpt.Enabled = True
'                   'MDI_IM.Toolbar1.Buttons(11).Enabled = True
'
'
'
'
''
''            ElseIf Trim(MyAuthoFile("mnucode")) = "Mnuutility" And MyAuthoFile("autho_status") = "Y" Then
''                    MDI_IM.Mnuinquiry.Enabled = True
''            ElseIf Trim(MyAuthoFile("mnucode")) = "Mnuinquiry" And MyAuthoFile("autho_status") <> "Y" Then
''                    MDI_IM.Mnuinquiry.Enabled = False
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "Mnuinquiry" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.Mnuinquiry.Enabled = True
'            ElseIf Trim(MyAuthoFile("mnucode")) = "mnusecurity" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.mnusecurity.Enabled = True
'
'            ElseIf Trim(MyAuthoFile("mnucode")) = "MnuOutsInq" And MyAuthoFile("autho_status") = "Y" Then
'                    MDI_IM.MnuOutsInq.Enabled = True
'
''
'
'            End If
'        MyAuthoFile.MoveNext
'        i = i + 1
'        Loop
'
'    Else
'
'
'        MDI_IM.mnuprodmf.Enabled = False
'        MDI_IM.mnuprodcat.Enabled = False
'        MDI_IM.mnulocationmf.Enabled = False
'        MDI_IM.mnuprodgrp.Enabled = False
'        MDI_IM.mnusuppmf.Enabled = False
'        MDI_IM.mnucustmf.Enabled = False
'        MDI_IM.mnusalesrep.Enabled = False
'        MDI_IM.MnuPriceBK.Enabled = False
'
'
'
'        MDI_IM.mnugrnadd.Enabled = False
'        MDI_IM.mnutfr.Enabled = False
'        MDI_IM.mnuinvoice.Enabled = False
'        MDI_IM.mnuopbentry.Enabled = False
'        MDI_IM.Mnusupprtn.Enabled = False
'        MDI_IM.mnucustrtns.Enabled = False
'        MDI_IM.mnuReceipts.Enabled = False
'        MDI_IM.mnuexshEntry.Enabled = False
'
'        MDI_IM.mnugrnlisting.Enabled = False
'        MDI_IM.mnusalesrpt.Enabled = False
'        MDI_IM.mnusalrprtdt.Enabled = False
'        MDI_IM.mnucostsaleRpt.Enabled = False
'        MDI_IM.mnustocksheet.Enabled = False
'        MDI_IM.mnubincard.Enabled = False
'        MDI_IM.MnuLoading.Enabled = False
'        MDI_IM.MnuIssuesRpt.Enabled = False
'        MDI_IM.mnuprintInv.Enabled = False
'        MDI_IM.mnutclist.Enabled = False
'        MDI_IM.MnuCollection.Enabled = False
'
'        MDI_IM.Mnuutility.Enabled = False
'        MDI_IM.Mnuinquiry.Enabled = False
'        MDI_IM.mnusecurity.Enabled = False
'
'
'        MDI_IM.Toolbar1.Buttons(1).Enabled = False
'        MDI_IM.Toolbar1.Buttons(2).Enabled = False
'        MDI_IM.Toolbar1.Buttons(3).Enabled = False
'        MDI_IM.Toolbar1.Buttons(4).Enabled = False
'        MDI_IM.Toolbar1.Buttons(5).Enabled = False
'        MDI_IM.Toolbar1.Buttons(6).Enabled = False
'        MDI_IM.Toolbar1.Buttons(7).Enabled = False
'        MDI_IM.Toolbar1.Buttons(8).Enabled = False
'        MsgBox "Your account has not yet authorized to use this module ", vbCritical, "Please contact your systems admnistrator"
'    End If
'
'End Sub

'Public Sub PrintReport()

    'frmPrint.Show vbModal, MDI_IM
    
'End Sub

Public Function GetTrnID(MLocCode, MItemNo, MQtyIn, MTrnId)

    Dim MyRS As New ADODB.Recordset
    '
    gstrsql = "select Max(trnID) as MaxID from M_hist where loc_code = '" & Trim(MLocCode) & "' and itemcode ='" & Trim(MItemNo) & "' "
    Set MyRS = SetRecordSet
    If MyRS.EOF = False Then
        MTrnId = IIf(IsNull(MyRS!MAxid) = False, MyRS!MAxid, 1)
    Else
        MTrnId = 1
    End If
    
    gstrsql = "select LOC_CODE,itemcode,QTY  from M_ITEMBL where loc_code = '" & Trim(MLocCode) & "' and itemcode ='" & Trim(MItemNo) & "'"
    Set MyRS = SetRecordSet
    If MyRS.EOF = False Then
        
        MQtyIn = MyRS!qty
        MQtyOut = 0
    Else
        MQtyIn = 0
        MQtyOut = 0

    End If
    'Debug.Print Mb4Qty
End Function

Public Function NameOfPC(MachineName As String) As Long
    Dim NameSize As Long
    Dim x As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    x = GetComputerName(MachineName, NameSize)
    PCName = Left(MachineName, NameSize)
End Function


Public Function CheckCodeExistance(ByVal strCode As String, ByVal strTable As String, ByVal strUserCode As String, ByVal objControl As Object) As Boolean
        If strUserCode <> "" Then
            gstrsql = " Select " & strCode & " From " & strTable & " Where " & strCode & "='" & strUserCode & "'"
            Call RetrieveRecordSet
            CheckCodeExistance = True
            If gdblRecCount = 0 Then
                CheckCodeExistance = False
                MsgBox "Invalid Entry", vbOKOnly + vbExclamation, "Existance Checking"
                objControl.SetFocus
            End If
        End If
End Function

Public Function InvalidAsciiChecking(ByVal Key As Integer) As Boolean
         If Key = 39 Or Key = 124 Then
             InvalidAsciiChecking = True
        End If
End Function


Public Function CheckCodeDuplicates(ByVal strCode As String, ByVal strTable As String, ByVal strUserCode As String, ByVal objControl As Object) As Boolean
        gstrsql = " Select " & strCode & " From " & strTable & " Where " & strCode & "='" & strUserCode & "'"
        Call RetrieveRecordSet
        CheckCodeDuplicates = False
        If gdblRecCount > 0 Then
            CheckCodeDuplicates = True
            MsgBox "Duplicates Are Not Allowed", vbOKOnly + vbExclamation, "Duplicate Checking"
            objControl.SetFocus
        End If
End Function



'Public Function FillOuts(MCustCode As String, BlnFound As Boolean)
'
'    'gstrsql = "SELECT  INV_NO, INV_DATE, INV_AMT, CUST_CODE, BAL_AMT, REC_AMT FROM D_DEBT where cust_code='" & Trim(MCustCode) & "' and (bal_amt-rec_amt) >0 "
'    'gstrsql = "SELECT     D_DEBT.INV_NO, D_DEBT.INV_DATE, D_DEBT.INV_AMT, D_DEBT.CUST_CODE, D_DEBT.BAL_AMT, D_DEBT.REC_AMT, m_entity.name,  m_entity.add1 + m_entity.add2 + m_entity.add3 AS address1 " _
'                & " FROM  D_DEBT INNER JOIN m_entity ON D_DEBT.CUST_CODE = m_entity.cust_code  WHERE     (D_DEBT.CUST_CODE = '" & Trim(MCustCode) & "') AND (D_DEBT.BAL_AMT - D_DEBT.REC_AMT > 0) "
'
'    FrmOuts.MSFlexGrid1.Clear
'    FrmOuts.MSFlexGrid1.Rows = 1
'    gstrsql = "SELECT     D_DEBT.INV_NO, D_DEBT.INV_DATE, D_DEBT.INV_AMT, D_DEBT.CUST_CODE, D_DEBT.BAL_AMT, D_DEBT.REC_AMT, m_entity.name," _
'                & " m_entity.add1 + m_entity.add2 + m_entity.add3 AS address1, D_DEBT.REP_CODE, m_salesrep.rep_name FROM         D_DEBT INNER JOIN  " _
'                & " m_entity ON D_DEBT.CUST_CODE = m_entity.cust_code INNER JOIN  m_salesrep ON D_DEBT.REP_CODE = m_salesrep.rep_code " _
'                & " WHERE     (D_DEBT.CUST_CODE = '" & Trim(MCustCode) & "') AND (D_DEBT.BAL_AMT - D_DEBT.REC_AMT > 0)"
'
'    Set MyRS = SetRecordSet
'
'    If MyRS.RecordCount > 0 Then
'        gblnFoundCde = True
'        FrmOuts.MSFlexGrid1.Rows = 1
'        FrmOuts.MSFlexGrid1.RowHeight(0) = 400
'        FrmOuts.MSFlexGrid1.Cols = 8
'        FrmOuts.MSFlexGrid1.ColWidth(0) = 500
'        FrmOuts.MSFlexGrid1.ColWidth(1) = 2000
'        FrmOuts.MSFlexGrid1.ColWidth(2) = 2000
'        FrmOuts.MSFlexGrid1.ColWidth(3) = 2000
'        FrmOuts.MSFlexGrid1.ColWidth(4) = 2000
'        FrmOuts.MSFlexGrid1.ColWidth(5) = 2000
'        FrmOuts.MSFlexGrid1.ColWidth(6) = 1000
'        FrmOuts.MSFlexGrid1.ColWidth(7) = 2000
'
'        FrmOuts.txtcustcode = MCustCode
'        FrmOuts.lblcustomer = MyRS!Name & " - " & Trim(MyRS!address1)
'        FrmOuts.MSFlexGrid1.TextMatrix(0, 1) = "Invoice No."
'        FrmOuts.MSFlexGrid1.TextMatrix(0, 2) = "Invoice date"
'        FrmOuts.MSFlexGrid1.TextMatrix(0, 3) = "Invoice Amount"
'        FrmOuts.MSFlexGrid1.TextMatrix(0, 4) = "Paid Amount"
'        FrmOuts.MSFlexGrid1.TextMatrix(0, 5) = "Balance Amount"
'        FrmOuts.MSFlexGrid1.TextMatrix(0, 6) = "Rep code"
'        FrmOuts.MSFlexGrid1.TextMatrix(0, 7) = "Rep Name"
'        Dim Invcount As Integer
'        Dim InvSum   As Double
'        i = 1
'        Do While Not MyRS.EOF
'
'            FrmOuts.MSFlexGrid1.AddItem i & vbTab & MyRS!inv_no & vbTab & Format(MyRS!inv_date, "dd/mm/yyyy") & vbTab & Format(MyRS!inv_amt, "#,###0.00") & vbTab & Format(MyRS!rec_amt, "#,###0.00") & vbTab & Format((MyRS!bal_amt - MyRS!rec_amt), "#,###0.00") & vbTab & MyRS!Rep_Code & vbTab & MyRS!rep_name
'            FrmOuts.MSFlexGrid1.RowHeight(i) = 400
'            InvSum = InvSum + (MyRS!bal_amt - MyRS!rec_amt)
'            MyRS.MoveNext
'            i = i + 1
'        Loop
'        FrmOuts.lblinvcount = i - 1
'        FrmOuts.lblvalue = Format(InvSum, "#,##0.00")
'
'    Else
'
'        gblnFoundCde = False
'        gstrSelectedCode = ""
'        gstrSelectedSecond = ""
'        gstrSelectedForth = ""
'        gstrSelectedThird = ""
'    End If
'
'End Function

Function CloseForm() As Boolean
    If Mydb.State = 1 Then
        Mydb.Close
    End If
    Set Mydb = Nothing
    CloseForm = True
    
End Function



Function ValidateText(mkey) As Boolean
    
    If Chr(mkey) = "," Or Chr(mkey) = ";" Or Chr(mkey) = ":" Or Chr(mkey) = "'" Then
        ValidateText = False
    Else
        ValidateText = True
    End If
    
End Function

Public Sub CloseRs()
    If MyRS.State = 1 Then MyRS.Close
       
'    End If
End Sub
