Sub Main()
    
    'CALL K4C445_@ which obtains output from WIP_TRX
    'v2.5.1 created on 8/5/2022: with additional retrieval of Upward Dipping Color (in VUM 861) for 842 process.
    
    Dim DSC As New DataSourceConnection
    Dim Error_Flag As Boolean
    Dim In_Use_By As String
    Dim PARM_Company As String
    Dim PARM_Div_1000 As String
    Dim PARM_No_Of_Weeks As Integer
    Dim PARM_Offset As String
    Dim PARM_Weekend As String
    Dim PARM_WIP As String
    Dim Work_Date As Date
    
    PARM_Company = Worksheets("Sheet1").Cells(1, 2)
    PARM_WIP = Worksheets("Sheet1").Cells(2, 2)
    PARM_Weekend = Worksheets("Sheet1").Cells(3, 2)
    PARM_Offset = Worksheets("Sheet1").Cells(4, 2)
    PARM_Div_1000 = Worksheets("Sheet1").Cells(5, 2)
    PARM_No_Of_Weeks = Worksheets("Sheet1").Cells(6, 2)
    
    If Trim(PARM_Company) = "" Or Trim(PARM_WIP) = "" Or Trim(PARM_Weekend) = "" Or PARM_No_Of_Weeks = 0 Then
        MsgBox "Company, WIP, Weekend and Number Of Weeks are mandatory."
        GoTo End_Program
    End If
    
    Work_Date = CDate(Mid(PARM_Weekend, 5, 2) & "/" & Mid(PARM_Weekend, 7, 2) & "/" & Mid(PARM_Weekend, 1, 4))
    If Weekday(Work_Date) <> 7 Then
        MsgBox "Weekend must be a Saturday date."
        GoTo End_Program
    End If
    
    DSC.EstCon_ASIAMFG
    
    In_Use_By = ""
    Error_Flag = False
    In_Use PARM_Company, PARM_Weekend, In_Use_By, Error_Flag, DSC
    
    If Error_Flag = True Then
        MsgBox "Plan currently being updated by " & In_Use_By, vbExclamation
        GoTo End_Program
    End If
    
    Application.ScreenUpdating = False
    
    GEN_REPORT PARM_Company, PARM_WIP, PARM_Weekend, PARM_Offset, PARM_Div_1000, PARM_No_Of_Weeks, _
               DSC
    
    DSC.EndCon_ASIAMFG
    
    Application.ScreenUpdating = True
    
End_Program:
End Sub

Sub GEN_REPORT(ByVal PARM_Company As String, ByVal PARM_WIP As String, ByVal PARM_Weekend As String, _
               ByVal PARM_Offset As String, ByVal PARM_Div_1000 As String, ByVal PARM_No_Of_Weeks As Integer, _
               ByRef DSC As DataSourceConnection)
    
    Dim ARR_CAT() As String
    Dim ARR_CUS() As String
    Dim ARR_OUT() As Double
    Dim ARR_PRC() As String
    Dim ARR_REQ() As Double
    Dim ARR_SKD() As String
    Dim ARR_SKU() As String
    Dim ARR_SUP() As String
    Dim ARR_TOOL() As String
    Dim BIWDDS As String
    Dim C As Long
    Dim CAT As Integer
    Dim COMMENTS As String
    Dim CUR_Bucket As Integer
    Dim CUR_DATETIME As Date
    Dim CUR_yyyyMMddHHmmss As String
    Dim CUS As Integer
    Dim D As Long
    Dim DIPUPWD As String
    Dim DLTFDS As String
    Dim DUDTDS() As String
    Dim DUDTTM() As String
    Dim Due As Integer
    Dim DUE_DTTM As String
    Dim DUTMDS() As String
    Dim FDDTDS As String
    Dim FDTMDS As String
    Dim File_Name As String
    Dim First_Time As Boolean
    Dim Found_DTTM As Boolean
    Dim i As Integer
    Dim JOB_DS As String
    Dim MIXCRW As String
    Dim NO_OF_TOOLS As Double
    Dim N4W() As String
    Dim OFDTDS As String
    Dim OFTMDS As String
    Dim PARTRW As String
    Dim PRC As Integer
    Dim PROD_ID As String
    Dim PROP_VALUE As String
    Dim R As Long
    Dim Rs As New ADODB.Recordset
    Dim Second_Time As Boolean
    Dim SKD As Integer
    Dim SKU As Integer
    Dim Sheet_Name As String
    Dim Stmt As String
    Dim SUP As Integer
    Dim SV_MIXCRW As String
    Dim SV_PARTRW As String
    Dim SV_TOYNRW As String
    Dim TDDTDS As String
    Dim TDDTTM As String
    Dim TDTMDS As String
    Dim TOOL As Integer
    Dim TOYNRW As String
    Dim VS_01 As String
    Dim VS_02 As String
    Dim VS_03 As String
    Dim VS_04 As String
    Dim WDDTTM As String
    Dim WIP_OUT() As Double
    Dim WIP_TYPT() As String
    Dim Work_Date As Date
    
    Work_Date = CDate(Mid(PARM_Weekend, 5, 2) & "/" & Mid(PARM_Weekend, 7, 2) & "/" & Mid(PARM_Weekend, 1, 4))
    
    Work_Date = DateAdd("d", -6, Work_Date)
    OFDTDS = (Year(Work_Date) * 10000) + (Month(Work_Date) * 100) + Day(Work_Date)
    OFTMDS = "190001"
    
    Work_Date = DateAdd("d", 1, Work_Date)
    FDDTDS = (Year(Work_Date) * 10000) + (Month(Work_Date) * 100) + Day(Work_Date)
    FDTMDS = "070000"
    
    Work_Date = DateAdd("d", ((PARM_No_Of_Weeks * 7) - 1), Work_Date)
    TDDTDS = (Year(Work_Date) * 10000) + (Month(Work_Date) * 100) + Day(Work_Date)
    TDTMDS = "190000"
    
    BIWDDS = PARM_Weekend
    
    Work_Date = DateAdd("d", -8, Work_Date)
    
    ReDim N4W(3)
    For i = 0 To 3
        Work_Date = DateAdd("d", 7, Work_Date)
        N4W(i) = (Year(Work_Date) * 10000) + (Month(Work_Date) * 100) + Day(Work_Date)
    Next
    
    WDDTTM = FDDTDS & FDTMDS
    TDDTTM = TDDTDS & TDTMDS
    Due = -1
    While WDDTTM <= TDDTTM
        Due = Due + 1
        ReDim Preserve DUDTDS(Due)
        ReDim Preserve DUTMDS(Due)
        ReDim Preserve DUDTTM(Due)
        DUDTDS(Due) = Mid(WDDTTM, 1, 8)
        DUTMDS(Due) = Mid(WDDTTM, 9, 6)
        DUDTTM(Due) = WDDTTM
        
        Work_Date = FormatDateTime(Mid(WDDTTM, 1, 4) & "/" & Mid(WDDTTM, 5, 2) & "/" & Mid(WDDTTM, 7, 2) & " " & _
                                   Mid(WDDTTM, 9, 2) & ":" & Mid(WDDTTM, 11, 2) & ":" & Mid(WDDTTM, 13, 2))
        Work_Date = DateAdd("h", 12, Work_Date)
        WDDTTM = Year(Work_Date) & Format(Month(Work_Date), "00") & Format(Day(Work_Date), "00") & _
                 Format(Hour(Work_Date), "00") & Format(Minute(Work_Date), "00") & Format(Second(Work_Date), "00")
    Wend
    
    JOB_DS = ""
    Second_Time = False
    DLTFDS = ""
    K4C445 PARM_Company, PARM_WIP, FDDTDS, FDTMDS, TDDTDS, TDTMDS, OFDTDS, OFTMDS, PARM_Offset, BIWDDS, DLTFDS, Second_Time, _
           JOB_DS, DSC
    
    CUR_DATETIME = Now
    CUR_yyyyMMddHHmmss = Format(CUR_DATETIME, "yyyyMMddHHmmss")
    i = 0
    While i <= Due
        If CUR_yyyyMMddHHmmss <= DUDTTM(i) Then
            GoTo Found_CUR
        End If
        i = i + 1
    Wend
    
Found_CUR:
    CUR_Bucket = i + 1
    
    File_Name = "K4DR" & JOB_DS
    'sql statement to retrieve K4XLSWF
    Stmt = "SELECT      * " & _
           "FROM        MMSBQGPL." & File_Name & " " & _
           "WHERE       XWRPID = 'REQUIREMNT' " & _
           "ORDER BY    XWRPID, XWIROW, XWICOL"
    
    Sheet_Name = "Requirement"
    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets(Sheet_Name).Delete
    Err.Clear
    Worksheets.Add.Name = Sheet_Name
    
    Worksheets(Sheet_Name).Activate
    Cells.Clear
    
    Cells.Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    
    'set the current cell to A1
    Range("A1").Activate
    
    'Write Header
    R = 0
    ActiveCell.Offset(R, 0).Value = "Subject:"
    ActiveCell.Offset(R, 1).Value = "Requirement By Workcentre Report"
    
    R = R + 1
    ActiveCell.Offset(R, 0).Value = "Date Generated:"
    ActiveCell.Offset(R, 1).Value = Year(CUR_DATETIME) & "/" & Month(CUR_DATETIME) & "/" & Day(CUR_DATETIME)
    
    R = R + 1
    ActiveCell.Offset(R, 0).Value = "Time Generated:"
    ActiveCell.Offset(R, 1).Value = Hour(CUR_DATETIME) & ":" & Minute(CUR_DATETIME) & ":" & Second(CUR_DATETIME)
    
    R = R + 1
    ActiveCell.Offset(R, 0).Value = "Current Bucket:"
    ActiveCell.Offset(R, 1).Value = CUR_Bucket
    
    For i = 0 To Due
        C = i + 14
        ActiveCell.Offset(R, C).Value = i + 1
        If (CUR_Bucket - 1) = i Then
            ActiveCell.Offset(R, C).Interior.ColorIndex = 3
        End If
    Next
    
    R = R + 1
    ActiveCell.Offset(R, 0).Value = "Toy-Part"
    ActiveCell.Offset(R, 1).Value = "Mix"
    ActiveCell.Offset(R, 2).Value = "Description"
    ActiveCell.Offset(R, 3).Value = "Class"
    ActiveCell.Offset(R, 4).Value = "Sub-Class"
    ActiveCell.Offset(R, 5).Value = "Category"
    ActiveCell.Offset(R, 6).Value = "Category Group"
    ActiveCell.Offset(R, 7).Value = "Supplier"
    ActiveCell.Offset(R, 8).Value = "Customer"
    ActiveCell.Offset(R, 9).Value = "Tool"
    ActiveCell.Offset(R, 10).Value = "SKU#"
    ActiveCell.Offset(R, 11).Value = "Current Inv."
    ActiveCell.Offset(R, 12).Value = "BI"
    ActiveCell.Offset(R, 13).Value = "Row Type"
    
    For i = 0 To Due
        C = i + 14
        ActiveCell.Offset(R, C).Value = Format(DUDTDS(i), "####/##/##") & " " & Format(DUTMDS(i), "##:##:##")
    Next
    
    C = C + 1
    ActiveCell.Offset(R, C).Value = "Total Requirement"
    
    If Trim(PARM_WIP) = "842" Or Trim(PARM_WIP) = "843" Then
        C = C + 1
        ActiveCell.Offset(R, C).Value = "Material"
        C = C + 1
        ActiveCell.Offset(R, C).Value = "Colour"
        C = C + 1
        ActiveCell.Offset(R, C).Value = "Holder"
        C = C + 1
        ActiveCell.Offset(R, C).Value = "Wheel Size"
        C = C + 1
        ActiveCell.Offset(R, C).Value = "Shot Weight"
    End If
    
    If Trim(PARM_WIP) = "833" Or Trim(PARM_WIP) = "842" Or Trim(PARM_WIP) = "843" Then
        C = C + 1
        ActiveCell.Offset(R, C).Value = "Actual Toolset"
        C = C + 1
        ActiveCell.Offset(R, C).Value = "Mold Toy-Part"
        C = C + 1
        ActiveCell.Offset(R, C).Value = "Cavity"
        C = C + 1
        ActiveCell.Offset(R, C).Value = "Cycle Time"
        C = C + 1
        ActiveCell.Offset(R, C).Value = "Mold Type"
        C = C + 1
        ActiveCell.Offset(R, C).Value = "Mold Availability"
    End If
    
    If Trim(PARM_WIP) = "863" Then
        C = C + 1
        ActiveCell.Offset(R, C).Value = "Tampo Family"
    End If
    
    If Trim(PARM_WIP) = "861" Then
        C = C + 1
        ActiveCell.Offset(R, C).Value = "Dipping"
        C = C + 1
        ActiveCell.Offset(R, C).Value = "Trolley Type"
    End If
    
    If Trim(PARM_WIP) = "855" Then
        C = C + 1
        ActiveCell.Offset(R, C).Value = "Axle Length"
        C = C + 1
        ActiveCell.Offset(R, C).Value = "Wheel Size"
    End If
    
    C = C + 1
    ActiveCell.Offset(R, C).Value = "Last Run"
    
    D = C + 1
    For i = 0 To 3
        C = D + i
        ActiveCell.Offset(R, C).Value = Format(N4W(i), "####/##/##")
    Next
    
    C = C + 1
    ActiveCell.Offset(R, C).Value = "Labour Hour"
    C = C + 1
    ActiveCell.Offset(R, C).Value = "Machine Hour"
    C = C + 1
    ActiveCell.Offset(R, C).Value = "Packaging Standard"
    If Trim(PARM_WIP) = "842" Then
        C = C + 1
        ActiveCell.Offset(R, C).Value = "Dipping"
    End If
    C = C + 1
    ActiveCell.Offset(R, C).Value = "SKU Description"
    C = C + 1
    ActiveCell.Offset(R, C).Value = "Bef.Assm. Process"
    C = C + 1
    ActiveCell.Offset(R, C).Value = "Bef.Assm. Component"
        
    DSC.Cmd_ASIAMFG.CommandText = Stmt
    Rs.CursorLocation = adUseClient
    Rs.CacheSize = 100
    Rs.Open DSC.Cmd_ASIAMFG
    
    SV_TOYNRW = ""
    SV_PARTRW = ""
    SV_MIXCRW = ""
    
    First_Time = True
    
    While Not Rs.EOF
        TOYNRW = Mid(Rs.Fields("XWIROW").Value, 1, 5)
        PARTRW = Mid(Rs.Fields("XWIROW").Value, 6, 5)
        MIXCRW = Mid(Rs.Fields("XWIROW").Value, 11, 2)
        PROP_ID = Mid(Rs.Fields("XWICOL").Value, 1, 10)
        PROP_VALUE = Mid(Rs.Fields("XWICOL").Value, 11, 40)
        DUE_DTTM = Mid(Rs.Fields("XWICOL").Value, 11, 14)
        
        If TOYNRW <> SV_TOYNRW Or PARTRW <> SV_PARTRW Or MIXCRW <> SV_MIXCRW Then
        
            If First_Time = False Then
                WRT_REQ SV_TOYNRW, SV_PARTRW, SV_MIXCRW, VS_01, VS_02, VS_03, VS_04, COMMENTS, DIPUPWD, Due, ARR_OUT, ARR_REQ, ARR_PRC, _
                        ARR_CAT, ARR_SUP, ARR_CUS, ARR_SKU, ARR_TOOL, ARR_SKD, R, PARM_Div_1000, PARM_WIP, NO_OF_TOOLS
            End If
            
            SV_TOYNRW = TOYNRW
            SV_PARTRW = PARTRW
            SV_MIXCRW = MIXCRW
            VS_01 = ""
            VS_02 = ""
            VS_03 = ""
            VS_04 = ""
            COMMENTS = ""
            DIPUPWD = ""
            ReDim ARR_CAT(0)
            CAT = -1
            ReDim ARR_REQ(Due)
            TOOL = -1
            ReDim ARR_TOOL(0)
            NO_OF_TOOLS = 0
            ReDim ARR_OUT(Due)
            SUP = -1
            ReDim ARR_SUP(0)
            SKU = -1
            ReDim ARR_SKU(0)
            SKD = -1
            ReDim ARR_SKD(0)
            PRC = -1
            ReDim ARR_PRC(0)
            CUS = -1
            ReDim ARR_CUS(0)
            First_Time = False
        End If

        Select Case Trim(PROP_ID)
            Case "DETL000001"
                VS_01 = Rs.Fields("XWVSTR").Value
            Case "CATEGORY"
                CAT = CAT + 1
                ReDim Preserve ARR_CAT(CAT)
                ARR_CAT(CAT) = Trim(PROP_VALUE)
            Case "REQUIREMNT"
                i = 0
                Found_DTTM = False
                While i <= Due And Found_DTTM = False
                    If DUE_DTTM = DUDTTM(i) Then
                        ARR_REQ(i) = ARR_REQ(i) + Rs.Fields("XWVNUM").Value
                        Found_DTTM = True
                    End If
                    i = i + 1
                Wend
            Case "FLOWCHART"
                VS_02 = Rs.Fields("XWVSTR").Value
            Case "TOOLINFO"
                TOOL = TOOL + 1
                ReDim Preserve ARR_TOOL(TOOL)
                ARR_TOOL(TOOL) = Rs.Fields("XWVSTR").Value
            Case "NOOFTOOLS"
                NO_OF_TOOLS = Rs.Fields("XWVNUM").Value
            Case "NXT4WKSCHD"
                VS_03 = Rs.Fields("XWVSTR").Value
            Case "OUTPUT"
                i = 0
                Found_DTTM = False
                While i <= Due And Found_DTTM = False
                    If DUE_DTTM <= DUDTTM(i) Then
                        ARR_OUT(i) = ARR_OUT(i) + Rs.Fields("XWVNUM").Value
                        Found_DTTM = True
                    End If
                    i = i + 1
                Wend
            Case "SUPPLIER"
                SUP = SUP + 1
                ReDim Preserve ARR_SUP(SUP)
                ARR_SUP(SUP) = Trim(Rs.Fields("XWVSTR").Value)
            Case "AFTDOWNWD"
                VS_04 = Rs.Fields("XWVSTR").Value
            Case "SKU#"
                SKU = SKU + 1
                ReDim Preserve ARR_SKU(SKU)
                ARR_SKU(SKU) = Trim(PROP_VALUE)
            Case "SKUDESC"
                SKD = SKD + 1
                ReDim Preserve ARR_SKD(SKD)
                ARR_SKD(SKD) = Rs.Fields("XWVSTR").Value
            Case "PROCSTRING"
                PRC = PRC + 1
                ReDim Preserve ARR_PRC(PRC)
                ARR_PRC(PRC) = Trim(Rs.Fields("XWVSTR").Value)
            Case "CUSTOMER"
                CUS = CUS + 1
                ReDim Preserve ARR_CUS(CUS)
                ARR_CUS(CUS) = Trim(Rs.Fields("XWVSTR").Value)
            Case "COMMENTS"
                COMMENTS = Rs.Fields("XWVSTR").Value
            Case "DIPUPWD"
                DIPUPWD = Rs.Fields("XWVSTR").Value
        End Select
              
        Rs.MoveNext
    Wend
    
    If First_Time = False Then
        WRT_REQ SV_TOYNRW, SV_PARTRW, SV_MIXCRW, VS_01, VS_02, VS_03, VS_04, COMMENTS, DIPUPWD, Due, ARR_OUT, ARR_REQ, ARR_PRC, _
                ARR_CAT, ARR_SUP, ARR_CUS, ARR_SKU, ARR_TOOL, ARR_SKD, R, PARM_Div_1000, PARM_WIP, NO_OF_TOOLS
    End If

    Rs.Close
    Set Rs = Nothing
       
    Sheet_Name = "Requirement"
    Worksheets(Sheet_Name).Activate
    Rows("5:5").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Selection.Font.ColorIndex = 0
    Selection.Font.Bold = True
    Selection.AutoFilter
    Selection.AutoFilter Field:=14, Criteria1:="<>MISC.", Operator:=xlAnd
    Worksheets(Sheet_Name).Columns("A:A").ColumnWidth = 12
    Worksheets(Sheet_Name).Columns("C:C").ColumnWidth = 25
    Worksheets(Sheet_Name).Columns("F:F").ColumnWidth = 10
    Worksheets(Sheet_Name).Columns("G:G").ColumnWidth = 10
    Worksheets(Sheet_Name).Columns("H:H").ColumnWidth = 15
    Worksheets(Sheet_Name).Columns("I:I").ColumnWidth = 15
    Worksheets(Sheet_Name).Columns("K:K").ColumnWidth = 15
           
    C = 14 + 1 + Due + 1
    Worksheets(Sheet_Name).Columns(C).ColumnWidth = 10
               
    If Trim(PARM_WIP) = "842" Or Trim(PARM_WIP) = "843" Then
        C = C + 1
        Worksheets(Sheet_Name).Columns(C).ColumnWidth = 30
        C = C + 1
        Worksheets(Sheet_Name).Columns(C).ColumnWidth = 25
        C = C + 1
        Worksheets(Sheet_Name).Columns(C).ColumnWidth = 10
    End If
      
    If Trim(PARM_WIP) = "833" Or Trim(PARM_WIP) = "842" Or Trim(PARM_WIP) = "843" Then
        If Trim(PARM_WIP) = "833" Then
            C = C + 2
        Else
            C = C + 4
        End If
        Worksheets(Sheet_Name).Columns(C).ColumnWidth = 15
    End If
    
    Select Case Trim(PARM_WIP)
        Case "833"
            C = 14 + Due + 17
        Case "842", "843"
            C = 14 + Due + 22
        Case "855"
            C = 14 + Due + 13
        Case "861"
            C = 14 + Due + 13
        Case "863"
            C = 14 + Due + 12
        Case Else
            C = 14 + Due + 11
    End Select
    Worksheets(Sheet_Name).Columns(C).ColumnWidth = 15
    
    Range("O6").Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.Zoom = 75
        
    Second_Time = True
    DLTFDS = "Y"
    K4C445 PARM_Company, PARM_WIP, FDDTDS, FDTMDS, TDDTDS, TDTMDS, OFDTDS, OFTMDS, PARM_Offset, BIWDDS, DLTFDS, Second_Time, _
           JOB_DS, DSC
    
End Sub

Sub In_Use(ByVal PARM_Company As String, ByVal PARM_Weekend As String, ByRef In_Use_By As String, _
           ByRef Error_Flag As Boolean, ByRef DSC As DataSourceConnection)
    
    Dim Rs As New ADODB.Recordset
    Dim Stmt As String
    
    Stmt = "SELECT      a.BMIUBY AS a_BMIUBY, b.BMIUBY AS b_BMIUBY " & _
           "FROM        MMSBQGPL.K4JBOBM a " & _
           "LEFT JOIN   MMSBQGPL.KJJBOBM b ON b.BMCMPY = a.BMCMPY AND b.BMBTID = a.BMBTID " & _
           "WHERE       a.BMCMPY = '" & PARM_Company & "' AND a.BMBTID = " & PARM_Weekend & " AND " & _
                       "(a.BMIUBY <> '          ' OR b.BMIUBY <> '          ') "
                       
    DSC.Cmd_ASIAMFG.CommandText = Stmt
    Rs.CursorLocation = adUseClient
    Rs.CacheSize = 100
    Rs.Open DSC.Cmd_ASIAMFG
    
    If Not Rs.EOF Then
        If Trim(Rs.Fields("a_BMIUBY").Value) <> "" Then
            In_Use_By = Rs.Fields("a_BMIUBY").Value
        Else
            In_Use_By = Rs.Fields("b_BMIUBY").Value
        End If
        Error_Flag = True
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub

Sub K4C445(ByVal PARM_Company As String, ByVal PARM_WIP As String, ByVal FDDTDS As String, ByVal FDTMDS As String, _
           ByVal TDDTDS As String, ByVal TDTMDS As String, ByVal OFDTDS As String, ByVal OFTMDS As String, _
           ByVal PARM_Offset As String, ByVal BIWDDS As String, ByVal DLTFDS As String, ByVal Second_Time As Boolean, _
           ByRef JOB_DS As String, ByRef DSC As DataSourceConnection)

    If Second_Time = True Then
        DSC.Cmd_ASIAMFG.Parameters.Delete ("CMPYDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("WIPCDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("FDDTDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("FDTMDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("TDDTDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("TDTMDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("OFDTDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("OFTMDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("OFSTDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("BIWDDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("DLTFDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("JOB_DS")
    End If
    
    DSC.Cmd_ASIAMFG.CommandText = "{{CALL MMSBQGPL.K4C445_@ (?,?,?,?,?,?,?,?,?,?,?,?)}}"
    'Add parameters
    DSC.Cmd_ASIAMFG.Parameters.Append DSC.Cmd_ASIAMFG.CreateParameter("CMPYDS", adChar, adParamInput, 4, PARM_Company)
    DSC.Cmd_ASIAMFG.Parameters.Append DSC.Cmd_ASIAMFG.CreateParameter("WIPCDS", adChar, adParamInput, 10, PARM_WIP)
    DSC.Cmd_ASIAMFG.Parameters.Append DSC.Cmd_ASIAMFG.CreateParameter("FDDTDS", adChar, adParamInput, 8, FDDTDS)
    DSC.Cmd_ASIAMFG.Parameters.Append DSC.Cmd_ASIAMFG.CreateParameter("FDTMDS", adChar, adParamInput, 6, FDTMDS)
    DSC.Cmd_ASIAMFG.Parameters.Append DSC.Cmd_ASIAMFG.CreateParameter("TDDTDS", adChar, adParamInput, 8, TDDTDS)
    DSC.Cmd_ASIAMFG.Parameters.Append DSC.Cmd_ASIAMFG.CreateParameter("TDTMDS", adChar, adParamInput, 6, TDTMDS)
    DSC.Cmd_ASIAMFG.Parameters.Append DSC.Cmd_ASIAMFG.CreateParameter("OFDTDS", adChar, adParamInput, 8, OFDTDS)
    DSC.Cmd_ASIAMFG.Parameters.Append DSC.Cmd_ASIAMFG.CreateParameter("OFTMDS", adChar, adParamInput, 6, OFTMDS)
    DSC.Cmd_ASIAMFG.Parameters.Append DSC.Cmd_ASIAMFG.CreateParameter("OFSTDS", adChar, adParamInput, 1, PARM_Offset)
    DSC.Cmd_ASIAMFG.Parameters.Append DSC.Cmd_ASIAMFG.CreateParameter("BIWDDS", adChar, adParamInput, 8, BIWDDS)
    DSC.Cmd_ASIAMFG.Parameters.Append DSC.Cmd_ASIAMFG.CreateParameter("DLTFDS", adChar, adParamInput, 1, DLTFDS)
    DSC.Cmd_ASIAMFG.Parameters.Append DSC.Cmd_ASIAMFG.CreateParameter("JOB_DS", adChar, adParamInputOutput, 6, JOB_DS)
    DSC.Cmd_ASIAMFG.Execute
    JOB_DS = DSC.Cmd_ASIAMFG.Parameters("JOB_DS").Value
    
    If Second_Time = True Then
        DSC.Cmd_ASIAMFG.Parameters.Delete ("CMPYDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("WIPCDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("FDDTDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("FDTMDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("TDDTDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("TDTMDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("OFDTDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("OFTMDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("OFSTDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("BIWDDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("DLTFDS")
        DSC.Cmd_ASIAMFG.Parameters.Delete ("JOB_DS")
    End If
    
End Sub

Sub WRT_REQ(ByVal SV_TOYNRW As String, ByVal SV_PARTRW As String, ByVal SV_MIXCRW As String, ByVal VS_01 As String, _
            ByVal VS_02 As String, ByVal VS_03 As String, ByVal VS_04 As String, ByVal COMMENTS As String, ByVal DIPUPWD As String, _
            ByVal Due As Integer, ByRef ARR_OUT() As Double, ByRef ARR_REQ() As Double, ByRef ARR_PRC() As String, _
            ByRef ARR_CAT() As String, ByRef ARR_SUP() As String, ByRef ARR_CUS() As String, ByRef ARR_SKU() As String, _
            ByRef ARR_TOOL() As String, ByRef ARR_SKD() As String, ByRef R As Long, ByVal PARM_Div_1000 As String, _
            ByVal PARM_WIP As String, ByVal NO_OF_TOOLS As Double)

    Dim BFA As Integer
    Dim C As Integer
    Dim cm As Integer
    Dim COMM() As String
    Dim COMP_BFA() As String
    Dim CUMM_VAR() As Double
    Dim D As Integer
    Dim Found_Flag As Boolean
    Dim GRP As String
    Dim i As Integer
    Dim j As Integer
    Dim Pos As Integer
    Dim PROC_BFA() As String
    Dim TOT_OUT As Double
    Dim TOT_REQ As Double
    Dim Work_Split() As String
    Dim Work_String_0 As String
    Dim Work_String_1 As String
    Dim Work_Sub_Split() As String
    Dim x As Integer
    Dim y As Integer
    
    Dim VS_01_DESXIT As String
    Dim VS_01_CICODE As String
    Dim VS_01_CISCOD As String
    Dim VS_01_BISG As String
    Dim VS_01_BI As Long
    Dim VS_01_CISG As String
    Dim VS_01_CI As Long
    Dim VS_01_BNBSTD As Long
    
    Dim VS_02_MATLDS As String
    Dim VS_02_COLODS As String
    Dim VS_02_HOLDDS As String
    Dim VS_02_SIZEDS As String
    Dim VS_02_SHWTDS As Double
    Dim VS_02_FAMLDS As String
    Dim VS_02_DIPCDS As String
    Dim VS_02_PCSHDS As Long
    Dim VS_02_AXLGDS As String
    Dim VS_02_LSTDDS As Double
    Dim VS_02_MSTDDS As Double
    Dim VS_02_NFMLDS As String
    
    Dim VS_03_LRUNDS As String
    Dim VS_03_SCH() As Long
    
    Dim VS_04_WHLSIZ As String
    Dim VS_04_TOOLDS As String
    
    'Breakdown VS_01 Data ---------------------------------------------------------------------------------------------------------------------
    VS_01_DESXIT = Mid(VS_01, 1, 25)
    VS_01_CICODE = Mid(VS_01, 26, 3)
    VS_01_CISCOD = Mid(VS_01, 29, 3)
    VS_01_BISG = Mid(VS_01, 32, 1)
    VS_01_BI = CLng(Mid(VS_01, 33, 9))
    If VS_01_BISG = "-" Then
        VS_01_BI = VS_01_BI * -1
    End If
    VS_01_CISG = Mid(VS_01, 42, 1)
    VS_01_CI = CLng(Mid(VS_01, 43, 9))
    If VS_01_CISG = "-" Then
        VS_01_CI = VS_01_CI * -1
    End If
    VS_01_BNBSTD = CLng(Mid(VS_01, 52, 4))
    
    'Breakdown VS_02 Data ---------------------------------------------------------------------------------------------------------------------
    VS_02_MATLDS = Mid(VS_02, 1, 30)
    VS_02_COLODS = Mid(VS_02, 31, 25)
    VS_02_HOLDDS = Mid(VS_02, 56, 15)
    VS_02_SIZEDS = Mid(VS_02, 71, 5)
    VS_02_SHWTDS = Mid(VS_02, 76, 5) / 100
    VS_02_FAMLDS = Mid(VS_02, 81, 2)
    VS_02_DIPCDS = Mid(VS_02, 83, 10)
    VS_02_PCSHDS = CLng(Mid(VS_02, 93, 9))
    VS_02_AXLGDS = Mid(VS_02, 102, 3)
    VS_02_LSTDDS = Mid(VS_02, 105, 9) / 1000
    VS_02_MSTDDS = Mid(VS_02, 114, 9) / 1000
    VS_02_NFMLDS = Mid(VS_02, 123, 50)
    
    'Breakdown VS_03 Data ---------------------------------------------------------------------------------------------------------------------
    VS_03_LRUNDS = Mid(VS_03, 1, 1)
    ReDim VS_03_SCH(3)
    For i = 0 To 3
        j = (i * 11) + 2
        VS_03_SCH(i) = CLng(Mid(VS_03, j, 11))
    Next
    'Breakdown VS_04 Data ---------------------------------------------------------------------------------------------------------------------
    VS_04_WHLSIZ = Mid(VS_04, 1, 5)
    VS_04_TOOLDS = Mid(VS_04, 6, 5)
    'Breakdow COMMENTS ------------------------------------------------------------------------------------------------------------------------
    If InStr(COMMENTS, ";") > 0 Then
        COMM = Split(COMMENTS, ";")
    Else
        ReDim COMM(0)
        COMM(0) = Trim(COMMENTS)
    End If
    '------------------------------------------------------------------------------------------------------------------------------------------
    
    '-----!!!! GetWIMP
    
    ReDim CUMM_VAR(Due)
    For i = 0 To Due
        If i = 0 Then
            CUMM_VAR(i) = VS_01_BI + ARR_OUT(i) - ARR_REQ(i)
        Else
            CUMM_VAR(i) = CUMM_VAR(i - 1) + ARR_OUT(i) - ARR_REQ(i)
        End If
    Next
    
    BFA = -1
    ReDim PROC_BFA(0)
    ReDim COMP_BFA(0)
    For i = 0 To UBound(ARR_PRC)
        Work_Split = Split(ARR_PRC(i), ";")
        Work_String_0 = ""
        Work_String_1 = ""
        For x = LBound(Work_Split) To UBound(Work_Split)
            Pos = InStr(Work_Split(x), ",")
            If Pos <> 0 Then
                Work_Sub_Split = Split(Work_Split(x), ",")
                If Trim(Work_Sub_Split(0)) = "851" Then
                    GoTo Found_851
                Else
                    Work_String_0 = Work_Sub_Split(0)
                    Work_String_1 = Work_Sub_Split(1)
                End If
            End If
        Next
        
Found_851:
        If Trim(Work_String_1) <> "" Then
            Found_Flag = False
            x = 0
            While x <= BFA And Found_Flag = False
                If Trim(Work_String_1) = Trim(COMP_BFA(x)) Then
                    Found_Flag = True
                End If
                x = x + 1
            Wend
            If Found_Flag = False Then
                BFA = BFA + 1
                ReDim Preserve PROC_BFA(BFA)
                ReDim Preserve COMP_BFA(BFA)
                PROC_BFA(BFA) = Work_String_0
                COMP_BFA(BFA) = Work_String_1
            End If
        End If
        
    Next
    
    Sheet_Name = "Requirement"
    Worksheets(Sheet_Name).Activate
    Range("A1").Activate
'START: Write 'REQ' line -----------------------------------------------------------------------------------------------------------------
    R = R + 1
    ActiveCell.Offset(R, 0).Value = Trim(SV_TOYNRW) & "-" & Trim(SV_PARTRW)
    ActiveCell.Offset(R, 1).Value = SV_MIXCRW
    ActiveCell.Offset(R, 2).Value = VS_01_DESXIT
    If Trim(PARM_WIP) = "855" Then
        For cm = 0 To UBound(COMM)
            If Trim(COMM(cm)) <> "" Then
                ActiveCell.Offset(R, 2).Value = Trim(ActiveCell.Offset(R, 2).Value) & Chr(10) & Trim(COMM(cm))
            End If
        Next
    End If
    ActiveCell.Offset(R, 3).Value = VS_01_CICODE
    ActiveCell.Offset(R, 4).Value = VS_01_CISCOD
    
    For i = 0 To UBound(ARR_CAT)
        If i = 0 Then
            ActiveCell.Offset(R, 5).Value = ARR_CAT(i)
        Else
            ActiveCell.Offset(R, 5).Value = ActiveCell.Offset(R, 5).Value & "," & ARR_CAT(i)
        End If
    Next
    
    GET_CATGRP ARR_CAT, GRP
    
    TOT_REQ = 0
    For i = 0 To Due
        TOT_REQ = TOT_REQ + ARR_REQ(i)
    Next
    
    If TOT_REQ = 0 Then
        GRP = "No Schedule"
    End If
    ActiveCell.Offset(R, 6).Value = GRP
    
    For i = 0 To UBound(ARR_SUP)
        If i = 0 Then
            ActiveCell.Offset(R, 7).Value = ARR_SUP(i)
        Else
            ActiveCell.Offset(R, 7).Value = ActiveCell.Offset(R, 7).Value & Chr(10) & ARR_SUP(i)
        End If
    Next
       
    For i = 0 To UBound(ARR_CUS)
        If i = 0 Then
            ActiveCell.Offset(R, 8).Value = ARR_CUS(i)
        Else
            ActiveCell.Offset(R, 8).Value = ActiveCell.Offset(R, 8).Value & Chr(10) & ARR_CUS(i)
        End If
    Next
    
    ActiveCell.Offset(R, 9).Value = VS_04_TOOLDS
    
    For i = 0 To UBound(ARR_SKU)
        If i = 0 Then
            ActiveCell.Offset(R, 10).Value = ARR_SKU(i)
        Else
            ActiveCell.Offset(R, 10).Value = ActiveCell.Offset(R, 10).Value & "," & ARR_SKU(i)
        End If
    Next
    
    If PARM_Div_1000 = "Y" Then
        ActiveCell.Offset(R, 11).Value = Round((VS_01_CI / 1000), 1)
        ActiveCell.Offset(R, 12).Value = Round((VS_01_BI / 1000), 1)
    Else
        ActiveCell.Offset(R, 11).Value = VS_01_CI
        ActiveCell.Offset(R, 12).Value = VS_01_BI
    End If
    ActiveCell.Offset(R, 13).Value = "REQ."
    ActiveCell.Offset(R, 13).Font.ColorIndex = 3
    
    TOT_REQ = 0
    For i = 0 To Due
        C = i + 14
        If PARM_Div_1000 = "Y" Then
            ActiveCell.Offset(R, C).Value = Round((ARR_REQ(i) / 1000), 1)
        Else
            ActiveCell.Offset(R, C).Value = Round(ARR_REQ(i), 3)
        End If
        TOT_REQ = TOT_REQ + ARR_REQ(i)
    Next
    
    C = C + 1
    If PARM_Div_1000 = "Y" Then
        ActiveCell.Offset(R, C).Value = Round((TOT_REQ / 1000), 1)
    Else
        ActiveCell.Offset(R, C).Value = TOT_REQ
    End If
    
    If Trim(PARM_WIP) = "842" Or Trim(PARM_WIP) = "843" Then
        C = C + 1
        ActiveCell.Offset(R, C).Value = VS_02_MATLDS
        C = C + 1
        ActiveCell.Offset(R, C).Value = VS_02_COLODS
        C = C + 1
        ActiveCell.Offset(R, C).Value = VS_02_HOLDDS
        C = C + 1
        ActiveCell.Offset(R, C).Value = VS_02_SIZEDS
        C = C + 1
        ActiveCell.Offset(R, C).Value = VS_02_SHWTDS
    End If
    
    If Trim(PARM_WIP) = "833" Or Trim(PARM_WIP) = "842" Or Trim(PARM_WIP) = "843" Then
        C = C + 1
        ActiveCell.Offset(R, C).Value = NO_OF_TOOLS
        If UBound(ARR_TOOL) > -1 Then
            Work_String_0 = ARR_TOOL(0)
            Work_String_0 = Replace(Work_String_0, ";", "")
            If InStr(Trim(Work_String_0), ",") Then
                Work_Split = Split(Work_String_0, ",")
                C = C + 1
                ActiveCell.Offset(R, C).Value = Work_Split(0)
                C = C + 1
                ActiveCell.Offset(R, C).Value = CDbl(Work_Split(1))
                C = C + 1
                ActiveCell.Offset(R, C).Value = CDbl(Work_Split(2))
                C = C + 1
                ActiveCell.Offset(R, C).Value = Work_Split(3)
                C = C + 1
                ActiveCell.Offset(R, C).Value = CDbl(Work_Split(4))
            Else
                C = C + 5
            End If
        Else
            C = C + 5
        End If
    End If
    
    If Trim(PARM_WIP) = "863" Then
        C = C + 1
        If VS_02_NFMLDS <> "" Then
            ActiveCell.Offset(R, C).Value = VS_02_NFMLDS
        Else
            ActiveCell.Offset(R, C).Value = VS_02_FAMLDS
        End If
    End If
    
    If Trim(PARM_WIP) = "861" Then
        C = C + 1
        ActiveCell.Offset(R, C).Value = VS_02_DIPCDS
        C = C + 1
        ActiveCell.Offset(R, C).Value = VS_02_PCSHDS
    End If
    
    If Trim(PARM_WIP) = "855" Then
        C = C + 1
        ActiveCell.Offset(R, C).Value = VS_02_AXLGDS
        C = C + 1
        ActiveCell.Offset(R, C).Value = VS_04_WHLSIZ
    End If
    
    C = C + 1
    ActiveCell.Offset(R, C).Value = VS_03_LRUNDS
    
    D = C + 1
    For i = 0 To 3
        C = D + i
        If PARM_Div_1000 = "Y" Then
            ActiveCell.Offset(R, C).Value = Round((VS_03_SCH(i) / 1000), 1)
        Else
            ActiveCell.Offset(R, C).Value = VS_03_SCH(i)
        End If
    Next
    
    C = C + 1
    ActiveCell.Offset(R, C).Value = VS_02_LSTDDS
    C = C + 1
    ActiveCell.Offset(R, C).Value = VS_02_MSTDDS
    
    C = C + 1
    ActiveCell.Offset(R, C).Value = VS_01_BNBSTD
    
    If Trim(PARM_WIP) = "842" Then
        C = C + 1
        ActiveCell.Offset(R, C).Value = DIPUPWD
    End If
    
    C = C + 1
    For i = 0 To UBound(ARR_SKD)
        If i = 0 Then
            ActiveCell.Offset(R, C).Value = ARR_SKD(i)
        Else
            ActiveCell.Offset(R, C).Value = ActiveCell.Offset(R, C).Value & "," & ARR_SKD(i)
        End If
    Next
    
    If BFA > -1 Then
        C = C + 1
        ActiveCell.Offset(R, C).Value = Trim(PROC_BFA(0))
        C = C + 1
        ActiveCell.Offset(R, C).Value = Trim(COMP_BFA(0))
    Else
        C = C + 2
    End If
    
    If UBound(ARR_PRC) > -1 Then
        Work_Split = Split(ARR_PRC(0), ";")
        For x = LBound(Work_Split) To UBound(Work_Split)
            Work_Sub_Split = Split(Work_Split(x), ",")
            For y = LBound(Work_Sub_Split) To UBound(Work_Sub_Split)
                C = C + 1
                ActiveCell.Offset(R, C).Value = Trim(Work_Sub_Split(y))
            Next
        Next
    End If
    
    Rows(R).RowHeight = 13.5
'END  : Write 'REQ' line -----------------------------------------------------------------------------------------------------------------
    
'START: Write 'OUTPUT' line --------------------------------------------------------------------------------------------------------------
    R = R + 1
    ActiveCell.Offset(R, 0).Value = Trim(SV_TOYNRW) & "-" & Trim(SV_PARTRW)
    ActiveCell.Offset(R, 0).Font.ColorIndex = 2
    ActiveCell.Offset(R, 1).Value = SV_MIXCRW
    ActiveCell.Offset(R, 1).Font.ColorIndex = 2
    ActiveCell.Offset(R, 2).Value = VS_01_DESXIT
    ActiveCell.Offset(R, 2).Font.ColorIndex = 2
    ActiveCell.Offset(R, 3).Value = VS_01_CICODE
    ActiveCell.Offset(R, 3).Font.ColorIndex = 2
    ActiveCell.Offset(R, 4).Value = VS_01_CISCOD
    ActiveCell.Offset(R, 4).Font.ColorIndex = 2
    
    For i = 0 To UBound(ARR_CAT)
        If i = 0 Then
            ActiveCell.Offset(R, 5).Value = ARR_CAT(i)
        Else
            ActiveCell.Offset(R, 5).Value = ActiveCell.Offset(R, 5).Value & "," & ARR_CAT(i)
        End If
    Next
    ActiveCell.Offset(R, 5).Font.ColorIndex = 2
    
    ActiveCell.Offset(R, 6).Value = GRP
    ActiveCell.Offset(R, 6).Font.ColorIndex = 2
    
    For i = 0 To UBound(ARR_SUP)
        If i = 0 Then
            ActiveCell.Offset(R, 7).Value = ARR_SUP(i)
        Else
            ActiveCell.Offset(R, 7).Value = ActiveCell.Offset(R, 7).Value & Chr(10) & ARR_SUP(i)
        End If
    Next
    ActiveCell.Offset(R, 7).Font.ColorIndex = 2
       
    For i = 0 To UBound(ARR_CUS)
        If i = 0 Then
            ActiveCell.Offset(R, 8).Value = ARR_CUS(i)
        Else
            ActiveCell.Offset(R, 8).Value = ActiveCell.Offset(R, 8).Value & Chr(10) & ARR_CUS(i)
        End If
    Next
    ActiveCell.Offset(R, 8).Font.ColorIndex = 2
    
    ActiveCell.Offset(R, 9).Value = VS_04_TOOLDS
    ActiveCell.Offset(R, 9).Font.ColorIndex = 2
    
    For i = 0 To UBound(ARR_SKU)
        If i = 0 Then
            ActiveCell.Offset(R, 10).Value = ARR_SKU(i)
        Else
            ActiveCell.Offset(R, 10).Value = ActiveCell.Offset(R, 10).Value & "," & ARR_SKU(i)
        End If
    Next
    ActiveCell.Offset(R, 10).Font.ColorIndex = 2
    
    ActiveCell.Offset(R, 13).Value = "OUTPUT"
    ActiveCell.Offset(R, 13).Font.ColorIndex = 43
    
    TOT_OUT = 0
    For i = 0 To Due
        C = i + 14
        If PARM_Div_1000 = "Y" Then
            ActiveCell.Offset(R, C).Value = Round((ARR_OUT(i) / 1000), 1)
        Else
            ActiveCell.Offset(R, C).Value = Round(ARR_OUT(i), 3)
        End If
        TOT_OUT = TOT_OUT + ARR_OUT(i)
    Next
    
    C = C + 1
    If PARM_Div_1000 = "Y" Then
        ActiveCell.Offset(R, C).Value = Round((TOT_OUT / 1000), 1)
    Else
        ActiveCell.Offset(R, C).Value = Round(TOT_OUT, 3)
    End If
    If TOT_REQ > TOT_OUT Then
        ActiveCell.Offset(R, C).Font.ColorIndex = 3
    End If
    
    
    If Trim(PARM_WIP) = "833" Or Trim(PARM_WIP) = "842" Or Trim(PARM_WIP) = "843" Then
        If UBound(ARR_TOOL) > 0 Then
            Work_String_0 = ARR_TOOL(1)
            Work_String_0 = Replace(Work_String_0, ";", "")
            Work_Split = Split(Work_String_0, ",")
            
            If Trim(PARM_WIP) = "842" Or Trim(PARM_WIP) = "843" Then
                C = C + 7
            Else
                C = C + 2
            End If
            ActiveCell.Offset(R, C).Value = Work_Split(0)
            C = C + 1
            ActiveCell.Offset(R, C).Value = CDbl(Work_Split(1))
            C = C + 1
            ActiveCell.Offset(R, C).Value = CDbl(Work_Split(2))
            C = C + 1
            ActiveCell.Offset(R, C).Value = Work_Split(3)
            C = C + 1
            ActiveCell.Offset(R, C).Value = CDbl(Work_Split(4))
        Else
            If Trim(PARM_WIP) = "842" Or Trim(PARM_WIP) = "843" Then
                C = C + 11
            Else
                C = C + 6
            End If
        End If
    End If
    
    Select Case Trim(PARM_WIP)
        Case "833"
            C = C + 7
        Case "842", "843"
            C = C + 7
        Case "855"
            C = C + 9
        Case "861"
            C = C + 9
        Case "863"
            C = C + 8
        Case Else
            C = C + 7
    End Select
    
    C = C + 2
    If Trim(PARM_WIP) = "842" Then
        C = C + 1
    End If
    
    For i = 0 To UBound(ARR_SKD)
        If i = 0 Then
            ActiveCell.Offset(R, C).Value = ARR_SKD(i)
        Else
            ActiveCell.Offset(R, C).Value = ActiveCell.Offset(R, C).Value & "," & ARR_SKD(i)
        End If
    Next
    
    If BFA > 0 Then
        C = C + 1
        ActiveCell.Offset(R, C).Value = Trim(PROC_BFA(1))
        C = C + 1
        ActiveCell.Offset(R, C).Value = Trim(COMP_BFA(1))
    Else
        C = C + 2
    End If
    
    If UBound(ARR_PRC) > 0 Then
        Work_Split = Split(ARR_PRC(1), ";")
        For x = LBound(Work_Split) To UBound(Work_Split)
            Work_Sub_Split = Split(Work_Split(x), ",")
            For y = LBound(Work_Sub_Split) To UBound(Work_Sub_Split)
                C = C + 1
                ActiveCell.Offset(R, C).Value = Trim(Work_Sub_Split(y))
            Next
        Next
    End If
    
    Rows(R).RowHeight = 13.5
'END  : Write 'OUTPUT' line --------------------------------------------------------------------------------------------------------------
    
'START: Write 'CUMVAR' line --------------------------------------------------------------------------------------------------------------
    R = R + 1
    ActiveCell.Offset(R, 0).Value = Trim(SV_TOYNRW) & "-" & Trim(SV_PARTRW)
    ActiveCell.Offset(R, 0).Font.ColorIndex = 2
    ActiveCell.Offset(R, 1).Value = SV_MIXCRW
    ActiveCell.Offset(R, 1).Font.ColorIndex = 2
    ActiveCell.Offset(R, 2).Value = VS_01_DESXIT
    ActiveCell.Offset(R, 2).Font.ColorIndex = 2
    ActiveCell.Offset(R, 3).Value = VS_01_CICODE
    ActiveCell.Offset(R, 3).Font.ColorIndex = 2
    ActiveCell.Offset(R, 4).Value = VS_01_CISCOD
    ActiveCell.Offset(R, 4).Font.ColorIndex = 2
    
    For i = 0 To UBound(ARR_CAT)
        If i = 0 Then
            ActiveCell.Offset(R, 5).Value = ARR_CAT(i)
        Else
            ActiveCell.Offset(R, 5).Value = ActiveCell.Offset(R, 5).Value & "," & ARR_CAT(i)
        End If
    Next
    ActiveCell.Offset(R, 5).Font.ColorIndex = 2
    
    ActiveCell.Offset(R, 6).Value = GRP
    ActiveCell.Offset(R, 6).Font.ColorIndex = 2
    
    For i = 0 To UBound(ARR_SUP)
        If i = 0 Then
            ActiveCell.Offset(R, 7).Value = ARR_SUP(i)
        Else
            ActiveCell.Offset(R, 7).Value = ActiveCell.Offset(R, 7).Value & Chr(10) & ARR_SUP(i)
        End If
    Next
    ActiveCell.Offset(R, 7).Font.ColorIndex = 2
       
    For i = 0 To UBound(ARR_CUS)
        If i = 0 Then
            ActiveCell.Offset(R, 8).Value = ARR_CUS(i)
        Else
            ActiveCell.Offset(R, 8).Value = ActiveCell.Offset(R, 8).Value & Chr(10) & ARR_CUS(i)
        End If
    Next
    ActiveCell.Offset(R, 8).Font.ColorIndex = 2
    
    ActiveCell.Offset(R, 9).Value = VS_04_TOOLDS
    ActiveCell.Offset(R, 9).Font.ColorIndex = 2
    
    For i = 0 To UBound(ARR_SKU)
        If i = 0 Then
            ActiveCell.Offset(R, 10).Value = ARR_SKU(i)
        Else
            ActiveCell.Offset(R, 10).Value = ActiveCell.Offset(R, 10).Value & "," & ARR_SKU(i)
        End If
    Next
    ActiveCell.Offset(R, 10).Font.ColorIndex = 2
    
    ActiveCell.Offset(R, 13).Value = "CUMVAR"
    ActiveCell.Offset(R, 13).Font.ColorIndex = 14
    
    For i = 0 To Due
        C = i + 14
        If PARM_Div_1000 = "Y" Then
            ActiveCell.Offset(R, C).Value = Round((CUMM_VAR(i) / 1000), 1)
        Else
            ActiveCell.Offset(R, C).Value = Round(CUMM_VAR(i), 3)
        End If
        If CUMM_VAR(i) < 0 Then
            ActiveCell.Offset(R, C).Font.ColorIndex = 3
        End If
    Next
    
    C = C + 1
    If PARM_Div_1000 = "Y" Then
        ActiveCell.Offset(R, C).Value = Round((CUMM_VAR(Due) / 1000), 1)
    Else
        ActiveCell.Offset(R, C).Value = Round(CUMM_VAR(Due), 3)
    End If
    If CUMM_VAR(Due) < 0 Then
        ActiveCell.Offset(R, C).Font.ColorIndex = 3
    End If
    
    
    If Trim(PARM_WIP) = "833" Or Trim(PARM_WIP) = "842" Or Trim(PARM_WIP) = "843" Then
        If UBound(ARR_TOOL) > 1 Then
            Work_String_0 = ARR_TOOL(2)
            Work_String_0 = Replace(Work_String_0, ";", "")
            Work_Split = Split(Work_String_0, ",")
            
            If Trim(PARM_WIP) = "842" Or Trim(PARM_WIP) = "843" Then
                C = C + 7
            Else
                C = C + 2
            End If
            ActiveCell.Offset(R, C).Value = Work_Split(0)
            C = C + 1
            ActiveCell.Offset(R, C).Value = CDbl(Work_Split(1))
            C = C + 1
            ActiveCell.Offset(R, C).Value = CDbl(Work_Split(2))
            C = C + 1
            ActiveCell.Offset(R, C).Value = Work_Split(3)
            C = C + 1
            ActiveCell.Offset(R, C).Value = CDbl(Work_Split(4))
        Else
            If Trim(PARM_WIP) = "842" Or Trim(PARM_WIP) = "843" Then
                C = C + 11
            Else
                C = C + 6
            End If
        End If
    End If
    
    Select Case Trim(PARM_WIP)
        Case "833"
            C = C + 7
        Case "842", "843"
            C = C + 7
        Case "855"
            C = C + 9
        Case "861"
            C = C + 9
        Case "863"
            C = C + 8
        Case Else
            C = C + 7
    End Select
    
    C = C + 2
    If Trim(PARM_WIP) = "842" Then
        C = C + 1
    End If
    
    For i = 0 To UBound(ARR_SKD)
        If i = 0 Then
            ActiveCell.Offset(R, C).Value = ARR_SKD(i)
        Else
            ActiveCell.Offset(R, C).Value = ActiveCell.Offset(R, C).Value & "," & ARR_SKD(i)
        End If
    Next
    
    If BFA > 1 Then
        C = C + 1
        ActiveCell.Offset(R, C).Value = Trim(PROC_BFA(2))
        C = C + 1
        ActiveCell.Offset(R, C).Value = Trim(COMP_BFA(2))
    Else
        C = C + 2
    End If

    If UBound(ARR_PRC) > 1 Then
        Work_Split = Split(ARR_PRC(2), ";")
        For x = LBound(Work_Split) To UBound(Work_Split)
            Work_Sub_Split = Split(Work_Split(x), ",")
            For y = LBound(Work_Sub_Split) To UBound(Work_Sub_Split)
                C = C + 1
                ActiveCell.Offset(R, C).Value = Trim(Work_Sub_Split(y))
            Next
        Next
    End If
    
    Rows(R).RowHeight = 13.5
'END  : Write 'CUMVAR' line --------------------------------------------------------------------------------------------------------------
    
'START: Write 'MISC' line ----------------------------------------------------------------------------------------------------------------
    If UBound(ARR_TOOL) > 2 Or UBound(ARR_PRC) > 2 Or BFA > 2 Then
        j = 3
        While j <= UBound(ARR_TOOL) Or j <= UBound(ARR_PRC) Or j <= BFA
            R = R + 1
            ActiveCell.Offset(R, 0).Value = Trim(SV_TOYNRW) & "-" & Trim(SV_PARTRW)
            ActiveCell.Offset(R, 0).Font.ColorIndex = 2
            ActiveCell.Offset(R, 1).Value = SV_MIXCRW
            ActiveCell.Offset(R, 1).Font.ColorIndex = 2
            ActiveCell.Offset(R, 2).Value = VS_01_DESXIT
            ActiveCell.Offset(R, 2).Font.ColorIndex = 2
            ActiveCell.Offset(R, 3).Value = VS_01_CICODE
            ActiveCell.Offset(R, 3).Font.ColorIndex = 2
            ActiveCell.Offset(R, 4).Value = VS_01_CISCOD
            ActiveCell.Offset(R, 4).Font.ColorIndex = 2
            
            For i = 0 To UBound(ARR_CAT)
                If i = 0 Then
                    ActiveCell.Offset(R, 5).Value = ARR_CAT(i)
                Else
                    ActiveCell.Offset(R, 5).Value = ActiveCell.Offset(R, 5).Value & "," & ARR_CAT(i)
                End If
            Next
            ActiveCell.Offset(R, 5).Font.ColorIndex = 2
            
            ActiveCell.Offset(R, 6).Value = GRP
            ActiveCell.Offset(R, 6).Font.ColorIndex = 2
            
            For i = 0 To UBound(ARR_SUP)
                If i = 0 Then
                    ActiveCell.Offset(R, 7).Value = ARR_SUP(i)
                Else
                    ActiveCell.Offset(R, 7).Value = ActiveCell.Offset(R, 7).Value & Chr(10) & ARR_SUP(i)
                End If
            Next
            ActiveCell.Offset(R, 7).Font.ColorIndex = 2

            For i = 0 To UBound(ARR_CUS)
                If i = 0 Then
                    ActiveCell.Offset(R, 8).Value = ARR_CUS(i)
                Else
                    ActiveCell.Offset(R, 8).Value = ActiveCell.Offset(R, 8).Value & Chr(10) & ARR_CUS(i)
                End If
            Next
            ActiveCell.Offset(R, 8).Font.ColorIndex = 2
            
            ActiveCell.Offset(R, 9).Value = VS_04_TOOLDS
            ActiveCell.Offset(R, 9).Font.ColorIndex = 2
            
            For i = 0 To UBound(ARR_SKU)
                If i = 0 Then
                    ActiveCell.Offset(R, 10).Value = ARR_SKU(i)
                Else
                    ActiveCell.Offset(R, 10).Value = ActiveCell.Offset(R, 10).Value & "," & ARR_SKU(i)
                End If
            Next
            ActiveCell.Offset(R, 10).Font.ColorIndex = 2
            
            ActiveCell.Offset(R, 13).Value = "MISC."
            
            If j <= UBound(ARR_TOOL) Then
                Work_String_0 = ARR_TOOL(j)
                Work_String_0 = Replace(Work_String_0, ";", "")
                Work_Split = Split(Work_String_0, ",")
            
                If Trim(PARM_WIP) = "842" Or Trim(PARM_WIP) = "843" Then
                    C = 14 + Due + 8
                Else
                    C = 14 + Due + 3
                End If
                ActiveCell.Offset(R, C).Value = Work_Split(0)
                C = C + 1
                ActiveCell.Offset(R, C).Value = CDbl(Work_Split(1))
                C = C + 1
                ActiveCell.Offset(R, C).Value = CDbl(Work_Split(2))
                C = C + 1
                ActiveCell.Offset(R, C).Value = Work_Split(3)
                C = C + 1
                ActiveCell.Offset(R, C).Value = CDbl(Work_Split(4))
            End If
            
            Select Case Trim(PARM_WIP)
                Case "833"
                    C = 14 + Due + 14
                Case "842", "843"
                    C = 14 + Due + 19
                Case "855"
                    C = 14 + Due + 10
                Case "861"
                    C = 14 + Due + 10
                Case "863"
                    C = 14 + Due + 9
                Case Else
                    C = 14 + Due + 8
            End Select
            
            C = C + 2
            If Trim(PARM_WIP) = "842" Then
                C = C + 1
            End If
            
            For i = 0 To UBound(ARR_SKD)
                If i = 0 Then
                    ActiveCell.Offset(R, C).Value = ARR_SKD(i)
                Else
                    ActiveCell.Offset(R, C).Value = ActiveCell.Offset(R, C).Value & "," & ARR_SKD(i)
                End If
            Next
            
            If j <= BFA Then
                Select Case Trim(PARM_WIP)
                    Case "833"
                        C = 14 + Due + 16
                    Case "842", "843"
                        C = 14 + Due + 21
                    Case "855"
                        C = 14 + Due + 12
                    Case "861"
                        C = 14 + Due + 12
                    Case "863"
                        C = 14 + Due + 11
                    Case Else
                        C = 14 + Due + 10
                End Select
                C = C + 1
                ActiveCell.Offset(R, C).Value = Trim(PROC_BFA(j))
                C = C + 1
                ActiveCell.Offset(R, C).Value = Trim(COMP_BFA(j))
            End If
                        
            If j <= UBound(ARR_PRC) Then
                Select Case Trim(PARM_WIP)
                    Case "833"
                        C = 14 + Due + 18
                    Case "842", "843"
                        C = 14 + Due + 23
                    Case "855"
                        C = 14 + Due + 14
                    Case "861"
                        C = 14 + Due + 14
                    Case "863"
                        C = 14 + Due + 13
                    Case Else
                        C = 14 + Due + 12
                End Select
                Work_Split = Split(ARR_PRC(j), ";")
                For x = LBound(Work_Split) To UBound(Work_Split)
                    Work_Sub_Split = Split(Work_Split(x), ",")
                    For y = LBound(Work_Sub_Split) To UBound(Work_Sub_Split)
                        C = C + 1
                        ActiveCell.Offset(R, C).Value = Trim(Work_Sub_Split(y))
                    Next
                Next
            End If
            
            Rows(R).RowHeight = 13.5
    
            j = j + 1
        Wend
    End If
'END  : Write 'MISC' line ----------------------------------------------------------------------------------------------------------------
    
    Rows(R + 1).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 45
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone

    
End Sub

Sub GET_CATGRP(ByRef ARR_CAT() As String, ByRef GRP As String)

    Dim A As Integer
    Dim Found_Flag As Boolean
    Dim Pos As Long
    
    A = 2
    GRP = "Others"
    Found_Flag = False
    While Not IsEmpty(Worksheets("CategoryGrouping").Cells(A, 1)) And Found_Flag = False
        i = 0
        While i <= UBound(ARR_CAT) And Found_Flag = False
            If Trim(ARR_CAT(i)) = Trim(Worksheets("CategoryGrouping").Cells(A, 1).Value) Then
                GRP = Trim(Worksheets("CategoryGrouping").Cells(A, 2).Value)
                Found_Flag = True
            End If
            i = i + 1
        Wend
        A = A + 1
    Wend
    
End Sub
