Imports Microsoft.Office.Tools.Ribbon
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Tools.Excel
Imports System.Data.SQLite
Imports System.Data


Public Class LabItemsRibbon

    Private Sub LabItemsRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
		'call
    End Sub

    Private Sub ButtonSearch_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ButtonSearch.Click
        Try
            Dim rng As Excel.Range = Globals.ThisAddIn.Application.Selection
            Dim xmmc = rng.Text.Trim()
            Dim dialog As New SearchDialog
            dialog.TextBox1.Text = xmmc
            If dialog.ShowDialog = DialogResult.OK Then
                Globals.ThisAddIn.Application.Cells(rng.Row, rng.Column - 5).Value = dialog.xmbh
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    Private Sub ButtonCalcu_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ButtonCalcu.Click, ButtonCalc.Click
        Try
            Dim rng As Range = Globals.ThisAddIn.Application.Selection
            Dim sum As Integer
            sum = 0
            For Each cell In rng.Cells
                If cell.Value.ToString() <> "" Then
                    sum = sum + CInt(cell.Value.ToString())
                    cell.value = ""
                End If
            Next
            rng.Cells(1, 1).Value = sum
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    Private Sub ButtonStatistics_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ButtonStatistics.Click
        Try
            Dim rng As Range = Globals.ThisAddIn.Application.Selection
            Dim cz, c1, c2, c3, c4 As Integer
            cz = c1 = c2 = c3 = c4 = 0
            For Each cell In rng.Cells
                Select Case cell.Value.ToString()
                    Case "验证"
                        c1 = c1 + 1
                    Case "综合"
                        c2 = c2 + 1
                    Case "设计研究"
                        c3 = c3 + 1
                    Case "演示"
                        c4 = c4 + 1
                End Select
            Next
            cz = c1 + c2 + c3 + c4
            Clipboard.SetText(cz & vbTab & c1 & vbTab & c2 & vbTab & c3)
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    Private Sub ButtonPaste_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ButtonPaste.Click
        Try
            Dim rng As Range = Globals.ThisAddIn.Application.Selection
            Dim text = Clipboard.GetText(TextDataFormat.Text)
            Dim lines = CountLines(text)
            Dim rc = rng.Row
            Dim cc = rng.Column
            Clipboard.Clear()
            Globals.ThisAddIn.Application.CutCopyMode = False
            For i = 1 To lines - 1
                Globals.ThisAddIn.Application.ActiveSheet.Rows(rc + 1).Insert()
            Next
            Globals.ThisAddIn.Application.CutCopyMode = True
            Clipboard.SetText(text)
            ''''rng.Cells(1, 1).Insert(XlInsertShiftDirection.xlShiftDown)
            Globals.ThisAddIn.Application.ActiveSheet.Cells(rc, cc).PasteSpecial()
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    Private Function CountLines(ByVal text As String)
        Dim count = 0
        count = text.Split(vbCrLf).Count - 1
        Return count
    End Function

    Private Sub ButtonSplitSum_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ButtonSplitSum.Click
        Try
            Dim rng As Range = Globals.ThisAddIn.Application.Selection
            Dim sum As Integer

            For Each cell In rng.Cells
                If cell.Value.ToString() <> "" Then
                    sum = 0
                    Dim arr = cell.Value.ToString().Split(",")
                    For i = 0 To UBound(arr)
                        sum = sum + CInt(arr(i).Trim())
                    Next
                    cell.value = sum
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    Private Sub ButtonDomain_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ButtonDomain.Click
        Try
            Dim c As Integer
            c = CInt(InputBox("查出的数据填入第几列", "查专业", "2"))
            Dim rng As Range = Globals.ThisAddIn.Application.Selection
            For Each cell In rng.Cells
                If cell.Value.ToString() <> "" Then
                    Dim bh = cell.Value.ToString().Split(",")(0).Trim()
                    Dim conn As New SQLiteConnection("Data Source=c:\items.sqlite;Pooling=true;FailIfMissing=false")
                    Dim cmd As New SQLiteCommand("select zymc from zy where zydm='" & bh.Substring(0, 4) & "' or zydm='" & bh.Substring(0, 5) & "'", conn)
                    Dim da As New SQLiteDataAdapter(cmd)
                    Dim dt As New System.Data.DataTable()
                    da.Fill(dt)
                    If dt.Rows.Count <> 0 Then
                        Dim zymc = dt.Rows(0).Item(0).ToString()
                        Dim cn = cell.Row
                        Globals.ThisAddIn.Application.ActiveSheet.Cells(cn, c).Value = zymc
                    End If
                    conn.Close()
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    Private Sub ButtonPinyin_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles ButtonPinyin.Click
        Try
            Dim rng As Range = Globals.ThisAddIn.Application.Selection
            Dim s As String
            Dim pytype As Integer
            Dim dialog As New PinyinDialog
            pytype = 1
            If dialog.ShowDialog() = DialogResult.OK Then
                pytype = dialog.pytype
            End If

            s = ""
            For Each cell In rng.Cells
                If cell.Value.ToString() <> "" Then
                    Dim py = PinYin(cell.Value.ToString())
                    Dim arr = py.Split(" ")
                    Dim rpy = ""
                    Select Case pytype
                        Case 1
                            rpy = py
                        Case 2
                            rpy = arr(0) & " "
                            For i = 1 To UBound(arr)
                                rpy = rpy & arr(i)
                            Next
                        Case 3
                            For i = 0 To UBound(arr)
                                rpy = rpy & arr(i)(0)
                            Next
                        Case 4
                            rpy = arr(0)
                            For i = 1 To UBound(arr)
                                rpy = rpy & arr(i)(0)
                            Next
                        Case 5
                            For i = 1 To UBound(arr)
                                rpy = rpy & arr(i)(0)
                            Next
                            rpy = rpy & arr(0)
                    End Select

                    s = s & rpy & vbCrLf
                End If
            Next
            Clipboard.SetText(s)
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    Private Function PinYin(ByVal Hz As String)
        Dim PinMa As String
        Dim MyPinMa As Object
        Dim Temp As Integer, i As Integer, j As Integer
        PinMa = "a,20319,ai,20317,an,20304,ang,20295,ao,20292,"
        PinMa = PinMa & "ba,20283,bai,20265,ban,20257,bang,20242,bao,20230,bei,20051,ben,20036,beng,20032,bi,20026,bian,20002,biao,19990,bie,19986,bin,19982,bing,19976,bo,19805,bu,19784,"
        PinMa = PinMa & "ca,19775,cai,19774,can,19763,cang,19756,cao,19751,ce,19746,ceng,19741,cha,19739,chai,19728,chan,19725,chang,19715,chao,19540,che,19531,chen,19525,cheng,19515,chi,19500,chong,19484,chou,19479,chu,19467,chuai,19289,chuan,19288,chuang,19281,chui,19275,chun,19270,chuo,19263,ci,19261,cong,19249,cou,19243,cu,19242,cuan,19238,cui,19235,cun,19227,cuo,19224,"
        PinMa = PinMa & "da,19218,dai,19212,dan,19038,dang,19023,dao,19018,de,19006,deng,19003,di,18996,dian,18977,diao,18961,die,18952,ding,18783,diu,18774,dong,18773,dou,18763,du,18756,duan,18741,dui,18735,dun,18731,duo,18722,"
        PinMa = PinMa & "e,18710,en,18697,er,18696,"
        PinMa = PinMa & "fa,18526,fan,18518,fang,18501,fei,18490,fen,18478,feng,18463,fo,18448,fou,18447,fu,18446,"
        PinMa = PinMa & "ga,18239,gai,18237,gan,18231,gang,18220,gao,18211,ge,18201,gei,18184,gen,18183,geng,18181,gong,18012,gou,17997,gu,17988,gua,17970,guai,17964,guan,17961,guang,17950,gui,17947,gun,17931,guo,17928,"
        PinMa = PinMa & "ha,17922,hai,17759,han,17752,hang,17733,hao,17730,he,17721,hei,17703,hen,17701,heng,17697,hong,17692,hou,17683,hu,17676,hua,17496,huai,17487,huan,17482,huang,17468,hui,17454,hun,17433,huo,17427,"
        PinMa = PinMa & "ji,17417,jia,17202,jian,17185,jiang,16983,jiao,16970,jie,16942,jin,16915,jing,16733,jiong,16708,jiu,16706,ju,16689,juan,16664,jue,16657,jun,16647,"
        PinMa = PinMa & "ka,16474,kai,16470,kan,16465,kang,16459,kao,16452,ke,16448,ken,16433,keng,16429,kong,16427,kou,16423,ku,16419,kua,16412,kuai,16407,kuan,16403,kuang,16401,kui,16393,kun,16220,kuo,16216,"
        PinMa = PinMa & "la,16212,lai,16205,lan,16202,lang,16187,lao,16180,le,16171,lei,16169,leng,16158,li,16155,lia,15959,lian,15958,liang,15944,liao,15933,lie,15920,lin,15915,ling,15903,liu,15889,long,15878,lou,15707,lu,15701,lv,15681,luan,15667,lue,15661,lun,15659,luo,15652,"
        PinMa = PinMa & "ma,15640,mai,15631,man,15625,mang,15454,mao,15448,me,15436,mei,15435,men,15419,meng,15416,mi,15408,mian,15394,miao,15385,mie,15377,min,15375,ming,15369,miu,15363,mo,15362,mou,15183,mu,15180,"
        PinMa = PinMa & "na,15165,nai,15158,nan,15153,nang,15150,nao,15149,ne,15144,nei,15143,nen,15141,neng,15140,ni,15139,nian,15128,niang,15121,niao,15119,nie,15117,nin,15110,ning,15109,niu,14941,nong,14937,nu,14933,nv,14930,nuan,14929,nue,14928,nuo,14926,"
        PinMa = PinMa & "o,14922,ou,14921,"
        PinMa = PinMa & "pa,14914,pai,14908,pan,14902,pang,14894,pao,14889,pei,14882,pen,14873,peng,14871,pi,14857,pian,14678,piao,14674,pie,14670,pin,14668,ping,14663,po,14654,pu,14645,"
        PinMa = PinMa & "qi,14630,qia,14594,qian,14429,qiang,14407,qiao,14399,qie,14384,qin,14379,qing,14368,qiong,14355,qiu,14353,qu,14345,quan,14170,que,14159,qun,14151,"
        PinMa = PinMa & "ran,14149,rang,14145,rao,14140,re,14137,ren,14135,reng,14125,ri,14123,rong,14122,rou,14112,ru,14109,ruan,14099,rui,14097,run,14094,ruo,14092,"
        PinMa = PinMa & "sa,14090,sai,14087,san,14083,sang,13917,sao,13914,se,13910,sen,13907,seng,13906,sha,13905,shai,13896,shan,13894,shang,13878,shao,13870,she,13859,shen,13847,sheng,13831,shi,13658,shou,13611,shu,13601,shua,13406,shuai,13404,shuan,13400,shuang,13398,shui,13395,shun,13391,shuo,13387,si,13383,song,13367,sou,13359,su,13356,suan,13343,sui,13340,sun,13329,suo,13326,"
        PinMa = PinMa & "ta,13318,tai,13147,tan,13138,tang,13120,tao,13107,te,13096,teng,13095,ti,13091,tian,13076,tiao,13068,tie,13063,ting,13060,tong,12888,tou,12875,tu,12871,tuan,12860,tui,12858,tun,12852,tuo,12849,"
        PinMa = PinMa & "wa,12838,wai,12831,wan,12829,wang,12812,wei,12802,wen,12607,weng,12597,wo,12594,wu,12585,"
        PinMa = PinMa & "xi,12556,xia,12359,xian,12346,xiang,12320,xiao,12300,xie,12120,xin,12099,xing,12089,xiong,12074,xiu,12067,xu,12058,xuan,12039,xue,11867,xun,11861,"
        PinMa = PinMa & "ya,11847,yan,11831,yang,11798,yao,11781,ye,11604,yi,11589,yin,11536,ying,11358,yo,11340,yong,11339,you,11324,yu,11303,yuan,11097,yue,11077,yun,11067,"
        PinMa = PinMa & "za,11055,zai,11052,zan,11045,zang,11041,zao,11038,ze,11024,zei,11020,zen,11019,zeng,11018,zha,11014,zhai,10838,zhan,10832,zhang,10815,zhao,10800,zhe,10790,zhen,10780,zheng,10764,zhi,10587,zhong,10544,zhou,10533,zhu,10519,zhua,10331,zhuai,10329,zhuan,10328,zhuang,10322,zhui,10315,zhun,10309,zhuo,10307,zi,10296,zong,10281,zou,10274,zu,10270,zuan,10262,zui,10260,zun,10256,zuo,10254"
        MyPinMa = Split(PinMa, ",")
        For i = 1 To Len(Hz)
            Temp = Asc(Mid(Hz, i, 1))
            If Temp < 0 Then
                Temp = Math.Abs(Temp)
                For j = 791 To 1 Step -2
                    If Temp <= Val(MyPinMa(j)) Then
                        PinYin = PinYin & MyPinMa(j - 1) & " "
                        Exit For
                    End If
                Next
            End If
        Next
        PinYin = Trim(PinYin)
    End Function
End Class
