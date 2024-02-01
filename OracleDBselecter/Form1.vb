Imports System.Text
Imports Oracle.ManagedDataAccess.Client



Public Class Form1
    Private da As OracleDataAdapter
    Private cb As OracleCommandBuilder
    Private ds As DataSet

    Dim s As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox1.Text = " " Then
            MsgBox("Nem adott meg dátumot a lekérdezés napjára", MsgBoxStyle.Critical, "HIBA")
        Else

            'kapcsolat az adatbázissal
            Dim oradb As String = "Data Source=(DESCRIPTION=" _
 + "(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=x.x.x.x)(PORT=xxxx)))" _
 + "(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=atomd)));" _
 + "User Id=test;Password=test;"

            Dim conn As New OracleConnection(oradb)
            conn.Open()
            Dim cmd As New OracleCommand
            ' cmd.Connection = conn
            'cmd.CommandText = "select * from dontes"
            ' cmd.CommandType = CommandType.Text
            '   Dim dr As OracleDataReader = cmd.ExecuteReader()
            ' dr.Read()
            '  Label1.Text &= dr.Item(2)

            Dim sql As String = " select 
  alk.id alk_id,alk.tszam PERSON_ID,vt.vzs_tipus_nev, don.dontes_nev, alk.d_datum, alk.a_datum, alk.lastmod,korl.mh_kor1, korl.mh_kor2, korl.mh_kor3, korl.mh_kor4, korl.mh_kor5, korl.mh_kor6
 from 
  view2 alk,
 alk_vzs_tipus_trz vt,
 dontes_trz don,

 --------
 (with k1 as
(select 
 k.id, t.mhk_nev, dense_rank() over (partition by k.id order by t.mhk_nev) x
from mh_korlatozas k, 
 mh_korlatozas_trz t
where 
 k.mhk_kod = t.mhk_kod(+)
)
select
* 
from k1
pivot
(max(mhk_nev)
 for x
 in ('1' as mh_kor1,'2' as mh_kor2,'3' as mh_kor3,'4' as mh_kor4,'5' as mh_kor5,'6' as mh_kor6) )
)korl
-----
where 
 --and alk.id > 1028427
to_char(alk.d_datum, 'yyyy-MM-dd') = " & "'" & TextBox1.Text & "'" &
"and alk.pactrz_ceg_id =1061
 and alk.vzs_tipus_kod = vt.vzs_tipus_kod(+)
 and alk.dontes_kod = don.dontes_kod(+)
 and alk.id = korl.id(+)
order by alk.id
"


            cmd = New OracleCommand(sql, conn) 'text kiolvasás az adatbázisból
            cmd.CommandType = CommandType.Text

            da = New OracleDataAdapter(cmd)
            cb = New OracleCommandBuilder(da)
            ds = New DataSet()

            da.Fill(ds)

            DataGridView1.DataSource = ds.Tables(0) 'text bevitel az a gridview-ba csak ott jelenik meg rendes formában ahogy kell neki

            '------------------------------------------------------------------------------------
            ' Itt adom meg azt hogy a dátum formátum az 1 es a 2 es illetve akár X -edik cellában az alábbi módon nézzen ki,
            ' a style.format csak megjelenítési formában írja át a kiíratásban nem általában
            Dim xRowCount As Integer = 0
            Do Until xRowCount = DataGridView1.RowCount
                DataGridView1.Rows(xRowCount).Cells(4).Style.Format = Format("yy-MMM-dd")
                DataGridView1.Rows(xRowCount).Cells(6).Style.Format = Format("yy-MMM-dd")
                DataGridView1.Rows(xRowCount).Cells(5).Style.Format = Format("yy-MMM-dd")
                xRowCount = xRowCount + 1
            Loop
            '------------------------------------------------------------------------------------
            DataGridView1.AllowUserToAddRows = False ' ez adja meg hogy ne jelenjen meg új sor a gridviewban
            DataGridView1.RowHeadersVisible = False
            conn.Dispose()
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Dim StrExport As String = ""


        For Each C As DataGridViewColumn In DataGridView1.Columns
            StrExport &= """" & C.HeaderText & ""","
        Next
        StrExport = StrExport.Substring(0, StrExport.Length - 1)
        StrExport &= Environment.NewLine


        For Each R As DataGridViewRow In DataGridView1.Rows
            For Each C As DataGridViewCell In R.Cells
                If Not C.Value Is Nothing Then
                    StrExport &= """" & C.Value.ToString & ""","
                Else
                    StrExport &= """" & "" & ""","
                End If
            Next
            StrExport = StrExport.Substring(0, StrExport.Length - 1)
            StrExport &= Environment.NewLine
        Next

        Dim tw As IO.TextWriter = New IO.StreamWriter("C:\Users\users\Desktop\Új mappa\orvosi_export.csv")
        tw.Write(StrExport)
        tw.Close()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        'create empty string
        Dim thecsvfile As String = String.Empty
        'get the column headers
        For Each column As DataGridViewColumn In DataGridView1.Columns
            thecsvfile = thecsvfile & column.HeaderText & ","
        Next
        'trim the last comma
        thecsvfile = thecsvfile.TrimEnd(",")
        'Add the line to the output
        thecsvfile = thecsvfile & vbCr & vbLf
        'get the rows
        For Each row As DataGridViewRow In DataGridView1.Rows
            'get the cells
            For Each cell As DataGridViewCell In row.Cells
                thecsvfile = thecsvfile & cell.FormattedValue.replace(",", "") & ","
            Next
            'trim the last comma
            thecsvfile = thecsvfile.TrimEnd(",")
            'Add the line to the output
            thecsvfile = thecsvfile & vbCr & vbLf
        Next
        'write the file
        My.Computer.FileSystem.WriteAllText("C:\Users\users\Desktop\Új mappa\orvosi_export.csv", thecsvfile, False)

    End Sub



    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        FolderBrowserDialog1.ShowDialog()
        ' Kiiratás fájlba, itt azért nem kell átkonvertálni a dátum formátumot mert itt a megjelenés alapján rakja bele a text-et 
        Dim strexp As String = String.Empty
        For Each column As DataGridViewColumn In DataGridView1.Columns
            strexp = strexp & """" & column.HeaderText & ""","
        Next
        strexp = strexp.TrimEnd(",""")
        strexp = strexp & vbCr & vbLf
        For Each row As DataGridViewRow In DataGridView1.Rows
            For Each cell As DataGridViewCell In row.Cells
                strexp = strexp & """" & cell.FormattedValue.replace(""",", """") & ""","
            Next
            strexp = strexp.TrimEnd(",""")
            strexp = strexp & vbCr & vbLf
        Next
        Dim a As String = Label2.Text
        Dim menthely As String = FolderBrowserDialog1.SelectedPath
        My.Computer.FileSystem.WriteAllText(menthely & "\orvosi_export.csv", strexp, False, Encoding.Default)

        'Encoding.ASCII.GetString(Encoding.UTF8.GetBytes(menthely & "\orvosi_export.csv"))

        '   Shell("robocopy " & a & "C:\Users\users\Desktop\csv-exportok" & a & " " & a & "\\nucleodc\ftps-megosztas" & a & " " & a & "teszt.csv" & a)

    End Sub

    '-------------------------A verzió itt benne vannak az " elválasztók
    ' Dim strexp As String = String.Empty
    '   For Each column As DataGridViewColumn In DataGridView1.Columns
    '      strexp = strexp & """" & column.HeaderText & """;"
    '  Next
    '  strexp = strexp.TrimEnd(";""")
    '  strexp = strexp & vbCr & vbLf
    '  For Each row As DataGridViewRow In DataGridView1.Rows
    '      For Each cell As DataGridViewCell In row.Cells
    '         strexp = strexp & """" & cell.FormattedValue.replace(""";", """") & """;"
    '      Next
    '      strexp = strexp.TrimEnd(";""")
    '       strexp = strexp & vbCr & vbLf
    '    Next
    '   My.Computer.FileSystem.WriteAllText("C:\Users\users\Desktop\Új mappa\teszt.csv", strexp, False)




    'Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    '    MsgBox("'" & Now.ToString("yyyy-MM-dd") & "'")
    'End Sub

    'karagtereket meg kell számlálni az első oszlopban, ha 6 számjegyű a törzsszám akkor szúrjon be 10 et ha 5 akkor 100-at 

    Public Function CountCharacter(ByVal value As String) As Integer 'Ez számolja meg a karaktereket 
        Dim cnt As Integer = 0
        For Each c As Char In value
            If c = c Then
                cnt += 1
            End If
        Next
        Return cnt
    End Function
    Dim t As String

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        'cellák kiválasztása és kiíratása a cellánkénti karakterszám
        Dim x As Integer

        For i = 0 To DataGridView1.Rows.Count - 1
            Try
                'Törzsszám megtoldása a pazrt taggal
                x = DataGridView1.Rows(i).Cells(1).Value 'Épp az adott értékkel egyenlő pl i = 12144 és ezt beleteszi az Xbe ez kell a karakterek megszámlálásához
                If DataGridView1.Rows(i).Cells(1).Value Then
                    t = ""
                    If CountCharacter(x) = 6 Then
                        t = "10" + DataGridView1.Rows(i).Cells(1).Value
                        DataGridView1.Rows(i).Cells(1).Value = t
                    End If
                    If CountCharacter(x) = 5 Then
                        t = "100" + DataGridView1.Rows(i).Cells(1).Value
                        DataGridView1.Rows(i).Cells(1).Value = t
                    End If
                    If CountCharacter(x) = 4 Then
                        t = "1000" + DataGridView1.Rows(i).Cells(1).Value
                        DataGridView1.Rows(i).Cells(1).Value = t
                    End If
                    If CountCharacter(x) = 3 Then
                        t = "10000" + DataGridView1.Rows(i).Cells(1).Value
                        DataGridView1.Rows(i).Cells(1).Value = t
                    End If
                    If CountCharacter(x) = 2 Then
                        t = "100000" + DataGridView1.Rows(i).Cells(1).Value
                        DataGridView1.Rows(i).Cells(1).Value = t
                    End If
                End If
                ' Dontes és a vizsgálat típusa kód átalakítás az SAP kódtábla alapjára

                '  If DataGridView1.Rows(i).Cells(3).Value Then 'Vtipus
                ' If DataGridView1.Rows(i).Cells(3).Value = "001" Then
                'DataGridView1.Rows(i).Cells(3).Value = "3"
                'ElseIf DataGridView1.Rows(i).Cells(3).Value = "002" Then
                'DataGridView1.Rows(i).Cells(3).Value = "1"
                'ElseIf DataGridView1.Rows(i).Cells(3).Value = "003" Then
                'DataGridView1.Rows(i).Cells(3).Value = "2"
                'ElseIf DataGridView1.Rows(i).Cells(3).Value = "004" Then
                'DataGridView1.Rows(i).Cells(3).Value = "5"
                'ElseIf DataGridView1.Rows(i).Cells(3).Value = "007" Then
                'DataGridView1.Rows(i).Cells(3).Value = "7"
                'ElseIf DataGridView1.Rows(i).Cells(3).Value = "008" Then
                'DataGridView1.Rows(i).Cells(3).Value = "4"
                'ElseIf DataGridView1.Rows(i).Cells(3).Value = "K00" Then
                'DataGridView1.Rows(i).Cells(3).Value = "6"
                'End If
                'End If


                'If DataGridView1.Rows(i).Cells(4).Value Then 'Dtipus
                'If DataGridView1.Rows(i).Cells(4).Value = "1" Then
                'DataGridView1.Rows(i).Cells(4).Value = "02"
                'ElseIf DataGridView1.Rows(i).Cells(4).Value = "2" Then
                'DataGridView1.Rows(i).Cells(4).Value = "03"
                'ElseIf DataGridView1.Rows(i).Cells(4).Value = "4" Then
                'DataGridView1.Rows(i).Cells(4).Value = "06"
                'ElseIf DataGridView1.Rows(i).Cells(4).Value = "5" Then
                'DataGridView1.Rows(i).Cells(4).Value = "04"
                'ElseIf DataGridView1.Rows(i).Cells(4).Value = "6" Then
                'DataGridView1.Rows(i).Cells(4).Value = "01"
                'ElseIf DataGridView1.Rows(i).Cells(4).Value = "7" Then
                'DataGridView1.Rows(i).Cells(4).Value = "05"
                'End If
                'End If
            Catch ex As Exception
                Continue For
            End Try
        Next
    End Sub







    '------------------------------------------------------------------------------------------------
    'AUTOMATIZÁLÁS RÉSZ 
    Private Sub tmr_lekerdezes_Tick(sender As Object, e As EventArgs) Handles tmr_lekerdezes.Tick


        Dim oradb As String = "Data Source=(DESCRIPTION=" _
 + "(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=x.x.x.x)(PORT=xxxx)))" _
 + "(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=atomd)));" _
 + "User Id=user;Password=test;"
        Dim conn As New OracleConnection(oradb)
        conn.Open()
        Dim cmd As New OracleCommand
        ' cmd.Connection = conn
        'cmd.CommandText = "select * from dontes"
        ' cmd.CommandType = CommandType.Text
        '   Dim dr As OracleDataReader = cmd.ExecuteReader()
        ' dr.Read()
        '  Label1.Text &= dr.Item(2)

        Dim sql As String = " select 
  alk.id alk_id,alk.tszam PERSON_ID,vt.vzs_tipus_nev, don.dontes_nev, alk.d_datum, alk.a_datum, alk.lastmod,korl.mh_kor1, korl.mh_kor2, korl.mh_kor3, korl.mh_kor4, korl.mh_kor5, korl.mh_kor6
 from 
  view2 alk,
 alk_vzs_tipus_trz vt,
 dontes_trz don,

 --------
 (with k1 as
(select 
 k.id, t.mhk_nev, dense_rank() over (partition by k.id order by t.mhk_nev) x
from mh_korlatozas k, 
 mh_korlatozas_trz t
where 
 k.mhk_kod = t.mhk_kod(+)
)
select
* 
from k1
pivot
(max(mhk_nev)
 for x
 in ('1' as mh_kor1,'2' as mh_kor2,'3' as mh_kor3,'4' as mh_kor4,'5' as mh_kor5,'6' as mh_kor6) )
)korl
-----
where 
 --and alk.id > 1028427
to_char(alk.d_datum, 'yyyy-MM-dd') = " & "'" & Now.ToString("yyyy-MM-dd") & "'" &
"and alk.pactrz_ceg_id =1061
 and alk.vzs_tipus_kod = vt.vzs_tipus_kod(+)
 and alk.dontes_kod = don.dontes_kod(+)
 and alk.id = korl.id(+)
order by alk.id

"


        cmd = New OracleCommand(sql, conn) 'text kiolvasás az adatbázisból
        cmd.CommandType = CommandType.Text

        da = New OracleDataAdapter(cmd)
        cb = New OracleCommandBuilder(da)
        ds = New DataSet()

        da.Fill(ds)

        DataGridView1.DataSource = ds.Tables(0) 'text bevitel az a gridview-ba csak ott jelenik meg rendes formában ahogy kell neki

        '------------------------------------------------------------------------------------
        ' Itt adom meg azt hogy a dátum formátum az 1 es a 2 es illetve akár X -edik cellában az alábbi módon nézzen ki,
        ' a style.format csak megjelenítési formában írja át a kiíratásban nem általában
        Dim xRowCount As Integer = 0
        Do Until xRowCount = DataGridView1.RowCount
            DataGridView1.Rows(xRowCount).Cells(4).Style.Format = Format("yy-MMM-dd")
            DataGridView1.Rows(xRowCount).Cells(6).Style.Format = Format("yy-MMM-dd")
            DataGridView1.Rows(xRowCount).Cells(5).Style.Format = Format("yy-MMM-dd")
            xRowCount = xRowCount + 1
        Loop
        '------------------------------------------------------------------------------------
        DataGridView1.AllowUserToAddRows = False ' ez adja meg hogy ne jelenjen meg új sor a gridviewban 
        DataGridView1.RowHeadersVisible = False
        conn.Dispose()
        Label1.Text = "2"
        tmr_atalakitas.Enabled = True
        tmr_lekerdezes.Enabled = False
    End Sub

    Private Sub tmr_atalakitas_Tick(sender As Object, e As EventArgs) Handles tmr_atalakitas.Tick
        Dim x As Integer

        For i = 0 To DataGridView1.Rows.Count - 1
            Try
                'Törzsszám megtoldása a pazrt taggal
                x = DataGridView1.Rows(i).Cells(1).Value 'Épp az adott értékkel egyenlő pl i = 12144 és ezt beleteszi az Xbe ez kell a karakterek megszámlálásához
                If DataGridView1.Rows(i).Cells(1).Value Then
                    t = ""
                    If CountCharacter(x) = 6 Then
                        t = "10" + DataGridView1.Rows(i).Cells(1).Value
                        DataGridView1.Rows(i).Cells(1).Value = t
                    End If
                    If CountCharacter(x) = 5 Then
                        t = "100" + DataGridView1.Rows(i).Cells(1).Value
                        DataGridView1.Rows(i).Cells(1).Value = t
                    End If
                    If CountCharacter(x) = 4 Then
                        t = "1000" + DataGridView1.Rows(i).Cells(1).Value
                        DataGridView1.Rows(i).Cells(1).Value = t
                    End If
                    If CountCharacter(x) = 3 Then
                        t = "10000" + DataGridView1.Rows(i).Cells(1).Value
                        DataGridView1.Rows(i).Cells(1).Value = t
                    End If
                    If CountCharacter(x) = 2 Then
                        t = "100000" + DataGridView1.Rows(i).Cells(1).Value
                        DataGridView1.Rows(i).Cells(1).Value = t
                    End If
                End If
                ' Dontes és a vizsgálat típusa kód átalakítás az SAP kódtábla alapjára

                '  If DataGridView1.Rows(i).Cells(3).Value Then 'Vtipus
                ' If DataGridView1.Rows(i).Cells(3).Value = "001" Then
                'DataGridView1.Rows(i).Cells(3).Value = "3"
                'ElseIf DataGridView1.Rows(i).Cells(3).Value = "002" Then
                'DataGridView1.Rows(i).Cells(3).Value = "1"
                'ElseIf DataGridView1.Rows(i).Cells(3).Value = "003" Then
                'DataGridView1.Rows(i).Cells(3).Value = "2"
                'ElseIf DataGridView1.Rows(i).Cells(3).Value = "004" Then
                'DataGridView1.Rows(i).Cells(3).Value = "5"
                'ElseIf DataGridView1.Rows(i).Cells(3).Value = "007" Then
                'DataGridView1.Rows(i).Cells(3).Value = "7"
                'ElseIf DataGridView1.Rows(i).Cells(3).Value = "008" Then
                'DataGridView1.Rows(i).Cells(3).Value = "4"
                'ElseIf DataGridView1.Rows(i).Cells(3).Value = "K00" Then
                'DataGridView1.Rows(i).Cells(3).Value = "6"
                'End If
                'End If


                'If DataGridView1.Rows(i).Cells(4).Value Then 'Dtipus
                'If DataGridView1.Rows(i).Cells(4).Value = "1" Then
                'DataGridView1.Rows(i).Cells(4).Value = "02"
                'ElseIf DataGridView1.Rows(i).Cells(4).Value = "2" Then
                'DataGridView1.Rows(i).Cells(4).Value = "03"
                'ElseIf DataGridView1.Rows(i).Cells(4).Value = "4" Then
                'DataGridView1.Rows(i).Cells(4).Value = "06"
                'ElseIf DataGridView1.Rows(i).Cells(4).Value = "5" Then
                'DataGridView1.Rows(i).Cells(4).Value = "04"
                'ElseIf DataGridView1.Rows(i).Cells(4).Value = "6" Then
                'DataGridView1.Rows(i).Cells(4).Value = "01"
                'ElseIf DataGridView1.Rows(i).Cells(4).Value = "7" Then
                'DataGridView1.Rows(i).Cells(4).Value = "05"
                'End If
                'End If
            Catch ex As Exception
                Continue For
            End Try
        Next
        Label1.Text = "3"
        tmr_kiir.Enabled = True
        tmr_atalakitas.Enabled = False
    End Sub

    Private Sub tmr_kiir_Tick(sender As Object, e As EventArgs) Handles tmr_kiir.Tick
        ' Kiiratás fájlba, itt azért nem kell átkonvertálni a dátum formátumot mert itt a megjelenés alapján rakja bele a text-et 
        Dim strexp As String = String.Empty
        For Each column As DataGridViewColumn In DataGridView1.Columns
            strexp = strexp & """" & column.HeaderText & ""","
        Next
        strexp = strexp.TrimEnd(",""")
        strexp = strexp & vbCr & vbLf
        For Each row As DataGridViewRow In DataGridView1.Rows
            For Each cell As DataGridViewCell In row.Cells
                strexp = strexp & """" & cell.FormattedValue.replace(""",", """") & ""","
            Next
            strexp = strexp.TrimEnd(",""")
            strexp = strexp & vbCr & vbLf
        Next
        Dim a As String = Label2.Text
        My.Computer.FileSystem.WriteAllText("D:\FTPS-Szerver\orvosi_export.csv", strexp, False, Encoding.Default)
        'Shell("robocopy " & a & "C:\Users\users\Desktop\csv-exportok" & a & " " & a & "\\nucleodc\ftps-megosztas" & a & " " & a & "orvosi_export.csv" & a)  '--> Kikerülésre került, mivel már nem hálózati meghajtóra megy a mentés 
        tmr_datumcheck.Enabled = True
        Label1.Text = "4"

        tmr_kiir.Enabled = False
    End Sub

    Private Sub tmr_datumcheck_Tick(sender As Object, e As EventArgs) Handles tmr_datumcheck.Tick
        If Now.ToString("HH:mm") = "22:00" Then
            tmr_lekerdezes.Enabled = True
            tmr_datumcheck.Enabled = False
            Label1.Text = "1"
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '  --> kikerülésre került, mivel nem hálóüzati meghajtóra kerül a mentés
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.Show()

    End Sub
    Private Sub Form1_resize(sender As Object, e As EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            Me.Hide()

            NotifyIcon1.ShowBalloonTip(0)
        End If

    End Sub
    Private Sub form1kilep(sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        e.Cancel = True
        Me.Hide()
        NotifyIcon1.ShowBalloonTip(0)
    End Sub
    Private Sub KilépésToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KilépésToolStripMenuItem.Click
        Shell("taskkill /f /im oracledbselecter.exe")

    End Sub


End Class
