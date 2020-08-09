using Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;

namespace 엑셀_관리_프로그램
{
    public partial class Main_window : Form
    {
        string name = null;

        public Main_window()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            string sDirPath;
            sDirPath = System.Windows.Forms.Application.StartupPath + "\\엑셀_관리_프로그램";
            DirectoryInfo di = new DirectoryInfo(sDirPath);
            if (di.Exists == false)
            {
                di.Create();
            }
            string strFile = System.Windows.Forms.Application.StartupPath + "\\엑셀_관리_프로그램\\엑셀_관리_프로그램_회원_개인.xlsx";
            FileInfo fileInfo = new FileInfo(strFile);
            string strFile2 = System.Windows.Forms.Application.StartupPath + "\\엑셀_관리_프로그램\\엑셀_관리_프로그램_회원_단체.xlsx";
            FileInfo fileInfo2 = new FileInfo(strFile2);

            //파일 있는지 확인 있을때(true), 없으면(false)
            if (fileInfo.Exists && fileInfo2.Exists)
            {
                Exl_load();
                Exl_load2();
            }
            else
            {
                not_exl_file();
                not_exl_file2();
            }
        }
        public void not_exl_file()
        {
            Excel.Application new_app = new Excel.Application();
            Excel.Worksheet worksheet = null;
            Workbook workbook = new_app.Workbooks.Add();
            DataTable dt = new DataTable();
            worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);

            dt.Columns.Add("회원번호"); // 내부 회원번호
            dt.Columns.Add("이름");
            dt.Columns.Add("성별");
            dt.Columns.Add("나이");
            dt.Columns.Add("전화번호");
            dt.Columns.Add("거주지");
            dt.Columns.Add("이메일");
            dt.Columns.Add("관심분야");
            dt.Columns.Add("방문경로");
            dt.Columns.Add("처음 방문");
            dt.Columns.Add("마지막 방문");
            dt.Columns.Add("총 방문 수");

            worksheet.Columns.AutoFit(); //자동 넒이
            new_app.DisplayAlerts = false;
            workbook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\엑셀_관리_프로그램\\엑셀_관리_프로그램_회원_개인.xlsx", XlFileFormat.xlWorkbookDefault);
            dataGridView1.DataSource = dt;
            new_app.Quit();

            releaseObject(new_app);
            releaseObject(worksheet);
            releaseObject(workbook);
        }
        public void not_exl_file2()
        {
            Excel.Application new_app = new Excel.Application();
            Excel.Worksheet worksheet = null;
            Workbook workbook = new_app.Workbooks.Add();
            DataTable dt = new DataTable();
            worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);

            dt.Columns.Add("단체번호"); // 내부 회원번호
            dt.Columns.Add("단체명");
            dt.Columns.Add("대표자");
            dt.Columns.Add("담당자");
            dt.Columns.Add("연락처");
            dt.Columns.Add("핸드폰");
            dt.Columns.Add("주소");
            dt.Columns.Add("이메일");
            dt.Columns.Add("관심분야");
            dt.Columns.Add("방문경로");
            dt.Columns.Add("처음 방문");
            dt.Columns.Add("마지막 방문");
            dt.Columns.Add("총 방문 수");

            worksheet.Columns.AutoFit(); //자동 넒이
            new_app.DisplayAlerts = false;
            workbook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\엑셀_관리_프로그램\\엑셀_관리_프로그램_회원_단체.xlsx", XlFileFormat.xlWorkbookDefault);
            dataGridView2.DataSource = dt;
            new_app.Quit();

            releaseObject(new_app);
            releaseObject(worksheet);
            releaseObject(workbook);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        { }


        public string user_info()
        {
            string str;
            str = name;
            return str;
        }

        // 함수명 releaseObject 입력인자 : 해제 원하는 Obj
        public void releaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
        private void button4_Click(object sender, EventArgs e) //save
        {
            Exl_Save();
            Exl_Save2();
        }
        public int loading_percent()
        {
            string timerun = null;
            int a = dataGridView1.RowCount, b = dataGridView2.RowCount;
            int sum = (a + b) / 10;

            return sum;
        }
        public void Exl_Save()
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            Excel.Worksheet worksheet = (Excel.Worksheet)app.Worksheets.get_Item(1);

            for (int r = 0; r < dataGridView1.RowCount; r++)
            {
                label2.Text = r.ToString();
                for (int c = 0; c < 12; c++)
                {
                    worksheet.Cells[r + 2, c + 1] = dataGridView1.Rows[r].Cells[c].Value;

                }
            }
            worksheet.Columns.AutoFit(); //자동 넒이
            app.DisplayAlerts = false;
            workbook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\엑셀_관리_프로그램\\엑셀_관리_프로그램_회원_개인.xlsx", XlFileFormat.xlWorkbookDefault);
            workbook.Close(System.Windows.Forms.Application.StartupPath + "\\엑셀_관리_프로그램" + Type.Missing);


            app.Quit();
            releaseObject(app);
            releaseObject(worksheet);
            releaseObject(workbook);
        }
        public void Exl_Save2()
        {
            Excel.Application app2 = new Excel.Application();
            Excel.Workbook workbook = app2.Workbooks.Add(Type.Missing);
            Excel.Worksheet worksheet = (Excel.Worksheet)app2.Worksheets.get_Item(1);


            for (int r = 0; r < dataGridView2.RowCount; r++)
            {
                label3.Text = r.ToString();
                for (int c = 0; c < 13; c++)
                {
                    worksheet.Cells[r + 2, c + 1] = dataGridView2.Rows[r].Cells[c].Value;

                }
            }
            worksheet.Columns.AutoFit(); //자동 넒이
            app2.DisplayAlerts = false;
            workbook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\엑셀_관리_프로그램\\엑셀_관리_프로그램_회원_단체.xlsx", XlFileFormat.xlWorkbookDefault);
            workbook.Close(System.Windows.Forms.Application.StartupPath + "\\엑셀_관리_프로그램" + Type.Missing);

            app2.Quit();
            releaseObject(app2);
            releaseObject(worksheet);
            releaseObject(workbook);
        }

        public void Exl_load()
        {

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            DataTable table = new DataTable();
            Excel.Application app = new Excel.Application();
            Excel.Worksheet worksheet = null;
            Excel.Range range;
            dataGridView1.Columns.Clear();
            Workbook workbook = app.Workbooks.Open(System.Windows.Forms.Application.StartupPath + "\\엑셀_관리_프로그램\\엑셀_관리_프로그램_회원_개인.xlsx", 0, false, 5, Missing.Value,
            Missing.Value, false, Missing.Value, Missing.Value, true, false, Missing.Value, false, false, false);

            worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1); // 시트 오픈

            range = worksheet.UsedRange;
            object[,] data = range.Value;


            table.Columns.Add("회원번호"); // 내부 회원번호
            table.Columns.Add("이름");
            table.Columns.Add("성별");
            table.Columns.Add("나이");
            table.Columns.Add("전화번호");
            table.Columns.Add("거주지");
            table.Columns.Add("이메일");
            table.Columns.Add("관심분야");
            table.Columns.Add("방문경로");
            table.Columns.Add("처음 방문");
            table.Columns.Add("마지막 방문");
            table.Columns.Add("총 방문 수");

            if (range.Rows.Count == null)
            {
                dataGridView1.DataSource = table;

                return;
            }
            else
            {
                for (int r = 0; r < range.Rows.Count; r++)
                {
                    DataRow dr = table.Rows.Add();

                    for (int c = 1; c < 13; c++)
                    {
                        dr[c - 1] = data[r + 1, c];
                    }
                }
            }

            workbook.Close(true);
            app.Quit();

            // 데이터그리드뷰에 데이터테이블 바인딩
            dataGridView1.DataSource = table;


            releaseObject(app);
            releaseObject(worksheet);
            releaseObject(workbook);
            releaseObject(range);
        }
        public void Exl_load2()
        {
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            DataTable table2 = new DataTable();
            Excel.Application app2 = new Excel.Application();
            Excel.Worksheet worksheet = null;
            Excel.Range range;
            dataGridView2.Columns.Clear();
            Workbook workbook = app2.Workbooks.Open(System.Windows.Forms.Application.StartupPath + "\\엑셀_관리_프로그램_엑셀\\엑셀_관리_프로그램_회원_단체.xlsx", 0, false, 5, Missing.Value,
            Missing.Value, false, Missing.Value, Missing.Value, true, false, Missing.Value, false, false, false);

            worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1); // 시트 오픈

            range = worksheet.UsedRange;
            object[,] data = range.Value;


            table2.Columns.Add("단체번호"); // 내부 회원번호
            table2.Columns.Add("단체명");
            table2.Columns.Add("대표자");
            table2.Columns.Add("담당자");
            table2.Columns.Add("연락처");
            table2.Columns.Add("핸드폰");
            table2.Columns.Add("주소");
            table2.Columns.Add("이메일");
            table2.Columns.Add("관심분야");
            table2.Columns.Add("방문경로");
            table2.Columns.Add("처음 방문");
            table2.Columns.Add("마지막 방문");
            table2.Columns.Add("총 방문 수");


            for (int r = 0; r < range.Rows.Count; r++)
            {
                DataRow dr = table2.Rows.Add();

                for (int c = 1; c < 14; c++)
                {
                    dr[c - 1] = data[r + 1, c];
                }
            }

            workbook.Close(true);
            app2.Quit();

            // 데이터그리드뷰에 데이터테이블 바인딩
            dataGridView2.DataSource = table2;


            releaseObject(app2);
            releaseObject(worksheet);
            releaseObject(workbook);
            releaseObject(range);
        }

        // 검색 변수 
        int iRowIdx = -1;

        private void button5_Click(object sender, EventArgs e)
        {
            int rows = dataGridView1.Rows.Count; // 전체row수
            int rows2 = dataGridView2.Rows.Count; // 전체row수
            string number = textBox1.Text; // 입력한 회원 이름

            // 개인 검색
            if (tabControl1.SelectedTab == tabPage1)
            {
                if (number == "")
                {
                    MessageBox.Show("검색할 대상을 입력해주세요.");
                    return;
                }

                for (int i = 0; i < rows - 1; i++)
                {
                    string value = dataGridView1.Rows[i].Cells[1].Value.ToString(); // 매 row의 특정열 값 가져오기

                    if (value == number) // 비교
                    {
                        iRowIdx = i;
                        if (iRowIdx >= 0)
                        {
                            dataGridView1.FirstDisplayedCell = dataGridView1.Rows[i].Cells[1];
                            dataGridView1.Rows[i].Cells[1].Selected = true;
                            return;
                        }
                    }
                }
                tabControl1.SelectedTab = tabPage2;
            }
            //단체 검색
            else if (tabControl1.SelectedTab == tabPage2)
            {
                for (int i = 0; i < rows2 - 1; i++)
                {
                    string value = dataGridView2.Rows[i].Cells[1].Value.ToString(); // 매 row의 특정열 값 가져오기

                    if (value == number) // 비교
                    {
                        iRowIdx = i;
                        if (iRowIdx >= 0)
                        {
                            dataGridView2.FirstDisplayedCell = dataGridView2.Rows[i].Cells[1];
                            dataGridView2.Rows[i].Cells[1].Selected = true;
                            return;
                        }
                    }
                }
                MessageBox.Show("검색할 대상이 없습니다.");
                tabControl1.SelectedTab = tabPage1;
                return;
            }
        }

        string pupu = null;
        private void dataGridView1_DoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            name = dataGridView1.Rows[e.RowIndex].Cells[1].FormattedValue.ToString();
            pupu = e.RowIndex.ToString();
            user_info();
            User_info u1 = new User_info(this);
            u1.Show();
        }

        public int haha()
        {
            return Convert.ToInt32(pupu);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Revisit r1 = new Revisit(this);
            r1.Show();
        }

        private void dataGridView2_DoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            name = dataGridView2.Rows[e.RowIndex].Cells[1].FormattedValue.ToString();
            pupu = e.RowIndex.ToString();
            user_info();
            group_info g1 = new group_info(this);
            g1.Show();
        }
        private void FormClosing(object sender, FormClosingEventArgs e)
        { }

        private void dataGridView_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
        {
        }

        private void ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

        }
        // 신규등록 창
        private void button2_Click(object sender, EventArgs e)
        {
            New_User f2 = new New_User(this);
            f2.Show();
        }
        // 단체등록 창
        private void button1_Click(object sender, EventArgs e)
        {
            New_group f3 = new New_group(this);
            f3.Show();
        }

        private void 저장ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Exl_Save();
            Exl_Save2();
        }

        private void 신규가입ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            New_User f2 = new New_User(this);
            f2.Show();
        }

        private void 단체회원ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            New_group f3 = new New_group(this);
            f3.Show();
        }

        private void 닫기ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Exl_Save();
            Exl_Save2();

            System.Windows.Forms.Application.OpenForms["Main_window"].Close();
        }

        private void 정보ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 정보ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            program_info f4 = new program_info(this);
            f4.Show();
        }
        private void button3_Click(object sender, EventArgs e)
        {
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button5_Click(sender, e);
            }
        }
    }
}