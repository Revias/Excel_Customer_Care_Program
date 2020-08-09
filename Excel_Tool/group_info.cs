using Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;

namespace 엑셀_관리_프로그램
{
    public partial class group_info : Form
    {
        DataTable table = new DataTable();
        Excel.Application app = new Excel.Application();
        Group group = new Group();

        Main_window f1 = new Main_window();

        public group_info(Main_window _form)
        {
            InitializeComponent();
            f1 = _form;
        }

        private void group_info_Load(object sender, EventArgs e)
        {
            string strFile = System.Windows.Forms.Application.StartupPath + "\\엑셀_관리_프로그램\\엑셀_관리_프로그램_회원_단체_방문내역.xlsx";
            FileInfo fileInfo = new FileInfo(strFile);
            //파일 있는지 확인 있을때(true), 없으면(false)
            if (fileInfo.Exists)
            {
                Start();
            }
            else
            {
                not_exl_file();

            }
            gidgvuser_info();
        }

        public void not_exl_file()
        {
            Excel.Application new_app = new Excel.Application();
            Excel.Worksheet worksheet = null;
            Excel.Range range = null;
            Workbook workbook = new_app.Workbooks.Add();
            worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);

            for (int i = 1; i < 13; i++)
            {

                Excel.Range cel0 = worksheet.Cells[1, i + 1];
                cel0.Value = i + "월";
            }
            Excel.Range cel1 = worksheet.Cells[2, 1];
            cel1.Value = "현장방문";
            Excel.Range cel2 = worksheet.Cells[3, 1];
            cel2.Value = "대관신청";
            Excel.Range cel3 = worksheet.Cells[4, 1];
            cel3.Value = "교육상담";
            Excel.Range cel4 = worksheet.Cells[5, 1];
            cel4.Value = "노동상담";
            Excel.Range cel5 = worksheet.Cells[6, 1];
            cel5.Value = "영화";
            Excel.Range cel6 = worksheet.Cells[7, 1];
            cel6.Value = "기타";

            workbook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\엑셀_관리_프로그램\\엑셀_관리_프로그램_단체_방문내역.xlsx");



            workbook.Close(true);
            new_app.Quit();

            f1.releaseObject(new_app);
            f1.releaseObject(worksheet);
            f1.releaseObject(workbook);
            Start();
        }

        public void Start()
        {
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            DataTable table = new DataTable();
            Excel.Application app = new Excel.Application();
            Excel.Worksheet worksheet = null;
            Excel.Range range;
            dataGridView1.Columns.Clear();
            Workbook workbook = app.Workbooks.Open(System.Windows.Forms.Application.StartupPath + "\\엑셀_관리_프로그램\\엑셀_관리_프로그램_회원_단체_방문내역.xlsx", 0, false, 5, Missing.Value,
            Missing.Value, false, Missing.Value, Missing.Value, true, false, Missing.Value, false, false, false);

            worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1); // 시트 오픈

            range = worksheet.UsedRange;

            label1.Text = f1.user_info();

            table.Columns.Add("0");
            table.Columns.Add("1");
            table.Columns.Add("2");
            table.Columns.Add("3");
            table.Columns.Add("4");
            table.Columns.Add("5");
            table.Columns.Add("6");
            table.Columns.Add("7");
            table.Columns.Add("8");
            table.Columns.Add("9");
            table.Columns.Add("10");
            table.Columns.Add("11");
            table.Columns.Add("12");

            object[,] data = range.Value;

            for (int r = 0; r < range.Rows.Count; r++)
            {
                DataRow dr = table.Rows.Add();

                for (int c = 1; c < 14; c++)
                {
                    dr[c - 1] = data[r + 1, c];
                }
            }
            workbook.Close(true);
            app.Quit();

            dataGridView1.DataSource = table;

            f1.releaseObject(app);
            f1.releaseObject(worksheet);
            f1.releaseObject(workbook);
            f1.releaseObject(range);
        }

        public void gidgvuser_info()
        {
            int value = f1.haha();
            label1.Text = f1.dataGridView2.Rows[value].Cells[1].FormattedValue.ToString();
            label4.Text = f1.dataGridView2.Rows[value].Cells[0].FormattedValue.ToString();
            label6.Text = f1.dataGridView2.Rows[value].Cells[2].FormattedValue.ToString();
            label8.Text = f1.dataGridView2.Rows[value].Cells[3].FormattedValue.ToString();
            label19.Text = f1.dataGridView2.Rows[value].Cells[4].FormattedValue.ToString();
            label10.Text = f1.dataGridView2.Rows[value].Cells[5].FormattedValue.ToString();
            label12.Text = f1.dataGridView2.Rows[value].Cells[7].FormattedValue.ToString();
            label21.Text = f1.dataGridView2.Rows[value].Cells[6].FormattedValue.ToString(); // 주소
            label15.Text = f1.dataGridView2.Rows[value].Cells[10].FormattedValue.ToString();
            label16.Text = f1.dataGridView2.Rows[value].Cells[11].FormattedValue.ToString();// 처음
            label18.Text = f1.dataGridView2.Rows[value].Cells[12].FormattedValue.ToString();
        }
    }
}