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
    public partial class User_info : Form
    {
        DataTable table = new DataTable();
        New_User user = new New_User();
        Main_window f1 = new Main_window();
        ExcelModule news = new ExcelModule();

        public User_info(Main_window _form)
        {
            InitializeComponent();
            f1 = _form;
        }

        private void label1_Click(object sender, EventArgs e)
        {
            
        }

        private void User_Load_1(object sender, EventArgs e)
        {
            string strFile = System.Windows.Forms.Application.StartupPath + "\\엑셀_관리_프로그램_엑셀\\엑셀_관리_프로그램_회원_개인_방문내역.xlsx";
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
            dgvuser_info();
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

            workbook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\엑셀_관리_프로그램_엑셀\\엑셀_관리_프로그램_회원_개인_방문내역.xlsx");


            workbook.Close(true);
            new_app.Quit();

            f1.releaseObject(range);
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
            Workbook workbook = app.Workbooks.Open(System.Windows.Forms.Application.StartupPath + "\\엑셀_관리_프로그램_엑셀\\엑셀_관리_프로그램_회원_개인_방문내역.xlsx", 0, false, 5, Missing.Value,
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

            int rows = range.Rows.Count;

            string number = (f1.haha() + 1).ToString();
            string reason = user.str_value();
            string cellmm = DateTime.Now.ToString("MM");
            string cellvalur = cellmm.Substring(1);
            int cell = Convert.ToInt32(cellvalur);

            /*
             *  신규가입시 새로 생성된 회원번호로 방문내역 생성 하고   
             */

            object[,] data = range.Value;
            MessageBox.Show("3");
            for (int i = 1; i < rows; i++)
            {
                string value = dataGridView1.Rows[i].Cells[1].Value.ToString();
                string value2 = dataGridView1.Rows[i].Cells[19].Value.ToString();
                MessageBox.Show("0");
                if (value == number) // 비교
                {
                    table.Rows.Add(dataGridView1.Rows[i].Cells[1]);

                    int row1 = 7;
                    for (int r = 0; r < row1; r++)
                    {
                        DataRow dr = table.Rows.Add();
                        for (int c = 1; c < 13; c++)
                        {
                            dr[c - 1] = data[r + 1, c];
                        }
                    }
                    MessageBox.Show("1");
                }
                else if (value2 == number)
                {
                    //dataGridView1.DataSource = dataGridView1.Rows[i].Cells[19];
                    MessageBox.Show("2");
                }
                else
                {
                    MessageBox.Show("정상적으로 등록이 되지 않은 회원입니다.");
                    return;
                }
                
            }
            dataGridView1.DataSource = table;

            user_info_save();
            app.Quit();
            f1.releaseObject(app);
            f1.releaseObject(worksheet);
            f1.releaseObject(workbook);
            f1.releaseObject(range);
            
        }

        public void dgvuser_info()
        {
            int value = f1.haha();
            label4.Text = f1.dataGridView1.Rows[value].Cells[0].FormattedValue.ToString();
            label6.Text = f1.dataGridView1.Rows[value].Cells[3].FormattedValue.ToString();
            label8.Text = f1.dataGridView1.Rows[value].Cells[4].FormattedValue.ToString();
            label10.Text = f1.dataGridView1.Rows[value].Cells[5].FormattedValue.ToString();
            label12.Text = f1.dataGridView1.Rows[value].Cells[6].FormattedValue.ToString();
            label15.Text = f1.dataGridView1.Rows[value].Cells[9].FormattedValue.ToString();
            label16.Text = f1.dataGridView1.Rows[value].Cells[10].FormattedValue.ToString();
            label18.Text = f1.dataGridView1.Rows[value].Cells[11].FormattedValue.ToString();

        }

        public void User_nbr()
        {
            Excel.Application new_app = new Excel.Application();
            Excel.Worksheet worksheet = null;
            Workbook workbook = new_app.Workbooks.Add();
            DataTable dt = new DataTable();
            worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);


            string number = (f1.haha()+1).ToString();
            string reason = user.str_value();
            string cellmm = DateTime.Now.ToString("MM");
            string cellvalur = cellmm.Substring(1);
            int cell = Convert.ToInt32(cellvalur);
            MessageBox.Show("10, "  + cell + cellmm);

            for (int i = 0; i < 74; i++)
            {
                MessageBox.Show("0");
                string value = dataGridView1.Rows[i].Cells[0].FormattedValue.ToString();
                if(value == reason)
                {
                    MessageBox.Show("1");
                    Excel.Range cel1 = worksheet.Cells[i, cell+1];

                    if (cel1 == null)
                    {
                        MessageBox.Show("2");
                        cel1.Value = 1;
                        return;
                    }
                    else
                    {
                        MessageBox.Show("3");
                        cel1.Value =+ 1;
                        return;
                    }

                }
            }
            MessageBox.Show("4");
            for (int i = 0; i < 74; i++)
            {
                string value = dataGridView1.Rows[i].Cells[0].FormattedValue.ToString();

                if (value == number) // 비교
                {
                    string vl = f1.dataGridView1.Rows[f1.haha()].Cells[10].FormattedValue.ToString();
                    string sp = vl.Substring(5,2);
                    string mm = DateTime.Now.ToString("MM");
                    

                    if (sp == mm)
                    {
                        MessageBox.Show(sp + mm + "성공");
                    }
                    else
                    {
                        MessageBox.Show(sp + mm + "실패");
                    }
                }
            }

            workbook.Close(true);
            new_app.Quit();

            f1.releaseObject(new_app);
            f1.releaseObject(worksheet);
            f1.releaseObject(workbook);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        public void user_info_save()
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            Excel.Worksheet worksheet = (Excel.Worksheet)app.Worksheets.get_Item(1);

            for (int r = 1; r < dataGridView1.RowCount; r++)
            {
                for (int c = 0; c < 75; c++)
                {
                    //worksheet.Cells[r + 2, c + 1] = dataGridView1.Rows[r].Cells[c].Value;

                    //label19.Text += dataGridView1.Rows[r].Cells[c].Value;
                }
            }


            worksheet.Columns.AutoFit(); //자동 넒이
            app.DisplayAlerts = false;

            f1.releaseObject(worksheet);
            f1.releaseObject(workbook);
        }

        private void label12_Click(object sender, EventArgs e)
        {

        }
    }
}
