using System;
using System.Data;
using System.Windows.Forms;

namespace 엑셀_관리_프로그램
{

    public partial class New_User : Form
    {
        Main_window f1 = new Main_window();
        ExcelModule news = new ExcelModule();


        string reason = null;
        public New_User()
        {
            InitializeComponent();

        }
        public New_User(Main_window _form)
        {

            InitializeComponent();
            f1 = _form;
        }
        private void Form_load(object sender, EventArgs e)
        {
            string[] data = { "1", "2", "3", "4", "5", "기타" };


            // 각 콤보박스에 데이타를 초기화
            comboBox1.Items.AddRange(data);

            // 처음 선택값 지정. 첫째 아이템 선택
            comboBox1.SelectedIndex = 0;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            int num = f1.dataGridView1.Rows.Count - 1;

            if (textBox1.Text == "" && textBox3.Text == "")
            {
                MessageBox.Show("이름을과 번호를 입력해주세요."); 
                return;
            }
            else
            {
                //overlap(); // 중복체크

                //1번 row 체크
                if (num == 0)
                {
                    news.newnum = "1";
                }
                else
                {
                    news.newnum = num.ToString();
                }

                //신규 추가
                news.name = textBox1.Text;
                //성별 체크
                if (radioButton1.Checked == true)
                {
                    news.sex = radioButton1.Text;
                }
                else if (radioButton2.Checked == true)
                {
                    news.sex = radioButton2.Text;
                }
                else { MessageBox.Show("성별을 체크 해주세요."); }

                news.age = textBox2.Text;
                news.tel = @"@" + textBox3.Text;
                news.address = textBox4.Text;
                news.email = textBox5.Text;

                // 관심분야
                if (checkBox1.Checked == true) { news.interest += checkBox1.Text + ", "; }
                if (checkBox2.Checked == true) { news.interest += checkBox2.Text + ", "; }
                if (checkBox3.Checked == true) { news.interest += checkBox3.Text + ", "; }
                if (checkBox4.Checked == true) { news.interest += checkBox4.Text + ", "; }
                if (checkBox5.Checked == true) { news.interest += checkBox5.Text + ", "; }
                // 방문 경로
                if (checkBox6.Checked == true) { news.path += checkBox6.Text + ", "; }
                if (checkBox7.Checked == true) { news.path += checkBox7.Text + ", "; }
                if (checkBox8.Checked == true) { news.path += checkBox8.Text + ", "; }
                if (checkBox9.Checked == true) { news.path += checkBox9.Text + ", "; }
                if (checkBox10.Checked == true) { news.path += checkBox10.Text + ", "; }
                if (checkBox11.Checked == true) { news.path += textBox6.Text + ", "; }

                DateTime dt = DateTime.Now;
                news.todaydate = dt.ToString();
                news.lastdate = dt.ToString();
                reason = comboBox1.SelectedItem.ToString();
                news.record = 1;

                newdata();
                Application.OpenForms["New_User"].Close();
            }
        }

        public string str_value()
        {
            return reason;
        }


        // 중복체크 
        public bool overlap(bool over)
        {
            string name = null, tel = null;
            int rows = f1.dataGridView1.Rows.Count; // 전체row수
            string number = textBox1.Text; // 입력한 회원 이름
            string tel2 = textBox3.Text;


            for (int i = 0; i < rows; i++)
            {
                string value = f1.dataGridView1.Rows[i].Cells[1].Value.ToString();
                if (value == number)
                {
                    name = value;
                    break;
                }
            }

            for (int i = 0; i < rows; i++)
            {
                string value = f1.dataGridView1.Rows[i].Cells[4].Value.ToString();
                if (value == tel2) // 비교
                {
                    tel = value;
                    break;
                }
            }
            if (name == textBox1.Text && tel == textBox3.Text)
            {
                MessageBox.Show("이미 등록된 회원(단체)입니다.");
                
            }
            return over;
        }


        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.OpenForms["New_User"].Close();
        }

        public void newdata()
        {
            Microsoft.Office.Interop.Excel.Application new_app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = new_app.Workbooks.Add();
            worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(1);


            DataTable dt = f1.dataGridView1.DataSource as DataTable;
            int num = f1.dataGridView1.Rows.Count + 1;
            int sum = 0;

            for (int i = 1; i < num; i++)
            {
                sum = i;
            }

            news.newnum = sum.ToString();

            if (dt == null)
            {
                for (int i = 1; i < 13; i++)
                {

                    Microsoft.Office.Interop.Excel.Range cel0 = worksheet.Cells[1, i + 1];
                    cel0.Value = i;
                }

                workbook.SaveAs(System.Windows.Forms.Application.StartupPath + "\\엑셀_관리_프로그램\\엑셀_관리_프로그램_회원_개인_방문내역.xlsx");


                workbook.Close(true);
                f1.releaseObject(worksheet);
                f1.releaseObject(workbook);

                f1.dataGridView1.DataSource = dt;

                return;
            }
            else
            {
                dt.Rows.Add(
                    news.newnum, 
                    news.name, 
                    news.sex,
                    news.age,
                    news.tel,
                    news.address,
                    news.email,
                    news.interest,
                    news.path,
                    news.todaydate,
                    news.lastdate,
                    news.record);
            }
            f1.dataGridView1.DataSource = dt;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}