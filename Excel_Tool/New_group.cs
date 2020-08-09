using System;
using System.Data;
using System.Windows.Forms;

namespace 엑셀_관리_프로그램
{
    public partial class New_group : Form
    {
        Main_window f1 = new Main_window();
        Group group = new Group();

        public New_group()
        {
            InitializeComponent();
        }
        public New_group(Main_window _form)
        {

            InitializeComponent();
            f1 = _form;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //overlap();
            int num = f1.dataGridView2.Rows.Count - 1;

            //1번 row 체크
            if (num == 0)
            {
                group.group_new_num = "1";
            }
            else
            {
                group.group_new_num = num.ToString();
            }

            //신규 추가
            group.group_name = textBox1.Text;
            group.ceo_name = textBox2.Text;
            group.charge = textBox3.Text;
            group.tel = textBox4.Text;
            group.tel2 = textBox5.Text;
            group.address = textBox6.Text;
            group.email = textBox7.Text;

            // 관심분야
            if (checkBox1.Checked == true) { group.interest += checkBox1.Text + ", "; }
            if (checkBox2.Checked == true) { group.interest += checkBox2.Text + ", "; }
            if (checkBox3.Checked == true) { group.interest += checkBox3.Text + ", "; }
            if (checkBox4.Checked == true) { group.interest += checkBox4.Text + ", "; }
            if (checkBox5.Checked == true) { group.interest += checkBox5.Text + ", "; }
            // 방문 경로
            if (checkBox6.Checked == true) { group.path += checkBox6.Text + ", "; }
            if (checkBox7.Checked == true) { group.path += checkBox7.Text + ", "; }
            if (checkBox8.Checked == true) { group.path += checkBox8.Text + ", "; }
            if (checkBox9.Checked == true) { group.path += checkBox9.Text + ", "; }
            if (checkBox10.Checked == true) { group.path += checkBox10.Text + ", "; }
            if (checkBox11.Checked == true) { group.path += textBox8.Text + ", "; }

            DateTime dt = DateTime.Now;
            group.todaydate = dt.ToString();
            group.lastdate = dt.ToString();

            group.record = 1;

            newdata();
            f1.tabControl1.SelectedTab = f1.tabPage2;
            Application.OpenForms["New_group"].Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.OpenForms["New_group"].Close();
        }
        public void newdata()
        {
            DataTable dt = f1.dataGridView2.DataSource as DataTable;
            int num = f1.dataGridView2.Rows.Count + 1;
            int sum = 0;
            textBox6.Text = num.ToString();

            for (int i = 1; i < num; i++)
            {
                sum = i;
            }

            group.group_new_num = sum.ToString();


            dt.Rows.Add(group.group_new_num, group.group_name, group.ceo_name, group.charge,
            group.tel,
            group.tel2,
            group.address,
            group.email,
            group.interest,
            group.path,
            group.todaydate,
            group.lastdate,
            group.record);

            f1.dataGridView2.DataSource = dt;
        }

        public void overlap()
        {
            string name = null;
            int rows = f1.dataGridView2.Rows.Count; // 전체row수
            string number = textBox1.Text; // 입력한 회원 이름

            if (number == "") { MessageBox.Show("단체명을 입력해주세요."); return; }

            for (int i = 0; i <= rows; i++)
            {
                string value = f1.dataGridView2.Rows[i].Cells[1].Value.ToString();
                if (value == number)
                {
                    name = number;
                    break;
                }
            }
            if (name == textBox1.Text)
            {
                MessageBox.Show("이미 등록된 단체입니다.");
                return;
            }
        }

        private void Form_load(object sender, EventArgs e)
        {
            string[] data = { "현장방문", "대관신청", "교육상담", "노동상담", "영화", "기타" };


            // 각 콤보박스에 데이타를 초기화
            comboBox1.Items.AddRange(data);

            // 처음 선택값 지정. 첫째 아이템 선택
            comboBox1.SelectedIndex = 0;
        }
    }
}
