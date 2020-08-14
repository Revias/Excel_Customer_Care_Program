using System;
using System.Windows.Forms;

namespace 엑셀_관리_프로그램
{
    public partial class Revisit : Form
    {
        Main_window f1 = new Main_window();
        string name = null;
        string tel = null;

        public Revisit()
        {
            InitializeComponent();

        }

        public Revisit(Main_window _form)
        {
            InitializeComponent();
            f1 = _form;
        }
        public string user_info(string test)
        {
            string str;
            str = test;
            return str;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            //개인 재방문
            if(checkBox1.Checked == false)
            {
                int rows = f1.dataGridView1.Rows.Count; // 전체row수
                int column = 0;
                string nametext = textBox1.Text; // 입력한 회원 이름
                string tel2 = textBox2.Text;

                if (nametext == "") { MessageBox.Show("이름을 입력해주세요."); return; }
                if (tel2 == "") { MessageBox.Show("전화번호를을 입력해주세요."); return; }

                for (int i = 0; i < rows-1; i++)
                {
                    string value = f1.dataGridView1.Rows[i].Cells[1].Value.ToString();
                    if (value == nametext)
                    {
                        name = nametext;
                        column = i; //방문 수 카운트
                        break; // 리턴을 하면 안된다..
                    }
                }

                for (int i = 0; i < rows-1; i++)
                {
                    string value = f1.dataGridView1.Rows[i].Cells[4].Value.ToString();
                    if (value == tel2) // 비교
                    {
                        tel = value;
                        break;
                    }
                }
                if (name == textBox1.Text && tel == textBox2.Text)
                {
                    DateTime dt = DateTime.Now;
                    f1.dataGridView1.Rows[column].Cells[10].Value = dt;
                    string str = f1.dataGridView1.Rows[column].Cells[11].Value.ToString();
                    int num = Convert.ToInt32(str);
                    num += 1;
                    f1.dataGridView1.Rows[column].Cells[11].Value = num.ToString();
                    MessageBox.Show("추가 되었습니다.");
                    int iRowIdx = column;
                    if (iRowIdx >= 0)
                    {
                        f1.dataGridView1.FirstDisplayedCell = f1.dataGridView1.Rows[column].Cells[1];
                        f1.dataGridView1.Rows[column].Cells[1].Selected = true;
                        Application.OpenForms["Revisit"].Close();
                        return;
                        
                    }
                    MessageBox.Show("재방문 확인");
                    
                }
                else
                { MessageBox.Show("등록되지 않은 회원입니다."); }

                Application.OpenForms["Revisit"].Close();
            }
            // 단체 재방문
            else if(checkBox1.Checked == true)
            {
                int rows = f1.dataGridView2.Rows.Count; // 전체row수
                int column = 0;
                string nametext = textBox1.Text; // 입력한 회원 이름
                string tel2 = textBox2.Text;

                if (nametext == "") { MessageBox.Show("이름을 입력해주세요."); return; }
                if (tel2 == "") { MessageBox.Show("전화번호를을 입력해주세요."); return; }

                for (int i = 0; i < rows-1; i++)
                {
                    string value = f1.dataGridView2.Rows[i].Cells[1].Value.ToString();
                    if (value == nametext)
                    {
                        name = nametext;
                        column = i;
                        break;
                    }
                }

                for (int i = 0; i < rows-1; i++)
                {
                    string value = f1.dataGridView2.Rows[i].Cells[5].Value.ToString();
                    if (value == tel2) // 비교
                    {
                        tel = value;
                        break;
                    }
                }

                if (name == textBox1.Text && tel == textBox2.Text)
                {
                    DateTime dt = DateTime.Now;
                    f1.dataGridView2.Rows[column].Cells[11].Value = dt;
                    string str = f1.dataGridView2.Rows[column].Cells[12].Value.ToString();
                    int num = Convert.ToInt32(str);
                    num += 1;
                    f1.dataGridView2.Rows[column].Cells[12].Value = num.ToString();
                    MessageBox.Show("추가 되었습니다.");
                    int iRowIdx = column;
                    if (iRowIdx >= 0)
                    {
                        f1.dataGridView2.FirstDisplayedCell = f1.dataGridView2.Rows[column].Cells[1];
                        f1.dataGridView2.Rows[column].Cells[1].Selected = true;
                        Application.OpenForms["Revisit"].Close();
                        return;
                    }
                }
                else
                { 
                    MessageBox.Show("등록되지 않은 단체입니다.");
                    Application.OpenForms["Revisit"].Close();
                }


                MessageBox.Show("추가 되었습니다.");
                Application.OpenForms["Revisit"].Close();
            }
            Application.OpenForms["Revisit"].Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.OpenForms["Revisit"].Close();
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))    //숫자와 백스페이스를 제외한 나머지를 바로 처리
            {
                e.Handled = true;
            }
        }

        private void Form_load(object sender, EventArgs e)
        {
            string[] data = { "1", "2", "3", "4", "5", "6" };


            // 각 콤보박스에 데이타를 초기화
            comboBox1.Items.AddRange(data);

            // 처음 선택값 지정. 첫째 아이템 선택
            comboBox1.SelectedIndex = 0;
        }
    }
}
