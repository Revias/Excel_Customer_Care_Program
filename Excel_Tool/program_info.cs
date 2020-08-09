using System.Windows.Forms;

namespace 엑셀_관리_프로그램
{
    public partial class program_info : Form
    {
        Main_window f1 = new Main_window();
        public program_info()
        {
            InitializeComponent();
        }
        public program_info(Main_window _form)
        {

            InitializeComponent();
            f1 = _form;
        }

        private void program_info_Load(object sender, System.EventArgs e)
        {

        }
    }
}
