using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StudentManagement
{
    public partial class FormLogin : Form
    {
        public FormLogin()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            string u = txtUserName.Text;
            string p = txtPassword.Text;
            if (u.Equals("admin") && p.Equals("123456"))
            {
                MessageBox.Show("Login successfully", "Notification", MessageBoxButtons.OK,MessageBoxIcon.Information);
                    //Đăng nhập thành công hiển thị form student
                    Form1 f = new Form1();
                f.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Login failed, try again", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }    
        }

        private void FormLogin_Load(object sender, EventArgs e)
        {

        }
    }
}
