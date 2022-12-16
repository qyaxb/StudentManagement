using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;
namespace StudentManagement
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //Khai báo mảng 100 đối tượng sinh viên
        Employee[] listEmployee = new Employee[100];
        //Khai báo biến lưu thứ tự sinh viên
        int n = 0;
        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }
        //Sự kiện khi click nút ADD
        private void btnAdd_Click(object sender, EventArgs e)
        {
            
            if(txtID.Text == "" && txtName.Text == "" && txtAddress.Text == "" )
            {
                MessageBox.Show("You haven't enter any value");
                return;
            }
            if(rdoMale.Checked == false&&rdoFemale.Checked == false)
            {
                MessageBox.Show("You haven't choose gender");
                return;
            }
            //Tạo mới 1 đối tượng sinh viên
            //Đưa vào mảng listStudent
            //Hiển thị ra datagridview
            Employee employee = new Employee();
            employee.ID = txtID.Text;
            employee.Name = txtName.Text;
            employee.Address = txtAddress.Text;
            employee.Birth = dtpBirth.Text;
            employee.Gender = EGender ;
            employee.Department = cbDepartment.Text;
            listEmployee[n] = employee;
            n++;
            display();
            txtID.Text = "";
            txtName.Text = "";
            txtAddress.Text = "";
            dtpBirth.Text = "";
            cbDepartment.Text = "";
            rdoFemale.Checked = false;
            rdoMale.Checked = false;
        }
        public void display()
        {
            dgvStudent.DataSource = null;
            dgvStudent.DataSource = listEmployee;
          //Mỗi lần thêm mới thì dgv lại hiển thị lại danh sách
            dgvStudent.Refresh();
        }
        int index;
        private void dgvStudent_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            index = e.RowIndex;

            if (listEmployee[index] == null)
            {
                txtID.Text = "";
                txtName.Text = "";
                txtAddress.Text = "";
                dtpBirth.Text = "";
                cbDepartment.Text = "";
                rdoFemale.Checked = false;
                rdoMale.Checked = false;
                return;
            }
            
            
                txtID.Text = listEmployee[index].ID.ToString();
                txtName.Text = listEmployee[index].Name.ToString();
                txtAddress.Text = listEmployee[index].Address.ToString();
                dtpBirth.Text = listEmployee[index].Birth.ToString();
            if (listEmployee[index].Gender.Equals(rdoFemale.Text))
            {
                rdoFemale.Checked = true;
            }
            else
            {
                rdoMale.Checked = true;
            }
            cbDepartment.Text = listEmployee[index].Department.ToString();
            
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            listEmployee[index].ID = txtID.Text;
            listEmployee[index].Name = txtName.Text;
            listEmployee[index].Address = txtAddress.Text;
            listEmployee[index].Birth = dtpBirth.Text;
            listEmployee[index].Gender = EGender;
            
            listEmployee[index].Department = cbDepartment.Text;
            dgvStudent.DataSource = null;
            dgvStudent.DataSource = listEmployee;
            dgvStudent.Refresh();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            bool alert;
          alert =  MessageBox.Show("Do you want to delete this employee?", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK;
            if (alert)
            {
                //Khi xóa thì phần tử vị trí thứ index sẽ bị thay thể bởi phần tử thứ index+1
                while (index < n)
                {
                    listEmployee[index] = listEmployee[index + 1];
                    index++;
                }

            }
            else
            {
                return;
            }
            listEmployee[n - 1] = null;
            n = n - 1;
            
            dgvStudent.DataSource = null;
            dgvStudent.DataSource = listEmployee;
            dgvStudent.Refresh();
            MessageBox.Show("Delete successfully", "Notification", MessageBoxButtons.OK);
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void dgvStudent_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void rdoMale_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoMale.Checked == true)
            {
                EGender = "Male";
            }
        }
        string EGender;
        private void rdoFemale_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoFemale.Checked == true)
            {
                EGender = "Female";

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void btnLogOut_Click(object sender, EventArgs e)
        {
            FormLogin L = new FormLogin();
            L.Show();
            this.Hide();
        }
        private void ExportExcel(string path)
        {
            Excel.Application application = new Excel.Application();
            application.Application.Workbooks.Add(Type.Missing);
            for (int i = 0; i < dgvStudent.Columns.Count; i++)
            {
                application.Cells[1, i + 1] = dgvStudent.Columns[i].HeaderText;
            }
            for (int i = 0; i < dgvStudent.Rows.Count; i++)
            {
                for (int j = 0; j < dgvStudent.Columns.Count; j++)
                {
                    application.Cells[i + 2, j + 1] = dgvStudent.Rows[i].Cells[j].Value;
                }
            }
            application.Columns.AutoFit();
            application.ActiveWorkbook.SaveCopyAs(path);
            application.ActiveWorkbook.Saved = true;
        }
    
        private void dgvStudent_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            return;
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form1 Form = new Form1();
            Form.Show();
            this.Hide();
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Export Excel";
            saveFileDialog.Filter = "Excel (*.xlsx)|*.xlsx| Excel 2003 (*xls)|*.xls";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    ExportExcel(saveFileDialog.FileName);
                    MessageBox.Show("File export successful!");
                }
                catch(Exception ex)
                {
                    MessageBox.Show("File export unsuccessful!\n"+ ex.Message);
                }
            }
        }
    }
}
