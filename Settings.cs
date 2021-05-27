using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace MLG_Fetch
{
    public partial class Settings : Form
    {
        public Settings()
        {
            InitializeComponent();
            textBox2.Text = Properties.Settings.Default["det_db_name"].ToString();
            textBox1.Text = Properties.Settings.Default["det_db_path"].ToString();
            textBox3.Text = Properties.Settings.Default["per_db_name"].ToString();
            textBox4.Text = Properties.Settings.Default["per_db_path"].ToString();
            textBox5.Text = Properties.Settings.Default["report1_template_path"].ToString();


            textBox7.Text = Properties.Settings.Default["login"].ToString();
            textBox6.Text = Properties.Settings.Default["password"].ToString();
            checkBox3.Checked = Convert.ToBoolean(Properties.Settings.Default["is_cookie_enabled"]);
            textBox8.Text = Properties.Settings.Default["cookie"].ToString();
            textBox9.Text = Properties.Settings.Default["regfile"].ToString();
            textBox10.Text = Properties.Settings.Default["divs"].ToString();

            checkBox1.Checked = Convert.ToBoolean(Properties.Settings.Default["create_new_det"]);
            checkBox2.Checked = Convert.ToBoolean(Properties.Settings.Default["create_new_per"]);

            checkBox4.Checked = Convert.ToBoolean(Properties.Settings.Default["dep_month_new"]);

            textBox12.Text = Properties.Settings.Default["dep_month_filename"].ToString();
            textBox11.Text = Properties.Settings.Default["dep_month_path"].ToString();
            textBox13.Text = Properties.Settings.Default["dep_month_list"].ToString();

            //==========================dep reg================================
            checkBox5.Checked = Convert.ToBoolean(Properties.Settings.Default["dep_reg_new"]);
            textBox16.Text = Properties.Settings.Default["dep_reg_filename"].ToString();
            textBox15.Text = Properties.Settings.Default["dep_reg_path"].ToString();
            textBox17.Text = Properties.Settings.Default["dep_reg_list"].ToString();
            textBox14.Text = Properties.Settings.Default["dep_reg_database"].ToString();
            radioButton1.Checked = Convert.ToBoolean(Properties.Settings.Default["reg_sort_by_indexes"]);
            radioButton2.Checked = !Convert.ToBoolean(Properties.Settings.Default["reg_sort_by_indexes"]);

            //==========================fsec================================
            checkBox6.Checked = Convert.ToBoolean(Properties.Settings.Default["fsec_new"]);
            textBox21.Text = Properties.Settings.Default["fsec_filename"].ToString();
            textBox20.Text = Properties.Settings.Default["fsec_path"].ToString();
            textBox18.Text = Properties.Settings.Default["fsec_list"].ToString();
            textBox19.Text = Properties.Settings.Default["fsec_database"].ToString();

            //=========================media==================================
            textBox22.Text = Properties.Settings.Default["media_database"].ToString();

            //========================rep6=================================
            textBox23.Text = Properties.Settings.Default["rep6_path"].ToString();
            textBox24.Text = Properties.Settings.Default["rep6_filename"].ToString();
            textBox25.Text = Properties.Settings.Default["rep6_list"].ToString();
            textBox26.Text = Properties.Settings.Default["rep6_database"].ToString();
            checkBox7.Checked = Convert.ToBoolean(Properties.Settings.Default["rep6_new"]);
            

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox1.Text = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath)+"\\"+textBox2.Text+".xlsx";
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked) {
                label4.Visible = true;
                textBox2.Visible = true;
            } else {
                label4.Visible = false;
                textBox2.Visible = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                label5.Visible = true;
                textBox3.Visible = true;
            }
            else
            {
                label5.Visible = false;
                textBox3.Visible = false;
            }
        }

        private void button1_Click(object sender, EventArgs e) //saver
        {
            Properties.Settings.Default["det_db_path"] = textBox1.Text;
            Properties.Settings.Default["det_db_name"] = textBox2.Text;
            Properties.Settings.Default["per_db_name"] = textBox3.Text;
            Properties.Settings.Default["per_db_path"] = textBox4.Text;

            Properties.Settings.Default["create_new_det"] = checkBox1.Checked;
            Properties.Settings.Default["create_new_per"] = checkBox2.Checked;

            Properties.Settings.Default["report1_template_path"] = textBox5.Text;

            Properties.Settings.Default["login"] = textBox7.Text;
            Properties.Settings.Default["password"] = textBox6.Text;
            Properties.Settings.Default["is_cookie_enabled"] = checkBox3.Checked;
            Properties.Settings.Default["cookie"] = textBox8.Text;
            Properties.Settings.Default["regfile"] = textBox9.Text;

            //==========================dep month================================
            Properties.Settings.Default["dep_month_new"] = checkBox4.Checked;
            
            Properties.Settings.Default["dep_month_filename"] = textBox12.Text;
            Properties.Settings.Default["dep_month_path"] = textBox11.Text;
            Properties.Settings.Default["dep_month_list"] = textBox13.Text;

            //==========================dep reg================================
            Properties.Settings.Default["dep_reg_new"] = checkBox5.Checked;
            Properties.Settings.Default["dep_reg_filename"] = textBox16.Text;
            Properties.Settings.Default["dep_reg_path"] = textBox15.Text;
            Properties.Settings.Default["dep_reg_list"] = textBox17.Text;
            Properties.Settings.Default["dep_reg_database"] = textBox14.Text;
            Properties.Settings.Default["reg_sort_by_indexes"] = radioButton1.Checked;
            


            //========================fsec====================================
            Properties.Settings.Default["fsec_new"] = checkBox6.Checked;
            Properties.Settings.Default["fsec_filename"] = textBox21.Text;
            Properties.Settings.Default["fsec_path"] = textBox20.Text;
            Properties.Settings.Default["fsec_list"] = textBox18.Text;
            Properties.Settings.Default["fsec_database"] = textBox19.Text;

            //========================media================================
            Properties.Settings.Default["media_database"] = textBox22.Text;

            //========================rep6=================================
            Properties.Settings.Default["rep6_path"] = textBox23.Text;
            Properties.Settings.Default["rep6_filename"] = textBox24.Text;
            Properties.Settings.Default["rep6_new"] = checkBox7.Checked;
            Properties.Settings.Default["rep6_list"] = textBox25.Text;
            Properties.Settings.Default["rep6_database"] = textBox26.Text;


            try
            {
                Properties.Settings.Default["divs"] = Convert.ToInt32(textBox10.Text);
                if (Convert.ToInt32(textBox10.Text) <= 0)
                {
                    MessageBox.Show("Число делений для кумулятивного графика указано неверно.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch {
                MessageBox.Show("Число делений для кумулятивного графика указано неверно.", "Ошибка",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }

            if (textBox6.Text == "" | textBox7.Text == "")
            {
                MessageBox.Show("Для онлайн-запросов необходимо указать логин и пароль\n", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }


            Properties.Settings.Default.Save(); 
            MessageBox.Show("Настройки обновлены и сохранены","Сохранение");
            Close();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            textBox4.Text = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + "\\" + textBox3.Text + ".xlsx";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog f1 = new OpenFileDialog();
            f1.Filter = "Excel files (*.xlsx)|*.xlsx";
            f1.FilterIndex = 0;
            f1.RestoreDirectory = true;

            if (f1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = f1.FileName;
            }
            else {
                MessageBox.Show("Файл не был выбран или не может быть открыт", "Ошибка");
            }
            }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog f2 = new OpenFileDialog();
            f2.Filter = "Excel files (*.xlsx)|*.xlsx";
            f2.FilterIndex = 0;
            f2.RestoreDirectory = true;

            if (f2.ShowDialog() == DialogResult.OK)
            {
                textBox4.Text = f2.FileName;
            }
            else
            {
                MessageBox.Show("Файл не был выбран или не может быть открыт", "Ошибка");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog f3 = new OpenFileDialog();
            f3.Filter = "Word template files (*.docx)|*.docx";
            f3.FilterIndex = 0;
            f3.RestoreDirectory = true;

            if (f3.ShowDialog() == DialogResult.OK)
            {
                textBox5.Text = f3.FileName;
            }
            else
            {
                MessageBox.Show("Файл не был выбран или не может быть открыт", "Ошибка");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog f3 = new OpenFileDialog();
            f3.Filter = "Text files (*.txt)|*.txt";
            f3.FilterIndex = 0;
            f3.RestoreDirectory = true;

            if (f3.ShowDialog() == DialogResult.OK)
            {
                textBox9.Text = f3.FileName;
            }
            else
            {
                MessageBox.Show("Файл не был выбран или не может быть открыт", "Ошибка");
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                label12.Visible = true;
                textBox12.Visible = true;
            }
            else
            {
                label12.Visible = false;
                textBox12.Visible = false;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog f3 = new OpenFileDialog();
            f3.Filter = "Excel book (*.xlsx)|*.xlsx|Excel xls (*.xls)|*.xls";
            f3.FilterIndex = 0;
            f3.RestoreDirectory = true;

            if (f3.ShowDialog() == DialogResult.OK)
            {
                textBox11.Text = f3.FileName;
            }
            else
            {
                MessageBox.Show("Файл не был выбран или не может быть открыт", "Ошибка");
            }
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            textBox11.Text = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + "\\" + textBox12.Text + ".xlsx";
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            OpenFileDialog f3 = new OpenFileDialog();
            f3.Filter = "Text files (*.txt)|*.txt";
            f3.FilterIndex = 0;
            f3.RestoreDirectory = true;

            if (f3.ShowDialog() == DialogResult.OK)
            {
                textBox13.Text = f3.FileName;
            }
            else
            {
                MessageBox.Show("Файл не был выбран или не может быть открыт", "Ошибка");
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            OpenFileDialog f3 = new OpenFileDialog();
            f3.Filter = "Excel files (*.xlsx)|*.xlsx";
            f3.FilterIndex = 0;
            f3.RestoreDirectory = true;

            if (f3.ShowDialog() == DialogResult.OK)
            {
                textBox15.Text = f3.FileName;
            }
            else
            {
                MessageBox.Show("Файл не был выбран или не может быть открыт", "Ошибка");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            OpenFileDialog f3 = new OpenFileDialog();
            f3.Filter = "Excel files (*.xlsx)|*.xlsx";
            f3.FilterIndex = 0;
            f3.RestoreDirectory = true;

            if (f3.ShowDialog() == DialogResult.OK)
            {
                textBox14.Text = f3.FileName;
            }
            else
            {
                MessageBox.Show("Файл не был выбран или не может быть открыт", "Ошибка");
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            OpenFileDialog f3 = new OpenFileDialog();
            f3.Filter = "Text files (*.txt)|*.txt";
            f3.FilterIndex = 0;
            f3.RestoreDirectory = true;

            if (f3.ShowDialog() == DialogResult.OK)
            {
                textBox17.Text = f3.FileName;
            }
            else
            {
                MessageBox.Show("Файл не был выбран или не может быть открыт", "Ошибка");
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked)
            {
                label16.Visible = true;
                textBox16.Visible = true;
            }
            else
            {
                label16.Visible = false;
                textBox16.Visible = false;
            }
        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {
            textBox15.Text = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + "\\" + textBox16.Text + ".xlsx";

        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked)
            {
                label21.Visible = true;
                textBox21.Visible = true;
            }
            else
            {
                label21.Visible = false;
                textBox21.Visible = false;
            }
        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            textBox20.Text = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + "\\" + textBox21.Text + ".xlsx";
        }

        private void button13_Click(object sender, EventArgs e)
        {
            OpenFileDialog f3 = new OpenFileDialog();
            f3.Filter = "Excel files (*.xlsx)|*.xlsx";
            f3.FilterIndex = 0;
            f3.RestoreDirectory = true;

            if (f3.ShowDialog() == DialogResult.OK)
            {
                textBox20.Text = f3.FileName;
            }
            else
            {
                MessageBox.Show("Файл не был выбран или не может быть открыт", "Ошибка");
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            OpenFileDialog f3 = new OpenFileDialog();
            f3.Filter = "Excel files (*.xlsx)|*.xlsx";
            f3.FilterIndex = 0;
            f3.RestoreDirectory = true;

            if (f3.ShowDialog() == DialogResult.OK)
            {
                textBox19.Text = f3.FileName;
            }
            else
            {
                MessageBox.Show("Файл не был выбран или не может быть открыт", "Ошибка");
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            OpenFileDialog f3 = new OpenFileDialog();
            f3.Filter = "Text files (*.txt)|*.txt";
            f3.FilterIndex = 0;
            f3.RestoreDirectory = true;

            if (f3.ShowDialog() == DialogResult.OK)
            {
                textBox18.Text = f3.FileName;
            }
            else
            {
                MessageBox.Show("Файл не был выбран или не может быть открыт", "Ошибка");
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            OpenFileDialog f3 = new OpenFileDialog();
            f3.Filter = "Excel files (*.xlsx)|*.xlsx";
            f3.FilterIndex = 0;
            f3.RestoreDirectory = true;

            if (f3.ShowDialog() == DialogResult.OK)
            {
                textBox22.Text = f3.FileName;
            }
            else
            {
                MessageBox.Show("Файл не был выбран или не может быть открыт", "Ошибка");
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            OpenFileDialog f3 = new OpenFileDialog();
            f3.Filter = "Excel files (*.xlsx)|*.xlsx";
            f3.FilterIndex = 0;
            f3.RestoreDirectory = true;

            if (f3.ShowDialog() == DialogResult.OK)
            {
                textBox23.Text = f3.FileName;
            }
            else
            {
                MessageBox.Show("Файл не был выбран или не может быть открыт", "Ошибка");
            }
        }

        private void textBox24_TextChanged(object sender, EventArgs e)
        {
            textBox23.Text = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + "\\" + textBox24.Text + ".xlsx";
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked)
            {
                label24.Visible = true;
                textBox24.Visible = true;
            }
            else
            {
                label24.Visible = false;
                textBox24.Visible = false;
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            OpenFileDialog f3 = new OpenFileDialog();
            f3.Filter = "Text files (*.txt)|*.txt";
            f3.FilterIndex = 0;
            f3.RestoreDirectory = true;

            if (f3.ShowDialog() == DialogResult.OK)
            {
                textBox25.Text = f3.FileName;
            }
            else
            {
                MessageBox.Show("Файл не был выбран или не может быть открыт", "Ошибка");
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            OpenFileDialog f3 = new OpenFileDialog();
            f3.Filter = "Excel files (*.xlsx)|*.xlsx";
            f3.FilterIndex = 0;
            f3.RestoreDirectory = true;

            if (f3.ShowDialog() == DialogResult.OK)
            {
                textBox26.Text = f3.FileName;
            }
            else
            {
                MessageBox.Show("Файл не был выбран или не может быть открыт", "Ошибка");
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            Welcome form = new Welcome(); //Show license dialog
            form.ShowDialog();
        }
    }
}
