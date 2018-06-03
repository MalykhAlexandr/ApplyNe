using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using F;
using FModel;

namespace Apply
{
    public partial class Form1 : Form
    {
        public Position Position { get; set; }
        public Form1()
        {
            InitializeComponent();
        }
        FootballDto GetModelFromUI()
        {
            return new FootballDto()
            {
                Filled = dateTimePicker1.Value,
                FullName = maskedTextBox1.Text,
                Citizenship = maskedTextBox2.Text,
                Age = (int)numericUpDown1.Value,
                Height = (int)numericUpDown2.Value,
                Weight = (int)numericUpDown3.Value,
                AgeStartCareer = (int)numericUpDown4.Value,
                ExperienceInFootball = (int)numericUpDown5.Value,
                CountTraums = (int)numericUpDown6.Value,
                TimeTraums = (int)numericUpDown7.Value,
                TraumаNow = maskedTextBox3.Text,
                Position = (Position)comboBox1.SelectedIndex,
                WorkingLeg = (WorkingLeg)comboBox2.SelectedIndex,
                Strength1 = (Strengths)comboBox3.SelectedIndex,
                WeakSides1 = (WeakSides)comboBox4.SelectedIndex,
                Strength2 = (Strengths)comboBox5.SelectedIndex,
                WeakSides2 = (WeakSides)comboBox6.SelectedIndex,
                Strength3 = (Strengths)comboBox7.SelectedIndex,
                WeakSides3 = (WeakSides)comboBox8.SelectedIndex,
                Traums = textBox1.Text,
            };
        }

        private void SetModelToUI(FootballDto dto)
        {
            dateTimePicker1.Value = dto.Filled;
            maskedTextBox1.Text = dto.FullName;
            maskedTextBox2.Text = dto.Citizenship;
            numericUpDown1.Value = dto.Age;
            numericUpDown2.Value = dto.Height;
            numericUpDown3.Value = dto.Weight;
            numericUpDown4.Value = dto.AgeStartCareer;
            numericUpDown5.Value = dto.ExperienceInFootball;
            numericUpDown6.Value = dto.CountTraums;
            numericUpDown7.Value = dto.TimeTraums;
            maskedTextBox3.Text = dto.TraumаNow;
            comboBox1.SelectedIndex = (int)dto.Position;
            comboBox2.SelectedIndex = (int)dto.WorkingLeg;
            comboBox3.SelectedIndex = (int)dto.Strength1;
            comboBox4.SelectedIndex = (int)dto.WeakSides1;
            comboBox5.SelectedIndex = (int)dto.Strength2;
            comboBox6.SelectedIndex = (int)dto.WeakSides2;
            comboBox7.SelectedIndex = (int)dto.Strength3;
            comboBox8.SelectedIndex = (int)dto.WeakSides3;
            textBox1.Text = dto.Traums;
        }

        private void button1_Click_1(object sender, System.EventArgs e)
        {
            var sfd = new SaveFileDialog() { Filter = "Файлы заявки|*.ftb" };
            var result = sfd.ShowDialog(this);
            if (result == DialogResult.OK)
            {
                var dto = GetModelFromUI();
                Helper.WriteToFile(sfd.FileName, dto);
            }
        }

        private void button2_Click(object sender, System.EventArgs e)
        {
            var ofd = new OpenFileDialog() { Filter = "Файл заявки|*.ftb" };
            var result = ofd.ShowDialog(this);
            if (result == DialogResult.OK)
            {
                var dto = Helper.LoadFromFile(ofd.FileName);
                SetModelToUI(dto);
            }
        }

        private void button3_Click(object sender, System.EventArgs e)
        {
            this.Close();
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            var lv = new LicenceValidator();
            if (!lv.HasLicense)
            {
                MessageBox.Show("Лицензия не найдена. Приобретите лицензию у владельца.");
                Application.Exit();
            }
            if (!lv.IsValid)
            {
                MessageBox.Show("Лицензия просрочена. Приобретите лицензию у владельца.");
                Application.Exit();
            }
        }
    }
}
