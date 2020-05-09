using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReDiveClanBattleHelper
{
    public partial class Form1 : Form
    {
        //^(2020-05-08 [0-5]+:[0-9]{2}:[0-9]{2})|(2020-05-07 [5-9]+:[0-9]{2}:[0-9]{2})|(^2020-05-07 [1-2]*[0-9]+:[0-9]{2}:[0-9]{2})
        static private bool Debug = false;
        static private Regex reg = new Regex("^[0-9]{4}-[0-9]{2}-[0-9]{2} [1-2]{0,1}[0-9]:[0-6][0-9]:[0-6][0-9] (.*)((\\([0-9]{5,12}\\))|(<(.*)>))$");
        static private Regex reg1 = new Regex("^(2020-05-08 [0-5]+:[0-9]{2}:[0-9]{2})|(2020-05-07 [5-9]+:[0-9]{2}:[0-9]{2})|(^2020-05-07 [1-2]*[0-9]+:[0-9]{2}:[0-9]{2}) (.*)((\\([0-9]{5,12}\\))|(<(.*)>))$");
        static private Regex reg2 = new Regex("^((完成)|(强行下树) [0-9]{1,7})|(完成 击杀)$");
        static private Regex regTime = new Regex("^(2020-05-08 [0-5]+:[0-9]{2}:[0-9]{2})|(2020-05-07 [5-9]+:[0-9]{2}:[0-9]{2})|(^2020-05-07 [1-2]*[0-9]+:[0-9]{2}:[0-9]{2})");
        static private Regex regName = new Regex(":[0-9]{2} (.*)((\\()|(<))");
        static private Regex regQQ = new Regex("(\\([0-9]{5,12}\\))|(<(.*)>)$");
        static private Regex regDmg = new Regex("([0-9]{1,7})|(击杀)");

        private Dictionary<string, string> nameList = new Dictionary<string, string>();
        private string filePath = "";
        private string defaultPath = @"E:\工会战\default.txt";

        private int Search(DataGridView dataGridView, string find, int index)
        {
            int row = dataGridView.Rows.Count;//得到总行数
            for (int i = 0; i < row; i++)//得到总行数并在之内循环
                if (find.Equals(dataGridView.Rows[i].Cells[index].Value))
                    return i;//定位到相同的单元格
            return -1;
        }

        private void GetNameList()
        {
            nameList = new Dictionary<string, string>();
            string line = "", name = "", qq = "";

            System.IO.StreamReader file = new System.IO.StreamReader(filePath);

            while ((line = file.ReadLine()) != null)
            {
                if (reg.IsMatch(line))
                {
                    name = regName.Match(line).ToString();
                    name = name.Substring(4, name.Length - 5);
                    qq = regQQ.Match(line).ToString();
                    qq = qq.Substring(1, qq.Length - 2);
                    if (nameList.ContainsKey(qq))
                        nameList[qq] = name;
                    else
                        nameList.Add(qq, name);
                }
            }
            file.Close();
        }

        private void FillBlank(DataGridView dataGridView)
        {
            int row = dataGridView.Rows.Count;
            int col = dataGridView.Columns.Count;
            for (int i = 0; i < row - 1; i++)//得到总行数并在之内循环
                nameList.Remove(dataGridView.Rows[i].Cells[1].Value.ToString());
            nameList.Remove("1000000");
            nameList.Remove("80000000");
            nameList.Remove("3386746424");
            nameList.Remove("10000");
            nameList.Remove("b@uid75.com");
            nameList.Remove("353211705");
            nameList.Remove("2948127973");
            nameList.Remove("190482951");
            nameList.Remove("1421447421");
            nameList.Remove("1339503382");
            nameList.Remove("365383062");
            foreach (KeyValuePair<string, string> p in nameList)
            {
                DataGridViewRow r = new DataGridViewRow();
                int index = dataGridView1.Rows.Add(r);
                dataGridView1.Rows[index].Cells[0].Value = index + 1;
                dataGridView1.Rows[index].Cells[1].Value = p.Key;
                dataGridView1.Rows[index].Cells[2].Value = p.Value;
                dataGridView1.Rows[index].Cells[3].Value = 0;
            }
        }

        private void Judge(DataGridView dataGridView)
        {
            int row = dataGridView.Rows.Count;
            int col = dataGridView.Columns.Count;
            for (int i = 0; i < row - 1; i++)//得到总行数并在之内循环
            {
                bool first = false, second = false, third = false, free = false;
                int count= Int32.Parse( dataGridView.Rows[i].Cells[3].Value.ToString());
                for(int j = 0; j < count; j++)
                {
                    int dmg =Int32.Parse( dataGridView.Rows[i].Cells[j*2+6].Value.ToString());
                    if (dmg <= 0)
                        continue;
                    else if (dmg < 400000)
                        free = true;
                    else if (dmg < 550000)
                        third = true;
                    else if (dmg < 700000)
                        second = true;
                    else
                        first = true;
                }
                string judge = "";
                if (first)
                    judge += "妈刀+";
                if (second)
                    judge += "充电狼+";
                if (third)
                    judge += "弟弟刀+";
                if (free)
                    judge += "补偿刀+";
                judge = judge.Substring(0, judge.Length - 1);
                dataGridView.Rows[i].Cells[4].Value = judge;
            }
        }

        private void Execute()
        {
            dataGridView1.Rows.Clear();
            GetNameList();
            int counter = 0;
            string line = "", name = "", qq = "";
            int damage = 0;
            DateTime time = new DateTime();
            bool match1 = false;

            System.IO.StreamReader file = new System.IO.StreamReader(filePath);

            while ((line = file.ReadLine()) != null)
            {
                if (match1)
                {
                    if (reg2.IsMatch(line))
                    {
                        string sdmg = regDmg.Match(line).ToString();
                        if (!sdmg.Equals("击杀"))
                        {
                            damage = Int32.Parse(sdmg);
                            if (damage == 0)
                                continue;
                        }
                        /* Log(time.ToString("T"));
                         Log(name);
                         Log(qq);
                         Log(damage);
                         Log("");*/
                        int rowIndex = Search(dataGridView1, qq, 1);
                        DataGridViewRow row;
                        if (rowIndex >= 0)
                        {
                            row = dataGridView1.Rows[rowIndex];
                            int index = (int)row.Cells[3].Value;
                            row.Cells[index * 2 + 6].Value = sdmg;
                            row.Cells[index * 2 + 7].Value = time.ToString("T");
                            row.Cells[3].Value = (int)row.Cells[3].Value + 1;
                            row.Cells[5].Value = (int)row.Cells[5].Value + damage;
                        }
                        else
                        {
                            row = new DataGridViewRow();
                            int index = dataGridView1.Rows.Add(row);
                            dataGridView1.Rows[index].Cells[0].Value = index + 1;
                            dataGridView1.Rows[index].Cells[1].Value = qq;
                            dataGridView1.Rows[index].Cells[2].Value = name;
                            dataGridView1.Rows[index].Cells[3].Value = 1;
                            dataGridView1.Rows[index].Cells[5].Value = damage;
                            dataGridView1.Rows[index].Cells[6].Value = sdmg;
                            dataGridView1.Rows[index].Cells[7].Value = time.ToString("T");
                        }
                        row = null;
                    }
                    name = "";
                    qq = "";
                    time = new DateTime();
                    damage = 0;
                    match1 = false;
                }
                else if (reg1.IsMatch(line))
                {
                    //Log(line);
                    string sTime = regTime.Match(line).ToString();
                    time = DateTime.Parse(regTime.Match(line).ToString());
                    name = regName.Match(line).ToString();
                    name = name.Substring(4, name.Length - 5);
                    qq = regQQ.Match(line).ToString();
                    qq = qq.Substring(1, qq.Length - 2);
                    match1 = true;
                }
                counter++;
            }

            file.Close();
            Judge(dataGridView1);
            FillBlank(dataGridView1);
            Log(counter);
        }


        private void Log(object s)
        {
            textBox1.Text += "\n\r" + s.ToString() + "\n\r";
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void InitDataGridView()
        {
            dataGridView1.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col1 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col2 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col3 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col4 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col5 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col6 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col7 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col8 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col9 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col10 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col11 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col12 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col13 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col14 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col15 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col16 = new DataGridViewTextBoxColumn();
            DataGridViewTextBoxColumn col17 = new DataGridViewTextBoxColumn();

            col.HeaderText = "序号";
            col1.HeaderText = "QQ号";
            col2.HeaderText = "出刀人";
            col3.HeaderText = "出刀数";
            col4.HeaderText = "智能识别出刀类型";
            col5.HeaderText = "总伤害";
            col6.HeaderText = "第一刀伤害";
            col7.HeaderText = "第一刀时间";
            col8.HeaderText = "第二刀伤害";
            col9.HeaderText = "第二刀时间";
            col10.HeaderText = "第三刀伤害";
            col11.HeaderText = "第三刀时间";
            col12.HeaderText = "第四刀伤害";
            col13.HeaderText = "第四刀时间";
            col14.HeaderText = "第五刀伤害";
            col15.HeaderText = "第五刀时间";
            col16.HeaderText = "第六刀伤害";
            col17.HeaderText = "第六刀时间";

            dataGridView1.Columns.Add(col);
            dataGridView1.Columns.Add(col1);
            dataGridView1.Columns.Add(col2);
            dataGridView1.Columns.Add(col3);
            dataGridView1.Columns.Add(col4);
            dataGridView1.Columns.Add(col5);
            dataGridView1.Columns.Add(col6);
            dataGridView1.Columns.Add(col7);
            dataGridView1.Columns.Add(col8);
            dataGridView1.Columns.Add(col9);
            dataGridView1.Columns.Add(col10);
            dataGridView1.Columns.Add(col11);
            dataGridView1.Columns.Add(col12);
            dataGridView1.Columns.Add(col13);
            dataGridView1.Columns.Add(col14);
            dataGridView1.Columns.Add(col15);
            dataGridView1.Columns.Add(col16);
            dataGridView1.Columns.Add(col17);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Visible = Debug;
            textBox2.Text = defaultPath;
            Log("start");
            InitDataGridView();
            //Execute();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件";
            dialog.Filter = "qq聊天记录(*.txt)|*.txt";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filePath = dialog.FileName;
                textBox2.Text = filePath;
                try
                {
                    button3_Click(null, null);
                }
                catch (Exception ex)
                {
                    Log(ex.Message);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                if (!File.Exists(defaultPath))
                {
                    MessageBox.Show("请先导入聊天记录");
                    return;

                }
                filePath = defaultPath;
            }
            string t = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string ta = dateTimePicker1.Value.AddDays(1).ToString("yyyy-MM-dd");
            //^(2020-05-08 [0-5]+:[0-9]{2}:[0-9]{2})|(2020-05-07 [5-9]+:[0-9]{2}:[0-9]{2})|(^2020-05-07 [1-2]*[0-9]+:[0-9]{2}:[0-9]{2})
            regTime = new Regex("^(" + ta + " [0-5]:[0-9]{2}:[0-9]{2})|(" + t + " [5-9]:[0-9]{2}:[0-9]{2})|(^" + t + " [1-2][0-9]:[0-9]{2}:[0-9]{2})");
            reg1 = new Regex("^((" + ta + " [0-5]:[0-9]{2}:[0-9]{2})|(" + t + " [5-9]:[0-9]{2}:[0-9]{2})|(^" + t + " [1-2][0-9]:[0-9]{2}:[0-9]{2}))(.*)$");
            Execute();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            //1141,561,dgv
            //1184,653,f1
            //653-561=f1-dgv    dgv=f1+561-653
            //1189,837
            //1146,745
        }

        private void Form1_ResizeEnd(object sender, EventArgs e)
        {
            dataGridView1.Height = this.Height + 561 - 653;
            dataGridView1.Width = this.Width + 1141 - 1184;
        }

        private void ExportToCsv(DataGridView dataGridView, string fp)
        {
            StreamWriter sw = new StreamWriter(fp, false, System.Text.Encoding.Default);

            int row = dataGridView.Rows.Count;
            int col = dataGridView.Columns.Count;

            string header = "";
            for (int i = 0; i < col; i++)
                header += dataGridView1.Columns[i].HeaderText + ",";
            sw.WriteLine(header);

            for (int i = 0; i < row; i++)//得到总行数并在之内循环
            {
                string line = "";
                for (int j = 0; j < col; j++)
                    line += dataGridView.Rows[i].Cells[j].Value + ",";
                sw.WriteLine(line);
            }

            sw.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件保存位置";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    ExportToCsv(dataGridView1, dialog.SelectedPath + "\\export.csv");
                    MessageBox.Show("导出完成，已导出为export.csv ，可用Excel打开");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导出失败\n\r" + ex.Message);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string target=comboBox1.SelectedItem.ToString();
            int row = dataGridView1.Rows.Count;
            int col = dataGridView1.Columns.Count;
            for (int i = 0; i < row - 1; i++)//得到总行数并在之内循环
            {
                object temp = dataGridView1.Rows[i].Cells[4].Value;
                if (null == temp)
                    continue;
                if (temp.ToString().Contains(target))
                {
                    dataGridView1.Rows.Remove(dataGridView1.Rows[i--]);
                    row--;
                }
            }
        }
    }
}
