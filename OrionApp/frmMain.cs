using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;

namespace OrionApp
{
    public partial class frmMain : Form
    {
        private const int WM_SETREDRAW = 11;

        public frmMain()
        {
            InitializeComponent();
        }

        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int wMsg, bool wParam, int lParam);

        private void button1_Click(object sender, EventArgs e)
        {
            toolStripProgressBar1.Visible = true;
            dataGridView1.DoubleBuffered(true);
            SendMessage(dataGridView1.Handle, WM_SETREDRAW, false, 0);
            dataGridView1.Rows.Clear();
            using (var orion2019Entities = new Orion2019Entities())
            {
                foreach (var item in employeesCheckedListBox.CheckedItems)
                {
                    var names = from pList in orion2019Entities.pList
                                from pDivision in orion2019Entities.PDivision
                                from pLogData in orion2019Entities.pLogData
                                where pDivision.ID == pList.Section && pDivision.Name == orgComboBox.Text &&
                                      pLogData.HozOrgan == pList.ID
                                      && pLogData.DeviceTime > dateTimePicker1.Value &&
                                      pLogData.DeviceTime <= dateTimePicker2.Value
                                      && pList.Name + " " + pList.FirstName + " " + pList.MidName ==
                                      item.ToString()
                                      && (pLogData.ZoneIndex == 0 || pLogData.ZoneIndex == 1)
                                      && pLogData.Event == 32
                                orderby pLogData.DeviceTime
                                select new
                                {
                                    pList.Name,
                                    pList.MidName,
                                    pList.FirstName,
                                    pList.ID,
                                    pLogData.DeviceTime,
                                    pLogData.ZoneIndex
                                };
                    dataGridView1.ColumnCount = 5;
                    //dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    var rows = new List<DataGridViewRow>();

                    foreach (var name in names)
                    {
                        var a = name.ZoneIndex == 0 ? "Выход" : "Вход";
                        var row = new DataGridViewRow();
                        row.CreateCells(dataGridView1);
                        row.Cells[0].Value = name.Name;
                        row.Cells[1].Value = name.FirstName;
                        row.Cells[2].Value = name.MidName;
                        row.Cells[3].Value = name.DeviceTime;
                        row.Cells[4].Value = a;
                        rows.Add(row);
                        toolStripStatusLabel1.Text = "Получаем данные по сотруднику: " + name.Name + " " +
                                                     name.FirstName + " " + name.MidName;
                        SendMessage(statusStrip1.Handle, WM_SETREDRAW, true, 0);
                        statusStrip1.Refresh();
                        Application.DoEvents();
                    }

                    dataGridView1.Rows.AddRange(rows.ToArray());
                    SendMessage(dataGridView1.Handle, WM_SETREDRAW, true, 0);
                    dataGridView1.Refresh();
                }
            }

            SendMessage(dataGridView1.Handle, WM_SETREDRAW, true, 0);
            dataGridView1.Refresh();
            toolStripStatusLabel1.Text = "Готово!";
            toolStripProgressBar1.Visible = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var color1 = ColorTranslator.ToOle(Color.LightGray);
            var color2 = ColorTranslator.ToOle(Color.Black);

            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show(@"Нет данных для экспорта!");
                return;
            }

            var exapp = new Microsoft.Office.Interop.Excel.Application { Visible = false };
            exapp.Workbooks.Add();
            var workSheet = (Worksheet)exapp.ActiveSheet;
            workSheet.Cells[1, 1] = "Фамилия";
            var range = workSheet.UsedRange;
            var cell = range.Cells[1, 1];
            cell.Interior.Color = color1;
            cell.Font.Color = color2;
            var border = cell.Borders;
            border.LineStyle = XlLineStyle.xlContinuous;
            border.Weight = 2d;
            workSheet.Cells[1, 2] = "Имя";
            range = workSheet.UsedRange;
            cell = range.Cells[1, 2];
            cell.Interior.Color = color1;
            cell.Font.Color = color2;
            border = cell.Borders;
            border.LineStyle = XlLineStyle.xlContinuous;
            border.Weight = 2d;
            workSheet.Cells[1, 3] = "Отчество";
            range = workSheet.UsedRange;
            cell = range.Cells[1, 3];
            cell.Interior.Color = color1;
            cell.Font.Color = color2;
            border = cell.Borders;
            border.LineStyle = XlLineStyle.xlContinuous;
            border.Weight = 2d;
            workSheet.Cells[1, 4] = "Дата";
            range = workSheet.UsedRange;
            cell = range.Cells[1, 4];
            cell.Interior.Color = color1;
            cell.Font.Color = color2;
            border = cell.Borders;
            border.LineStyle = XlLineStyle.xlContinuous;
            border.Weight = 2d;
            workSheet.Cells[1, 5] = "Событие";
            range = workSheet.UsedRange;
            cell = range.Cells[1, 5];
            cell.Interior.Color = color1;
            cell.Font.Color = color2;
            border = cell.Borders;
            border.LineStyle = XlLineStyle.xlContinuous;
            border.Weight = 2d;
            var rowExcel = 2;
            for (var i = 0; i < dataGridView1.Rows.Count; i++)
                if (dataGridView1.Rows[i].Visible)
                {
                    workSheet.Cells[rowExcel, "A"] = dataGridView1.Rows[i].Cells[0].Value;
                    range = workSheet.UsedRange;
                    cell = range.Cells[rowExcel, "A"];
                    border = cell.Borders;
                    border.LineStyle = XlLineStyle.xlContinuous;
                    border.Weight = 2d;
                    workSheet.Cells[rowExcel, "B"] = dataGridView1.Rows[i].Cells[1].Value;
                    range = workSheet.UsedRange;
                    cell = range.Cells[rowExcel, "B"];
                    border = cell.Borders;
                    border.LineStyle = XlLineStyle.xlContinuous;
                    border.Weight = 2d;
                    workSheet.Cells[rowExcel, "C"] = dataGridView1.Rows[i].Cells[2].Value;
                    range = workSheet.UsedRange;
                    cell = range.Cells[rowExcel, "C"];
                    border = cell.Borders;
                    border.LineStyle = XlLineStyle.xlContinuous;
                    border.Weight = 2d;
                    workSheet.Cells[rowExcel, "D"] = dataGridView1.Rows[i].Cells[3].Value;
                    range = workSheet.UsedRange;
                    cell = range.Cells[rowExcel, "D"];
                    border = cell.Borders;
                    border.LineStyle = XlLineStyle.xlContinuous;
                    border.Weight = 2d;
                    workSheet.Cells[rowExcel, "E"] = dataGridView1.Rows[i].Cells[4].Value;
                    range = workSheet.UsedRange;
                    cell = range.Cells[rowExcel, "E"];
                    border = cell.Borders;
                    border.LineStyle = XlLineStyle.xlContinuous;
                    border.Weight = 2d;
                    ++rowExcel;
                }

            workSheet.Columns.AutoFit();

            exapp.Visible = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show(@"Нет данных для экспорта!");
                return;
            }

            var header = "Отчет по событиям ОрионПРО.\nОрганизация: " + comboBox1.Text + "\nПодразделение: " +
                         orgComboBox.Text +
                         "\nДата с " + dateTimePicker1.Value.ToShortDateString() + " по " +
                         dateTimePicker2.Value.ToShortDateString() + ".";
            var clsPrint = new ClsPrint(dataGridView1, header);
            clsPrint.PrintForm();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var photoForm = new photoForm();
            photoForm.ShowDialog();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                for (var i = employeesCheckedListBox.Items.Count - 1; i >= 0; i--) employeesCheckedListBox.SetItemChecked(i, true);

                employeesCheckedListBox.Enabled = false;
            }
            else
            {
                for (var i = 0; i < employeesCheckedListBox.Items.Count; i++) employeesCheckedListBox.SetItemChecked(i, false);

                employeesCheckedListBox.Enabled = true;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            orgComboBox.Items.Clear();
            using (var orion2019Entities = new Orion2019Entities())
            {
                var names = (from r in orion2019Entities.pList
                             from pDivision in orion2019Entities.PDivision
                             from pCompany in orion2019Entities.PCompany
                             where r.Company == pCompany.ID && r.Section == pDivision.ID && pCompany.Name == comboBox1.Text
                             select pDivision.Name).Distinct();

                foreach (var name in names) orgComboBox.Items.Add(name);
            }
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "";
            toolStripProgressBar1.Visible = false;
            dateTimePicker1.Value = DateTime.Today.AddDays(-1);
            comboBox1.Items.Clear();
            using (var orion2019Entities = new Orion2019Entities())
            {
                var names = (
                    from u in orion2019Entities.PCompany
                    select u.Name).Distinct();

                foreach (var name in names) comboBox1.Items.Add(name);
            }
        }

        private void orgComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            employeesCheckedListBox.Items.Clear();
            using (var orion2019Entities = new Orion2019Entities())
            {
                var names = from pList in orion2019Entities.pList
                            from pDivision in orion2019Entities.PDivision
                            from pCompany in orion2019Entities.PCompany
                            where pList.Company == pCompany.ID && pList.Section == pDivision.ID &&
                                  pCompany.Name == comboBox1.Text && pDivision.Name == orgComboBox.Text
                            orderby pList.Name
                            select new { pList.Name, pList.MidName, pList.FirstName };

                foreach (var name in names)
                {
                    employeesCheckedListBox.Items.Add(name.Name + " " + name.FirstName + " " + name.MidName);
                    employeesCheckedListBox.DrawMode = DrawMode.OwnerDrawFixed;
                }
            }
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            var bmp = new Bitmap(dataGridView1.Size.Width + 10, dataGridView1.Size.Height + 10);
            dataGridView1.DrawToBitmap(bmp, dataGridView1.Bounds);
            e.Graphics.DrawImage(bmp, 0, 0);
        }
    }

    public static class ExtensionMethods
    {
        public static void DoubleBuffered(this DataGridView dgv, bool setting)
        {
            var dgvType = dgv.GetType();
            var pi = dgvType.GetProperty("DoubleBuffered",
                BindingFlags.Instance | BindingFlags.NonPublic);
            if (pi != null) pi.SetValue(dgv, setting, null);
        }
    }
}