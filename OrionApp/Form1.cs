using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OrionApp
{
    public partial class photoForm : Form
    {
        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int wMsg, bool wParam, int lParam);

        public photoForm()
        {
            InitializeComponent();
        }

        private void photoForm_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            using (var orion2019Entities = new Orion2019Entities())
            {
                var names = (
                    from u in orion2019Entities.PCompany
                    select u.Name).Distinct();

                foreach (var name in names) comboBox1.Items.Add(name);
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

        private void orgComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
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
                    comboBox2.Items.Add(name.Name + " " + name.FirstName + " " + name.MidName);
                }
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            using (var orion2019Entities = new Orion2019Entities())
            {
                var names = from pList in orion2019Entities.pList
                            from pDivision in orion2019Entities.PDivision
                            from pCompany in orion2019Entities.PCompany
                            where pList.Company == pCompany.ID && pList.Section == pDivision.ID &&
                                  pCompany.Name == comboBox1.Text && pDivision.Name == orgComboBox.Text
                                  && pList.Name + " " + pList.FirstName + " " + pList.MidName ==
                                  comboBox2.Text
                            orderby pList.Name
                            select new { pList.Name, pList.MidName, pList.FirstName, pList.Picture };

                foreach (var iName in names)
                {
                    try
                    {
                        var ms1 = new MemoryStream(iName.Picture);
                        pictureBox1.Image = Image.FromStream(ms1);
                    }
                    catch (Exception)
                    {
                        pictureBox1.Image = null;
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            pictureBox1.Image.Save(folderBrowserDialog1.SelectedPath + "\\" + comboBox2.Text + ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {
        }
    }
}