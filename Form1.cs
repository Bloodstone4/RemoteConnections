using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.DirectoryServices;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TerminalTools;

namespace RemoteConnections
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
                       
        }

        //Кнопка "Подключиться"
        private void Button1_Click(object sender, EventArgs e)
        {
            Connect();
        }

        //Событие при нажатии ENTER
        private void TextBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                Connect();
        }

        //Подключение 
        private void Connect()
        {
            this.Cursor = Cursors.WaitCursor;

            string server = comboBox1.Text;
            string message = "";
            var x = TermServicesManager.ListSessions(server);
            if (x.Count > 0)
                foreach (var item in x)
                {
                    if (item.ConnectionState.ToString() == "Active")
                    {
                        if (checkBox1.Checked)
                        {
                            message = TermServicesManager.GetSessionInfo(server, item.SessionId).UserName;
                            Process.Start("mstsc.exe", $@"/shadow:{item.SessionId} /control /v:{server}");
                        }
                        else
                        {
                            message = TermServicesManager.GetSessionInfo(server, item.SessionId).UserName;
                            Process.Start("mstsc.exe", $@"/shadow:{item.SessionId} /control /noConsentPrompt /v:{server}");
                        }
                    }
                }
            else
                message = "Недоступен";
            if (message == "")
                message = "Не подключено";
            textBox1.Text = message;

            this.Cursor = Cursors.Default;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.ActiveControl = comboBox1;
            GetComputersFromAD();
            //comboBox1.DroppedDown = true;


        }

        private void GetComputersFromAD()
        {
            var root = new DirectoryEntry("LDAP://" + "oilpro");
            var departmentSet = root.Children.Find("OU=Departments");
            List<string> strList = new List<string>();
            foreach (DirectoryEntry depart in departmentSet.Children)
            {
                if (depart.Children != null)
                {
                    bool isSuccess = true;
                    DirectoryEntry compSet = new DirectoryEntry();
                    try
                    {
                        compSet = depart.Children.Find("OU=Computers");
                    }
                    catch
                    {
                                
                        try
                        {
                            compSet = depart.Children.Find("OU=GAPR");
                            foreach (DirectoryEntry comp in compSet.Children)
                            {
                                                             
                                    foreach (DirectoryEntry c in comp.Children)
                                    {
                                        var offset = c.Name.IndexOf("=") + 1;
                                        strList.Add(c.Name.Substring(offset));
                                    }
                                
                            }
                        }
                        catch
                        {
                            isSuccess = false;
                        }
                    }
                    if (isSuccess)
                    {
                        if (compSet != null)
                        {

                            foreach (DirectoryEntry comp in compSet.Children)
                            {
                                var offset = comp.Name.IndexOf("=") + 1;
                                strList.Add(comp.Name.Substring(offset));
                            }
                        }
                            
                           
                        
                    }
                }
            }
            var strListOr = strList.OrderBy(x => x);
            foreach (var str in strListOr)
            {
                comboBox1.Items.Add(str);
            }
            
        }

        int count = 1;
        string text = string.Empty;
        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            
            object[] originalList = (object[])comboBox1.Tag;
            if (originalList == null)
            {
                // backup original list
                originalList = new object[comboBox1.Items.Count];
                comboBox1.Items.CopyTo(originalList, 0);
                comboBox1.Tag = originalList;
            }

            // prepare list of matching items
            string s = comboBox1.Text;
            IEnumerable<object> newList = originalList;
            if (s.Length > 0)
            {
                newList = originalList.Where(item => item.ToString().ToLower().Contains(s));
            }

            // clear list (loop through it, otherwise the cursor would move to the beginning of the textbox...)
            while (comboBox1.Items.Count > 0)
            {
                comboBox1.Items.RemoveAt(0);
            }

            //comboBox1.Items.Clear();
            // re-set list
            comboBox1.Items.AddRange(newList.ToArray());
            comboBox1.DroppedDown = true;
            //text= comboBox1.Text;

            //comboBox1.Text = text;

        }

        private void comboBox1_MouseClick(object sender, MouseEventArgs e)
        {
            comboBox1.Text = comboBox1.SelectedItem.ToString();
        }

        private void comboBox1_Click(object sender, EventArgs e)
        {
            comboBox1.Text = comboBox1.SelectedItem.ToString();
        }

        private void comboBox1_DragEnter(object sender, DragEventArgs e)
        {
            MessageBox.Show("Enter");
        }

        //private void comboBox1_KeyDown(object sender, KeyEventArgs e)
        //{
        //  //  var index= (string)comboBox1.SelectedIndex.ToString();
        //   // comboBox1.Text = 
        //}
    }
}
