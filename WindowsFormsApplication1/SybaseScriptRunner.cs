//Developer: Erwin Macalalad
//Email: retractsoulution@gmail.com
//If you have hundred of Sybase Database Server and you want to run single script across to all server
//and get the return SERVER STATUS or RESULT SET this simple SYBASE_SCRIPT_RUNNER is a big help for you.
//This Program is written is C# 2015,
//With dependencies: ClosedXML.0.88.0 but ClosedXML also have its own needed package DocumentFormat.OpenXml.2.7.2 and FastMember.Signed.1.1.0 all available in NuGeT.
//Last Modified 08-14-17

using ClosedXML.Excel;
using Sybase.Data.AseClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class SybaseScriptRunner : Form
    {
        private DoEvents doEvent = new DoEvents();
        private BackgroundWorker bw = new BackgroundWorker();
        private Queue<object> m_processQueue = new Queue<object>();
        private object m_syncObject = new object();
        DataTable dtServer = new DataTable();

        public SybaseScriptRunner()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dtServer.ReadXml("SEVERLIST.XML");

            listView1.Clear();
            listView1.Columns.Add("CODE", 50);
            listView1.Columns.Add("IP ADDRESS", 115);
            listView1.Columns.Add("STATUS", 80);

            fillServerList();

            doEvent.DoCompleted += DoCompleted;
            bw.DoWork += new DoWorkEventHandler(this.bw_DoWork);
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(this.bw_RunWorkerCompleted);

        }

        private void fillServerList()
        {
            listView1.Items.Clear();
            foreach (DataRow r in dtServer.Rows)
            {
                ListViewItem new_item = new ListViewItem();
                new_item.Text = r["SERVER_NAME"].ToString();
                new_item.Checked = r["CHECKED"].ToString() == "TRUE" ? true : false;

                new_item.SubItems.Add(
                        new ListViewItem.ListViewSubItem { Name = "IP_ADDRESS", Text = r["IP_ADDRESS"].ToString() }
                );
                new_item.SubItems.Add(
                        new ListViewItem.ListViewSubItem { Name = "STATUS", Text = "PENDING" }
                );
                new_item.Checked = true;
                listView1.Items.Add(new_item);
            }

            foreach(DataRow r in dtServer.Rows)
            {
                r["MSG"] = "";
                r["CHECKED"] = "FALSE";
                r["USER_NAME"] = "";
                r["PASSWORD"] = "";
                r["DIRECTORY"] = "";
                r["SCRIPT1"] = "";
                r["FILENAME"] = "";
            }

            File.Delete("SEVERLIST.XML");
            dtServer.WriteXml("SEVERLIST.XML", XmlWriteMode.WriteSchema);
        }

        private void DoCompleted(object sender, DoCompletedEventArgs e)
        {
            lock (this.m_syncObject)
                this.m_processQueue.Enqueue(e.office);
            if (!this.bw.IsBusy)
                this.bw.RunWorkerAsync();
        }

        void statusBar(string msg)
        {
            statusStrip1.Items["toolStripStatusLabel1"].Text = "Please wait processing " + msg;
        }
        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            if (this.m_processQueue.Count > 0)
            {
                DataRow dr;
                lock (this.m_syncObject)
                {
                    dr = (DataRow)m_processQueue.Dequeue();
                }

                sbLog sbLog = new sbLog(Path.Combine(dr["DIRECTORY"].ToString(), "log.txt"));
                string slog = "";

                slog = "Start Logging " + dr["SERVER_NAME"].ToString();
                sbLog.AppendLine(slog);
                statusBar(slog);

                try
                {
                    Stopwatch stopWatch = new Stopwatch();
                    stopWatch.Start();
                    

                    if (dr["CHECKED"].ToString() == "TRUE")
                    {
                        string conString = "DataSource='" + dr["IP_ADDRESS"] + "';UID='"+ dr["USER_NAME"].ToString() + "';PWD='" + dr["PASSWORD"].ToString() + "';";
                        using (AseConnection AseCon = new AseConnection(conString))
                        {
                            DataTable mainTable = new DataTable();

                            slog = dr["SERVER_NAME"].ToString() + " Start Connection.";
                            Console.WriteLine(slog);
                            sbLog.AppendLine(slog);
                            statusBar(slog);

                            AseCon.Open();
                            using (AseTransaction aseTrans = AseCon.BeginTransaction())
                            {
                                slog = dr["SERVER_NAME"].ToString() + " Start Transaction.";
                                Console.WriteLine(slog);
                                sbLog.AppendLine(slog);
                                statusBar(slog);


                                if (dr["SCRIPT1"].ToString().Trim().Length > 0)
                                {
                                    using (AseCommand comMain = new AseCommand())
                                    {
                                        comMain.CommandType = CommandType.Text;
                                        comMain.Connection = AseCon;
                                        comMain.CommandText = dr["SCRIPT1"].ToString().Trim();
                                        comMain.Transaction = aseTrans;
                                        comMain.CommandTimeout = 0;

                                        slog = dr["SERVER_NAME"].ToString() + " Start Main Query Retrieval.";
                                        Console.WriteLine(slog);
                                        sbLog.AppendLine(slog);
                                        statusBar(slog);

                                        using (AseDataAdapter da = new AseDataAdapter(comMain))
                                        {
                                            da.Fill(mainTable);

                                            slog = dr["SERVER_NAME"].ToString() + " Main Table Count: " + mainTable.Rows.Count;
                                            Console.WriteLine(slog);
                                            sbLog.AppendLine(slog);
                                            statusBar(slog);
                                        }
                                    }
                                }

                                if (mainTable.Rows.Count > 0)
                                {
                                    string fileResult = dr["SERVER_NAME"].ToString() + "-" + mainTable.Rows.Count;
                                    string fileName = Path.Combine(dr["DIRECTORY"].ToString(), fileResult + ".xlsx");
                                    dr["FILENAME"] = fileName;

                                    XLWorkbook wb = new XLWorkbook();
                                    wb.Worksheets.Add(mainTable, dr["SERVER_NAME"].ToString());
                                    wb.SaveAs(fileName);
                                    wb.Dispose();
                                    dr["MSG"] = "SUCCESSFULL - Record Count: " + mainTable.Rows.Count;
                                }
                                else
                                {
                                    dr["MSG"] = "NO RECORD";
                                }

                                mainTable.Dispose();
                                aseTrans.Commit();

                            } //using (AseTransaction aseTrans = AseCon.BeginTransaction())

                            stopWatch.Stop();
                            dr["MSG"] = dr["MSG"].ToString() + " Time Elapsed: " + stopWatch.Elapsed.ToString() + " Retrieval DateTime: " + DateTime.Now.ToString();

                            slog = dr["SERVER_NAME"].ToString() + " Done processing with no error.";
                            Console.WriteLine(slog);
                            sbLog.AppendLine(slog);
                            statusBar(slog);
                            
                            slog = dr["SERVER_NAME"].ToString() + " Time Elapsed : " + stopWatch.Elapsed.ToString();
                            Console.WriteLine(slog);
                            sbLog.AppendLine(slog);
                            statusBar(slog);
                            AseCon.Open();

                        } //using (AseConnection AseCon = new AseConnection(conString))
                    }
                    else
                    {
                        dr["MSG"] = "SKIP";

                        slog = dr["SERVER_NAME"].ToString() + " SKIP";
                        Console.WriteLine(slog);
                        sbLog.AppendLine(slog);
                        statusBar(slog);

                    }
                    
                    dr["CHECKED"] = "FALSE";
                }
                catch (AseException ex)
                {
                    dr["MSG"] = "AseException " + ex.Message;
                    dr["CHECKED"] = "TRUE";

                    slog = dr["SERVER_NAME"].ToString() + " Error AseException " + ex.Message;
                    Console.WriteLine(slog);
                    sbLog.AppendLine(slog);
                    //statusBar(slog);

                }
                catch (NullReferenceException ex)
                {
                    dr["MSG"] = "NullReferenceException " + ex.Message;
                    dr["CHECKED"] = "TRUE";
                    
                    slog = dr["SERVER_NAME"].ToString() + " Error NullReferenceException " + ex.Message;
                    Console.WriteLine(slog);
                    sbLog.AppendLine(slog);
                    //statusBar(slog);

                }

                e.Result = dr;

            } //if (this.m_processQueue.Count > 0)
        }

        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            DataRow r = (DataRow)e.Result;
            ListViewItem new_item = new ListViewItem();
            new_item.Text = r["SERVER_NAME"].ToString();
            new_item.SubItems.Add(
                    new ListViewItem.ListViewSubItem { Name = "IP_ADDRESS", Text = r["IP_ADDRESS"].ToString() }
            );
            new_item.SubItems.Add(
                    new ListViewItem.ListViewSubItem { Name = "FILENAME", Text = r["FILENAME"].ToString() }
            );
            new_item.SubItems.Add(
                    new ListViewItem.ListViewSubItem { Name = "MSG", Text = r["MSG"].ToString() }
            );
            new_item.Checked = true;
            listView2.Items.Add(new_item);

            ListViewItem item = listView1.FindItemWithText(r["SERVER_NAME"].ToString());
            item.Checked = r["CHECKED"].ToString() == "TRUE" ? true : false;
            item.SubItems["STATUS"].Text = r["MSG"].ToString();

            if (this.m_processQueue.Count > 0)
            {
                Application.DoEvents();
                this.bw.RunWorkerAsync();
            }
            else
            {
                button2.Text = "Run";
                button2.Enabled = true;
                button1.Enabled = true;
                button3.Enabled = true;
                listView1.CheckBoxes = true;
                textBox1.ReadOnly = false;
                button5.Enabled = true;
                this.Text = "Sybase Script Runner [DONE]";
                statusBar("Ready");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(textBox3.Text.Trim() == "")
            {
                MessageBox.Show("UserName cannot be empty.");
                return;
            }

            if (textBox4.Text.Trim() == "")
            {
                MessageBox.Show("Password cannot be empty.");
                return;
            }

            if (textBox1.Text.Trim() == "")
            {
                MessageBox.Show("Script to RUN cannot be empty.");
                return;
            }

            if(listView1.Items.Count == 0)
            {
                MessageBox.Show("Please create atleast one server.");
                return;
            }

            int serverCount = 0;
            foreach(ListViewItem i  in listView1.Items)
            {
                if(i.Checked)
                {
                    serverCount += 1;
                }
            }

            if (serverCount == 0)
            {
                MessageBox.Show("Please checked atleast one server.");
                return;
            }

            listView2.Clear();
            listView2.Columns.Add("SERVER NAME", 50);
            listView2.Columns.Add("IP ADDRESS", 115);
            listView2.Columns.Add("FILENAME", 200);
            listView2.Columns.Add("MSG", 700);
            button2.Enabled = false;
            button1.Enabled = false;
            button2.Text = "Please Wait...";
            button3.Enabled = false;
            listView1.CheckBoxes= false;
            button5.Enabled = false;
            this.Text = "Sybase Script Runner [Processing ...]";

            string dirResult = DateTime.Now.ToString("yyyyMMddHHmmss");
            if (!Directory.Exists(dirResult))
            {
                Directory.CreateDirectory(dirResult);
            }

            foreach (DataRow r in dtServer.Rows)
            {
                ListViewItem item = listView1.FindItemWithText(r["SERVER_NAME"].ToString());

                if (item != null)
                {
                    r["DIRECTORY"] = dirResult;
                    r["FILENAME"] = "";
                    r["MSG"] = "";
                    r["USER_NAME"] = "";
                    r["PASSWORD"] = "";
                    r["SCRIPT1"] = "";

                    if (item.Checked == true)
                    {
                        r["CHECKED"] = "TRUE";
                        r["USER_NAME"] = textBox3.Text.Trim();
                        r["PASSWORD"] = textBox4.Text.Trim();
                        r["SCRIPT1"] = textBox1.Text;
                        textBox1.ReadOnly = true;
                        item.SubItems["STATUS"].Text = "On Queue";
                    }
                    else
                    {
                        item.SubItems["STATUS"].Text = "SKIP";
                    }

                    textBox4.Text = "";
                }
            }
            doEvent.doAsync(dtServer);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            foreach(ListViewItem i in listView1.Items)
            {
                i.Checked = !i.Checked;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (listView2.Items.Count == 0) return;
            ListViewToCSV csv = new ListViewToCSV();
            string filename = "Report - " + DateTime.Now.ToString("yyyyMMddHHmmss") + ".CSV";
            csv.ToCSV(listView2,filename, true);
            MessageBox.Show("Report Save!\r\n" + filename);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            foreach(ListViewItem item in listView1.Items)
            {
                item.Checked = false;
            }
        }

        private void addServerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            formServer f = new formServer();

            DialogResult result = f.ShowDialog();
            if (result != DialogResult.OK)
            {
                f.Dispose();
                return;
            }

            ListViewItem item = listView1.FindItemWithText(f.serverName);

            if(item != null)
            {
                MessageBox.Show(f.serverName + " is already exists.");
                return;
            }

            DataRow r = dtServer.NewRow();
            r["SERVER_NAME"] = f.serverName;
            r["IP_ADDRESS"] = f.ipAddress;
            dtServer.Rows.Add(r);
            f.Dispose();

            fillServerList();
        }

        private void deleteServerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0)
            {
                MessageBox.Show("No server selected.");
                return;
            }

            DialogResult result;
            result = MessageBox.Show("Do you want to delete this server/s?", "Deleting...", MessageBoxButtons.YesNo);

            if (result == DialogResult.No)
            {
                return;
            }

            foreach (ListViewItem d in listView1.SelectedItems)
            {
                d.Remove();
                DataRow[] rR =  dtServer.Select("SERVER_NAME='" + d.Text.Trim() + "'");
                foreach(DataRow r in rR)
                {
                    dtServer.Rows.Remove(r);
                }
            }

            fillServerList();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
                "Developer: Erwin Macalalad\r\n\r\n" +
                "Email: retractsoulution@gmail.com\r\n\r\n" +
                "If you have hundred of Sybase Database Server and you want to run single script across to all server " +
                "and get the return SERVER STATUS or RESULT SET this simple SYBASE_SCRIPT_RUNNER is a big help for you. " +
                "This Program is written is C# 2015, "+
                "With dependencies: ClosedXML.0.88.0 but ClosedXML also have its own needed package DocumentFormat.OpenXml.2.7.2 and FastMember.Signed.1.1.0 all available in NuGeT.\r\n\r\n" +
                "Last Modified 08-14-17"
                );
        }
    }


    public class DoCompletedEventArgs : EventArgs
    {
        public DataRow office { get; private set; }
        public DoCompletedEventArgs(DataRow dr)
        {
            office = dr;
        }
    }

    public class DoEvents
    {
        public event EventHandler<DoCompletedEventArgs> DoCompleted;

        public void doAsync(DataTable dt)
        {
            foreach (DataRow r in dt.Rows)
            {
                if (DoCompleted != null)
                {
                    this.DoCompleted(this, new DoCompletedEventArgs(r));
                }
            }
        }
    }

    class ListViewToCSV
    {
        public void ToCSV(ListView listView, string filePath, bool includeHidden)
        {
            //make header string
            StringBuilder result = new StringBuilder();
            WriteCSVRow(result, listView.Columns.Count, i => includeHidden || listView.Columns[i].Width > 0, i => listView.Columns[i].Text);

            //export data rows
            foreach (ListViewItem listItem in listView.Items)
                WriteCSVRow(result, listView.Columns.Count, i => includeHidden || listView.Columns[i].Width > 0, i => listItem.SubItems[i].Text);

            File.WriteAllText(filePath, result.ToString());
        }

        private void WriteCSVRow(StringBuilder result, int itemsCount, Func<int, bool> isColumnNeeded, Func<int, string> columnValue)
        {
            bool isFirstTime = true;
            for (int i = 0; i < itemsCount; i++)
            {
                if (!isColumnNeeded(i))
                    continue;

                if (!isFirstTime)
                    result.Append(",");
                isFirstTime = false;

                result.Append(String.Format("\"{0}\"", columnValue(i)));
            }
            result.AppendLine();
        }
    }

    class sbLog
    {
        private string fileName = "";

        public sbLog(string f)
        {
            fileName = f;
        }

        public void AppendLine(string log)
        {
            File.AppendAllText(fileName, DateTime.Now.ToString() + "\t" + log + "\r\n");
        }
    }
}