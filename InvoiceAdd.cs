using System;
using System.Windows.Forms;
using System.Data;
using System.IO;
using Interop.QBFC13;
using System.Xml.Linq;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace InvoiceAdd
{
    public class frm1_InvoiceAdd : System.Windows.Forms.Form
    {
        private System.ComponentModel.Container components = null;
        private System.Windows.Forms.Button btn1_Send;
        private System.Windows.Forms.Button btn2_Exit;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Button btnOpenFile_Reset;
        private DataGridView dataGridView1;
        public int CurrentRow = 0;

        DataTable dt = new DataTable();
        DataTable dtTemp2 = new DataTable();
        string fileName = string.Empty;
        private CheckBox checkBox1;
        private CheckBox checkBox2;
        private CheckBox checkBox3;
        private CheckBox checkBox4;
        private CheckBox checkBox5;
        private CheckBox checkBox6;
        private CheckBox checkBox7;
        private CheckBox checkBox9;
        private CheckBox checkBox10;
        private CheckBox checkBox11;
        private CheckBox checkBox12;
        private CheckBox checkBox8;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private TextBox textBox1;
        private Label label5;
        private TextBox tbProgramLog;
        bool ifError = false;

        public frm1_InvoiceAdd()
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.btn1_Send = new System.Windows.Forms.Button();
            this.btn2_Exit = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.btnOpenFile_Reset = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.checkBox3 = new System.Windows.Forms.CheckBox();
            this.checkBox4 = new System.Windows.Forms.CheckBox();
            this.checkBox5 = new System.Windows.Forms.CheckBox();
            this.checkBox6 = new System.Windows.Forms.CheckBox();
            this.checkBox7 = new System.Windows.Forms.CheckBox();
            this.checkBox9 = new System.Windows.Forms.CheckBox();
            this.checkBox10 = new System.Windows.Forms.CheckBox();
            this.checkBox11 = new System.Windows.Forms.CheckBox();
            this.checkBox12 = new System.Windows.Forms.CheckBox();
            this.checkBox8 = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.tbProgramLog = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // btn1_Send
            // 
            this.btn1_Send.BackColor = System.Drawing.SystemColors.Control;
            this.btn1_Send.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F);
            this.btn1_Send.Location = new System.Drawing.Point(482, 575);
            this.btn1_Send.Name = "btn1_Send";
            this.btn1_Send.Size = new System.Drawing.Size(80, 32);
            this.btn1_Send.TabIndex = 57;
            this.btn1_Send.Text = "Send";
            this.btn1_Send.UseVisualStyleBackColor = false;
            this.btn1_Send.Click += new System.EventHandler(this.btn1_Send_Click_1);
            // 
            // btn2_Exit
            // 
            this.btn2_Exit.BackColor = System.Drawing.SystemColors.Control;
            this.btn2_Exit.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn2_Exit.Location = new System.Drawing.Point(570, 575);
            this.btn2_Exit.Name = "btn2_Exit";
            this.btn2_Exit.Size = new System.Drawing.Size(75, 32);
            this.btn2_Exit.TabIndex = 58;
            this.btn2_Exit.Text = "Exit";
            this.btn2_Exit.UseVisualStyleBackColor = false;
            this.btn2_Exit.Click += new System.EventHandler(this.btn2_Exit_Click);
            // 
            // btnOpenFile_Reset
            // 
            this.btnOpenFile_Reset.BackColor = System.Drawing.SystemColors.Control;
            this.btnOpenFile_Reset.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOpenFile_Reset.Location = new System.Drawing.Point(529, 534);
            this.btnOpenFile_Reset.Name = "btnOpenFile_Reset";
            this.btnOpenFile_Reset.Size = new System.Drawing.Size(80, 32);
            this.btnOpenFile_Reset.TabIndex = 63;
            this.btnOpenFile_Reset.Text = "Open File";
            this.btnOpenFile_Reset.UseVisualStyleBackColor = false;
            this.btnOpenFile_Reset.Click += new System.EventHandler(this.btnOpenFile_Reset_Click);
            // 
            // dataGridView1
            // 
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView1.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView1.Location = new System.Drawing.Point(7, 82);
            this.dataGridView1.Name = "dataGridView1";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridView1.Size = new System.Drawing.Size(637, 400);
            this.dataGridView1.TabIndex = 64;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(138, 488);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(44, 17);
            this.checkBox1.TabIndex = 65;
            this.checkBox1.Text = "Yes";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(247, 488);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(40, 17);
            this.checkBox2.TabIndex = 66;
            this.checkBox2.Text = "No";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // checkBox3
            // 
            this.checkBox3.AutoSize = true;
            this.checkBox3.Location = new System.Drawing.Point(333, 488);
            this.checkBox3.Name = "checkBox3";
            this.checkBox3.Size = new System.Drawing.Size(119, 17);
            this.checkBox3.TabIndex = 67;
            this.checkBox3.Text = "ItemCode not found";
            this.checkBox3.UseVisualStyleBackColor = true;
            // 
            // checkBox4
            // 
            this.checkBox4.AutoSize = true;
            this.checkBox4.Location = new System.Drawing.Point(138, 511);
            this.checkBox4.Name = "checkBox4";
            this.checkBox4.Size = new System.Drawing.Size(86, 17);
            this.checkBox4.TabIndex = 68;
            this.checkBox4.Text = "Item Number";
            this.checkBox4.UseVisualStyleBackColor = true;
            // 
            // checkBox5
            // 
            this.checkBox5.AutoSize = true;
            this.checkBox5.Location = new System.Drawing.Point(247, 511);
            this.checkBox5.Name = "checkBox5";
            this.checkBox5.Size = new System.Drawing.Size(71, 17);
            this.checkBox5.TabIndex = 69;
            this.checkBox5.Text = "ItemCode";
            this.checkBox5.UseVisualStyleBackColor = true;
            // 
            // checkBox6
            // 
            this.checkBox6.AutoSize = true;
            this.checkBox6.Location = new System.Drawing.Point(333, 511);
            this.checkBox6.Name = "checkBox6";
            this.checkBox6.Size = new System.Drawing.Size(85, 17);
            this.checkBox6.TabIndex = 70;
            this.checkBox6.Text = "Part Number";
            this.checkBox6.UseVisualStyleBackColor = true;
            // 
            // checkBox7
            // 
            this.checkBox7.AutoSize = true;
            this.checkBox7.Location = new System.Drawing.Point(419, 511);
            this.checkBox7.Name = "checkBox7";
            this.checkBox7.Size = new System.Drawing.Size(79, 17);
            this.checkBox7.TabIndex = 71;
            this.checkBox7.Text = "Description";
            this.checkBox7.UseVisualStyleBackColor = true;
            // 
            // checkBox9
            // 
            this.checkBox9.AutoSize = true;
            this.checkBox9.Location = new System.Drawing.Point(138, 534);
            this.checkBox9.Name = "checkBox9";
            this.checkBox9.Size = new System.Drawing.Size(44, 17);
            this.checkBox9.TabIndex = 73;
            this.checkBox9.Text = "Yes";
            this.checkBox9.UseVisualStyleBackColor = true;
            // 
            // checkBox10
            // 
            this.checkBox10.AutoSize = true;
            this.checkBox10.Location = new System.Drawing.Point(247, 534);
            this.checkBox10.Name = "checkBox10";
            this.checkBox10.Size = new System.Drawing.Size(40, 17);
            this.checkBox10.TabIndex = 74;
            this.checkBox10.Text = "No";
            this.checkBox10.UseVisualStyleBackColor = true;
            // 
            // checkBox11
            // 
            this.checkBox11.AutoSize = true;
            this.checkBox11.Location = new System.Drawing.Point(138, 559);
            this.checkBox11.Name = "checkBox11";
            this.checkBox11.Size = new System.Drawing.Size(44, 17);
            this.checkBox11.TabIndex = 75;
            this.checkBox11.Text = "Yes";
            this.checkBox11.UseVisualStyleBackColor = true;
            // 
            // checkBox12
            // 
            this.checkBox12.AutoSize = true;
            this.checkBox12.Location = new System.Drawing.Point(247, 559);
            this.checkBox12.Name = "checkBox12";
            this.checkBox12.Size = new System.Drawing.Size(40, 17);
            this.checkBox12.TabIndex = 76;
            this.checkBox12.Text = "No";
            this.checkBox12.UseVisualStyleBackColor = true;
            // 
            // checkBox8
            // 
            this.checkBox8.AutoSize = true;
            this.checkBox8.Location = new System.Drawing.Point(505, 511);
            this.checkBox8.Name = "checkBox8";
            this.checkBox8.Size = new System.Drawing.Size(65, 17);
            this.checkBox8.TabIndex = 72;
            this.checkBox8.Text = "Quantity";
            this.checkBox8.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 492);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 13);
            this.label1.TabIndex = 77;
            this.label1.Text = "ICodes match?";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 515);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(124, 13);
            this.label2.TabIndex = 78;
            this.label2.Text = "Columns in correct order:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(7, 537);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(73, 13);
            this.label3.TabIndex = 79;
            this.label3.Text = "Correct input?";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(7, 562);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(98, 13);
            this.label4.TabIndex = 80;
            this.label4.Text = "Row order correct?";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(7, 40);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(637, 36);
            this.textBox1.TabIndex = 83;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(242, 9);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(130, 25);
            this.label5.TabIndex = 84;
            this.label5.Text = "BOM INFO.";
            // 
            // tbProgramLog
            // 
            this.tbProgramLog.Location = new System.Drawing.Point(675, 40);
            this.tbProgramLog.Multiline = true;
            this.tbProgramLog.Name = "tbProgramLog";
            this.tbProgramLog.Size = new System.Drawing.Size(288, 442);
            this.tbProgramLog.TabIndex = 85;
            // 
            // frm1_InvoiceAdd
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.ClientSize = new System.Drawing.Size(976, 617);
            this.Controls.Add(this.tbProgramLog);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.checkBox12);
            this.Controls.Add(this.checkBox11);
            this.Controls.Add(this.checkBox10);
            this.Controls.Add(this.checkBox9);
            this.Controls.Add(this.checkBox8);
            this.Controls.Add(this.checkBox7);
            this.Controls.Add(this.checkBox6);
            this.Controls.Add(this.checkBox5);
            this.Controls.Add(this.checkBox4);
            this.Controls.Add(this.checkBox3);
            this.Controls.Add(this.checkBox2);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btnOpenFile_Reset);
            this.Controls.Add(this.btn2_Exit);
            this.Controls.Add(this.btn1_Send);
            this.Name = "frm1_InvoiceAdd";
            this.Text = "Add an Item Inventory to QuickBooks";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        [STAThread]
        static void Main()
        {
            Application.Run(new frm1_InvoiceAdd());
        }

        private void btn2_Exit_Click(object sender, System.EventArgs e)
        {
            Dispose();
        }
        
        private double QBFCLatestVersion(QBSessionManager SessionManager)
        {
            // Use oldest version to ensure that this application work with any QuickBooks (US)
            IMsgSetRequest msgset = SessionManager.CreateMsgSetRequest("US", 1, 0);
            msgset.AppendHostQueryRq();
            IMsgSetResponse QueryResponse = SessionManager.DoRequests(msgset);
            
            // The response list contains only one response, which corresponds to our single HostQuery request
            IResponse response = QueryResponse.ResponseList.GetAt(0);

            // Please refer to QBFC Developers Guide for details on why "as" clause was used to link this derrived class to its base class
            IHostRet HostResponse = response.Detail as IHostRet;
            IBSTRList supportedVersions = HostResponse.SupportedQBXMLVersionList as IBSTRList;

            int i;
            double vers;
            double LastVers = 0;
            string svers = null;
            for (i = 0; i <= supportedVersions.Count - 1; i++)
            {
                svers = supportedVersions.GetAt(i);
                vers = Convert.ToDouble(svers);
                if (vers > LastVers)
                {
                    LastVers = vers;
                }
            }
            return LastVers;
        }

        public IMsgSetRequest getLatestMsgSetRequest(QBSessionManager sessionManager)
        {
            double supportedVersion = QBFCLatestVersion(sessionManager);
            short qbXMLMajorVer = 0;
            short qbXMLMinorVer = 0;

            if (supportedVersion >= 6.0)
            {
                qbXMLMajorVer = 6;
                qbXMLMinorVer = 0;
            }
            else if (supportedVersion >= 5.0)
            {
                qbXMLMajorVer = 5;
                qbXMLMinorVer = 0;
            }
            else if (supportedVersion >= 4.0)
            {
                qbXMLMajorVer = 4;
                qbXMLMinorVer = 0;
            }
            else if (supportedVersion >= 3.0)
            {
                qbXMLMajorVer = 3;
                qbXMLMinorVer = 0;
            }
            else if (supportedVersion >= 2.0)
            {
                qbXMLMajorVer = 2;
                qbXMLMinorVer = 0;
            }
            else if (supportedVersion >= 1.1)
            {
                qbXMLMajorVer = 1;
                qbXMLMinorVer = 1;
            }
            else
            {
                qbXMLMajorVer = 1;
                qbXMLMinorVer = 0;
                tbProgramLog.AppendText(Environment.NewLine + "It seems that you are running QuickBooks 2002 Release 1. We strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements");
            }
            IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", qbXMLMajorVer, qbXMLMinorVer);
            return requestMsgSet;
        }
        
        void SaveXML(string xmlstring)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                StreamWriter sr = new StreamWriter(saveFileDialog1.FileName);
                sr.Write(xmlstring);
                sr.Close();
            }
        }

        private void btnOpenFile_Reset_Click(object sender, System.EventArgs e)
        {
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;
            checkBox7.Checked = false;
            checkBox8.Checked = false;
            checkBox9.Checked = false;
            checkBox10.Checked = false;
            checkBox11.Checked = false;
            checkBox12.Checked = false;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = @"C:\Users\Elizabeth.Earl\source\repos\ConsoleApp2\ConsoleApp2\";
                openFileDialog.Filter = "xml files (*.xml)|*.xml";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    fileName = openFileDialog.FileName;
                    startErrorChecking(sender, e);
                }
            }
        }

        private void startErrorChecking(object sender, EventArgs e)
        {
            dt = new DataTable();
            dtTemp2 = new DataTable();
            XDocument doc = XDocument.Load(fileName);
            string matchFirst = @"^[1-2][0-9][0-9][0-9][0-9][0-9]00";
            foreach (XElement bom in doc.Descendants("bom"))
            {
                string nameData = bom.Attribute("name").Value;
                string docPath = bom.Attribute("document_path").Value;
                string input = System.Text.RegularExpressions.Regex.Match(nameData, matchFirst).Value;
                string cut = docPath.Substring(Math.Max(0, docPath.Length - 15), 8);

                SqlDataAdapter dataAdapter = new SqlDataAdapter();
                String connectionString = @"Data Source=SQLSERVER\ITEMCODE;Initial Catalog=dat8121;Integrated Security=True";
                DataTable dtReturnValue = new DataTable();
                DataTable dtTemp = new DataTable();
                int dtReturnValueCount = dtReturnValue.Rows.Count;
                int dtTempCount = dtTemp.Rows.Count;
                DataSet DataSet1 = new DataSet();
                DataView DataView1 = new DataView();
                //DataTable dtTemp2 = new DataTable();
                dataAdapter = new SqlDataAdapter("SELECT  [ItemCode], [Description]" +
                    " FROM[dat8121].[dbo].[I_ItemCode]" +
                    " where TRY_CAST(ItemCode as nvarchar) = '" + input +
                    "' OR TRY_CAST(ItemCode as nvarchar) = '" + cut + "'", connectionString);
                dataAdapter.Fill(dtTemp2);
                dtReturnValue = dtTemp2.Clone();

                bool isTrue = dtTemp2.Rows.Count > 0;
                if (isTrue == true)
                {
                    string fromData = dtTemp2.Rows[0][0].ToString();
                    if (input.Equals(cut) == true)
                    {
                        if (input.Equals(fromData) == true)
                        {
                            textBox1.Text = (dtTemp2.Rows[0][0].ToString() + "\r\n" + dtTemp2.Rows[0][1].ToString());
                            checkBox1.Checked = true;
                            columnCorrectOrder(sender, e, dtTemp2);
                        }
                    }
                    else if (String.Equals(input, cut) == false)
                    {
                        ifError = true;
                        columnCorrectOrder(sender, e, ifError, input, cut, dtTemp2);
                    }
                }
                else
                {
                    ifError = true;
                    columnCorrectOrder(sender, e, ifError, input, cut, dtTemp2);
                }
            }
        }

        private void columnCorrectOrder(object sender, EventArgs e, bool ifError, string input, string cut, DataTable dtTemp2)
        {
            if (ifError == true)
            {
                if (String.Equals(input, cut) == false)
                {
                    checkBox2.Checked = true;
                    columnCorrectOrder(sender, e, dtTemp2);
                }
                else
                {
                    checkBox3.Checked = true;
                    columnCorrectOrder(sender, e, dtTemp2);
                }
            }
        }

        private void columnCorrectOrder(object sender, EventArgs e, DataTable dtTemp2)
        {
            XDocument doc = XDocument.Load(fileName);
            Dictionary<int, string> openWith = new Dictionary<int, string>();
            foreach (XElement bom in doc.Descendants("bomcol"))
            {
                int currentColumn = Int32.Parse(bom.Attribute("col_no").Value);
                string colHeader = bom.Attribute("name").Value ?? "n/a";
                openWith.Add(currentColumn, colHeader);
            }

            var lines = openWith.Select(kv => kv.Key + " " + kv.Value.ToString());
            string itemNoMatching = @"(?i)ITEM NO*|ITEMNO*|ITMN*|ITM N*";
            string itemCodeMatching = @"(?i)ITEM CO*|ITEMCO*|ITMC*|ITM C*";
            string partNoMatching = @"(?i)PART*|PRT*";
            string descriptionMatching = @"(?i)DESC*";
            string qtyMatching = @"(?i)QTY*|QU*|Qt*";
            int itemNoCol, itemCodeCol, partNumberCol, descriptionCol, qtyCol;

            if (Regex.IsMatch(openWith[0], itemNoMatching))
            {
                itemNoCol = openWith.Keys.ElementAt(0);
                checkBox4.Checked = true;
                dt.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").Where(x => (int)x.Attribute("col_no") == itemNoCol).FirstOrDefault() == null ?
                      null : "ITEM NO.");
            }
            else
            {
                dt.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").Where(x => (int)x.Attribute("col_no") == 0).FirstOrDefault() == null ?
                     null : "ITEM NO.");
            }
            if (Regex.IsMatch(openWith[1], itemCodeMatching))
            {
                itemCodeCol = openWith.Keys.ElementAt(1);
                checkBox5.Checked = true;
                dt.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").Where(x => (int)x.Attribute("col_no") == itemCodeCol).FirstOrDefault() == null ?
                      null : "ITEMCODE");
            }
            else
            {
                dt.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").Where(x => (int)x.Attribute("col_no") == 1).FirstOrDefault() == null ?
                     null : "ITEMCODE");
            }
            if (Regex.IsMatch(openWith[2], partNoMatching))
            {
                partNumberCol = openWith.Keys.ElementAt(2);
                checkBox6.Checked = true;
                dt.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").Where(x => (int)x.Attribute("col_no") == partNumberCol).FirstOrDefault() == null ?
                      null : "PART NUMBER");
            }
            else
            {
                dt.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").Where(x => (int)x.Attribute("col_no") == 2).FirstOrDefault() == null ?
                     null : "PART NUMBER");
            }
            if (Regex.IsMatch(openWith[3], descriptionMatching))
            {
                descriptionCol = openWith.Keys.ElementAt(3);
                checkBox7.Checked = true;
                dt.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").Where(x => (int)x.Attribute("col_no") == descriptionCol).FirstOrDefault() == null ?
                      null : "DESCRIPTION");
            }
            else
            {
                dt.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").Where(x => (int)x.Attribute("col_no") == 3).FirstOrDefault() == null ?
                     null : "DESCRIPTION");
            }
            if (Regex.IsMatch(openWith[4], qtyMatching))
            {
                qtyCol = openWith.Keys.ElementAt(4);
                checkBox8.Checked = true;
                dt.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").Where(x => (int)x.Attribute("col_no") == qtyCol).FirstOrDefault() == null ?
                      null : "QTY.");
            }
            else
            {
                dt.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").Where(x => (int)x.Attribute("col_no") == 4).FirstOrDefault() == null ?
                     null : "QTY.");
            }
            addRows(dt, doc, openWith, dtTemp2);
        }

        private void addRows(DataTable dt, XDocument doc, Dictionary<int, string> openWith, DataTable dtTemp2)
        {
            bool error = false;
            foreach (XElement bomrow in doc.Descendants("bomrow"))
            {
                string itemNo = @"^\d+$";
                string itemCode = @"^[1-2][0-9][0-9][0-9][0-9][0-9]00";
                string qty = @"^\d+$";
                string iNoGiven = (string)bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                            y => y.Value.ToUpper().Contains("ITEM NO")).Key
                            ).FirstOrDefault().Attribute("value");
                string iCODEgiven = (string)bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                            y => y.Value.ToUpper().Contains("CODE") || y.Value.ToUpper().Contains("CD")).Key
                            ).FirstOrDefault().Attribute("value");
                string qtyGiven = (string)bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                            y => y.Value.ToUpper().Contains("Q")).Key
                            ).FirstOrDefault().Attribute("value");
                bool numberTest = Regex.IsMatch(iNoGiven, itemNo);
                bool codeTest = Regex.IsMatch(iCODEgiven, itemCode);
                bool qtyTest = Regex.IsMatch(qtyGiven, qty);
                if (numberTest == false)
                {
                    checkBox10.Checked = true;
                    error = true;
                }
                else if (codeTest == false)
                {
                    checkBox10.Checked = true;
                    error = true;
                }
                else if (qtyTest == false)
                {
                    checkBox10.Checked = true;
                    error = true;
                }
            }

            if (error == false)
            {
                createATable(dt, doc, openWith, dtTemp2);
            }
            else
            {
                createATableWithErrors(dt, doc, openWith, dtTemp2);
            }
        }

        private void createATableWithErrors(DataTable dt, XDocument doc, Dictionary<int, string> openWith, DataTable dtTemp2)
        {
            foreach (XElement bomrow in doc.Descendants("bomrow"))
            {
                dt.Rows.Add(
                bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                    y => y.Value.ToUpper().Contains("ITEM NO")).Key
                    ).FirstOrDefault() == null ?
                    null : (string)bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                        y => y.Value.ToUpper().Contains("ITEM NO")).Key
                        ).FirstOrDefault().Attribute("value") ?? null,

                bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                    y => y.Value.ToUpper().Contains("CODE") || y.Value.ToUpper().Contains("CD")).Key
                    ).FirstOrDefault() == null ?
                    null : (string)bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                        y => y.Value.ToUpper().Contains("CODE") || y.Value.ToUpper().Contains("CD")).Key
                        ).FirstOrDefault().Attribute("value") ?? null,

                bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                    y => y.Value.ToUpper().Contains("PART")).Key
                    ).FirstOrDefault() == null ?
                    null : (string)bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                        y => y.Value.ToUpper().Contains("PART")).Key
                        ).FirstOrDefault().Attribute("value") ?? "N/A",

                bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                    y => y.Value.ToUpper().Contains("DES")).Key
                    ).FirstOrDefault() == null ?
                    null : (string)bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                        y => y.Value.ToUpper().Contains("DES")).Key
                        ).FirstOrDefault().Attribute("value") ?? "N/A",

                bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                    y => y.Value.ToUpper().Contains("Q")).Key
                    ).FirstOrDefault() == null ?
                    null : (string)bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                        y => y.Value.ToUpper().Contains("Q")).Key
                        ).FirstOrDefault().Attribute("value") ?? null
                );
            }
            string[] firstRow = dt.AsEnumerable().Select(r => r.Field<string>("ITEM NO.")).ToArray();
            string has = String.Join(",", firstRow);
            int starting = 1;
            IEnumerable<int> totalColumns = Enumerable.Range(starting, Int32.Parse(firstRow.Last()));
            string shouldBe = String.Join(",", totalColumns);
            showTableErrors(dt, has, shouldBe, dtTemp2);
        }

        private void createATable(DataTable dt, XDocument doc, Dictionary<int, string> openWith, DataTable dtTemp2)
        {
            foreach (XElement bomrow in doc.Descendants("bomrow"))
            {
                dt.Rows.Add(
                bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                    y => y.Value.ToUpper().Contains("ITEM NO")).Key
                    ).FirstOrDefault() == null ?
                    null : (int?)bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                        y => y.Value.ToUpper().Contains("ITEM NO")).Key
                        ).FirstOrDefault().Attribute("value") ?? null,

                bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                    y => y.Value.ToUpper().Contains("CODE") || y.Value.ToUpper().Contains("CD")).Key
                    ).FirstOrDefault() == null ?
                    null : (int?)bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                        y => y.Value.ToUpper().Contains("CODE") || y.Value.ToUpper().Contains("CD")).Key
                        ).FirstOrDefault().Attribute("value") ?? null,

                bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                    y => y.Value.ToUpper().Contains("PART")).Key
                    ).FirstOrDefault() == null ?
                    null : (string)bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                        y => y.Value.ToUpper().Contains("PART")).Key
                        ).FirstOrDefault().Attribute("value") ?? "N/A",

                bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                    y => y.Value.ToUpper().Contains("DES")).Key
                    ).FirstOrDefault() == null ?
                    null : (string)bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                        y => y.Value.ToUpper().Contains("DES")).Key
                        ).FirstOrDefault().Attribute("value") ?? "N/A",

                bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                    y => y.Value.ToUpper().Contains("Q")).Key
                    ).FirstOrDefault() == null ?
                    null : (int?)bomrow.Elements("bomcell").Where(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                        y => y.Value.ToUpper().Contains("Q")).Key
                        ).FirstOrDefault().Attribute("value") ?? null
                );
            }
            checkBox9.Checked = true;
            string[] firstRow = dt.AsEnumerable().Select(r => r.Field<string>("ITEM NO.")).ToArray();
            string has = String.Join(",", firstRow);
            int starting = 1;
            IEnumerable<int> totalColumns = Enumerable.Range(starting, Int32.Parse(firstRow.Last()));
            string shouldBe = String.Join(",", totalColumns);
            showTableErrors(dt, has, shouldBe, dtTemp2);
        }

        private void showTableErrors(DataTable dt, string has, string shouldBe, DataTable dtTemp2)
        {
            if (has == shouldBe)
            {
                checkBox11.Checked = true;
                ending(dt, dtTemp2);
            }
            else
            {
                checkBox12.Checked = true;
                ending(dt, dtTemp2);
            }

        }

        private void ending(DataTable dt, DataTable dtTemp2)
        {
            dataGridView1.DataSource = dt;
            if (checkBox1.Checked == true && checkBox9.Checked == true && checkBox11.Checked == true)
            {
                dataGridView1.DataSource = dt;
            }
            else
            {
                tbProgramLog.AppendText(Environment.NewLine + "Cannot import a table with errors!");
            }
        }
       
        private void QBFC_AddInvoice()
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                //Create the session Manager object
                sessionManager = new QBSessionManager();
                //Create the message set request object to hold our request
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                getValuesForQB(requestMsgSet, dt, dtTemp2);
                //Connect to QuickBooks and begin a session
                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;
                //Send the request and get the response from QuickBooks
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                //End the session and close the connection to QuickBooks
                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;
                WalkItemInventoryAssemblyQueryRs(responseMsgSet);
            }
            catch (Exception e)
            {
               tbProgramLog.AppendText(e.Message);
                if (sessionBegun)
                {
                    sessionManager.EndSession();
                }
                if (connectionOpen)
                {
                    sessionManager.CloseConnection();
                }
            }
        }

        //private void getValuesForQB(IMsgSetRequest requestMsgSet)
        //private void getValuesForQB(out IMsgSetRequest requestMsgSet, DataTable dt, DataTable dtTemp2)
        void getValuesForQB(IMsgSetRequest requestMsgSet, DataTable dt, DataTable dtTemp2)
        {
            //MessageBox.Show(dtTemp2.Rows[0][0].ToString());//itemcode full bom
            //MessageBox.Show(dtTemp2.Rows[0][1].ToString());//desc full bom
            //MessageBox.Show(dt.Rows[0][0].ToString());//itemno for assembly
            //MessageBox.Show(dt.Rows[0][1].ToString());//itemcode for assembly
            //MessageBox.Show(dt.Rows[0][4].ToString());//qty for assembly
            IItemInventoryAssemblyQuery ItemInventoryAssemblyQueryRq = requestMsgSet.AppendItemInventoryAssemblyQueryRq();
            IListWithClassFilter listWithClassFilter = ItemInventoryAssemblyQueryRq.ORListQueryWithOwnerIDAndClass.ListWithClassFilter;
            //Set field value for MatchCriterion
            listWithClassFilter.ORNameFilter.NameFilter.MatchCriterion.SetValue(ENMatchCriterion.mcContains);
            //Set field value for Name
            ItemInventoryAssemblyQueryRq.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ORNameFilter.NameFilter.Name.SetValue(dtTemp2.Rows[0][0].ToString());
            //Set field value for Name with > 1 result
            //ItemInventoryAssemblyQueryRq.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ORNameFilter.NameFilter.Name.SetValue("2000");

            //IItemInventoryAssemblyMod ItemInventoryAssemblyModRq = requestMsgSet.AppendItemInventoryAssemblyModRq();
            //ItemInventoryAssemblyModRq.ListID.SetValue("80013E9B-1527005428");
            //ItemInventoryAssemblyModRq.EditSequence.SetValue("1543594308");
            //ItemInventoryAssemblyModRq.SalesDesc.SetValue("Assy, Yoke Arm End Lifting");
            //ItemInventoryAssemblyModRq.PurchaseDesc.SetValue(dtTemp2.Rows[0][1].ToString());

            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    IItemInventoryAssemblyLine ItemInventoryAssemblyLine1 = ItemInventoryAssemblyModRq.ORItemInventoryAssemblyLine.ItemInventoryAssemblyLineList.Append();
            //    ItemInventoryAssemblyLine1.ItemInventoryRef.ListID.SetValue("80000A97-1511318515");
            //    ItemInventoryAssemblyLine1.ItemInventoryRef.FullName.SetValue(dt.Rows[i][1].ToString());
            //    ItemInventoryAssemblyLine1.Quantity.SetValue(Convert.ToDouble(dt.Rows[i][4]));
            //}
        }

        void WalkItemInventoryAssemblyQueryRs(IMsgSetResponse responseMsgSet)
        {
            if (responseMsgSet == null) return;
            IResponseList responseList = responseMsgSet.ResponseList;
            if (responseList == null) return;
            //if we sent only one request, there is only one response, we'll walk the list for this sample
            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);
                //MessageBox.Show(response.StatusCode.ToString() + ": " + response.StatusMessage.ToString(), "Error code");
                tbProgramLog.AppendText(response.StatusCode.ToString() + ": " + response.StatusMessage.ToString());
                //check the status code of the response, 0=ok, >0 is warning
                if (response.StatusCode >= 0)
                {
                    if (response.StatusCode == 1)
                    {
                        
                            tbProgramLog.AppendText(Environment.NewLine + "Call add method");
                            //addItem(requestMsgSet);
                    }
                    else if (response.StatusCode == 0)
                    {
                        //the request-specific response is in the details, make sure we have some
                        if (response.Detail != null)
                        {
                            //make sure the response is the type we're expecting
                            ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                            if (responseType == ENResponseType.rtItemInventoryAssemblyQueryRs)
                            {
                                //upcast to more specific type here, this is safe because we checked with response.Type check above
                                IItemInventoryAssemblyRetList ItemInventoryAssemblyRetList = (IItemInventoryAssemblyRetList)response.Detail;
                                WalkItemInventoryAssemblyRet(ItemInventoryAssemblyRetList);
                            }
                        }
                    }
                }
            }
        }

        void modifyItem(IMsgSetRequest requestMsgSet)
        {
            IItemInventoryAssemblyMod ItemInventoryAssemblyModRq = requestMsgSet.AppendItemInventoryAssemblyModRq();
            ItemInventoryAssemblyModRq.ListID.SetValue(dtTemp2.Rows[0][0].ToString());
            ItemInventoryAssemblyModRq.EditSequence.SetValue("1535167356");
            ItemInventoryAssemblyModRq.SalesDesc.SetValue("Assy, Yoke Arm End Lifting");
            ItemInventoryAssemblyModRq.PurchaseDesc.SetValue(dtTemp2.Rows[0][1].ToString());

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IItemInventoryAssemblyLine ItemInventoryAssemblyLine1 = ItemInventoryAssemblyModRq.ORItemInventoryAssemblyLine.ItemInventoryAssemblyLineList.Append();
                ItemInventoryAssemblyLine1.ItemInventoryRef.ListID.SetValue("");
                ItemInventoryAssemblyLine1.ItemInventoryRef.FullName.SetValue(dt.Rows[i][1].ToString());
                ItemInventoryAssemblyLine1.Quantity.SetValue(Convert.ToDouble(dt.Rows[i][4]));
            }
            tbProgramLog.AppendText(Environment.NewLine + "HELLO WORLD");
        }

        void addItem(IMsgSetRequest requestMsgSet)
        {
            IItemInventoryAssemblyAdd ItemInventoryAssemblyAddRq = requestMsgSet.AppendItemInventoryAssemblyAddRq();
            ItemInventoryAssemblyAddRq.Name.SetValue(dtTemp2.Rows[0][0].ToString());
            ItemInventoryAssemblyAddRq.SalesDesc.SetValue("Assy, Yoke Arm End Lifting");
            ItemInventoryAssemblyAddRq.PurchaseDesc.SetValue(dtTemp2.Rows[0][1].ToString());

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IItemInventoryAssemblyLine ItemInventoryAssemblyLine1 = ItemInventoryAssemblyAddRq.ItemInventoryAssemblyLineList.Append();
                ItemInventoryAssemblyLine1.ItemInventoryRef.ListID.SetValue(dt.Rows[i][0].ToString());
                ItemInventoryAssemblyLine1.ItemInventoryRef.FullName.SetValue(dt.Rows[i][1].ToString());
                ItemInventoryAssemblyLine1.Quantity.SetValue(Convert.ToDouble(dt.Rows[i][4]));
            }
            
        }

        void getEditSequence(IItemInventoryAssemblyRetList ItemInventoryAssemblyRetList)
        {
            WalkItemInventoryAssemblyRet(ItemInventoryAssemblyRetList);
        }
        void WalkItemInventoryAssemblyRet(IItemInventoryAssemblyRetList ItemInventoryAssemblyRetList)
        {
            if (ItemInventoryAssemblyRetList == null) return;
            for (int x = 0; x < ItemInventoryAssemblyRetList.Count; x++)
            {
                IItemInventoryAssemblyRet ItemInventoryAssemblyRet = ItemInventoryAssemblyRetList.GetAt(x);
                string sequence = (string)ItemInventoryAssemblyRet.EditSequence.GetValue();
                string listID = (string)ItemInventoryAssemblyRet.ListID.GetValue();
                tbProgramLog.AppendText(Environment.NewLine + "Edit sequence: " + sequence + Environment.NewLine + "List ID: " + listID);
            }
           // modifyItem(requestMsgSet);
            //string icode = dtTemp2.Rows[0][0].ToString();
            //tbProgramLog.AppendText(Environment.NewLine + "No. of matches to " + icode +": "+ ItemInventoryAssemblyRetList.Count.ToString());
        }

        private void btn1_Send_Click_1(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true && checkBox9.Checked == true && checkBox11.Checked == true)
            {
                QBFC_AddInvoice();
            }
            else
            {
                tbProgramLog.AppendText(Environment.NewLine + "Cannot import a table with errors!");
            }
        } 
    }
}