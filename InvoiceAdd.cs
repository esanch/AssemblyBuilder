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
    public class Frm1InvoiceAdd : System.Windows.Forms.Form
    {
        private System.ComponentModel.Container components = null;
        private System.Windows.Forms.Button btn1_Send;
        private System.Windows.Forms.Button btn2_Exit;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Button btnOpenFile_Reset;
        private DataGridView dataGridView1;
        public int CurrentRow = 0;

        DataTable secondLevelTbl = new DataTable();
        DataTable topLevelTbl = new DataTable();
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

        public Frm1InvoiceAdd()
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
            this.btn1_Send.Location = new System.Drawing.Point(394, 717);
            this.btn1_Send.Name = "btn1_Send";
            this.btn1_Send.Size = new System.Drawing.Size(80, 32);
            this.btn1_Send.TabIndex = 57;
            this.btn1_Send.Text = "Send";
            this.btn1_Send.UseVisualStyleBackColor = false;
            this.btn1_Send.Click += new System.EventHandler(this.Btn1_Send_Click_1);
            // 
            // btn2_Exit
            // 
            this.btn2_Exit.BackColor = System.Drawing.SystemColors.Control;
            this.btn2_Exit.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn2_Exit.Location = new System.Drawing.Point(482, 717);
            this.btn2_Exit.Name = "btn2_Exit";
            this.btn2_Exit.Size = new System.Drawing.Size(75, 32);
            this.btn2_Exit.TabIndex = 58;
            this.btn2_Exit.Text = "Exit";
            this.btn2_Exit.UseVisualStyleBackColor = false;
            this.btn2_Exit.Click += new System.EventHandler(this.Btn2_Exit_Click);
            // 
            // btnOpenFile_Reset
            // 
            this.btnOpenFile_Reset.BackColor = System.Drawing.SystemColors.Control;
            this.btnOpenFile_Reset.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOpenFile_Reset.Location = new System.Drawing.Point(441, 676);
            this.btnOpenFile_Reset.Name = "btnOpenFile_Reset";
            this.btnOpenFile_Reset.Size = new System.Drawing.Size(80, 32);
            this.btnOpenFile_Reset.TabIndex = 63;
            this.btnOpenFile_Reset.Text = "Open File";
            this.btnOpenFile_Reset.UseVisualStyleBackColor = false;
            this.btnOpenFile_Reset.Click += new System.EventHandler(this.BtnOpenFile_Reset_Click);
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
            this.dataGridView1.Location = new System.Drawing.Point(7, 124);
            this.dataGridView1.Name = "dataGridView1";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridView1.Size = new System.Drawing.Size(563, 483);
            this.dataGridView1.TabIndex = 64;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(135, 627);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(44, 17);
            this.checkBox1.TabIndex = 65;
            this.checkBox1.Text = "Yes";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(244, 627);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(40, 17);
            this.checkBox2.TabIndex = 66;
            this.checkBox2.Text = "No";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // checkBox3
            // 
            this.checkBox3.AutoSize = true;
            this.checkBox3.Location = new System.Drawing.Point(330, 627);
            this.checkBox3.Name = "checkBox3";
            this.checkBox3.Size = new System.Drawing.Size(119, 17);
            this.checkBox3.TabIndex = 67;
            this.checkBox3.Text = "ItemCode not found";
            this.checkBox3.UseVisualStyleBackColor = true;
            // 
            // checkBox4
            // 
            this.checkBox4.AutoSize = true;
            this.checkBox4.Location = new System.Drawing.Point(135, 650);
            this.checkBox4.Name = "checkBox4";
            this.checkBox4.Size = new System.Drawing.Size(86, 17);
            this.checkBox4.TabIndex = 68;
            this.checkBox4.Text = "Item Number";
            this.checkBox4.UseVisualStyleBackColor = true;
            // 
            // checkBox5
            // 
            this.checkBox5.AutoSize = true;
            this.checkBox5.Location = new System.Drawing.Point(244, 650);
            this.checkBox5.Name = "checkBox5";
            this.checkBox5.Size = new System.Drawing.Size(71, 17);
            this.checkBox5.TabIndex = 69;
            this.checkBox5.Text = "ItemCode";
            this.checkBox5.UseVisualStyleBackColor = true;
            // 
            // checkBox6
            // 
            this.checkBox6.AutoSize = true;
            this.checkBox6.Location = new System.Drawing.Point(330, 650);
            this.checkBox6.Name = "checkBox6";
            this.checkBox6.Size = new System.Drawing.Size(85, 17);
            this.checkBox6.TabIndex = 70;
            this.checkBox6.Text = "Part Number";
            this.checkBox6.UseVisualStyleBackColor = true;
            // 
            // checkBox7
            // 
            this.checkBox7.AutoSize = true;
            this.checkBox7.Location = new System.Drawing.Point(416, 650);
            this.checkBox7.Name = "checkBox7";
            this.checkBox7.Size = new System.Drawing.Size(79, 17);
            this.checkBox7.TabIndex = 71;
            this.checkBox7.Text = "Description";
            this.checkBox7.UseVisualStyleBackColor = true;
            // 
            // checkBox9
            // 
            this.checkBox9.AutoSize = true;
            this.checkBox9.Location = new System.Drawing.Point(135, 673);
            this.checkBox9.Name = "checkBox9";
            this.checkBox9.Size = new System.Drawing.Size(44, 17);
            this.checkBox9.TabIndex = 73;
            this.checkBox9.Text = "Yes";
            this.checkBox9.UseVisualStyleBackColor = true;
            // 
            // checkBox10
            // 
            this.checkBox10.AutoSize = true;
            this.checkBox10.Location = new System.Drawing.Point(244, 673);
            this.checkBox10.Name = "checkBox10";
            this.checkBox10.Size = new System.Drawing.Size(40, 17);
            this.checkBox10.TabIndex = 74;
            this.checkBox10.Text = "No";
            this.checkBox10.UseVisualStyleBackColor = true;
            // 
            // checkBox11
            // 
            this.checkBox11.AutoSize = true;
            this.checkBox11.Location = new System.Drawing.Point(135, 698);
            this.checkBox11.Name = "checkBox11";
            this.checkBox11.Size = new System.Drawing.Size(44, 17);
            this.checkBox11.TabIndex = 75;
            this.checkBox11.Text = "Yes";
            this.checkBox11.UseVisualStyleBackColor = true;
            // 
            // checkBox12
            // 
            this.checkBox12.AutoSize = true;
            this.checkBox12.Location = new System.Drawing.Point(244, 698);
            this.checkBox12.Name = "checkBox12";
            this.checkBox12.Size = new System.Drawing.Size(40, 17);
            this.checkBox12.TabIndex = 76;
            this.checkBox12.Text = "No";
            this.checkBox12.UseVisualStyleBackColor = true;
            // 
            // checkBox8
            // 
            this.checkBox8.AutoSize = true;
            this.checkBox8.Location = new System.Drawing.Point(502, 650);
            this.checkBox8.Name = "checkBox8";
            this.checkBox8.Size = new System.Drawing.Size(65, 17);
            this.checkBox8.TabIndex = 72;
            this.checkBox8.Text = "Quantity";
            this.checkBox8.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 631);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 13);
            this.label1.TabIndex = 77;
            this.label1.Text = "ICodes match?";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(4, 654);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(124, 13);
            this.label2.TabIndex = 78;
            this.label2.Text = "Columns in correct order:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(4, 676);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(73, 13);
            this.label3.TabIndex = 79;
            this.label3.Text = "Correct input?";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(4, 701);
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
            this.textBox1.Size = new System.Drawing.Size(563, 78);
            this.textBox1.TabIndex = 83;
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(0, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(100, 23);
            this.label5.TabIndex = 86;
            // 
            // tbProgramLog
            // 
            this.tbProgramLog.Location = new System.Drawing.Point(587, 40);
            this.tbProgramLog.Multiline = true;
            this.tbProgramLog.Name = "tbProgramLog";
            this.tbProgramLog.Size = new System.Drawing.Size(483, 709);
            this.tbProgramLog.TabIndex = 85;
            // 
            // Frm1InvoiceAdd
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.ClientSize = new System.Drawing.Size(1082, 755);
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
            this.Name = "Frm1InvoiceAdd";
            this.Text = "Add an Item Inventory to QuickBooks";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        [STAThread]
        static void Main()
        {
            Application.Run(new Frm1InvoiceAdd());
        }

        private void Btn2_Exit_Click(object sender, System.EventArgs e)
        {
            Dispose();
        }

        //private double QBFCLatestVersion(QBSessionManager SessionManager)
        //{
        //    // Use oldest version to ensure that this application work with any QuickBooks (US)
        //    IMsgSetRequest msgset = SessionManager.CreateMsgSetRequest("US", 1, 0);
        //    msgset.AppendHostQueryRq();
        //    IMsgSetResponse QueryResponse = SessionManager.DoRequests(msgset);

        //    // The response list contains only one response, which corresponds to our single HostQuery request
        //    IResponse response = QueryResponse.ResponseList.GetAt(0);

        //    // Please refer to QBFC Developers Guide for details on why "as" clause was used to link this derrived class to its base class
        //    IHostRet HostResponse = response.Detail as IHostRet;
        //    IBSTRList supportedVersions = HostResponse.SupportedQBXMLVersionList as IBSTRList;

        //    int i;
        //    double vers;
        //    double LastVers = 0;
        //    string svers = null;
        //    for (i = 0; i <= supportedVersions.Count - 1; i++)
        //    {
        //        svers = supportedVersions.GetAt(i);
        //        vers = Convert.ToDouble(svers);
        //        if (vers > LastVers)
        //        {
        //            LastVers = vers;
        //        }
        //    }
        //    return LastVers;
        //}

        //public IMsgSetRequest getLatestMsgSetRequest(QBSessionManager sessionManager)
        //{
        //    double supportedVersion = QBFCLatestVersion(sessionManager);
        //    short qbXMLMajorVer = 0;
        //    short qbXMLMinorVer = 0;

        //    if (supportedVersion >= 6.0)
        //    {
        //        qbXMLMajorVer = 6;
        //        qbXMLMinorVer = 0;
        //    }
        //    else if (supportedVersion >= 5.0)
        //    {
        //        qbXMLMajorVer = 5;
        //        qbXMLMinorVer = 0;
        //    }
        //    else if (supportedVersion >= 4.0)
        //    {
        //        qbXMLMajorVer = 4;
        //        qbXMLMinorVer = 0;
        //    }
        //    else if (supportedVersion >= 3.0)
        //    {
        //        qbXMLMajorVer = 3;
        //        qbXMLMinorVer = 0;
        //    }
        //    else if (supportedVersion >= 2.0)
        //    {
        //        qbXMLMajorVer = 2;
        //        qbXMLMinorVer = 0;
        //    }
        //    else if (supportedVersion >= 1.1)
        //    {
        //        qbXMLMajorVer = 1;
        //        qbXMLMinorVer = 1;
        //    }
        //    else
        //    {
        //        qbXMLMajorVer = 1;
        //        qbXMLMinorVer = 0;
        //        tbProgramLog.AppendText(Environment.NewLine + "It seems that you are running QuickBooks 2002 Release 1. We strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements");
        //    }
        //    IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", qbXMLMajorVer, qbXMLMinorVer);
        //    return requestMsgSet;
        //}

        //void SaveXML(string xmlstring)
        //{
        //    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
        //    {
        //        StreamWriter sr = new StreamWriter(saveFileDialog1.FileName);
        //        sr.Write(xmlstring);
        //        sr.Close();
        //    }
        //}

        private void BtnOpenFile_Reset_Click(object sender, System.EventArgs e)
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
                //openFileDialog.InitialDirectory = @"C:\Users\Elizabeth.Earl\source\repos\ConsoleApp2\ConsoleApp2\";
                openFileDialog.Filter = "xml files (*.xml)|*.xml";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    fileName = openFileDialog.FileName;
                    StartErrorChecking(sender, e);
                }
            }
        }

        private void StartErrorChecking(object sender, EventArgs e)
        {
            secondLevelTbl = new DataTable();
            topLevelTbl = new DataTable();
            XDocument doc = XDocument.Load(fileName);
            string matchFirst = @"^[1-2][0-9][0-9][0-9][0-9][0-9]00";
            foreach (XElement bom in doc.Descendants("bom"))
            {
                string nameData = bom.Attribute("name").Value;
                string docPath = bom.Attribute("document_path").Value;
                string input = System.Text.RegularExpressions.Regex.Match(nameData, matchFirst).Value;
                string cut = docPath.Substring(Math.Max(0, docPath.Length - 15), 8);

                //SqlDataAdapter dataAdapter = new SqlDataAdapter();
                string connectionString = @"Data Source=SQLSERVER\ITEMCODE;Initial Catalog=dat8121;Integrated Security=True";
                DataTable dtReturnValue = new DataTable();
                DataTable dtTemp = new DataTable();
                int dtReturnValueCount = dtReturnValue.Rows.Count;
                int dtTempCount = dtTemp.Rows.Count;
                //DataSet DataSet1 = new DataSet();
                //DataView DataView1 = new DataView();
                //DataTable topLevelTbl = new DataTable();
                SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT a.ItemCode, a.Description, c.[IncomeAccountRefListID], c.[COGSAccountRefListID], c.[AssetAccountRefListID]" +
                " FROM [dat8121].[dbo].[v_ItemCode_QB] b" +
                " RIGHT JOIN [dat8121].[dbo].[I_ItemCode] a ON a.itemcode = b.ItemCode" +
                " RIGHT JOIN [QODBC].[dbo].[Tbl_Item] c ON a.ItemCode = TRY_CAST(c.fullname AS int) AND b.ItemCode = TRY_CAST(c.fullname AS int)" +
                " WHERE a.itemcode = '" + input +
                "' OR a.itemcode = '" + cut + "'"
                    /*
                    "SELECT  [ItemCode], [Description]" +
                     " FROM[dat8121].[dbo].[I_ItemCode]" +
                     " where TRY_CAST(ItemCode as nvarchar) = '" + input +
                     "' OR TRY_CAST(ItemCode as nvarchar) = '" + cut + "'"
                     */
                    , connectionString);
                /*select a.itemcode,a.description,b.itemcode as qbname from v_itemcode_qb b
right join i_itemcode a on a.itemcode=b.itemcode
where a.itemcode =25000000*/

                dataAdapter.Fill(topLevelTbl);
                dtReturnValue = topLevelTbl.Clone();

                bool isTrue = topLevelTbl.Rows.Count > 0;
                if (isTrue == true)
                {
                    string fromData = topLevelTbl.Rows[0][0].ToString();
                    if (input.Equals(cut) == true)
                    {
                        if (input.Equals(fromData) == true)
                        {
                            textBox1.Text = (topLevelTbl.Rows[0][0].ToString() + "\r\n" + topLevelTbl.Rows[0][1].ToString() + "\r\n" +
                                             topLevelTbl.Rows[0][2].ToString() + "\r\n" + topLevelTbl.Rows[0][3].ToString() + "\r\n" +
                                             topLevelTbl.Rows[0][4].ToString());
                            checkBox1.Checked = true;
                            ColumnCorrectOrder(sender, e, topLevelTbl);
                        }
                    }
                    else if (String.Equals(input, cut) == false)
                    {
                        ifError = true;
                        ColumnCorrectOrder(sender, e, ifError, input, cut, topLevelTbl);
                    }
                }
                else
                {
                    ifError = true;
                    ColumnCorrectOrder(sender, e, ifError, input, cut, topLevelTbl);
                }
            }
        }

        private void ColumnCorrectOrder(object sender, EventArgs e, bool ifError, string input, string cut, DataTable topLevelTbl)
        {
            if (ifError == true)
            {
                if (String.Equals(input, cut) == false)
                {
                    checkBox2.Checked = true;
                    ColumnCorrectOrder(sender, e, topLevelTbl);
                }
                else
                {
                    checkBox3.Checked = true;
                    ColumnCorrectOrder(sender, e, topLevelTbl);
                }
            }
        }

        private void ColumnCorrectOrder(object sender, EventArgs e, DataTable topLevelTbl)
        {
            XDocument doc = XDocument.Load(fileName);
            Dictionary<int, string> openWith = new Dictionary<int, string>();
            foreach (XElement bom in doc.Descendants("bomcol"))
            {
                int currentColumn = Int32.Parse(bom.Attribute("col_no")?.Value);
                string colHeader = bom.Attribute("name")?.Value ?? "n/a";
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
                secondLevelTbl.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").FirstOrDefault(x => (int)x.Attribute("col_no") == itemNoCol) == null ?
                      null : "ITEM NO.");
            }
            else
            {
                secondLevelTbl.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").FirstOrDefault(x => (int)x.Attribute("col_no") == 0) == null ?
                     null : "ITEM NO.");
            }
            if (Regex.IsMatch(openWith[1], itemCodeMatching))
            {
                itemCodeCol = openWith.Keys.ElementAt(1);
                checkBox5.Checked = true;
                secondLevelTbl.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").FirstOrDefault(x => (int)x.Attribute("col_no") == itemCodeCol) == null ?
                      null : "ITEMCODE");
            }
            else
            {
                secondLevelTbl.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").FirstOrDefault(x => (int)x.Attribute("col_no") == 1) == null ?
                     null : "ITEMCODE");
            }
            if (Regex.IsMatch(openWith[2], partNoMatching))
            {
                partNumberCol = openWith.Keys.ElementAt(2);
                checkBox6.Checked = true;
                secondLevelTbl.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").FirstOrDefault(x => (int)x.Attribute("col_no") == partNumberCol) == null ?
                      null : "PART NUMBER");
            }
            else
            {
                secondLevelTbl.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").FirstOrDefault(x => (int)x.Attribute("col_no") == 2) == null ?
                     null : "PART NUMBER");
            }
            if (Regex.IsMatch(openWith[3], descriptionMatching))
            {
                descriptionCol = openWith.Keys.ElementAt(3);
                checkBox7.Checked = true;
                secondLevelTbl.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").FirstOrDefault(x => (int)x.Attribute("col_no") == descriptionCol) == null ?
                      null : "DESCRIPTION");
            }
            else
            {
                secondLevelTbl.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").FirstOrDefault(x => (int)x.Attribute("col_no") == 3) == null ?
                     null : "DESCRIPTION");
            }
            if (Regex.IsMatch(openWith[4], qtyMatching))
            {
                qtyCol = openWith.Keys.ElementAt(4);
                checkBox8.Checked = true;
                secondLevelTbl.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").FirstOrDefault(x => (int)x.Attribute("col_no") == qtyCol) == null ?
                      null : "QTY.");
            }
            else
            {
                secondLevelTbl.Columns.Add(doc.Descendants("bomheader").Elements("bomcol").FirstOrDefault(x => (int)x.Attribute("col_no") == 4) == null ?
                     null : "QTY.");
            }
            AddRows(secondLevelTbl, doc, openWith, topLevelTbl);
        }

        private void AddRows(DataTable secondLevelTbl, XDocument doc, Dictionary<int, string> openWith, DataTable topLevelTbl)
        {
            bool error = false;
            foreach (XElement bomrow in doc.Descendants("bomrow"))
            {
                string itemNo = @"^\d+$";
                string itemCode = @"^[1-2][0-9][0-9][0-9][0-9][0-9]00";
                string qty = @"^\d+$";
                string iNoGiven = (string)bomrow.Elements("bomcell"
                            ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                      y => y.Value.ToUpper().Contains("ITEM NO")).Key)
                    ?.Attribute("value");
                string iCODEgiven = (string)bomrow.Elements("bomcell"
                            ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                      y => y.Value.ToUpper().Contains("CODE") || y.Value.ToUpper().Contains("CD")).Key)
                    ?.Attribute("value");
                string qtyGiven = (string)bomrow.Elements("bomcell"
                            ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                      y => y.Value.ToUpper().Contains("Q")).Key)
                    ?.Attribute("value");
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
                CreateATable(secondLevelTbl, doc, openWith, topLevelTbl);
            }
            else
            {
                CreateATableWithErrors(secondLevelTbl, doc, openWith, topLevelTbl);
            }
        }

        private void CreateATableWithErrors(DataTable secondLevelTbl, XDocument doc, Dictionary<int, string> openWith, DataTable topLevelTbl)
        {
            foreach (XElement bomrow in doc.Descendants("bomrow"))
            {
                secondLevelTbl.Rows.Add(
                bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("ITEM NO")).Key) == null ?
                    null : (string)bomrow.Elements("bomcell"
                        ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                  y => y.Value.ToUpper().Contains("ITEM NO")).Key)
                               ?.Attribute("value") ?? null,

                bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("CODE") || y.Value.ToUpper().Contains("CD")).Key) == null ?
                    null : (string)bomrow.Elements("bomcell"
                        ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                  y => y.Value.ToUpper().Contains("CODE") || y.Value.ToUpper().Contains("CD")).Key)
                               ?.Attribute("value") ?? null,

                bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("PART")).Key) == null ?
                    null : (string)bomrow.Elements("bomcell"
                        ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                  y => y.Value.ToUpper().Contains("PART")).Key)
                               ?.Attribute("value") ?? "N/A",

                bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("DES")).Key) == null ?
                    null : (string)bomrow.Elements("bomcell"
                        ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                  y => y.Value.ToUpper().Contains("DES")).Key)
                               ?.Attribute("value") ?? "N/A",

                bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("Q")).Key) == null ?
                    null : (string)bomrow.Elements("bomcell"
                        ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                  y => y.Value.ToUpper().Contains("Q")).Key)
                               ?.Attribute("value") ?? null
                );
            }
            string[] firstRow = secondLevelTbl.AsEnumerable().Select(r => r.Field<string>("ITEM NO.")).ToArray();
            string has = String.Join(",", firstRow);
            int starting = 1;
            IEnumerable<int> totalColumns = Enumerable.Range(starting, Int32.Parse(firstRow.Last()));
            string shouldBe = String.Join(",", totalColumns);
            ShowTableErrors(secondLevelTbl, has, shouldBe);
        }

        private void CreateATable(DataTable secondLevelTbl, XDocument doc, Dictionary<int, string> openWith, DataTable topLevelTbl)
        {
            foreach (XElement bomrow in doc.Descendants("bomrow"))
            {
                secondLevelTbl.Rows.Add(
                bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("ITEM NO")).Key) == null ?
                    null : (int?)bomrow.Elements("bomcell"
                        ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                  y => y.Value.ToUpper().Contains("ITEM NO")).Key)
                               ?.Attribute("value") ?? null,

                bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("CODE") || y.Value.ToUpper().Contains("CD")).Key) == null ?
                    null : (int?)bomrow.Elements("bomcell"
                        ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                  y => y.Value.ToUpper().Contains("CODE") || y.Value.ToUpper().Contains("CD")).Key)
                               ?.Attribute("value") ?? null,

                bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("PART")).Key) == null ?
                    null : (string)bomrow.Elements("bomcell"
                        ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                  y => y.Value.ToUpper().Contains("PART")).Key)
                               ?.Attribute("value") ?? "N/A",

                bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("DES")).Key) == null ?
                    null : (string)bomrow.Elements("bomcell"
                        ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                  y => y.Value.ToUpper().Contains("DES")).Key)
                               ?.Attribute("value") ?? "N/A",

                bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("Q")).Key) == null ?
                    null : (int?)bomrow.Elements("bomcell"
                        ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                  y => y.Value.ToUpper().Contains("Q")).Key)
                               ?.Attribute("value") ?? null
                );
            }
            checkBox9.Checked = true;
            string[] firstRow = secondLevelTbl.AsEnumerable().Select(r => r.Field<string>("ITEM NO.")).ToArray();
            string has = String.Join(",", firstRow);
            int starting = 1;
            IEnumerable<int> totalColumns = Enumerable.Range(starting, Int32.Parse(firstRow.Last()));
            string shouldBe = String.Join(",", totalColumns);
            ShowTableErrors(secondLevelTbl, has, shouldBe);
        }

        private void ShowTableErrors(DataTable secondLevelTbl, string has, string shouldBe)
        {
            if (has == shouldBe)
            {
                checkBox11.Checked = true;
                Ending(secondLevelTbl);
            }
            else
            {
                checkBox12.Checked = true;
                Ending(secondLevelTbl);
            }

        }

        private void Ending(DataTable secondLevelTbl)
        {
            dataGridView1.DataSource = secondLevelTbl;
            if ((checkBox1.Checked == true && checkBox9.Checked == true && checkBox11.Checked == true) || (checkBox1.Checked == false && checkBox9.Checked == true && checkBox11.Checked == true))
            {
                dataGridView1.DataSource = secondLevelTbl;
            }
            else
            {
                tbProgramLog.AppendText(Environment.NewLine + "Table contains errors!");
            }
        }

        private void Btn1_Send_Click_1(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true && checkBox9.Checked == true && checkBox11.Checked == true)
            {
                //    QBFC_InventoryAssemblyQuery();
                //}
                //else if (checkBox3.Checked == true && checkBox9.Checked == true && checkBox11.Checked == true)
                //{
                /*ADD regardless if ItemCode exists or not*/
                tbProgramLog.AppendText(Environment.NewLine + "Add then Modify the Item");
                AddThenModify();
            }
            else
            {
                tbProgramLog.AppendText(Environment.NewLine + "Cannot import a table with errors!");
            }
        }

        private void AddThenModify()
        {
            QBFC_ItemAdd();
            QBFC_InventoryAssemblyQuery();
        }

        private void QBFC_ItemAdd()
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                sessionManager = new QBSessionManager();

                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                BuildItemInventoryAssemblyAddRq(requestMsgSet);

                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                //print xml string
                tbProgramLog.AppendText(Environment.NewLine +requestMsgSet.ToXMLString());
                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;

                WalkItemInventoryAssemblyAddRs(responseMsgSet);
            }
            catch (Exception e)
            {
                tbProgramLog.AppendText(Environment.NewLine +e.Message);
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

        private void BuildItemInventoryAssemblyAddRq(IMsgSetRequest requestMsgSet)
        {
            IItemInventoryAssemblyAdd itemInventoryAssemblyAddRq = requestMsgSet.AppendItemInventoryAssemblyAddRq();
            DataRow row = topLevelTbl.Rows[0];
            itemInventoryAssemblyAddRq.Name.SetValue(row[0].ToString());
            itemInventoryAssemblyAddRq.SalesDesc.SetValue(row[1].ToString());
            itemInventoryAssemblyAddRq.PurchaseDesc.SetValue(row[1].ToString());
            itemInventoryAssemblyAddRq.IncomeAccountRef.ListID.SetValue(row[2].ToString());
            itemInventoryAssemblyAddRq.COGSAccountRef.ListID.SetValue(row[3].ToString());
            itemInventoryAssemblyAddRq.AssetAccountRef.ListID.SetValue(row[4].ToString());
        }

        private void WalkItemInventoryAssemblyAddRs(IMsgSetResponse responseMsgSet)
        {
            if (responseMsgSet == null) return;
            tbProgramLog.AppendText(Environment.NewLine + "before loop in walkiteminventoryassemblyaddrs");
            IResponseList responseList = responseMsgSet.ResponseList;
            if (responseList == null) return;

            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);

                if (response.StatusCode >= 0)
                {
                    if (response.Detail != null)
                    {
                        ENResponseType responseType = (ENResponseType) response.Type.GetValue();
                        if (responseType == ENResponseType.rtItemInventoryAssemblyAddRs)
                        {
                            IItemInventoryAssemblyRet itemInventoryAssemblyRet = (IItemInventoryAssemblyRet) response.Detail;
                            WalkItemInventoryAssemblyRet(itemInventoryAssemblyRet);
                        }
                    }
                }
            }
        }

        private void WalkItemInventoryAssemblyRet(IItemInventoryAssemblyRet itemInventoryAssemblyRet)
        {
            tbProgramLog.AppendText(Environment.NewLine + "Before error");
            if (itemInventoryAssemblyRet == null) return;
            tbProgramLog.AppendText(Environment.NewLine + "Error fixed");
            string sequence = (string)itemInventoryAssemblyRet.EditSequence.GetValue();
            string listId = (string)itemInventoryAssemblyRet.ListID.GetValue();
            tbProgramLog.AppendText(Environment.NewLine + "Edit sequence: " + sequence + Environment.NewLine + "List ID: " + listId);
        }

        private void QBFC_InventoryAssemblyQuery()
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                sessionManager = new QBSessionManager();

                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                QueryItemAssembly(requestMsgSet, topLevelTbl);

                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;
                WalkItemInventoryAssemblyQueryRs(responseMsgSet);
            }
            catch (Exception e)
            {
                tbProgramLog.AppendText(Environment.NewLine +e.Message);
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

        void QueryItemAssembly(IMsgSetRequest requestMsgSet, DataTable topLevelTbl)
        {
            IItemInventoryAssemblyQuery itemInventoryAssemblyQueryRq = requestMsgSet.AppendItemInventoryAssemblyQueryRq();
            IListWithClassFilter listWithClassFilter = itemInventoryAssemblyQueryRq.ORListQueryWithOwnerIDAndClass.ListWithClassFilter;
            listWithClassFilter.ORNameFilter.NameFilter.MatchCriterion.SetValue(ENMatchCriterion.mcContains);
            itemInventoryAssemblyQueryRq.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ORNameFilter.NameFilter.Name.SetValue(topLevelTbl.Rows[0][0].ToString());
        }

        void WalkItemInventoryAssemblyQueryRs(IMsgSetResponse responseMsgSet)
        {
            IResponseList responseList = responseMsgSet?.ResponseList;
            if (responseList == null) return;

            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);
                tbProgramLog.AppendText(Environment.NewLine +response.StatusCode.ToString() + ": " + response.StatusMessage.ToString());

                if (response.StatusCode >= 0)
                {
                    if (response.StatusCode == 0)
                    {
                        if (response.Detail != null)
                        {
                            ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                            if (responseType == ENResponseType.rtItemInventoryAssemblyQueryRs)
                            {
                                IItemInventoryAssemblyRetList itemInventoryAssemblyRetList = (IItemInventoryAssemblyRetList)response.Detail;
                                WalkItemInventoryAssemblyRet(itemInventoryAssemblyRetList);
                            }
                        }
                    }
                }
            }
        }

        void WalkItemInventoryAssemblyRet(IItemInventoryAssemblyRetList itemInventoryAssemblyRetList)
        {
            if (itemInventoryAssemblyRetList == null) return;
            string sequence = string.Empty;
            string listId = string.Empty; ;
            for (int x = 0; x < itemInventoryAssemblyRetList.Count; x++)
            {
                IItemInventoryAssemblyRet itemInventoryAssemblyRet = itemInventoryAssemblyRetList.GetAt(x);
                sequence = (string)itemInventoryAssemblyRet.EditSequence.GetValue();
                listId = (string)itemInventoryAssemblyRet.ListID.GetValue();
                tbProgramLog.AppendText(Environment.NewLine + "Edit sequence: " + sequence + Environment.NewLine + "List ID: " + listId);
            }
            QBFC_ItemQuery(sequence, listId);
        }

        private void QBFC_ItemQuery(string sequence, string listId)
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                sessionManager = new QBSessionManager();

                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                QueryAllItems(requestMsgSet, secondLevelTbl);

                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                tbProgramLog.AppendText(Environment.NewLine +requestMsgSet.ToXMLString());
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;

                WalkAllItemsQueryRs(responseMsgSet, sequence, listId);
            }
            catch (Exception e)
            {
                tbProgramLog.AppendText(Environment.NewLine +e.Message);
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

        private void QueryAllItems(IMsgSetRequest requestMsgSet, DataTable secondLevelTbl)
        {
            List<string> itemCodes = secondLevelTbl.AsEnumerable().Select(r => r.Field<string>("ItemCode")).ToList();
            foreach (string itemCode in itemCodes)
            {
                IItemQuery itemQueryRq = requestMsgSet.AppendItemQueryRq();
                itemQueryRq.ORListQuery.ListFilter.ORNameFilter.NameFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                itemQueryRq.ORListQuery.ListFilter.ORNameFilter.NameFilter.Name.SetValue(itemCode);
            }
        }

        private void WalkAllItemsQueryRs(IMsgSetResponse responseMsgSet, string sequence, string listId)
        {
            if (responseMsgSet == null) return;
            IResponseList responseList = responseMsgSet.ResponseList;
            if (responseList == null) return;

            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);
                tbProgramLog.AppendText(Environment.NewLine +response.StatusCode.ToString() + ": " + response.StatusMessage.ToString());

                if (response.StatusCode >= 0)
                {
                    if (response.StatusCode == 1)
                    {
                        string itemNoError = response.RequestID;
                       // tbProgramLog.AppendText(Environment.NewLine+"RequestID:  " + itemNoError);
                        FindTheNeededValues(itemNoError);
                       // throw new NotImplementedException();
                    }
                    else if (response.StatusCode ==0)
                    {
                        if (response.Detail != null)
                        {
                            ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                            if (responseType == ENResponseType.rtItemQueryRs)
                            {
                                IORItemRetList itemRetList = (IORItemRetList)response.Detail;
                                WalkAllItemsQueryRet(itemRetList, sequence, listId);
                            }
                        }
                    }
                }
            }
        }

        private void FindTheNeededValues(string itemNoError)
        {
            tbProgramLog.AppendText(Environment.NewLine + "The item with no existing value has itemNo value of: " + itemNoError);
            
            //string itemList = string.Empty;
            ////AddTheItemThatIsNotFoundHere
            //List<string> itemCodes = secondLevelTbl.AsEnumerable().Select(r => r.Field<string>("ItemCode")).ToList();
            //foreach (string itemCode in itemCodes)
            //{
            //    itemList = itemCode;
            //}
            //topLevelTbl = new DataTable();
            
            //string subItem = itemList;
            //String connectionString = @"Data Source=SQLSERVER\ITEMCODE;Initial Catalog=dat8121;Integrated Security=True";
            //SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT  [ItemCode], [Description]" +
            //                                                " FROM[dat8121].[dbo].[I_ItemCode]" +
            //                                                " where TRY_CAST(ItemCode as nvarchar) = '" + subItem  +"'"
            //    , connectionString);
            //dataAdapter.Fill(topLevelTbl);
            //tbProgramLog.AppendText(Environment.NewLine + "Line Reached");
            //DataRow row = topLevelTbl.Rows[0];
            //topLevelTbl.Columns.Add("IncomeAccountRef", typeof(string));
            //topLevelTbl.Columns.Add("COGSAccountRef", typeof(string));
            //topLevelTbl.Columns.Add("AssetAccountRef", typeof(string));
            
            ////row[2] = "570000-1136323777";
            ////row[3] = "800001E1-1537737142";
            ////row[4] = "800001A9-1511318480";
            //tbProgramLog.AppendText(Environment.NewLine +row[0] + "  " + row[1]);
            //tbProgramLog.AppendText(Environment.NewLine +itemList);
            //QBFC_ItemAdd();
            //throw new NotImplementedException();
        }

        private void WalkAllItemsQueryRet(IORItemRetList itemRetList, string sequence, string listId)
        {
            tbProgramLog.AppendText(Environment.NewLine + "WalkAllItemsQueryRet method was reached");
            string itemListId = string.Empty;
            string itemName = string.Empty;
            string itemSequence = string.Empty;
            if (itemRetList == null)
            {
                QBFC_ItemAdd();
            }
            for (int y = 0; y < itemRetList.Count; y++)
            {
                IORItemRet itemRet = itemRetList.GetAt(y);
                if (itemRet.ItemInventoryAssemblyRet != null)
                {
                    itemListId = itemRet.ItemInventoryAssemblyRet.ListID.GetValue();
                    itemSequence = (string)itemRet.ItemInventoryAssemblyRet.EditSequence.GetValue();
                    itemName = (string)itemRet.ItemInventoryAssemblyRet.Name.GetValue();
                }
                else if (itemRet.ItemInventoryRet != null)
                {
                    itemListId = itemRet.ItemInventoryRet.ListID.GetValue();
                    itemSequence = (string)itemRet.ItemInventoryRet.EditSequence.GetValue();
                    itemName = (string)itemRet.ItemInventoryRet.Name.GetValue();
                }
                else if (itemRet.ItemNonInventoryRet != null)
                {
                    itemListId = itemRet.ItemNonInventoryRet.ListID.GetValue();
                    itemSequence = (string)itemRet.ItemNonInventoryRet.EditSequence.GetValue();
                    itemName = (string)itemRet.ItemNonInventoryRet.Name.GetValue();
                }
                else if (itemRet.ItemSubtotalRet != null)
                {
                    itemListId = itemRet.ItemSubtotalRet.ListID.GetValue();
                    itemSequence = (string)itemRet.ItemSubtotalRet.EditSequence.GetValue();
                    itemName = (string)itemRet.ItemSubtotalRet.Name.GetValue();
                }
                else if (itemRet.ItemDiscountRet != null)
                {
                    itemListId = itemRet.ItemDiscountRet.ListID.GetValue();
                    itemSequence = (string)itemRet.ItemDiscountRet.EditSequence.GetValue();
                    itemName = (string)itemRet.ItemDiscountRet.Name.GetValue();
                }
                else if (itemRet.ItemPaymentRet != null)
                {
                    itemListId = itemRet.ItemPaymentRet.ListID.GetValue();
                    itemSequence = (string)itemRet.ItemPaymentRet.EditSequence.GetValue();
                    itemName = (string)itemRet.ItemPaymentRet.Name.GetValue();
                }
                else if (itemRet.ItemSalesTaxRet != null)
                {
                    itemListId = itemRet.ItemSalesTaxRet.ListID.GetValue();
                    itemSequence = (string)itemRet.ItemSalesTaxRet.EditSequence.GetValue();
                    itemName = (string)itemRet.ItemSalesTaxRet.Name.GetValue();
                }
                else if (itemRet.ItemSalesTaxGroupRet != null)
                {
                    itemListId = itemRet.ItemSalesTaxGroupRet.ListID.GetValue();
                    itemSequence = (string)itemRet.ItemSalesTaxGroupRet.EditSequence.GetValue();
                    itemName = (string)itemRet.ItemSalesTaxGroupRet.Name.GetValue();
                }
                else if (itemRet.ItemGroupRet != null)
                {
                    itemListId = itemRet.ItemGroupRet.ListID.GetValue();
                    itemSequence = (string)itemRet.ItemGroupRet.EditSequence.GetValue();
                    itemName = (string)itemRet.ItemGroupRet.Name.GetValue();
                }
                else if (itemRet.ItemServiceRet != null)
                {
                    itemListId = itemRet.ItemServiceRet.ListID.GetValue();
                    itemSequence = (string)itemRet.ItemServiceRet.EditSequence.GetValue();
                    itemName = (string)itemRet.ItemServiceRet.Name.GetValue();
                }
                else if (itemRet.ItemOtherChargeRet != null)
                {
                    itemListId = itemRet.ItemOtherChargeRet.ListID.GetValue();
                    itemSequence = (string)itemRet.ItemOtherChargeRet.EditSequence.GetValue();
                    itemName = (string)itemRet.ItemOtherChargeRet.Name.GetValue();
                }
                tbProgramLog.AppendText(Environment.NewLine + "Edit sequence: " + itemSequence + Environment.NewLine + "List ID: " + itemListId);
                tbProgramLog.AppendText(Environment.NewLine + "Name: " + itemName);
            }
            QBFC_ItemModify(sequence, listId, itemListId);
        }

        private void QBFC_ItemModify(string sequence, string listID, string itemListID)
        {
            tbProgramLog.AppendText(Environment.NewLine + "QBFC_ItemModify method was reached");
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                sessionManager = new QBSessionManager();

                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                ModifyItem(requestMsgSet, secondLevelTbl, topLevelTbl, sequence, listID);

                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                tbProgramLog.AppendText(Environment.NewLine +requestMsgSet.ToXMLString());
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;

                WalkItemsModifyRs(responseMsgSet, itemListID);
            }
            catch (Exception e)
            {
                tbProgramLog.AppendText(Environment.NewLine +e.Message);
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

        private void ModifyItem(IMsgSetRequest requestMsgSet, DataTable secondLevelTbl, DataTable topLevelTbl, string sequence, string listId)
        {
            IItemInventoryAssemblyMod itemInventoryAssemblyModRq = requestMsgSet.AppendItemInventoryAssemblyModRq();
            itemInventoryAssemblyModRq.ListID.SetValue(listId);
            itemInventoryAssemblyModRq.SalesDesc.SetValue(topLevelTbl.Rows[0][1].ToString());
            itemInventoryAssemblyModRq.PurchaseDesc.SetValue(topLevelTbl.Rows[0][1].ToString());
            itemInventoryAssemblyModRq.EditSequence.SetValue(sequence);

            for (int i = 0; i < secondLevelTbl.Rows.Count; i++)
            {
                IItemInventoryAssemblyLine itemInventoryAssemblyLine1 = itemInventoryAssemblyModRq.ORItemInventoryAssemblyLine.ItemInventoryAssemblyLineList.Append();
                itemInventoryAssemblyLine1.ItemInventoryRef.FullName.SetValue(secondLevelTbl.Rows[i][1].ToString());
                itemInventoryAssemblyLine1.Quantity.SetValue(Convert.ToDouble(secondLevelTbl.Rows[i][4]));
            }
        }

        private void WalkItemsModifyRs(IMsgSetResponse responseMsgSet, string itemListId)
        {
            if (responseMsgSet == null) return;
            IResponseList responseList = responseMsgSet.ResponseList;
            if (responseList == null) return;

            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);

                if (response.StatusCode >= 0)
                {
                    if (response.Detail != null)
                    {
                        ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                        if (responseType == ENResponseType.rtItemInventoryAssemblyModRs)
                        {
                            IItemInventoryAssemblyRet itemModifyRet = (IItemInventoryAssemblyRet)response.Detail;
                            WalkItemModifyRet(itemModifyRet);
                        }
                    }
                }
            }
        }

        private void WalkItemModifyRet(IItemInventoryAssemblyRet itemModifyRet)
        {
            if (itemModifyRet == null) return;
            string sequence = (string)itemModifyRet.EditSequence.GetValue();
            string listId = (string)itemModifyRet.ListID.GetValue();
            tbProgramLog.AppendText(Environment.NewLine + "Edit sequence: " + sequence + Environment.NewLine + "List ID: " + listId);
        }

    }
}