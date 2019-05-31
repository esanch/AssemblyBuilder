/*************************************************************************
 * 
 * Elizabeth Earl
 * __________________
 * 
 *  [2018] - [2019] 
 * 
 ************************************************************************/
using System;
using System.Windows.Forms;
using System.Data;
using Interop.QBFC13;
using System.Xml.Linq;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace InvoiceAdd
{
    public class Frm1InvoiceAdd : Form
    {
        private readonly System.ComponentModel.Container components = null;
        private Button btn1_Send;
        private Button btn2_Exit;
        private Button btnOpenFile_Reset;
        private DataGridView dataGridView1;
        string IncomeAccount = "570000-1136323777";
        string COGSAccount = "800001E1-1537737142";
        string InventoryAssetAccount = "800001A9-1511318480";
        DataTable secondLevelTbl = new DataTable();
        DataTable topLevelTbl = new DataTable();
        DataTable subAssembly = new DataTable();
        DataTable toplvlBOMtbl = new DataTable();
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
        private RichTextBox textBox1;
        private RichTextBox tbProgramLog;
        private RichTextBox txtBox;
        protected internal Label label6;
        private PictureBox pictureBox1;
        private Button BOM;
        private Label label5;
        private CheckBox checkBox14;
        private CheckBox checkBox13;
        private Label label7;
        private CheckBox checkBox16;
        private CheckBox checkBox15;
        private Label label8;
        private CheckBox checkBox18;
        private CheckBox checkBox17;
        private DataGridView dataGridView2;
        private Label label9;

        bool ifError;

        private Frm1InvoiceAdd()
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                components?.Dispose();
            }

            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            this.btn1_Send = new System.Windows.Forms.Button();
            this.btn2_Exit = new System.Windows.Forms.Button();
            this.btnOpenFile_Reset = new System.Windows.Forms.Button();
            this.BOM = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.checkBox3 = new System.Windows.Forms.CheckBox();
            this.checkBox4 = new System.Windows.Forms.CheckBox();
            this.checkBox5 = new System.Windows.Forms.CheckBox();
            this.checkBox6 = new System.Windows.Forms.CheckBox();
            this.checkBox7 = new System.Windows.Forms.CheckBox();
            this.checkBox8 = new System.Windows.Forms.CheckBox();
            this.checkBox9 = new System.Windows.Forms.CheckBox();
            this.checkBox10 = new System.Windows.Forms.CheckBox();
            this.checkBox11 = new System.Windows.Forms.CheckBox();
            this.checkBox12 = new System.Windows.Forms.CheckBox();
            this.checkBox13 = new System.Windows.Forms.CheckBox();
            this.checkBox14 = new System.Windows.Forms.CheckBox();
            this.checkBox15 = new System.Windows.Forms.CheckBox();
            this.checkBox16 = new System.Windows.Forms.CheckBox();
            this.checkBox17 = new System.Windows.Forms.CheckBox();
            this.checkBox18 = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.RichTextBox();
            this.tbProgramLog = new System.Windows.Forms.RichTextBox();
            this.txtBox = new System.Windows.Forms.RichTextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // btn1_Send
            // 
            this.btn1_Send.BackColor = System.Drawing.SystemColors.Control;
            this.btn1_Send.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F);
            this.btn1_Send.Location = new System.Drawing.Point(406, 711);
            this.btn1_Send.Name = "btn1_Send";
            this.btn1_Send.Size = new System.Drawing.Size(87, 32);
            this.btn1_Send.TabIndex = 57;
            this.btn1_Send.Text = "Send";
            this.btn1_Send.UseVisualStyleBackColor = false;
            this.btn1_Send.Click += new System.EventHandler(this.Btn1_Send_Click_1);
            // 
            // btn2_Exit
            // 
            this.btn2_Exit.BackColor = System.Drawing.SystemColors.Control;
            this.btn2_Exit.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn2_Exit.Location = new System.Drawing.Point(496, 711);
            this.btn2_Exit.Name = "btn2_Exit";
            this.btn2_Exit.Size = new System.Drawing.Size(80, 32);
            this.btn2_Exit.TabIndex = 58;
            this.btn2_Exit.Text = "Exit";
            this.btn2_Exit.UseVisualStyleBackColor = false;
            this.btn2_Exit.Click += new System.EventHandler(this.Btn2_Exit_Click);
            // 
            // btnOpenFile_Reset
            // 
            this.btnOpenFile_Reset.BackColor = System.Drawing.SystemColors.Control;
            this.btnOpenFile_Reset.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOpenFile_Reset.Location = new System.Drawing.Point(496, 673);
            this.btnOpenFile_Reset.Name = "btnOpenFile_Reset";
            this.btnOpenFile_Reset.Size = new System.Drawing.Size(80, 32);
            this.btnOpenFile_Reset.TabIndex = 63;
            this.btnOpenFile_Reset.Text = "Open File";
            this.btnOpenFile_Reset.UseVisualStyleBackColor = false;
            this.btnOpenFile_Reset.Click += new System.EventHandler(this.BtnOpenFile_Reset_Click);
            // 
            // BOM
            // 
            this.BOM.BackColor = System.Drawing.SystemColors.Control;
            this.BOM.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BOM.Location = new System.Drawing.Point(406, 673);
            this.BOM.Name = "BOM";
            this.BOM.Size = new System.Drawing.Size(87, 32);
            this.BOM.TabIndex = 90;
            this.BOM.Text = "Show BOM";
            this.BOM.UseVisualStyleBackColor = false;
            this.BOM.Click += new System.EventHandler(this.BOM_Click);
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
            this.dataGridView1.Location = new System.Drawing.Point(12, 164);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridView1.Size = new System.Drawing.Size(591, 451);
            this.dataGridView1.TabIndex = 64;
            // 
            // dataGridView2
            // 
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView2.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle4;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView2.DefaultCellStyle = dataGridViewCellStyle5;
            this.dataGridView2.Location = new System.Drawing.Point(12, 89);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.ReadOnly = true;
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView2.RowHeadersDefaultCellStyle = dataGridViewCellStyle6;
            this.dataGridView2.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.dataGridView2.Size = new System.Drawing.Size(591, 69);
            this.dataGridView2.TabIndex = 100;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(154, 621);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(44, 17);
            this.checkBox1.TabIndex = 65;
            this.checkBox1.Text = "Yes";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(263, 621);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(40, 17);
            this.checkBox2.TabIndex = 66;
            this.checkBox2.Text = "No";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // checkBox3
            // 
            this.checkBox3.AutoSize = true;
            this.checkBox3.Location = new System.Drawing.Point(349, 621);
            this.checkBox3.Name = "checkBox3";
            this.checkBox3.Size = new System.Drawing.Size(119, 17);
            this.checkBox3.TabIndex = 67;
            this.checkBox3.Text = "ItemCode not found";
            this.checkBox3.UseVisualStyleBackColor = true;
            // 
            // checkBox4
            // 
            this.checkBox4.AutoSize = true;
            this.checkBox4.Location = new System.Drawing.Point(154, 644);
            this.checkBox4.Name = "checkBox4";
            this.checkBox4.Size = new System.Drawing.Size(86, 17);
            this.checkBox4.TabIndex = 68;
            this.checkBox4.Text = "Item Number";
            this.checkBox4.UseVisualStyleBackColor = true;
            // 
            // checkBox5
            // 
            this.checkBox5.AutoSize = true;
            this.checkBox5.Location = new System.Drawing.Point(263, 644);
            this.checkBox5.Name = "checkBox5";
            this.checkBox5.Size = new System.Drawing.Size(71, 17);
            this.checkBox5.TabIndex = 69;
            this.checkBox5.Text = "ItemCode";
            this.checkBox5.UseVisualStyleBackColor = true;
            // 
            // checkBox6
            // 
            this.checkBox6.AutoSize = true;
            this.checkBox6.Location = new System.Drawing.Point(349, 644);
            this.checkBox6.Name = "checkBox6";
            this.checkBox6.Size = new System.Drawing.Size(85, 17);
            this.checkBox6.TabIndex = 70;
            this.checkBox6.Text = "Part Number";
            this.checkBox6.UseVisualStyleBackColor = true;
            // 
            // checkBox7
            // 
            this.checkBox7.AutoSize = true;
            this.checkBox7.Location = new System.Drawing.Point(435, 644);
            this.checkBox7.Name = "checkBox7";
            this.checkBox7.Size = new System.Drawing.Size(79, 17);
            this.checkBox7.TabIndex = 71;
            this.checkBox7.Text = "Description";
            this.checkBox7.UseVisualStyleBackColor = true;
            // 
            // checkBox8
            // 
            this.checkBox8.AutoSize = true;
            this.checkBox8.Location = new System.Drawing.Point(521, 644);
            this.checkBox8.Name = "checkBox8";
            this.checkBox8.Size = new System.Drawing.Size(65, 17);
            this.checkBox8.TabIndex = 72;
            this.checkBox8.Text = "Quantity";
            this.checkBox8.UseVisualStyleBackColor = true;
            // 
            // checkBox9
            // 
            this.checkBox9.AutoSize = true;
            this.checkBox9.Location = new System.Drawing.Point(154, 667);
            this.checkBox9.Name = "checkBox9";
            this.checkBox9.Size = new System.Drawing.Size(44, 17);
            this.checkBox9.TabIndex = 73;
            this.checkBox9.Text = "Yes";
            this.checkBox9.UseVisualStyleBackColor = true;
            // 
            // checkBox10
            // 
            this.checkBox10.AutoSize = true;
            this.checkBox10.Location = new System.Drawing.Point(263, 667);
            this.checkBox10.Name = "checkBox10";
            this.checkBox10.Size = new System.Drawing.Size(40, 17);
            this.checkBox10.TabIndex = 74;
            this.checkBox10.Text = "No";
            this.checkBox10.UseVisualStyleBackColor = true;
            // 
            // checkBox11
            // 
            this.checkBox11.AutoSize = true;
            this.checkBox11.Location = new System.Drawing.Point(154, 692);
            this.checkBox11.Name = "checkBox11";
            this.checkBox11.Size = new System.Drawing.Size(44, 17);
            this.checkBox11.TabIndex = 75;
            this.checkBox11.Text = "Yes";
            this.checkBox11.UseVisualStyleBackColor = true;
            // 
            // checkBox12
            // 
            this.checkBox12.AutoSize = true;
            this.checkBox12.Location = new System.Drawing.Point(263, 692);
            this.checkBox12.Name = "checkBox12";
            this.checkBox12.Size = new System.Drawing.Size(40, 17);
            this.checkBox12.TabIndex = 76;
            this.checkBox12.Text = "No";
            this.checkBox12.UseVisualStyleBackColor = true;
            // 
            // checkBox13
            // 
            this.checkBox13.AutoSize = true;
            this.checkBox13.Location = new System.Drawing.Point(446, 8);
            this.checkBox13.Name = "checkBox13";
            this.checkBox13.Size = new System.Drawing.Size(44, 17);
            this.checkBox13.TabIndex = 91;
            this.checkBox13.Text = "Yes";
            this.checkBox13.UseVisualStyleBackColor = true;
            // 
            // checkBox14
            // 
            this.checkBox14.AutoSize = true;
            this.checkBox14.Location = new System.Drawing.Point(495, 8);
            this.checkBox14.Name = "checkBox14";
            this.checkBox14.Size = new System.Drawing.Size(40, 17);
            this.checkBox14.TabIndex = 92;
            this.checkBox14.Text = "No";
            this.checkBox14.UseVisualStyleBackColor = true;
            // 
            // checkBox15
            // 
            this.checkBox15.AutoSize = true;
            this.checkBox15.Location = new System.Drawing.Point(496, 35);
            this.checkBox15.Name = "checkBox15";
            this.checkBox15.Size = new System.Drawing.Size(70, 17);
            this.checkBox15.TabIndex = 94;
            this.checkBox15.Text = "Assembly";
            this.checkBox15.UseVisualStyleBackColor = true;
            // 
            // checkBox16
            // 
            this.checkBox16.AutoSize = true;
            this.checkBox16.Location = new System.Drawing.Point(445, 35);
            this.checkBox16.Name = "checkBox16";
            this.checkBox16.Size = new System.Drawing.Size(45, 17);
            this.checkBox16.TabIndex = 95;
            this.checkBox16.Text = "Part";
            this.checkBox16.UseVisualStyleBackColor = true;
            // 
            // checkBox17
            // 
            this.checkBox17.AutoSize = true;
            this.checkBox17.Location = new System.Drawing.Point(445, 58);
            this.checkBox17.Name = "checkBox17";
            this.checkBox17.Size = new System.Drawing.Size(44, 17);
            this.checkBox17.TabIndex = 97;
            this.checkBox17.Text = "Yes";
            this.checkBox17.UseVisualStyleBackColor = true;
            // 
            // checkBox18
            // 
            this.checkBox18.AutoSize = true;
            this.checkBox18.Location = new System.Drawing.Point(495, 57);
            this.checkBox18.Name = "checkBox18";
            this.checkBox18.Size = new System.Drawing.Size(40, 17);
            this.checkBox18.TabIndex = 98;
            this.checkBox18.Text = "No";
            this.checkBox18.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 625);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 13);
            this.label1.TabIndex = 77;
            this.label1.Text = "ICodes match?";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(23, 648);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(124, 13);
            this.label2.TabIndex = 78;
            this.label2.Text = "Columns in correct order:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(23, 670);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(73, 13);
            this.label3.TabIndex = 79;
            this.label3.Text = "Correct input?";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(23, 695);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(98, 13);
            this.label4.TabIndex = 80;
            this.label4.Text = "Row order correct?";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(332, 9);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(99, 13);
            this.label5.TabIndex = 93;
            this.label5.Text = "Item already exists?";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.ForeColor = System.Drawing.SystemColors.ButtonFace;
            this.label6.Location = new System.Drawing.Point(1028, 730);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(47, 13);
            this.label6.TabIndex = 88;
            this.label6.Text = "E.Earl ©";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(332, 39);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(102, 13);
            this.label7.TabIndex = 96;
            this.label7.Text = "Is Part or Assembly?";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(307, 61);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(127, 13);
            this.label8.TabIndex = 99;
            this.label8.Text = "Showing complete BOM?";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(205, 9);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(58, 13);
            this.label9.TabIndex = 101;
            this.label9.Text = "Top Level:";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(171, 25);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(130, 58);
            this.textBox1.TabIndex = 83;
            this.textBox1.Text = "";
            // 
            // tbProgramLog
            // 
            this.tbProgramLog.Location = new System.Drawing.Point(609, 372);
            this.tbProgramLog.Name = "tbProgramLog";
            this.tbProgramLog.Size = new System.Drawing.Size(466, 348);
            this.tbProgramLog.TabIndex = 85;
            this.tbProgramLog.Text = "";
            // 
            // txtBox
            //
            this.txtBox.Location = new System.Drawing.Point(609, 5);
            this.txtBox.Name = "txtBox";
            this.txtBox.Size = new System.Drawing.Size(466, 348);
            this.txtBox.TabIndex = 87;
            this.txtBox.Text = "";
            // 
            // pictureBox1
            // 
            this.pictureBox1.ErrorImage = global::InvoiceAdd.Properties.Resources.UBELogo;
            this.pictureBox1.Image = global::InvoiceAdd.Properties.Resources.UBELogo;
            this.pictureBox1.InitialImage = global::InvoiceAdd.Properties.Resources.UBELogo;
            this.pictureBox1.Location = new System.Drawing.Point(-26, 5);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(225, 81);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 89;
            this.pictureBox1.TabStop = false;
            // 
            // Frm1InvoiceAdd
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.ClientSize = new System.Drawing.Size(1082, 746);
            this.Controls.Add(this.btn1_Send);
            this.Controls.Add(this.btn2_Exit);
            this.Controls.Add(this.btnOpenFile_Reset);
            this.Controls.Add(this.BOM);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.checkBox2);
            this.Controls.Add(this.checkBox3);
            this.Controls.Add(this.checkBox4);
            this.Controls.Add(this.checkBox5);
            this.Controls.Add(this.checkBox6);
            this.Controls.Add(this.checkBox7);
            this.Controls.Add(this.checkBox8);
            this.Controls.Add(this.checkBox9);
            this.Controls.Add(this.checkBox10);
            this.Controls.Add(this.checkBox11);
            this.Controls.Add(this.checkBox12);
            this.Controls.Add(this.checkBox13);
            this.Controls.Add(this.checkBox14);
            this.Controls.Add(this.checkBox15);
            this.Controls.Add(this.checkBox16);
            this.Controls.Add(this.checkBox17);
            this.Controls.Add(this.checkBox18);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.tbProgramLog);
            this.Controls.Add(this.txtBox);
            this.Controls.Add(this.pictureBox1);
            this.Name = "Frm1InvoiceAdd";
            this.Text = "Add an Item Inventory to QuickBooks";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        [STAThread]
        static void Main()
        {
            Application.Run(new Frm1InvoiceAdd());
        }

        //[STAThread]
        //static void Main(string args[])
        //{
        //    MessageBox.Show("Running {0} through workflow", args[1]);
        //    Application.Run(new Frm1InvoiceAdd());
        //}


        private void Btn2_Exit_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void BtnOpenFile_Reset_Click(object sender, EventArgs e)
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
            checkBox13.Checked = false;
            checkBox14.Checked = false;
            checkBox15.Checked = false;
            checkBox16.Checked = false;
            checkBox17.Checked = false;
            checkBox18.Checked = false;
            dataGridView1.DataSource = null;
            dataGridView2.DataSource = null;
            textBox1.Clear();
            tbProgramLog.Clear();
            txtBox.Clear();
            /*
             fileName = @"\\SQLSERVER\bom\24902200-20180907.xml";


                        try
                        {
                            StartErrorChecking();
                        }
                        catch (FileNotFoundException ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
  */

            ///*
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = @"xml files (*.xml)|*.xml";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    fileName = openFileDialog.FileName;
                    try
                    {
                        StartErrorChecking();
                    }
                    catch (System.Xml.XmlException ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            // */
        }

        private void StartErrorChecking()
        {
            secondLevelTbl = new DataTable();
            topLevelTbl = new DataTable();
            XDocument doc = XDocument.Load(fileName);
            string matchFirst = @"^[1-2][0-9][0-9][0-9][0-9][0-9]00";
            foreach (XElement bom in doc.Descendants("bom"))
            {
                string nameData = bom.Attribute("name")?.Value;
                string docPath = bom.Attribute("document_path")?.Value;
                string input = Regex.Match(nameData, matchFirst).Value;
                string cut = docPath?.Substring(Math.Max(0, docPath.Length - 15), 8);
                string connectionString =
                    @"Data Source=SQLSERVER\ITEMCODE;Initial Catalog=dat8121;Integrated Security=True";
                SqlDataAdapter dataAdapter = new SqlDataAdapter(
                    "SELECT a.ItemCode, a.Description, ItemType, c.[IncomeAccountRefListID], c.[COGSAccountRefListID], c.[AssetAccountRefListID]" +
                    " FROM [dat8121].[dbo].[v_ItemCode_QB] b" +
                    " RIGHT JOIN [dat8121].[dbo].[I_ItemCode] a ON a.itemcode = b.ItemCode" +
                    " LEFT JOIN [QODBC].[dbo].[Tbl_Item] c ON a.ItemCode = TRY_CAST(c.fullname AS int) AND b.ItemCode = TRY_CAST(c.fullname AS int)" +
                    " WHERE a.itemcode = '" + input +
                    "' OR a.itemcode = '" + cut + "'"
                    , connectionString);
                dataAdapter.Fill(topLevelTbl);

                bool isTrue = topLevelTbl.Rows.Count > 0;
                if (isTrue)
                {
                    string fromData = topLevelTbl.Rows[0][0].ToString();
                    if (input.Equals(cut))
                    {
                        if (input.Equals(fromData))
                        {
                            textBox1.Text = (topLevelTbl.Rows[0][0] + Environment.NewLine);
                            dataGridView2.DataSource = topLevelTbl;
                            checkBox1.Checked = true;
                            ColumnCorrectOrder();
                        }
                    }
                    else if (String.Equals(input, cut) == false)
                    {
                        ifError = true;
                        ColumnCorrectOrder(ifError, input, cut);
                    }
                }
                else
                {
                    ifError = true;
                    ColumnCorrectOrder(ifError, input, cut);
                }
            }
        }

        private void ColumnCorrectOrder(bool ifError, string input, string cut)
        {
            if (ifError)
            {
                if (String.Equals(input, cut) == false)
                {
                    checkBox2.Checked = true;
                    ColumnCorrectOrder();
                }
                else
                {
                    checkBox3.Checked = true;
                    ColumnCorrectOrder();
                }
            }
        }

        private void ColumnCorrectOrder()
        {
            XDocument doc = XDocument.Load(fileName);
            Dictionary<int, string> openWith = new Dictionary<int, string>();
            foreach (XElement bom in doc.Descendants("bomcol"))
            {
                int currentColumn = int.Parse(bom.Attribute("col_no")?.Value);
                string colHeader = bom.Attribute("name")?.Value ?? "n/a";
                openWith.Add(currentColumn, colHeader);
            }
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
                secondLevelTbl.Columns.Add(
                    doc.Descendants("bomheader").Elements("bomcol")
                        .FirstOrDefault(x => (int)x.Attribute("col_no") == itemNoCol) == null
                        ? null
                        : "ITEM NO.");
            }
            else
            {
                secondLevelTbl.Columns.Add(
                    doc.Descendants("bomheader").Elements("bomcol")
                        .FirstOrDefault(x => (int)x.Attribute("col_no") == 0) == null
                        ? null
                        : "ITEM NO.");
            }

            if (Regex.IsMatch(openWith[1], itemCodeMatching))
            {
                itemCodeCol = openWith.Keys.ElementAt(1);
                checkBox5.Checked = true;
                secondLevelTbl.Columns.Add(
                    doc.Descendants("bomheader").Elements("bomcol")
                        .FirstOrDefault(x => (int)x.Attribute("col_no") == itemCodeCol) == null
                        ? null
                        : "ITEMCODE");
            }
            else
            {
                secondLevelTbl.Columns.Add(
                    doc.Descendants("bomheader").Elements("bomcol")
                        .FirstOrDefault(x => (int)x.Attribute("col_no") == 1) == null
                        ? null
                        : "ITEMCODE");
            }

            if (Regex.IsMatch(openWith[2], partNoMatching))
            {
                partNumberCol = openWith.Keys.ElementAt(2);
                checkBox6.Checked = true;
                secondLevelTbl.Columns.Add(
                    doc.Descendants("bomheader").Elements("bomcol")
                        .FirstOrDefault(x => (int)x.Attribute("col_no") == partNumberCol) == null
                        ? null
                        : "PART NUMBER");
            }
            else
            {
                secondLevelTbl.Columns.Add(
                    doc.Descendants("bomheader").Elements("bomcol")
                        .FirstOrDefault(x => (int)x.Attribute("col_no") == 2) == null
                        ? null
                        : "PART NUMBER");
            }

            if (Regex.IsMatch(openWith[3], descriptionMatching))
            {
                descriptionCol = openWith.Keys.ElementAt(3);
                checkBox7.Checked = true;
                secondLevelTbl.Columns.Add(
                    doc.Descendants("bomheader").Elements("bomcol")
                        .FirstOrDefault(x => (int)x.Attribute("col_no") == descriptionCol) == null
                        ? null
                        : "DESCRIPTION");
            }
            else
            {
                secondLevelTbl.Columns.Add(
                    doc.Descendants("bomheader").Elements("bomcol")
                        .FirstOrDefault(x => (int)x.Attribute("col_no") == 3) == null
                        ? null
                        : "DESCRIPTION");
            }

            if (Regex.IsMatch(openWith[4], qtyMatching))
            {
                qtyCol = openWith.Keys.ElementAt(4);
                checkBox8.Checked = true;
                secondLevelTbl.Columns.Add(
                    doc.Descendants("bomheader").Elements("bomcol")
                        .FirstOrDefault(x => (int)x.Attribute("col_no") == qtyCol) == null
                        ? null
                        : "QTY.");
            }
            else
            {
                secondLevelTbl.Columns.Add(
                    doc.Descendants("bomheader").Elements("bomcol")
                        .FirstOrDefault(x => (int)x.Attribute("col_no") == 4) == null
                        ? null
                        : "QTY.");
            }

            AddRows(secondLevelTbl, doc, openWith);
        }

        private void AddRows(DataTable secondLevelTbl, XDocument doc, Dictionary<int, string> openWith)
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
                string iCodeGiven = (string)bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("CODE") ||
                                                   y.Value.ToUpper().Contains("CD")).Key)
                    ?.Attribute("value");
                string qtyGiven = (string)bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("Q")).Key)
                    ?.Attribute("value");
                bool numberTest = Regex.IsMatch(iNoGiven, itemNo);
                bool codeTest = Regex.IsMatch(iCodeGiven, itemCode);
                bool qtyTest = Regex.IsMatch(qtyGiven, qty);
                if (numberTest == false)
                {
                    txtBox.AppendText(Environment.NewLine + "Error in the following ItemNo provided: " + iNoGiven);
                    checkBox10.Checked = true;
                    error = true;
                }
                else if (codeTest == false)
                {
                    txtBox.AppendText(Environment.NewLine + "Error in the following ItemCode provided in line " +
                                      iNoGiven + ": " + iCodeGiven);
                    checkBox10.Checked = true;
                    error = true;
                }
                else if (qtyTest == false)
                {
                    txtBox.AppendText(Environment.NewLine + "Error in the following Quantity provided in line " +
                                      iNoGiven + ": " + qtyGiven);
                    checkBox10.Checked = true;
                    error = true;
                }
            }
            if (error == false)
            {
                CreateATable(secondLevelTbl, doc, openWith);
            }
            else
            {
                CreateATableWithErrors(secondLevelTbl, doc, openWith);

            }
        }

        private void CreateATableWithErrors(DataTable secondLevelTbl, XDocument doc, Dictionary<int, string> openWith)
        {
            foreach (XElement bomrow in doc.Descendants("bomrow"))
            {
                secondLevelTbl.Rows.Add(
                    bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("ITEM NO")).Key) == null
                        ? null
                        : (string)bomrow.Elements("bomcell"
                            ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                      y => y.Value.ToUpper().Contains("ITEM NO")).Key)
                            ?.Attribute("value"),

                    bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("CODE") ||
                                                   y.Value.ToUpper().Contains("CD")).Key) == null
                        ? null
                        : (string)bomrow.Elements("bomcell"
                            ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                      y => y.Value.ToUpper().Contains("CODE") ||
                                                           y.Value.ToUpper().Contains("CD")).Key)
                            ?.Attribute("value"),

                    bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("PART")).Key) == null
                        ? null
                        : (string)bomrow.Elements("bomcell"
                              ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                        y => y.Value.ToUpper().Contains("PART")).Key)
                              ?.Attribute("value") ?? "N/A",

                    bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("DES")).Key) == null
                        ? null
                        : (string)bomrow.Elements("bomcell"
                              ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                        y => y.Value.ToUpper().Contains("DES")).Key)
                              ?.Attribute("value") ?? "N/A",

                    bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("Q")).Key) == null
                        ? null
                        : (string)bomrow.Elements("bomcell"
                            ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                      y => y.Value.ToUpper().Contains("Q")).Key)
                            ?.Attribute("value")
                );
            }
            string[] firstRow = secondLevelTbl.AsEnumerable().Select(r => r.Field<string>("ITEM NO.")).ToArray();
            string has = String.Join(",", firstRow);
            int starting = 1;
            IEnumerable<int> totalColumns = Enumerable.Range(starting, Int32.Parse(firstRow.Last()));
            string shouldBe = String.Join(",", totalColumns);
            ShowTableErrors(secondLevelTbl, has, shouldBe);
        }

        private void CreateATable(DataTable secondLevelTbl, XDocument doc, Dictionary<int, string> openWith)
        {
            foreach (XElement bomrow in doc.Descendants("bomrow"))
            {
                secondLevelTbl.Rows.Add(
                    bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("ITEM NO")).Key) == null
                        ? null
                        : (int?)bomrow.Elements("bomcell"
                            ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                      y => y.Value.ToUpper().Contains("ITEM NO")).Key)
                            ?.Attribute("value"),

                    bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("CODE") ||
                                                   y.Value.ToUpper().Contains("CD")).Key) == null
                        ? null
                        : (int?)bomrow.Elements("bomcell"
                            ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                      y => y.Value.ToUpper().Contains("CODE") ||
                                                           y.Value.ToUpper().Contains("CD")).Key)
                            ?.Attribute("value"),

                    bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("PART")).Key) == null
                        ? null
                        : (string)bomrow.Elements("bomcell"
                              ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                        y => y.Value.ToUpper().Contains("PART")).Key)
                              ?.Attribute("value") ?? "N/A",

                    bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("DES")).Key) == null
                        ? null
                        : (string)bomrow.Elements("bomcell"
                              ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                        y => y.Value.ToUpper().Contains("DES")).Key)
                              ?.Attribute("value") ?? "N/A",

                    bomrow.Elements("bomcell"
                    ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                              y => y.Value.ToUpper().Contains("Q")).Key) == null
                        ? null
                        : (int?)bomrow.Elements("bomcell"
                            ).FirstOrDefault(x => (int)x.Attribute("col_no") == openWith.FirstOrDefault(
                                                      y => y.Value.ToUpper().Contains("Q")).Key)
                            ?.Attribute("value")
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
                DoesItemExist();
                Ending(secondLevelTbl);
            }
            else
            {
                checkBox12.Checked = true;
                DoesItemExist();
                Ending(secondLevelTbl);
            }
        }

        private void Ending(DataTable secondLevelTbl)
        {
            checkBox18.Checked = true;
            dataGridView1.DataSource = secondLevelTbl;
            if ((checkBox1.Checked && checkBox9.Checked && checkBox11.Checked) ||
                (checkBox1.Checked == false && checkBox9.Checked && checkBox11.Checked))
            {
                dataGridView1.DataSource = secondLevelTbl;
            }
            else
            {
                txtBox.AppendText(Environment.NewLine + "Table contains errors!");
            }
        }

        private void Btn1_Send_Click_1(object sender, EventArgs e)
        {
            tbProgramLog.Clear();
            txtBox.Clear();
            if (checkBox16.Checked)
            {
                txtBox.AppendText(Environment.NewLine + "Cannot import a pre-existing Part Item as an Assembly");
            }
            else if (checkBox1.Checked && checkBox9.Checked && checkBox11.Checked && checkBox14.Checked)
            {
                txtBox.AppendText(Environment.NewLine + "START OF PROGRAM" + Environment.NewLine);
                AddThenModify();

            }
            else if (checkBox1.Checked && checkBox9.Checked && checkBox11.Checked && checkBox13.Checked & checkBox15.Checked)
            {
                txtBox.AppendText(Environment.NewLine + "START OF PROGRAM" + Environment.NewLine);
                AddThenModify();
            }
            else
            {
                txtBox.AppendText(Environment.NewLine + "Cannot import a table with errors" +
                                  Environment.NewLine + "or a complete BOM with sub BOMs!");
            }
        }

        private void AddThenModify()
        {
            InventoryAssemblyQuery();
            ItemAddAssembly();
            tbProgramLog.AppendText(Environment.NewLine);
            txtBox.AppendText(Environment.NewLine);
            txtBox.AppendText(Environment.NewLine + "Query assembly again");
            InventoryAssemblyQuery();
            txtBox.AppendText(Environment.NewLine + "END OF PROGRAM");
        }

        private void DoesItemExist()
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                sessionManager = new QBSessionManager();

                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                DoesItemAssemblyExistParameters(requestMsgSet);

                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;

                DoesItemAssemblyExistResponse(responseMsgSet);
            }
            catch (Exception e)
            {
                txtBox.AppendText(Environment.NewLine + e.Message);
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
        void DoesItemAssemblyExistParameters(IMsgSetRequest requestMsgSet)
        {
            IItemQuery itemQueryRq = requestMsgSet.AppendItemQueryRq();
            itemQueryRq.ORListQuery.ListFilter.ORNameFilter.NameFilter.MatchCriterion.SetValue(ENMatchCriterion
                .mcEndsWith);
            itemQueryRq.ORListQuery.ListFilter.ORNameFilter.NameFilter.Name.SetValue(topLevelTbl.Rows[0][0].ToString());
        }
        void DoesItemAssemblyExistResponse(IMsgSetResponse responseMsgSet)
        {
            if (responseMsgSet == null) return;
            IResponseList responseList = responseMsgSet.ResponseList;
            if (responseList == null) return;


            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);
                if (response.StatusCode >= 0)
                {
                    if (response.StatusCode.Equals(1))
                    {
                        checkBox14.Checked = true;
                    }
                    else if (response.StatusCode.Equals(0))
                    {
                        if (response.Detail != null)
                        {
                            ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                            if (responseType == ENResponseType.rtItemQueryRs)
                            {
                                IORItemRetList itemRetList = (IORItemRetList)response.Detail;
                                checkBox13.Checked = true;
                                ItemExistsAsPartOrAssembly(itemRetList);
                            }
                        }
                    }
                }
            }
        }
        void ItemExistsAsPartOrAssembly(IORItemRetList itemRetList)
        {
            if (itemRetList == null) return;
            for (int y = 0; y < itemRetList.Count; y++)
            {
                IORItemRet itemRet = itemRetList.GetAt(y);
                string topLevelSequence;
                string topLevelListId;
                if (itemRet.ItemInventoryAssemblyRet != null)
                {
                    checkBox15.Checked = true;
                }
                else if (itemRet.ItemInventoryRet != null)
                {
                    checkBox16.Checked = true;
                }
            }
        }

        private void InventoryAssemblyQuery()
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                sessionManager = new QBSessionManager();

                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                BuildAssemblyQuery(requestMsgSet, topLevelTbl);

                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                //tbProgramLog.AppendText(Environment.NewLine + "INVENTORY ASSEMBLY QUERY: " + requestMsgSet.ToXMLString());
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;
                WalkItemAssemblyQueryRs(responseMsgSet);
            }
            catch (Exception e)
            {
                tbProgramLog.AppendText(Environment.NewLine + e.Message);
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
        void BuildAssemblyQuery(IMsgSetRequest requestMsgSet, DataTable topLevelTbl)
        {
            IItemInventoryAssemblyQuery itemInventoryAssemblyQueryRq =
                requestMsgSet.AppendItemInventoryAssemblyQueryRq();
            IListWithClassFilter listWithClassFilter =
                itemInventoryAssemblyQueryRq.ORListQueryWithOwnerIDAndClass.ListWithClassFilter;
            listWithClassFilter.ORNameFilter.NameFilter.MatchCriterion.SetValue(ENMatchCriterion.mcEndsWith);
            //listWithClassFilter.ORNameFilter.NameFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
            //listWithClassFilter.ORNameFilter.NameFilter.MatchCriterion.SetValue(ENMatchCriterion.mcContains);
            listWithClassFilter.ORNameFilter.NameFilter.Name.SetValue(topLevelTbl.Rows[0][0].ToString());
        }
        void WalkItemAssemblyQueryRs(IMsgSetResponse responseMsgSet)
        {
            IResponseList responseList = responseMsgSet?.ResponseList;
            if (responseList == null) return;

            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);
                tbProgramLog.AppendText(Environment.NewLine + response.StatusCode + ": " + response.StatusMessage);

                if (response.StatusCode >= 0)
                {
                    if (response.Detail != null)
                    {
                        ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                        if (responseType == ENResponseType.rtItemInventoryAssemblyQueryRs)
                        {
                            IItemInventoryAssemblyRetList itemInventoryAssemblyRetList =
                                (IItemInventoryAssemblyRetList)response.Detail;
                            WalkItemAssemblyQueryRet(itemInventoryAssemblyRetList);
                        }
                    }
                }
            }
        }
        void WalkItemAssemblyQueryRet(IItemInventoryAssemblyRetList itemInventoryAssemblyRetList)
        {
            if (itemInventoryAssemblyRetList == null) return;
            string sequence = string.Empty;
            string listId = string.Empty;
            for (int x = 0; x < itemInventoryAssemblyRetList.Count; x++)
            {
                IItemInventoryAssemblyRet itemInventoryAssemblyRet = itemInventoryAssemblyRetList.GetAt(x);
                sequence = itemInventoryAssemblyRet.EditSequence.GetValue();
                listId = itemInventoryAssemblyRet.ListID.GetValue();
                tbProgramLog.AppendText(Environment.NewLine + "Assembly:" + Environment.NewLine + "Edit Sequence: " +
                                        sequence + Environment.NewLine + "List ID: " + listId + Environment.NewLine);
                txtBox.AppendText(Environment.NewLine + "End of Assembly query" + Environment.NewLine);
            }
            ItemQuery(sequence, listId);
        }

        private void ItemQuery(string sequence, string listId)
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                sessionManager = new QBSessionManager();

                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                BuildItemQuery(requestMsgSet, secondLevelTbl);

                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                //tbProgramLog.AppendText(Environment.NewLine + "ITEM QUERY: " + requestMsgSet.ToXMLString());
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;

                WalkItemQueryRs(responseMsgSet, sequence, listId);
            }
            catch (Exception e)
            {
                tbProgramLog.AppendText(Environment.NewLine + e.Message);
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
        void BuildItemQuery(IMsgSetRequest requestMsgSet, DataTable secondLevelTbl)
        {
            List<string> itemCodes = secondLevelTbl.AsEnumerable().Select(r => r.Field<string>("ItemCode")).ToList();
            foreach (string itemCode in itemCodes)
            {
                IItemQuery itemQueryRq = requestMsgSet.AppendItemQueryRq();
                itemQueryRq.ORListQuery.ListFilter.ORNameFilter.NameFilter.MatchCriterion.SetValue(ENMatchCriterion
                    .mcEndsWith);
                //itemQueryRq.ORListQuery.ListFilter.ORNameFilter.NameFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                //itemQueryRq.ORListQuery.ListFilter.ORNameFilter.NameFilter.MatchCriterion.SetValue(ENMatchCriterion.mcContains);
                itemQueryRq.ORListQuery.ListFilter.ORNameFilter.NameFilter.Name.SetValue(itemCode);
            }
        }
        void WalkItemQueryRs(IMsgSetResponse responseMsgSet, string sequence, string listId)
        {
            if (responseMsgSet == null) return;
            IResponseList responseList = responseMsgSet.ResponseList;
            if (responseList == null) return;

            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);
                if (response.StatusCode >= 0)
                {
                    if (response.StatusCode.Equals(1))
                    {
                        string itemNoError = response.RequestID;
                        FindTheNeededValues(itemNoError);
                    }
                    else if (response.StatusCode.Equals(0))
                    {
                        if (response.Detail != null)
                        {
                            ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                            if (responseType == ENResponseType.rtItemQueryRs)
                            {
                                IORItemRetList itemRetList = (IORItemRetList)response.Detail;
                                WalkItemQueryRet(itemRetList, sequence, listId);
                            }
                        }
                    }
                }
            }
        }
        void WalkItemQueryRet(IORItemRetList itemRetList, string sequence, string listId)
        {
            string itemListId = string.Empty;
            string itemName = string.Empty;
            string itemSequence = string.Empty;
            if (itemRetList == null) return;
            for (int y = 0; y < itemRetList.Count; y++)
            {
                IORItemRet itemRet = itemRetList.GetAt(y);
                if (itemRet.ItemInventoryAssemblyRet != null)
                {
                    itemListId = itemRet.ItemInventoryAssemblyRet.ListID.GetValue();
                    itemSequence = itemRet.ItemInventoryAssemblyRet.EditSequence.GetValue();
                    itemName = itemRet.ItemInventoryAssemblyRet.Name.GetValue();
                }
                else if (itemRet.ItemInventoryRet != null)
                {
                    itemListId = itemRet.ItemInventoryRet.ListID.GetValue();
                    itemSequence = itemRet.ItemInventoryRet.EditSequence.GetValue();
                    itemName = itemRet.ItemInventoryRet.Name.GetValue();
                }
                else if (itemRet.ItemNonInventoryRet != null)
                {
                    itemListId = itemRet.ItemNonInventoryRet.ListID.GetValue();
                    itemSequence = itemRet.ItemNonInventoryRet.EditSequence.GetValue();
                    itemName = itemRet.ItemNonInventoryRet.Name.GetValue();
                }
                else if (itemRet.ItemSubtotalRet != null)
                {
                    itemListId = itemRet.ItemSubtotalRet.ListID.GetValue();
                    itemSequence = itemRet.ItemSubtotalRet.EditSequence.GetValue();
                    itemName = itemRet.ItemSubtotalRet.Name.GetValue();
                }
                else if (itemRet.ItemDiscountRet != null)
                {
                    itemListId = itemRet.ItemDiscountRet.ListID.GetValue();
                    itemSequence = itemRet.ItemDiscountRet.EditSequence.GetValue();
                    itemName = itemRet.ItemDiscountRet.Name.GetValue();
                }
                else if (itemRet.ItemPaymentRet != null)
                {
                    itemListId = itemRet.ItemPaymentRet.ListID.GetValue();
                    itemSequence = itemRet.ItemPaymentRet.EditSequence.GetValue();
                    itemName = itemRet.ItemPaymentRet.Name.GetValue();
                }
                else if (itemRet.ItemSalesTaxRet != null)
                {
                    itemListId = itemRet.ItemSalesTaxRet.ListID.GetValue();
                    itemSequence = itemRet.ItemSalesTaxRet.EditSequence.GetValue();
                    itemName = itemRet.ItemSalesTaxRet.Name.GetValue();
                }
                else if (itemRet.ItemSalesTaxGroupRet != null)
                {
                    itemListId = itemRet.ItemSalesTaxGroupRet.ListID.GetValue();
                    itemSequence = itemRet.ItemSalesTaxGroupRet.EditSequence.GetValue();
                    itemName = itemRet.ItemSalesTaxGroupRet.Name.GetValue();
                }
                else if (itemRet.ItemGroupRet != null)
                {
                    itemListId = itemRet.ItemGroupRet.ListID.GetValue();
                    itemSequence = itemRet.ItemGroupRet.EditSequence.GetValue();
                    itemName = itemRet.ItemGroupRet.Name.GetValue();
                }
                else if (itemRet.ItemServiceRet != null)
                {
                    itemListId = itemRet.ItemServiceRet.ListID.GetValue();
                    itemSequence = itemRet.ItemServiceRet.EditSequence.GetValue();
                    itemName = itemRet.ItemServiceRet.Name.GetValue();
                }
                else if (itemRet.ItemOtherChargeRet != null)
                {
                    itemListId = itemRet.ItemOtherChargeRet.ListID.GetValue();
                    itemSequence = itemRet.ItemOtherChargeRet.EditSequence.GetValue();
                    itemName = itemRet.ItemOtherChargeRet.Name.GetValue();
                }
                tbProgramLog.AppendText(Environment.NewLine + "Lower level edit sequence: " + itemSequence +
                                        Environment.NewLine + "List ID: " + itemListId + Environment.NewLine);
                txtBox.AppendText(Environment.NewLine + "Item " + itemName + " will be modified");
            }
            ItemModify(sequence, listId);
        }

        private void ItemModify(string sequence, string listId)
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                sessionManager = new QBSessionManager();

                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                BuildItemModify(requestMsgSet, secondLevelTbl, topLevelTbl, sequence, listId);

                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                //tbProgramLog.AppendText(Environment.NewLine + "ITEM MODIFY: " + requestMsgSet.ToXMLString());
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;

                WalkItemModifyRs(responseMsgSet);
            }
            catch (Exception e)
            {
                tbProgramLog.AppendText(Environment.NewLine + e.Message);
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
        void BuildItemModify(IMsgSetRequest requestMsgSet, DataTable secondLevelTbl, DataTable topLevelTbl, string sequence, string listId)
        {
            IItemInventoryAssemblyMod itemInventoryAssemblyModRq = requestMsgSet.AppendItemInventoryAssemblyModRq();
            itemInventoryAssemblyModRq.ListID.SetValue(listId);
            itemInventoryAssemblyModRq.SalesDesc.SetValue(topLevelTbl.Rows[0][1].ToString());
            itemInventoryAssemblyModRq.PurchaseDesc.SetValue(topLevelTbl.Rows[0][1].ToString());
            itemInventoryAssemblyModRq.EditSequence.SetValue(sequence);

            for (int i = 0; i < secondLevelTbl.Rows.Count; i++)
            {
                IItemInventoryAssemblyLine itemInventoryAssemblyLine1 = itemInventoryAssemblyModRq
                    .ORItemInventoryAssemblyLine.ItemInventoryAssemblyLineList.Append();
                itemInventoryAssemblyLine1.ItemInventoryRef.FullName.SetValue(secondLevelTbl.Rows[i][1].ToString());
                itemInventoryAssemblyLine1.Quantity.SetValue(Convert.ToDouble(secondLevelTbl.Rows[i][4]));
            }
        }
        void WalkItemModifyRs(IMsgSetResponse responseMsgSet)
        {
            IResponseList responseList = responseMsgSet?.ResponseList;
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
        void WalkItemModifyRet(IItemInventoryAssemblyRet itemModifyRet)
        {
            if (itemModifyRet == null) return;
            string sequence = itemModifyRet.EditSequence.GetValue();
            string listId = itemModifyRet.ListID.GetValue();
            tbProgramLog.AppendText(Environment.NewLine + "Top level edit sequence: " + sequence + Environment.NewLine +
                                    "List ID: " + listId);
        }

        private void FindTheNeededValues(string itemNoError)
        {
            int col = Int32.Parse(itemNoError);
            subAssembly = new DataTable();
            string subItem = secondLevelTbl.Rows[col][1].ToString();
            string connectionString =
                @"Data Source=SQLSERVER\ITEMCODE;Initial Catalog=dat8121;Integrated Security=True";
            SqlDataAdapter dataAdapter = new SqlDataAdapter("SELECT a.[ItemCode], b.[Description], a.[itemType]" +
                                                            " FROM [PDMengineeringVault].[dbo].[v_Documents] a" +
                                                            " RIGHT JOIN [PDMengineeringVault].[dbo].[v_BOMData] b" +
                                                            " ON a.[Itemcode] = b.[ItemCode]" +
                                                            " WHERE (a.[ItemType] LIKE 'ass%' OR a.[ItemType] LIKE 'par%')" +
                                                            " AND(a.[Itemcode] = '" + subItem +
                                                            "' OR b.[ItemCode] = '" + subItem + "')"
                , connectionString);
            dataAdapter.Fill(subAssembly);
            subAssembly.Columns.Add("IncomeAccountRef", typeof(string));
            subAssembly.Columns.Add("COGSAccountRef", typeof(string));
            subAssembly.Columns.Add("AssetAccountRef", typeof(string));
            DataRow row = subAssembly.Rows[0];
            row[3] = IncomeAccount;
            row[4] = COGSAccount;
            row[5] = InventoryAssetAccount;
            tbProgramLog.AppendText(Environment.NewLine + "item in the query: " + subItem);
            string A = "Assembly";
            string P = "Part";
            if (row[2].ToString() == A)
            {
                tbProgramLog.AppendText(Environment.NewLine + "col1: " + row[0] + " col2: " + row[1] + " col3: " +
                                        row[2] + " col4: " + row[3] + " col5: " + row[4] + " col6: " + row[5]);
                ItemAddSubAssembly(subItem);
            }
            else if (row[2].ToString() == P)
            {
                tbProgramLog.AppendText(Environment.NewLine + "col1: " + row[0] + " col2: " + row[1] + " col3: " +
                                        row[2] + " col4: " + row[3] + " col5: " + row[4] + " col6: " + row[5]);
                ItemAddPart(subItem);
            }
            else
            {
                txtBox.AppendText(Environment.NewLine + "Cannot continue"
                                                      + Environment.NewLine + "Check the itemType");
            }
        }

        private void ItemAddAssembly()
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                sessionManager = new QBSessionManager();

                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                BuildItemAssemblyAddRq(requestMsgSet);

                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                //tbProgramLog.AppendText(Environment.NewLine + "ITEM ADD ASSEMBLY: " + requestMsgSet.ToXMLString());
                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;

                WalkItemAssemblyAddRs(responseMsgSet);
            }
            catch (Exception e)
            {
                tbProgramLog.AppendText(Environment.NewLine + e.Message);
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
        void BuildItemAssemblyAddRq(IMsgSetRequest requestMsgSet)
        {
            IItemInventoryAssemblyAdd itemInventoryAssemblyAddRq = requestMsgSet.AppendItemInventoryAssemblyAddRq();
            DataRow row = topLevelTbl.Rows[0];
            itemInventoryAssemblyAddRq.Name.SetValue(row[0].ToString());
            itemInventoryAssemblyAddRq.SalesDesc.SetValue(row[1].ToString());
            itemInventoryAssemblyAddRq.PurchaseDesc.SetValue(row[1].ToString());
            itemInventoryAssemblyAddRq.IncomeAccountRef.ListID.SetValue(IncomeAccount);
            itemInventoryAssemblyAddRq.COGSAccountRef.ListID.SetValue(COGSAccount);
            itemInventoryAssemblyAddRq.AssetAccountRef.ListID.SetValue(InventoryAssetAccount);
        }
        void WalkItemAssemblyAddRs(IMsgSetResponse responseMsgSet)
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
                        if (responseType == ENResponseType.rtItemInventoryAssemblyAddRs)
                        {
                            IItemInventoryAssemblyRet itemInventoryAssemblyRet =
                                (IItemInventoryAssemblyRet)response.Detail;
                            WalkItemAssemblyAddRet(itemInventoryAssemblyRet);
                        }
                    }
                }
            }
        }
        void WalkItemAssemblyAddRet(IItemInventoryAssemblyRet itemInventoryAssemblyRet)
        {
            if (itemInventoryAssemblyRet == null) return;
            string sequence = itemInventoryAssemblyRet.EditSequence.GetValue();
            string listId = itemInventoryAssemblyRet.ListID.GetValue();
            tbProgramLog.AppendText(Environment.NewLine + "Edit sequence: " + sequence + Environment.NewLine +
                                    "List ID: " + listId);
            InventoryAssemblyQuery();
        }

        private void ItemAddSubAssembly(string subItem)
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                sessionManager = new QBSessionManager();

                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                BuildItemSubAssemblyAddRq(requestMsgSet);

                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                //tbProgramLog.AppendText(Environment.NewLine + "ITEM ADD ASSEMBLY: " + Environment.NewLine + requestMsgSet.ToXMLString());
                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;

                WalkItemSubAssemblyAddRs(responseMsgSet, subItem);
            }
            catch (Exception e)
            {
                tbProgramLog.AppendText(Environment.NewLine + e.Message);
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
        void BuildItemSubAssemblyAddRq(IMsgSetRequest requestMsgSet)
        {
            IItemInventoryAssemblyAdd itemInventoryAssemblyAddRq = requestMsgSet.AppendItemInventoryAssemblyAddRq();
            DataRow row = subAssembly.Rows[0];
            itemInventoryAssemblyAddRq.Name.SetValue(row[0].ToString());
            itemInventoryAssemblyAddRq.SalesDesc.SetValue(row[1].ToString());
            itemInventoryAssemblyAddRq.PurchaseDesc.SetValue(row[1].ToString());
            itemInventoryAssemblyAddRq.IncomeAccountRef.ListID.SetValue(IncomeAccount);
            itemInventoryAssemblyAddRq.COGSAccountRef.ListID.SetValue(COGSAccount);
            itemInventoryAssemblyAddRq.AssetAccountRef.ListID.SetValue(InventoryAssetAccount);
        }
        void WalkItemSubAssemblyAddRs(IMsgSetResponse responseMsgSet, string subItem)
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
                        if (responseType == ENResponseType.rtItemInventoryAssemblyAddRs)
                        {
                            IItemInventoryAssemblyRet itemInventoryAssemblyRet =
                                (IItemInventoryAssemblyRet)response.Detail;
                            WalkItemSubAssemblyAddRet(itemInventoryAssemblyRet, subItem);
                        }
                    }
                }
            }
        }
        void WalkItemSubAssemblyAddRet(IItemInventoryAssemblyRet itemInventoryAssemblyRet, string subItem)
        {
            if (itemInventoryAssemblyRet == null) return;
            string sequence = itemInventoryAssemblyRet.EditSequence.GetValue();
            string listId = itemInventoryAssemblyRet.ListID.GetValue();
            tbProgramLog.AppendText(Environment.NewLine + "Edit sequence: " + sequence + Environment.NewLine +
                                    "List ID: " + listId);
            SubQuery(subItem);
        }


        private void ItemAddPart(string subItem)
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                sessionManager = new QBSessionManager();

                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                BuildItemPartAddRq(requestMsgSet);

                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                // tbProgramLog.AppendText(Environment.NewLine + "ITEM ADD PART: " + Environment.NewLine + requestMsgSet.ToXMLString());
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;

                WalkItemPartAddRs(responseMsgSet, subItem);
            }
            catch (Exception e)
            {
                tbProgramLog.AppendText(Environment.NewLine + e.Message);
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
        void BuildItemPartAddRq(IMsgSetRequest requestMsgSet)
        {
            IItemInventoryAdd itemInventoryAddRq = requestMsgSet.AppendItemInventoryAddRq();
            DataRow row = subAssembly.Rows[0];
            itemInventoryAddRq.Name.SetValue(row[0].ToString());
            itemInventoryAddRq.SalesDesc.SetValue(row[1].ToString());
            itemInventoryAddRq.PurchaseDesc.SetValue(row[1].ToString());
            itemInventoryAddRq.IncomeAccountRef.ListID.SetValue(IncomeAccount);
            itemInventoryAddRq.COGSAccountRef.ListID.SetValue(COGSAccount);
            itemInventoryAddRq.AssetAccountRef.ListID.SetValue(InventoryAssetAccount);

        }
        void WalkItemPartAddRs(IMsgSetResponse responseMsgSet, string subItem)
        {
            IResponseList responseList = responseMsgSet?.ResponseList;
            if (responseList == null) return;

            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);

                if (response.StatusCode >= 0)
                {
                    if (response.Detail != null)
                    {
                        ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                        if (responseType == ENResponseType.rtItemInventoryAddRs)
                        {
                            IItemInventoryRet itemInventoryRet = (IItemInventoryRet)response.Detail;
                            WalkItemPartAddRet(itemInventoryRet, subItem);
                        }
                    }
                }
            }
        }
        void WalkItemPartAddRet(IItemInventoryRet itemInventoryRet, string subItem)
        {
            if (itemInventoryRet == null) return;
            string partSequence = itemInventoryRet.EditSequence.GetValue();
            string partListId = itemInventoryRet.ListID.GetValue();
            tbProgramLog.AppendText(Environment.NewLine + "Edit sequence: " + partSequence + Environment.NewLine +
                                    "List ID: " + partListId);
            SubQuery(subItem);
        }

        private void SubQuery(string subItem)
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                sessionManager = new QBSessionManager();

                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 13, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                BuildSubQuery(requestMsgSet, subItem);

                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                //tbProgramLog.AppendText(Environment.NewLine + "ITEM QUERY: " + requestMsgSet.ToXMLString());
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;

                WalkSubQueryRs(responseMsgSet);
            }
            catch (Exception e)
            {
                tbProgramLog.AppendText(Environment.NewLine + e.Message);
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
        void BuildSubQuery(IMsgSetRequest requestMsgSet, string subItem)
        {
            IItemQuery itemQueryRq = requestMsgSet.AppendItemQueryRq();
            itemQueryRq.ORListQuery.ListFilter.ORNameFilter.NameFilter.MatchCriterion.SetValue(ENMatchCriterion
                .mcEndsWith);
            //itemQueryRq.ORListQuery.ListFilter.ORNameFilter.NameFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
            //itemQueryRq.ORListQuery.ListFilter.ORNameFilter.NameFilter.MatchCriterion.SetValue(ENMatchCriterion.mcContains);
            itemQueryRq.ORListQuery.ListFilter.ORNameFilter.NameFilter.Name.SetValue(subItem);

        }
        void WalkSubQueryRs(IMsgSetResponse responseMsgSet)
        {
            if (responseMsgSet == null) return;
            IResponseList responseList = responseMsgSet.ResponseList;
            if (responseList == null) return;

            for (int i = 0; i < responseList.Count; i++)
            {
                IResponse response = responseList.GetAt(i);
                if (response.StatusCode >= 0)
                {
                    if (response.StatusCode.Equals(1))
                    {
                        string itemNoError = response.RequestID;
                        FindTheNeededValues(itemNoError);
                    }
                    else if (response.StatusCode.Equals(0))
                    {
                        if (response.Detail != null)
                        {
                            ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                            if (responseType == ENResponseType.rtItemQueryRs)
                            {
                                IORItemRetList itemRetList = (IORItemRetList)response.Detail;
                                WalkSubQueryRet(itemRetList);
                            }
                        }
                    }
                }
            }
        }
        void WalkSubQueryRet(IORItemRetList itemRetList)
        {
            string itemListId = string.Empty;
            string itemName = string.Empty;
            string itemSequence = string.Empty;
            if (itemRetList == null) return;
            for (int y = 0; y < itemRetList.Count; y++)
            {
                IORItemRet itemRet = itemRetList.GetAt(y);
                if (itemRet.ItemInventoryAssemblyRet != null)
                {
                    itemListId = itemRet.ItemInventoryAssemblyRet.ListID.GetValue();
                    itemSequence = itemRet.ItemInventoryAssemblyRet.EditSequence.GetValue();
                    itemName = itemRet.ItemInventoryAssemblyRet.Name.GetValue();
                }
                else if (itemRet.ItemInventoryRet != null)
                {
                    itemListId = itemRet.ItemInventoryRet.ListID.GetValue();
                    itemSequence = itemRet.ItemInventoryRet.EditSequence.GetValue();
                    itemName = itemRet.ItemInventoryRet.Name.GetValue();
                }
                tbProgramLog.AppendText(Environment.NewLine + "Lower level edit sequence: " + itemSequence +
                                        Environment.NewLine + "List ID: " + itemListId + Environment.NewLine);
                txtBox.AppendText(Environment.NewLine + "Item " + itemName + " will be modified");
            }
        }

        private void BOM_Click(object sender, EventArgs e)
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
            checkBox13.Checked = false;
            checkBox14.Checked = false;
            checkBox15.Checked = false;
            checkBox16.Checked = false;
            checkBox17.Checked = false;
            checkBox18.Checked = false;
            tbProgramLog.Clear();
            txtBox.Clear();
            textBox1.Clear();
            dataGridView1.DataSource = null;
            dataGridView2.DataSource = null;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = @"xml files (*.xml)|*.xml";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    fileName = openFileDialog.FileName;
                    try
                    {
                        XDocument doc = XDocument.Load(fileName);
                        string matchFirst = @"^[1-2][0-9][0-9][0-9][0-9][0-9]00";
                        foreach (XElement bom in doc.Descendants("bom"))
                        {
                            string nameData = bom.Attribute("name")?.Value;
                            string v = Regex.Match(nameData, matchFirst).Value;
                            int input = int.Parse(v);
                            toplvlBOM(input);
                            entireBom(input);
                        }
                    }
                    catch (System.Xml.XmlException ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }

        private void toplvlBOM(int input)
        {
            toplvlBOMtbl = new DataTable();
            string connectionString =
                @"Data Source=SQLSERVER\ITEMCODE;Initial Catalog=dat8121;Integrated Security=True";
            SqlDataAdapter dataAdapter = new SqlDataAdapter(
                "SELECT ItemCode AS 'Item Code', QBSalesDesc AS Description," +
                " Cast(QBQtyOnHand AS decimal(10)) AS QTY" +
                " FROM[dat8121].[dbo].[v_ItemCode_QB]" +
                " WHERE ItemCode = '" + input + "'"
                , connectionString);

            dataAdapter.Fill(toplvlBOMtbl);

            bool isTrue = toplvlBOMtbl.Rows.Count > 0;
            if (isTrue)
            {
                textBox1.Text = (toplvlBOMtbl.Rows[0][0] + Environment.NewLine);
                dataGridView2.DataSource = toplvlBOMtbl;
            }
            txtBox.AppendText(Environment.NewLine + "Complete BOM displayed may contain errors!");
        }

        private DataTable entireBom(int input)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter();
            String connectionString =
                @"Data Source=SQLSERVER\ITEMCODE;Initial Catalog=dat8121;Integrated Security=True";
            DataTable dtReturnValue = new DataTable();
            DataTable dtTemp2 = new DataTable();
            dataAdapter = new SqlDataAdapter(
                "SELECT cast(v_ItemInventoryAssemblyLine.ItemInventoryAssemblyLnSeqNo as nvarchar(50)) 'NO.' " +
                ", v_ItemInventoryAssemblyLine.ItemInventoryAssemblyLnItemInventoryRefFullName as 'ITEM CODE'" +
                ",COALESCE(v_ItemInventoryAssemblyLine.PartNumber, 'Unavailable') as 'PART NUMBER'" +
                //",COALESCE(v_ItemInventoryAssemblyLine.Description, 'Unavailable') as 'DESCRIPTION'" +
                ",COALESCE((SELECT Description FROM [dat8121].[dbo].[v_ItemCode] AS iCode " +
                "WHERE iCode.itemcode = v_ItemInventoryAssemblyLine.ItemInventoryAssemblyLnItemInventoryRefFullName), 'Unavailable') AS DESCRIPTION" +
                ",cast(v_ItemInventoryAssemblyLine.ItemInventoryAssemblyLnQuantity as decimal(10)) 'QTY'" +
                //",((select (A.QBSalesPrice - A.QBPurchaseCost)" +
                //" FROM [dat8121].[dbo].[v_ItemCode_QB] AS A " +
                //" WHERE TRY_CAST(A.ItemCode As int) = TRY_CAST(v_iteminventoryassemblyline.ItemInventoryAssemblyLnItemInventoryRefFullName AS INT) " +
                //" ) * ItemInventoryAssemblyLnQuantity) AS '+Gain/-Loss' " +
                ",v_ItemInventoryAssemblyLine.counts" +
                " FROM [QODBC].[dbo].v_ItemInventoryAssemblyLine AS v_ItemInventoryAssemblyLine" +
                " LEFT JOIN [dat8121].[dbo].[v_ItemCode_QB] AS v_itemcode_qb" +
                " ON v_ItemInventoryAssemblyLine.itemcode = v_itemcode_qb.itemcode" +
                " WHERE v_iteminventoryassemblyline.itemcode = " + input +
                " ORDER BY ItemInventoryAssemblyLnSeqNo", connectionString);
            dataAdapter.Fill(dtTemp2);
            dtReturnValue = dtTemp2.Clone();
            foreach (DataRow row in dtTemp2.Rows)
            {
                dtReturnValue.ImportRow(row);
                if (row["counts"].ToString() != "0")
                {
                    string seqno = row["NO."].ToString();
                    double val = Double.Parse(row["QTY"].ToString());
                    DataTable dtTemp = new DataTable();
                    dtTemp = entireBom(int.Parse(row["ITEM CODE"].ToString()));
                    foreach (DataRow innerRow in dtTemp.Rows)
                    {
                        double qty = Double.Parse(innerRow["QTY"].ToString());
                        innerRow["NO."] = seqno + "." + innerRow["NO."];
                        innerRow["QTY"] = val * qty;
                        dtReturnValue.ImportRow(innerRow);
                    }

                }
            }
            dtReturnValue.Columns.Remove("counts");
            dataGridView1.DataSource = dtReturnValue;
            checkBox17.Checked = true;
            return dtReturnValue;
        }
    }
}