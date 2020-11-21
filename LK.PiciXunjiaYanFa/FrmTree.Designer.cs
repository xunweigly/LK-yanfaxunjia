namespace LKU8.shoukuan
{
    partial class FrmTree
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.产品 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.ddate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cmaker = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.采购员 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.供应商 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CAS = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.名称 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.规格 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.数量1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.数量2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.数量3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.询价1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.询价2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.询价3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.备注1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.备注2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.备注3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.HighlightText;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.产品);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1072, 114);
            this.panel1.TabIndex = 0;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(729, 29);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(62, 18);
            this.label5.TabIndex = 5;
            this.label5.Text = "label5";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(661, 29);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(62, 18);
            this.label4.TabIndex = 4;
            this.label4.Text = "规格：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(482, 29);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(62, 18);
            this.label3.TabIndex = 3;
            this.label3.Text = "label3";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(387, 29);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 18);
            this.label2.TabIndex = 2;
            this.label2.Text = "名称：";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(195, 29);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(62, 18);
            this.label1.TabIndex = 1;
            this.label1.Text = "label1";
            // 
            // 产品
            // 
            this.产品.AutoSize = true;
            this.产品.Location = new System.Drawing.Point(64, 29);
            this.产品.Name = "产品";
            this.产品.Size = new System.Drawing.Size(98, 18);
            this.产品.TabIndex = 0;
            this.产品.Text = "产品编码：";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.HighlightText;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ddate,
            this.cmaker,
            this.采购员,
            this.供应商,
            this.CAS,
            this.名称,
            this.规格,
            this.数量1,
            this.数量2,
            this.数量3,
            this.询价1,
            this.询价2,
            this.询价3,
            this.备注1,
            this.备注2,
            this.备注3});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 114);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 30;
            this.dataGridView1.Size = new System.Drawing.Size(1072, 578);
            this.dataGridView1.TabIndex = 3;
            // 
            // ddate
            // 
            this.ddate.DataPropertyName = "ddate";
            this.ddate.HeaderText = "询价日期";
            this.ddate.Name = "ddate";
            this.ddate.ReadOnly = true;
            // 
            // cmaker
            // 
            this.cmaker.DataPropertyName = "cmaker";
            this.cmaker.HeaderText = "业务员";
            this.cmaker.Name = "cmaker";
            this.cmaker.ReadOnly = true;
            // 
            // 采购员
            // 
            this.采购员.DataPropertyName = "gpxjrxm";
            this.采购员.HeaderText = "采购员";
            this.采购员.Name = "采购员";
            this.采购员.ReadOnly = true;
            // 
            // 供应商
            // 
            this.供应商.DataPropertyName = "gys";
            this.供应商.HeaderText = "供应商";
            this.供应商.Name = "供应商";
            this.供应商.ReadOnly = true;
            // 
            // CAS
            // 
            this.CAS.DataPropertyName = "cinvcode";
            this.CAS.HeaderText = "CAS";
            this.CAS.Name = "CAS";
            this.CAS.ReadOnly = true;
            // 
            // 名称
            // 
            this.名称.DataPropertyName = "cinvname";
            this.名称.HeaderText = "名称";
            this.名称.Name = "名称";
            this.名称.ReadOnly = true;
            // 
            // 规格
            // 
            this.规格.DataPropertyName = "cinvstd";
            this.规格.HeaderText = "规格";
            this.规格.Name = "规格";
            this.规格.ReadOnly = true;
            // 
            // 数量1
            // 
            this.数量1.DataPropertyName = "cqty1";
            this.数量1.HeaderText = "数量1";
            this.数量1.Name = "数量1";
            this.数量1.ReadOnly = true;
            // 
            // 数量2
            // 
            this.数量2.DataPropertyName = "cqty2";
            this.数量2.HeaderText = "数量2";
            this.数量2.Name = "数量2";
            this.数量2.ReadOnly = true;
            // 
            // 数量3
            // 
            this.数量3.DataPropertyName = "cqty3";
            this.数量3.HeaderText = "数量3";
            this.数量3.Name = "数量3";
            this.数量3.ReadOnly = true;
            // 
            // 询价1
            // 
            this.询价1.DataPropertyName = "xunjia1";
            dataGridViewCellStyle1.NullValue = null;
            this.询价1.DefaultCellStyle = dataGridViewCellStyle1;
            this.询价1.HeaderText = "询价1";
            this.询价1.Name = "询价1";
            this.询价1.ReadOnly = true;
            // 
            // 询价2
            // 
            this.询价2.DataPropertyName = "xunjia2";
            this.询价2.HeaderText = "询价2";
            this.询价2.Name = "询价2";
            this.询价2.ReadOnly = true;
            // 
            // 询价3
            // 
            this.询价3.DataPropertyName = "xunjia3";
            this.询价3.HeaderText = "询价3";
            this.询价3.Name = "询价3";
            this.询价3.ReadOnly = true;
            // 
            // 备注1
            // 
            this.备注1.DataPropertyName = "bz1";
            this.备注1.HeaderText = "备注1";
            this.备注1.Name = "备注1";
            this.备注1.ReadOnly = true;
            // 
            // 备注2
            // 
            this.备注2.DataPropertyName = "bz2";
            this.备注2.HeaderText = "备注2";
            this.备注2.Name = "备注2";
            this.备注2.ReadOnly = true;
            // 
            // 备注3
            // 
            this.备注3.DataPropertyName = "bz3";
            this.备注3.HeaderText = "备注3";
            this.备注3.Name = "备注3";
            this.备注3.ReadOnly = true;
            // 
            // FrmTree
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1072, 692);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.panel1);
            this.Name = "FrmTree";
            this.Text = "询价历史";
            this.Load += new System.EventHandler(this.FrmTree_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label 产品;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DataGridViewTextBoxColumn ddate;
        private System.Windows.Forms.DataGridViewTextBoxColumn cmaker;
        private System.Windows.Forms.DataGridViewTextBoxColumn 采购员;
        private System.Windows.Forms.DataGridViewTextBoxColumn 供应商;
        private System.Windows.Forms.DataGridViewTextBoxColumn CAS;
        private System.Windows.Forms.DataGridViewTextBoxColumn 名称;
        private System.Windows.Forms.DataGridViewTextBoxColumn 规格;
        private System.Windows.Forms.DataGridViewTextBoxColumn 数量1;
        private System.Windows.Forms.DataGridViewTextBoxColumn 数量2;
        private System.Windows.Forms.DataGridViewTextBoxColumn 数量3;
        private System.Windows.Forms.DataGridViewTextBoxColumn 询价1;
        private System.Windows.Forms.DataGridViewTextBoxColumn 询价2;
        private System.Windows.Forms.DataGridViewTextBoxColumn 询价3;
        private System.Windows.Forms.DataGridViewTextBoxColumn 备注1;
        private System.Windows.Forms.DataGridViewTextBoxColumn 备注2;
        private System.Windows.Forms.DataGridViewTextBoxColumn 备注3;

    }
}