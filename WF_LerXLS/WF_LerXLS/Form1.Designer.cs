namespace WF_LerXLS
{
    partial class Form1
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
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btn_Arquivo_Original = new System.Windows.Forms.Button();
            this.lblarquivo_origem = new System.Windows.Forms.Label();
            this.lblarquivo_comparativo = new System.Windows.Forms.Label();
            this.btn_Arquivo_Comparativo = new System.Windows.Forms.Button();
            this.btn_comparar = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lstbox_campos = new System.Windows.Forms.ListBox();
            this.dtgv1 = new System.Windows.Forms.DataGridView();
            this.btn_Exporta_Excel = new System.Windows.Forms.Button();
            this.dtgv2 = new System.Windows.Forms.DataGridView();
            this.lblerros = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dtgv1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtgv2)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btn_Arquivo_Original
            // 
            this.btn_Arquivo_Original.Location = new System.Drawing.Point(80, 24);
            this.btn_Arquivo_Original.Name = "btn_Arquivo_Original";
            this.btn_Arquivo_Original.Size = new System.Drawing.Size(192, 23);
            this.btn_Arquivo_Original.TabIndex = 1;
            this.btn_Arquivo_Original.Text = "Selecione o Arquivo Original";
            this.btn_Arquivo_Original.UseVisualStyleBackColor = true;
            this.btn_Arquivo_Original.Click += new System.EventHandler(this.btn_Arquivo_Original_Click);
            // 
            // lblarquivo_origem
            // 
            this.lblarquivo_origem.AutoSize = true;
            this.lblarquivo_origem.Location = new System.Drawing.Point(16, 56);
            this.lblarquivo_origem.MaximumSize = new System.Drawing.Size(250, 20);
            this.lblarquivo_origem.Name = "lblarquivo_origem";
            this.lblarquivo_origem.Size = new System.Drawing.Size(0, 13);
            this.lblarquivo_origem.TabIndex = 2;
            // 
            // lblarquivo_comparativo
            // 
            this.lblarquivo_comparativo.AutoSize = true;
            this.lblarquivo_comparativo.Location = new System.Drawing.Point(504, 56);
            this.lblarquivo_comparativo.MaximumSize = new System.Drawing.Size(250, 20);
            this.lblarquivo_comparativo.Name = "lblarquivo_comparativo";
            this.lblarquivo_comparativo.Size = new System.Drawing.Size(0, 13);
            this.lblarquivo_comparativo.TabIndex = 3;
            // 
            // btn_Arquivo_Comparativo
            // 
            this.btn_Arquivo_Comparativo.Location = new System.Drawing.Point(496, 24);
            this.btn_Arquivo_Comparativo.Name = "btn_Arquivo_Comparativo";
            this.btn_Arquivo_Comparativo.Size = new System.Drawing.Size(192, 23);
            this.btn_Arquivo_Comparativo.TabIndex = 4;
            this.btn_Arquivo_Comparativo.Text = "Selecione o Arquivo de Verificação";
            this.btn_Arquivo_Comparativo.UseVisualStyleBackColor = true;
            this.btn_Arquivo_Comparativo.Click += new System.EventHandler(this.btn_Arquivo_Comparativo_Click);
            // 
            // btn_comparar
            // 
            this.btn_comparar.Location = new System.Drawing.Point(272, 128);
            this.btn_comparar.Name = "btn_comparar";
            this.btn_comparar.Size = new System.Drawing.Size(224, 23);
            this.btn_comparar.TabIndex = 6;
            this.btn_comparar.Text = "Comparar";
            this.btn_comparar.UseVisualStyleBackColor = true;
            this.btn_comparar.Click += new System.EventHandler(this.btn_comparar_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(272, 128);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(224, 23);
            this.progressBar1.TabIndex = 7;
            this.progressBar1.Visible = false;
            // 
            // lstbox_campos
            // 
            this.lstbox_campos.FormattingEnabled = true;
            this.lstbox_campos.Location = new System.Drawing.Point(272, 24);
            this.lstbox_campos.Name = "lstbox_campos";
            this.lstbox_campos.Size = new System.Drawing.Size(224, 95);
            this.lstbox_campos.TabIndex = 8;
            // 
            // dtgv1
            // 
            this.dtgv1.AllowUserToAddRows = false;
            this.dtgv1.AllowUserToDeleteRows = false;
            this.dtgv1.AllowUserToResizeRows = false;
            this.dtgv1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.ColumnHeader;
            this.dtgv1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllHeaders;
            this.dtgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtgv1.Location = new System.Drawing.Point(8, 160);
            this.dtgv1.Name = "dtgv1";
            this.dtgv1.Size = new System.Drawing.Size(768, 152);
            this.dtgv1.TabIndex = 9;
            // 
            // btn_Exporta_Excel
            // 
            this.btn_Exporta_Excel.Location = new System.Drawing.Point(8, 336);
            this.btn_Exporta_Excel.Name = "btn_Exporta_Excel";
            this.btn_Exporta_Excel.Size = new System.Drawing.Size(224, 23);
            this.btn_Exporta_Excel.TabIndex = 10;
            this.btn_Exporta_Excel.Text = "Exportar para Excel";
            this.btn_Exporta_Excel.UseVisualStyleBackColor = true;
            this.btn_Exporta_Excel.Click += new System.EventHandler(this.btn_Exporta_Excel_Click);
            // 
            // dtgv2
            // 
            this.dtgv2.AllowUserToAddRows = false;
            this.dtgv2.AllowUserToDeleteRows = false;
            this.dtgv2.AllowUserToResizeColumns = false;
            this.dtgv2.AllowUserToResizeRows = false;
            this.dtgv2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.ColumnHeader;
            this.dtgv2.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllHeaders;
            this.dtgv2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtgv2.Location = new System.Drawing.Point(232, 336);
            this.dtgv2.Name = "dtgv2";
            this.dtgv2.Size = new System.Drawing.Size(320, 120);
            this.dtgv2.TabIndex = 11;
            // 
            // lblerros
            // 
            this.lblerros.AutoSize = true;
            this.lblerros.Location = new System.Drawing.Point(16, 144);
            this.lblerros.Name = "lblerros";
            this.lblerros.Size = new System.Drawing.Size(115, 13);
            this.lblerros.TabIndex = 12;
            this.lblerros.Text = "Detalhamento de Erros";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(232, 320);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(88, 13);
            this.label1.TabIndex = 13;
            this.label1.Text = "Subtotal de Erros";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 561);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblerros);
            this.Controls.Add(this.dtgv2);
            this.Controls.Add(this.btn_Exporta_Excel);
            this.Controls.Add(this.dtgv1);
            this.Controls.Add(this.lstbox_campos);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btn_comparar);
            this.Controls.Add(this.btn_Arquivo_Comparativo);
            this.Controls.Add(this.lblarquivo_comparativo);
            this.Controls.Add(this.lblarquivo_origem);
            this.Controls.Add(this.btn_Arquivo_Original);
            this.Name = "Form1";
            this.Text = "Leitura de Arquivo xls";
            ((System.ComponentModel.ISupportInitialize)(this.dtgv1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dtgv2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btn_Arquivo_Original;
        private System.Windows.Forms.Label lblarquivo_origem;
        private System.Windows.Forms.Label lblarquivo_comparativo;
        private System.Windows.Forms.Button btn_Arquivo_Comparativo;
        private System.Windows.Forms.Button btn_comparar;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ListBox lstbox_campos;
        private System.Windows.Forms.DataGridView dtgv1;
        private System.Windows.Forms.Button btn_Exporta_Excel;
        private System.Windows.Forms.DataGridView dtgv2;
        private System.Windows.Forms.Label lblerros;
        private System.Windows.Forms.Label label1;
    }
}

