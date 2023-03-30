
namespace WindowsFormsApp1
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
            this.btnAbrirSW = new System.Windows.Forms.Button();
            this.Visivel = new System.Windows.Forms.CheckBox();
            this.btnFecharSW = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnExportar = new System.Windows.Forms.Button();
            this.btnFechar = new System.Windows.Forms.Button();
            this.btnAbrir = new System.Windows.Forms.Button();
            this.labelSW = new System.Windows.Forms.Label();
            this.btnBuscaProp = new System.Windows.Forms.Button();
            this.btnCria_AtualizaProp = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnAbrirSW
            // 
            this.btnAbrirSW.Location = new System.Drawing.Point(24, 21);
            this.btnAbrirSW.Name = "btnAbrirSW";
            this.btnAbrirSW.Size = new System.Drawing.Size(118, 65);
            this.btnAbrirSW.TabIndex = 0;
            this.btnAbrirSW.Text = "Abrir SW";
            this.btnAbrirSW.UseVisualStyleBackColor = true;
            this.btnAbrirSW.Click += new System.EventHandler(this.btnAbrirSW_Click);
            // 
            // Visivel
            // 
            this.Visivel.AutoSize = true;
            this.Visivel.Location = new System.Drawing.Point(203, 44);
            this.Visivel.Name = "Visivel";
            this.Visivel.Size = new System.Drawing.Size(70, 21);
            this.Visivel.TabIndex = 1;
            this.Visivel.Text = "Visível";
            this.Visivel.UseVisualStyleBackColor = true;
            this.Visivel.CheckedChanged += new System.EventHandler(this.chkinv_CheckedChanged);
            // 
            // btnFecharSW
            // 
            this.btnFecharSW.Location = new System.Drawing.Point(24, 92);
            this.btnFecharSW.Name = "btnFecharSW";
            this.btnFecharSW.Size = new System.Drawing.Size(118, 65);
            this.btnFecharSW.TabIndex = 2;
            this.btnFecharSW.Text = "Fechar SW";
            this.btnFecharSW.UseVisualStyleBackColor = true;
            this.btnFecharSW.Click += new System.EventHandler(this.btnFecharSW_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnAbrirSW);
            this.groupBox1.Controls.Add(this.Visivel);
            this.groupBox1.Controls.Add(this.btnFecharSW);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(313, 176);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "SW";
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnCria_AtualizaProp);
            this.groupBox2.Controls.Add(this.btnBuscaProp);
            this.groupBox2.Controls.Add(this.btnExportar);
            this.groupBox2.Controls.Add(this.btnFechar);
            this.groupBox2.Controls.Add(this.btnAbrir);
            this.groupBox2.Location = new System.Drawing.Point(12, 194);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(313, 225);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Arquivo";
            // 
            // btnExportar
            // 
            this.btnExportar.Location = new System.Drawing.Point(24, 157);
            this.btnExportar.Name = "btnExportar";
            this.btnExportar.Size = new System.Drawing.Size(118, 43);
            this.btnExportar.TabIndex = 6;
            this.btnExportar.Text = "Exportar";
            this.btnExportar.UseVisualStyleBackColor = true;
            this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
            // 
            // btnFechar
            // 
            this.btnFechar.Location = new System.Drawing.Point(24, 80);
            this.btnFechar.Name = "btnFechar";
            this.btnFechar.Size = new System.Drawing.Size(118, 50);
            this.btnFechar.TabIndex = 1;
            this.btnFechar.Text = "Fechar Arquivo";
            this.btnFechar.UseVisualStyleBackColor = true;
            this.btnFechar.Click += new System.EventHandler(this.btnFechar_Click);
            // 
            // btnAbrir
            // 
            this.btnAbrir.Location = new System.Drawing.Point(24, 21);
            this.btnAbrir.Name = "btnAbrir";
            this.btnAbrir.Size = new System.Drawing.Size(118, 53);
            this.btnAbrir.TabIndex = 0;
            this.btnAbrir.Text = "Abrir Arquivo";
            this.btnAbrir.UseVisualStyleBackColor = true;
            this.btnAbrir.Click += new System.EventHandler(this.btnAbrir_Click);
            // 
            // labelSW
            // 
            this.labelSW.AutoSize = true;
            this.labelSW.Location = new System.Drawing.Point(9, 424);
            this.labelSW.Name = "labelSW";
            this.labelSW.Size = new System.Drawing.Size(53, 17);
            this.labelSW.TabIndex = 5;
            this.labelSW.Text = "Pronto!";
            // 
            // btnBuscaProp
            // 
            this.btnBuscaProp.Location = new System.Drawing.Point(148, 21);
            this.btnBuscaProp.Name = "btnBuscaProp";
            this.btnBuscaProp.Size = new System.Drawing.Size(159, 53);
            this.btnBuscaProp.TabIndex = 6;
            this.btnBuscaProp.Text = "Busca Propriedade";
            this.btnBuscaProp.UseVisualStyleBackColor = true;
            this.btnBuscaProp.Click += new System.EventHandler(this.btnBuscaProp_Click);
            // 
            // btnCria_AtualizaProp
            // 
            this.btnCria_AtualizaProp.Location = new System.Drawing.Point(148, 80);
            this.btnCria_AtualizaProp.Name = "btnCria_AtualizaProp";
            this.btnCria_AtualizaProp.Size = new System.Drawing.Size(159, 50);
            this.btnCria_AtualizaProp.TabIndex = 7;
            this.btnCria_AtualizaProp.Text = "Cria/Alt Propriedade";
            this.btnCria_AtualizaProp.UseVisualStyleBackColor = true;
            this.btnCria_AtualizaProp.Click += new System.EventHandler(this.btnCria_AtualizaProp_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(337, 450);
            this.Controls.Add(this.labelSW);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnAbrirSW;
        private System.Windows.Forms.CheckBox Visivel;
        private System.Windows.Forms.Button btnFecharSW;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnAbrir;
        private System.Windows.Forms.Button btnFechar;
        private System.Windows.Forms.Label labelSW;
        private System.Windows.Forms.Button btnExportar;
        private System.Windows.Forms.Button btnBuscaProp;
        private System.Windows.Forms.Button btnCria_AtualizaProp;
    }
}

