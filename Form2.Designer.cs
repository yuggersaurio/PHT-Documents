namespace Exporter
{
    partial class Form2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form2));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.contraseñaTXT = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.logoApp = new System.Windows.Forms.PictureBox();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.logoApp)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.contraseñaTXT);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(5, 196);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(240, 138);
            this.groupBox1.TabIndex = 134;
            this.groupBox1.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(43, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(162, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Digite su clave de acceso";
            // 
            // contraseñaTXT
            // 
            this.contraseñaTXT.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.contraseñaTXT.Location = new System.Drawing.Point(28, 37);
            this.contraseñaTXT.Name = "contraseñaTXT";
            this.contraseñaTXT.PasswordChar = '*';
            this.contraseñaTXT.Size = new System.Drawing.Size(189, 30);
            this.contraseñaTXT.TabIndex = 1;
            this.contraseñaTXT.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.contraseñaTXT.TextChanged += new System.EventHandler(this.contraseñaTXT_TextChanged);
            this.contraseñaTXT.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.contraseñaTXT_KeyPress);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(28, 78);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(189, 40);
            this.button1.TabIndex = 2;
            this.button1.Text = "Ingresar";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // logoApp
            // 
            this.logoApp.Image = ((System.Drawing.Image)(resources.GetObject("logoApp.Image")));
            this.logoApp.Location = new System.Drawing.Point(5, 6);
            this.logoApp.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.logoApp.Name = "logoApp";
            this.logoApp.Size = new System.Drawing.Size(240, 183);
            this.logoApp.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.logoApp.TabIndex = 133;
            this.logoApp.TabStop = false;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(249, 338);
            this.Controls.Add(this.logoApp);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Iniciar usuario";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.Form2_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.logoApp)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox contraseñaTXT;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox logoApp;
    }
}