namespace PowerPointTTS
{
    partial class MainFormTTS
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
            this.groupBox9 = new System.Windows.Forms.GroupBox();
            this.listBox2Language = new System.Windows.Forms.ListBox();
            this.numericUpDown2Speed = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.numericUpDown1Vol = new System.Windows.Forms.NumericUpDown();
            this.button4 = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.groupBox9.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown2Speed)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1Vol)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox9
            // 
            this.groupBox9.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox9.Controls.Add(this.listBox2Language);
            this.groupBox9.Controls.Add(this.numericUpDown2Speed);
            this.groupBox9.Controls.Add(this.label3);
            this.groupBox9.Controls.Add(this.label2);
            this.groupBox9.Controls.Add(this.numericUpDown1Vol);
            this.groupBox9.Location = new System.Drawing.Point(12, 265);
            this.groupBox9.Name = "groupBox9";
            this.groupBox9.Size = new System.Drawing.Size(189, 153);
            this.groupBox9.TabIndex = 61;
            this.groupBox9.TabStop = false;
            this.groupBox9.Text = "音声設定";
            // 
            // listBox2Language
            // 
            this.listBox2Language.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listBox2Language.FormattingEnabled = true;
            this.listBox2Language.ItemHeight = 12;
            this.listBox2Language.Location = new System.Drawing.Point(11, 76);
            this.listBox2Language.Name = "listBox2Language";
            this.listBox2Language.Size = new System.Drawing.Size(164, 64);
            this.listBox2Language.TabIndex = 5;
            // 
            // numericUpDown2Speed
            // 
            this.numericUpDown2Speed.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.numericUpDown2Speed.Location = new System.Drawing.Point(46, 43);
            this.numericUpDown2Speed.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numericUpDown2Speed.Minimum = new decimal(new int[] {
            10,
            0,
            0,
            -2147483648});
            this.numericUpDown2Speed.Name = "numericUpDown2Speed";
            this.numericUpDown2Speed.Size = new System.Drawing.Size(129, 19);
            this.numericUpDown2Speed.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 45);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(29, 12);
            this.label3.TabIndex = 2;
            this.label3.Text = "音速";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "音量";
            // 
            // numericUpDown1Vol
            // 
            this.numericUpDown1Vol.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.numericUpDown1Vol.Location = new System.Drawing.Point(46, 18);
            this.numericUpDown1Vol.Name = "numericUpDown1Vol";
            this.numericUpDown1Vol.Size = new System.Drawing.Size(129, 19);
            this.numericUpDown1Vol.TabIndex = 0;
            this.numericUpDown1Vol.Value = new decimal(new int[] {
            100,
            0,
            0,
            0});
            // 
            // button4
            // 
            this.button4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button4.Location = new System.Drawing.Point(241, 317);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(184, 31);
            this.button4.TabIndex = 71;
            this.button4.Text = "■ Stop";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // label5
            // 
            this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(270, 397);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(155, 12);
            this.label5.TabIndex = 70;
            this.label5.Text = "Groove Direction 2017 v 1.0.4";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.label5.Click += new System.EventHandler(this.label5_Click);
            // 
            // button3
            // 
            this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button3.ForeColor = System.Drawing.Color.Red;
            this.button3.Location = new System.Drawing.Point(241, 355);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(184, 31);
            this.button3.TabIndex = 69;
            this.button3.Text = "● 録音";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(241, 276);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(184, 31);
            this.button1.TabIndex = 68;
            this.button1.Text = "▶ Play";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox1.Location = new System.Drawing.Point(12, 12);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox1.Size = new System.Drawing.Size(413, 247);
            this.textBox1.TabIndex = 72;
            // 
            // MainFormTTS
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(437, 423);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox9);
            this.Name = "MainFormTTS";
            this.Text = "文字から音声ファイルを作成";
            this.Load += new System.EventHandler(this.MainFormTTS_Load);
            this.groupBox9.ResumeLayout(false);
            this.groupBox9.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown2Speed)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1Vol)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox9;
        private System.Windows.Forms.ListBox listBox2Language;
        private System.Windows.Forms.NumericUpDown numericUpDown2Speed;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown numericUpDown1Vol;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
    }
}