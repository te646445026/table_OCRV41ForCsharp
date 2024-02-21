namespace table_OCRV41ForCsharp_winform
{
    partial class myForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            workSpace = new TabControl();
            work = new TabPage();
            keyMessage = new TabPage();
            REQUEST_URL_label = new Label();
            REQUEST_URL_textBox = new TextBox();
            SECRET_KEY_label = new Label();
            SECRET_KEY_textBox = new TextBox();
            API_KEY_label = new Label();
            API_KEY_textBox = new TextBox();
            workPath_label = new Label();
            workPath_textBox = new TextBox();
            textBox1 = new TextBox();
            dataFilePath_label = new Label();
            workSpace.SuspendLayout();
            work.SuspendLayout();
            keyMessage.SuspendLayout();
            SuspendLayout();
            // 
            // workSpace
            // 
            workSpace.Controls.Add(work);
            workSpace.Controls.Add(keyMessage);
            workSpace.Location = new Point(16, 12);
            workSpace.Name = "workSpace";
            workSpace.RightToLeftLayout = true;
            workSpace.SelectedIndex = 0;
            workSpace.Size = new Size(885, 488);
            workSpace.TabIndex = 0;
            // 
            // work
            // 
            work.Controls.Add(textBox1);
            work.Controls.Add(dataFilePath_label);
            work.Controls.Add(workPath_textBox);
            work.Controls.Add(workPath_label);
            work.Location = new Point(4, 26);
            work.Name = "work";
            work.Padding = new Padding(3);
            work.Size = new Size(877, 458);
            work.TabIndex = 0;
            work.Text = "工作";
            work.UseVisualStyleBackColor = true;
            // 
            // keyMessage
            // 
            keyMessage.Controls.Add(REQUEST_URL_label);
            keyMessage.Controls.Add(REQUEST_URL_textBox);
            keyMessage.Controls.Add(SECRET_KEY_label);
            keyMessage.Controls.Add(SECRET_KEY_textBox);
            keyMessage.Controls.Add(API_KEY_label);
            keyMessage.Controls.Add(API_KEY_textBox);
            keyMessage.Location = new Point(4, 26);
            keyMessage.Name = "keyMessage";
            keyMessage.Padding = new Padding(3);
            keyMessage.Size = new Size(877, 458);
            keyMessage.TabIndex = 1;
            keyMessage.Text = "密钥信息";
            keyMessage.UseVisualStyleBackColor = true;
            // 
            // REQUEST_URL_label
            // 
            REQUEST_URL_label.AutoSize = true;
            REQUEST_URL_label.Location = new Point(141, 199);
            REQUEST_URL_label.Name = "REQUEST_URL_label";
            REQUEST_URL_label.Size = new Size(91, 17);
            REQUEST_URL_label.TabIndex = 5;
            REQUEST_URL_label.Text = "REQUEST_URL";
            // 
            // REQUEST_URL_textBox
            // 
            REQUEST_URL_textBox.Location = new Point(238, 196);
            REQUEST_URL_textBox.Name = "REQUEST_URL_textBox";
            REQUEST_URL_textBox.Size = new Size(428, 23);
            REQUEST_URL_textBox.TabIndex = 4;
            REQUEST_URL_textBox.Text = "https://aip.baidubce.com/rest/2.0/ocr/v1/table";
            // 
            // SECRET_KEY_label
            // 
            SECRET_KEY_label.AutoSize = true;
            SECRET_KEY_label.Location = new Point(153, 160);
            SECRET_KEY_label.Name = "SECRET_KEY_label";
            SECRET_KEY_label.Size = new Size(79, 17);
            SECRET_KEY_label.TabIndex = 3;
            SECRET_KEY_label.Text = "SECRET_KEY";
            // 
            // SECRET_KEY_textBox
            // 
            SECRET_KEY_textBox.Location = new Point(238, 157);
            SECRET_KEY_textBox.Name = "SECRET_KEY_textBox";
            SECRET_KEY_textBox.Size = new Size(428, 23);
            SECRET_KEY_textBox.TabIndex = 2;
            SECRET_KEY_textBox.Text = "505cd0eiUZt22mPzelDGVrWzN7ELwteh";
            // 
            // API_KEY_label
            // 
            API_KEY_label.AutoSize = true;
            API_KEY_label.Location = new Point(178, 119);
            API_KEY_label.Name = "API_KEY_label";
            API_KEY_label.Size = new Size(54, 17);
            API_KEY_label.TabIndex = 1;
            API_KEY_label.Text = "API_KEY";
            // 
            // API_KEY_textBox
            // 
            API_KEY_textBox.Location = new Point(238, 116);
            API_KEY_textBox.Name = "API_KEY_textBox";
            API_KEY_textBox.Size = new Size(428, 23);
            API_KEY_textBox.TabIndex = 0;
            API_KEY_textBox.Text = "Et4nGdx8ecc5chOnoilbxEyX";
            // 
            // workPath_label
            // 
            workPath_label.AutoSize = true;
            workPath_label.Location = new Point(40, 9);
            workPath_label.Name = "workPath_label";
            workPath_label.Size = new Size(56, 17);
            workPath_label.TabIndex = 0;
            workPath_label.Text = "工作目录";
            // 
            // workPath_textBox
            // 
            workPath_textBox.Location = new Point(102, 6);
            workPath_textBox.Name = "workPath_textBox";
            workPath_textBox.Size = new Size(147, 23);
            workPath_textBox.TabIndex = 1;
            // 
            // textBox1
            // 
            textBox1.Location = new Point(102, 47);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(147, 23);
            textBox1.TabIndex = 3;
            textBox1.TextChanged += textBox1_TextChanged;
            // 
            // dataFilePath_label
            // 
            dataFilePath_label.AutoSize = true;
            dataFilePath_label.Location = new Point(6, 50);
            dataFilePath_label.Name = "dataFilePath_label";
            dataFilePath_label.Size = new Size(92, 17);
            dataFilePath_label.TabIndex = 2;
            dataFilePath_label.Text = "需识别图片路径";
            dataFilePath_label.Click += this.label1_Click;
            // 
            // myForm
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(924, 523);
            Controls.Add(workSpace);
            Name = "myForm";
            Text = "限速器自动化办公工具";
            workSpace.ResumeLayout(false);
            work.ResumeLayout(false);
            work.PerformLayout();
            keyMessage.ResumeLayout(false);
            keyMessage.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private TabControl workSpace;
        private TabPage work;
        private TabPage keyMessage;
        private Label SECRET_KEY_label;
        private TextBox SECRET_KEY_textBox;
        private Label API_KEY_label;
        private TextBox API_KEY_textBox;
        private Label REQUEST_URL_label;
        private TextBox REQUEST_URL_textBox;
        private Label workPath_label;
        private TextBox workPath_textBox;
        private TextBox textBox1;
        private Label dataFilePath_label;
    }
}
