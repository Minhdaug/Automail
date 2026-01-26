namespace AutoMailWinForm
{
    partial class Form1
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
			txtInput = new RichTextBox();
			btnStart = new Button();
			gridResult = new DataGridView();
			btnExport = new Button();
			((System.ComponentModel.ISupportInitialize)gridResult).BeginInit();
			SuspendLayout();
			// 
			// txtInput
			// 
			txtInput.Location = new Point(12, 27);
			txtInput.Name = "txtInput";
			txtInput.Size = new Size(544, 191);
			txtInput.TabIndex = 0;
			txtInput.Text = "";
			txtInput.TextChanged += richTextBox1_TextChanged;
			// 
			// btnStart
			// 
			btnStart.Location = new Point(632, 69);
			btnStart.Name = "btnStart";
			btnStart.Size = new Size(122, 23);
			btnStart.TabIndex = 1;
			btnStart.Text = "Bắt đầu chạy";
			btnStart.UseVisualStyleBackColor = true;
			btnStart.Click += btnStart_Click;
			// 
			// gridResult
			// 
			gridResult.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
			gridResult.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			gridResult.Location = new Point(0, 224);
			gridResult.Name = "gridResult";
			gridResult.Size = new Size(1295, 361);
			gridResult.TabIndex = 2;
			// 
			// btnExport
			// 
			btnExport.Location = new Point(632, 117);
			btnExport.Name = "btnExport";
			btnExport.Size = new Size(122, 23);
			btnExport.TabIndex = 3;
			btnExport.Text = "Xuất File Excel";
			btnExport.UseVisualStyleBackColor = true;
			btnExport.Click += btnExport_Click;
			// 
			// Form1
			// 
			AutoScaleDimensions = new SizeF(7F, 15F);
			AutoScaleMode = AutoScaleMode.Font;
			ClientSize = new Size(1293, 586);
			Controls.Add(btnExport);
			Controls.Add(gridResult);
			Controls.Add(btnStart);
			Controls.Add(txtInput);
			Name = "Form1";
			Text = "Form1";
			Load += Form1_Load;
			((System.ComponentModel.ISupportInitialize)gridResult).EndInit();
			ResumeLayout(false);
		}

		#endregion

		private RichTextBox txtInput;
		private Button btnStart;
		private DataGridView gridResult;
		private Button btnExport;
	}
}
