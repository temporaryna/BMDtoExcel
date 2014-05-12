namespace Bmd_To_Excel_project
{
    partial class Form1
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
			this.button4 = new System.Windows.Forms.Button();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.progressBar2 = new System.Windows.Forms.ProgressBar();
			this.progressBar3 = new System.Windows.Forms.ProgressBar();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.SuspendLayout();
			// 
			// button4
			// 
			this.button4.Location = new System.Drawing.Point(12, 184);
			this.button4.Name = "button4";
			this.button4.Size = new System.Drawing.Size(335, 66);
			this.button4.TabIndex = 1;
			this.button4.Text = "Load file";
			this.button4.UseVisualStyleBackColor = true;
			this.button4.Click += new System.EventHandler(this.button4_Click);
			this.button4.DragDrop += new System.Windows.Forms.DragEventHandler(this.button4_DragDrop);
			this.button4.DragEnter += new System.Windows.Forms.DragEventHandler(this.button4_DragEnter);
			this.button4.DragLeave += new System.EventHandler(this.button4_DragLeave);
			// 
			// openFileDialog1
			// 
			this.openFileDialog1.FileName = "Item.bmd";
			this.openFileDialog1.Filter = "Program Item files|*.bmd; *.xls; *.xlsx";
			// 
			// progressBar2
			// 
			this.progressBar2.Location = new System.Drawing.Point(12, 120);
			this.progressBar2.Maximum = 512;
			this.progressBar2.Name = "progressBar2";
			this.progressBar2.Size = new System.Drawing.Size(335, 23);
			this.progressBar2.Step = 1;
			this.progressBar2.TabIndex = 2;
			// 
			// progressBar3
			// 
			this.progressBar3.Location = new System.Drawing.Point(12, 149);
			this.progressBar3.Maximum = 16;
			this.progressBar3.Name = "progressBar3";
			this.progressBar3.Size = new System.Drawing.Size(335, 23);
			this.progressBar3.Step = 1;
			this.progressBar3.TabIndex = 2;
			// 
			// saveFileDialog1
			// 
			this.saveFileDialog1.FileName = "Item.bmd";
			this.saveFileDialog1.Filter = "Item BMD file|*.bmd";
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(359, 262);
			this.Controls.Add(this.progressBar3);
			this.Controls.Add(this.progressBar2);
			this.Controls.Add(this.button4);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Name = "Form1";
			this.Text = "BMD <-> Excel";
			this.ResumeLayout(false);

        }

        #endregion

		private System.Windows.Forms.Button button4;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.ProgressBar progressBar2;
		private System.Windows.Forms.ProgressBar progressBar3;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    }
}

