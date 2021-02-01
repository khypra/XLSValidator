
namespace XLSValidator
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
            this.Comp_XLS = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // Comp_XLS
            // 
            this.Comp_XLS.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.Comp_XLS.Location = new System.Drawing.Point(103, 122);
            this.Comp_XLS.Name = "Comp_XLS";
            this.Comp_XLS.Size = new System.Drawing.Size(94, 39);
            this.Comp_XLS.TabIndex = 0;
            this.Comp_XLS.Text = "Selecionar Arquivo";
            this.Comp_XLS.UseVisualStyleBackColor = true;
            this.Comp_XLS.Click += new System.EventHandler(this.ObterXLS);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "xlsInput";
            this.openFileDialog1.Filter = "xls files (*.xls)|*.xls";
            this.openFileDialog1.InitialDirectory = "\"C:\\\"";
            this.openFileDialog1.RestoreDirectory = true;
            this.openFileDialog1.Title = "Buscar XLS";
            this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog1_FileOk);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(317, 173);
            this.Controls.Add(this.Comp_XLS);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button Comp_XLS;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}

