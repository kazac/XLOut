namespace XLOut
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
            btnCopy = new Button();
            label1 = new Label();
            SuspendLayout();
            // 
            // btnCopy
            // 
            btnCopy.Location = new Point(220, 55);
            btnCopy.Name = "btnCopy";
            btnCopy.Size = new Size(94, 29);
            btnCopy.TabIndex = 0;
            btnCopy.Text = "Copy";
            btnCopy.UseVisualStyleBackColor = true;
            btnCopy.Click += button1_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(100, 105);
            label1.Name = "label1";
            label1.Size = new Size(335, 20);
            label1.TabIndex = 1;
            label1.Text = "Copy data to the Clipboard as an Excel  fragment";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(534, 172);
            Controls.Add(label1);
            Controls.Add(btnCopy);
            Name = "Form1";
            Text = "XL Out";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnCopy;
        private Label label1;
    }
}
