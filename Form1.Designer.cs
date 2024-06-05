namespace ContentsGenerator
{
    partial class frmMain
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
            components = new System.ComponentModel.Container();
            openFileDialog1 = new OpenFileDialog();
            button1 = new Button();
            lblFileName = new Label();
            lblGenerate = new Button();
            imageList1 = new ImageList(components);
            pictureBox1 = new PictureBox();
            colorDialog1 = new ColorDialog();
            lblBGColour = new Label();
            label2 = new Label();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            SuspendLayout();
            // 
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialog1";
            // 
            // button1
            // 
            button1.Location = new Point(12, 46);
            button1.Name = "button1";
            button1.Size = new Size(75, 23);
            button1.TabIndex = 0;
            button1.Text = "Load File";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // lblFileName
            // 
            lblFileName.Location = new Point(12, 19);
            lblFileName.MaximumSize = new Size(200, 0);
            lblFileName.MinimumSize = new Size(200, 0);
            lblFileName.Name = "lblFileName";
            lblFileName.Size = new Size(200, 15);
            lblFileName.TabIndex = 1;
            lblFileName.Text = "No File Loaded";
            // 
            // lblGenerate
            // 
            lblGenerate.Location = new Point(93, 46);
            lblGenerate.Name = "lblGenerate";
            lblGenerate.Size = new Size(75, 23);
            lblGenerate.TabIndex = 2;
            lblGenerate.Text = "Generate";
            lblGenerate.UseVisualStyleBackColor = true;
            lblGenerate.Click += lblGenerate_Click;
            // 
            // imageList1
            // 
            imageList1.ColorDepth = ColorDepth.Depth32Bit;
            imageList1.ImageSize = new Size(16, 16);
            imageList1.TransparentColor = Color.Transparent;
            // 
            // pictureBox1
            // 
            pictureBox1.Location = new Point(231, 27);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(641, 622);
            pictureBox1.TabIndex = 3;
            pictureBox1.TabStop = false;
            pictureBox1.Click += pictureBox1_Click;
            // 
            // lblBGColour
            // 
            lblBGColour.AutoSize = true;
            lblBGColour.BackColor = SystemColors.MenuHighlight;
            lblBGColour.Location = new Point(12, 87);
            lblBGColour.MinimumSize = new Size(20, 20);
            lblBGColour.Name = "lblBGColour";
            lblBGColour.Size = new Size(20, 20);
            lblBGColour.TabIndex = 4;
            lblBGColour.Click += label1_Click;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(48, 92);
            label2.Name = "label2";
            label2.Size = new Size(110, 15);
            label2.TabIndex = 5;
            label2.Text = "Background Colour";
            // 
            // frmMain
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(884, 661);
            Controls.Add(label2);
            Controls.Add(lblBGColour);
            Controls.Add(pictureBox1);
            Controls.Add(lblGenerate);
            Controls.Add(lblFileName);
            Controls.Add(button1);
            Name = "frmMain";
            Text = "Contents Creator";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private OpenFileDialog openFileDialog1;
        private Button button1;
        private Label lblFileName;
        private Button lblGenerate;
        private ImageList imageList1;
        private PictureBox pictureBox1;
        private ColorDialog colorDialog1;
        private Label lblBGColour;
        private Label label2;
    }
}
