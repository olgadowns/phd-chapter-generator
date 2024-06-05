namespace ContentsGenerator
{
    using System.Diagnostics.Contracts;
    using Excel = Microsoft.Office.Interop.Excel;
    public partial class frmMain : Form
    {

        public frmMain()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            lblFileName.Text = "C:\\Users\\WilliamHarrington\\OneDrive - HM & W Harrington\\Desktop\\phd-chapter-generator\\PhD Contents.xlsx";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();

            lblFileName.Text = openFileDialog1.FileName;

        }

        private void label1_Click(object sender, EventArgs e)
        {
            colorDialog1.ShowDialog();
            lblBGColour.BackColor = colorDialog1.Color;
        }

        private void lblGenerate_Click(object sender, EventArgs e)
        {
            List<List<string>> list = new List<List<string>>();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(lblFileName.Text);

            foreach (Excel._Worksheet sheet in xlWorkbook.Sheets)
            {

                Excel.Range xlRange = sheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;


                List<string> items = new List<string>();

                items.Add(sheet.Name);

                for (int a = 1; a <= rowCount; a++)
                {
                    if (xlRange.Cells[a, 1] != null)
                    {
                        items.Add(xlRange.Cells[a, 1].Value2.ToString());
                    }
                    else
                    {
                        break;
                    }
                }
                list.Add(items);
            }

            int current_id = 1;
            foreach (var chapter in list)
            {
                foreach (var section in chapter)
                {
                    generateIamge(list, current_id.ToString(), chapter[0], section);
                    current_id++;
                }
            }

            xlApp.Quit();

        }

        public string getNextItem(List<string> sections, int index)
        {
            if (index == 0)
            {
                return sections[1];
            }
            else
            {
                return sections[index + 1];
            }
        }

        public void generateIamge(List<List<string>> chapters,string prepend_filename, string current_chapter, string highlited_item = "")
        {
            string the_chapter = "";
            int boxWidth = 110;
            int boxHeight = 50;
            int spacing = 20;
            int boxHeaderExtraHeight = 50;

            int totalImageWidth = ((chapters.Count) * boxWidth) + ((chapters.Count + 1) * spacing);

            int totalImageHeight = 0;

            int numberofboxes = 1;

            foreach (var l in chapters)
            {
                if (l.Count > numberofboxes)
                {
                    numberofboxes = l.Count;
                }
            }

            totalImageHeight = ((spacing + boxHeight) * numberofboxes) + spacing;

            Bitmap bitmap = new Bitmap(totalImageWidth, totalImageHeight, System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
            Graphics graphics = Graphics.FromImage(bitmap);

            //Draw frames
            Pen pen = new Pen(lblBGColour.BackColor, 2);
            Brush fillBruish = new SolidBrush(lblBGColour.BackColor);
            Brush brushWhite = new SolidBrush(Color.White);
            Brush brushShaddow = new SolidBrush(Color.CadetBlue);
            Brush brushBlack = new SolidBrush(Color.Black);
            Font font = new Font("Arial", 8, FontStyle.Regular);
            Font fontUnderline = new Font("Arial", 8, FontStyle.Underline);

            int x = 0;            
           

            foreach (var section in chapters)
            {
                //y = 0;
                
                //foreach (var section in chapters)
                for (int y = 0; y < section.Count - 1; y++) 
                {
                   
                    txtLog.Text += getNextItem(section, y) + "\r\n";

                    int centreX = (int)(x * boxWidth + spacing + (spacing * x) + 10);
                    int centreY = (int)y * boxHeight + spacing + (y * spacing);


                    if (y == 0)
                    {
                        the_chapter = section[0];
                        graphics.FillRectangle(fillBruish, x * boxWidth + spacing + (spacing * x), y * (boxHeight + boxHeaderExtraHeight) + spacing + (y * spacing), boxWidth, (boxHeight + boxHeaderExtraHeight));

                        Rectangle r = new Rectangle(x * boxWidth + spacing + (spacing * x), y * (boxHeight + boxHeaderExtraHeight) + spacing + (y * spacing), boxWidth, (boxHeight + boxHeaderExtraHeight) / 2);

                        StringFormat sf = new StringFormat();
                        sf.LineAlignment = StringAlignment.Center;
                        sf.Alignment = StringAlignment.Center;
                        graphics.DrawString(section[0], fontUnderline, brushWhite, r, sf);

                        //Second text
                         r = new Rectangle(x * boxWidth + spacing + (spacing * x), y * (boxHeight + boxHeaderExtraHeight) + 5 + spacing + (y * spacing) + ((boxHeight + boxHeaderExtraHeight) / 3), boxWidth, (boxHeight + boxHeaderExtraHeight) / 2 );

                        graphics.DrawString(getNextItem(section, y), font , brushWhite, r, sf);
                        
                    }
                  else
                    {
                        if (highlited_item == getNextItem(section, y) && the_chapter == current_chapter)
                        {
                            graphics.FillRectangle(brushShaddow, x * boxWidth + spacing + (spacing * x), y * boxHeight + spacing + boxHeaderExtraHeight + (y * spacing), boxWidth, boxHeight);

                            graphics.DrawRectangle(pen, x * boxWidth + spacing + (spacing * x), y * boxHeight + spacing + boxHeaderExtraHeight + (y * spacing), boxWidth, boxHeight);

                            Rectangle r = new Rectangle(x * boxWidth + spacing + (spacing * x), y * boxHeight + spacing + boxHeaderExtraHeight + (y * spacing), boxWidth, boxHeight);

                            StringFormat sf = new StringFormat();
                            sf.LineAlignment = StringAlignment.Center;
                            sf.Alignment = StringAlignment.Center;
                            graphics.DrawString(getNextItem(section, y), font, brushWhite, r, sf);

                        }
                        else
                        {
                            graphics.FillRectangle(brushWhite, x * boxWidth + spacing + (spacing * x), y * boxHeight + spacing + boxHeaderExtraHeight + (y * spacing), boxWidth, boxHeight);

                            graphics.DrawRectangle(pen, x * boxWidth + spacing + (spacing * x), y * boxHeight + spacing + boxHeaderExtraHeight + (y * spacing), boxWidth, boxHeight);

                            Rectangle r = new Rectangle(x * boxWidth + spacing + (spacing * x), y * boxHeight + spacing + boxHeaderExtraHeight + (y * spacing), boxWidth, boxHeight);

                            StringFormat sf = new StringFormat();
                            sf.LineAlignment = StringAlignment.Center;
                            sf.Alignment = StringAlignment.Center;
                            graphics.DrawString(getNextItem(section, y), font, brushBlack, r, sf);

                        }

                    }

                    //y++;
                    var filename = "images\\" + prepend_filename + " - " + cleanFilename(current_chapter + " - " + highlited_item + ".png");
                
                    bitmap.Save(filename);

                }
                x++;
                the_chapter = current_chapter;
            }





            //bitmap.Save(@"images\\base_image.png");

           // pictureBox1.Image = Image.FromFile(@"images\\base_image.png");

           
        }

        public string cleanFilename(string fileName)
        {
            return Path.GetInvalidFileNameChars().Aggregate(fileName, (current, c) => current.Replace(c.ToString(), string.Empty));

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}