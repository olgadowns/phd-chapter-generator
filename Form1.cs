namespace ContentsGenerator
{
    using System.Diagnostics.Contracts;
    using static System.Collections.Specialized.BitVector32;
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
            int current_chapter = 1;
            foreach (var chapter in list)
            {
                foreach (var section in chapter)
                {
                    generateIamge(list, current_id.ToString(), chapter[0], section, current_chapter);
                    current_id++;
                    //break;
                }
                current_chapter++;
                //break; 
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

        public void generateIamge(List<List<string>> chapters,string prepend_filename, string current_chapter, string highlited_item, int current_chapter_id)
        {
            string the_chapter = "";
            int boxWidth = 130;
            int boxHeight = 40;
            int spacing = 10;
            int boxHeaderExtraHeight = 50;
            int y_offset = 50;
            
            int up_to_box = int.Parse(prepend_filename);
            int current_box = 0;

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

            totalImageHeight = ((spacing + boxHeight) * numberofboxes) + spacing + 40;

            Bitmap bitmap = new Bitmap(totalImageWidth, totalImageHeight, System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
            Graphics graphics = Graphics.FromImage(bitmap);

            //Draw frames
            Pen pen = new Pen(lblBGColour.BackColor, 2);
            Pen penHeader = new Pen(lblBGColour.BackColor, 1);
            Brush fillBruish = new SolidBrush(lblBGColour.BackColor);
            Brush brushWhite = new SolidBrush(Color.White);
            Brush brushShaddow = new SolidBrush(Color.CadetBlue);
            Brush brushPast = new SolidBrush(Color.LightGray);
            Brush brushBlack = new SolidBrush(Color.Black);
            Font font = new Font("Arial", 8, FontStyle.Regular);
            Font fontBold = new Font("Arial", 8, FontStyle.Bold);
            Font fontUnderline = new Font("Arial", 8, FontStyle.Underline);
            Font fontBoldUnderline = new Font("Arial", 10, FontStyle.Bold | FontStyle.Underline);
            Font fontHeader = new Font("Arial", 11);

            int x = 0;


            StringFormat sfHeader = new StringFormat();
            sfHeader.LineAlignment = StringAlignment.Center;
            sfHeader.Alignment = StringAlignment.Center;                     

            Rectangle headerRectangle = new Rectangle(spacing, spacing, ((chapters.Count - 1) * (boxWidth + spacing)) + boxWidth, 40);

            graphics.DrawRectangle(penHeader, headerRectangle);

            graphics.DrawString("Measuring and quantifying the benefits of improved Internet connectivity in regional and remote Australia and its effect on adoption of technology", fontHeader, brushBlack, headerRectangle, sfHeader);

            foreach (var section in chapters)
            {
                
                
                //foreach (var section in chapters)
                for (int y = 0; y < section.Count - 1; y++) 
                {
                   
                    txtLog.Text += getNextItem(section, y) + "\r\n";

                    int centreX = (int)(x * boxWidth + spacing + (spacing * x) + 10);
                    int centreY = (int)y * boxHeight + spacing + (y * spacing);


                    if (y == 0)
                    {
                        the_chapter = section[0];
                        graphics.FillRectangle(fillBruish, x * boxWidth + spacing + (spacing * x), y * (boxHeight + boxHeaderExtraHeight) + spacing + (y * spacing) + y_offset, boxWidth, (boxHeight + boxHeaderExtraHeight));

                        Rectangle r = new Rectangle(x * boxWidth + spacing + (spacing * x), y * (boxHeight + boxHeaderExtraHeight) + spacing + (y * spacing) - 5 + y_offset, boxWidth, (boxHeight + boxHeaderExtraHeight) / 2);

                        StringFormat sf = new StringFormat();
                        sf.LineAlignment = StringAlignment.Center;
                        sf.Alignment = StringAlignment.Center;
                        graphics.DrawString(section[0], fontBoldUnderline, brushWhite, r, sf);

                        //Second text
                        r = new Rectangle(x * boxWidth + spacing + (spacing * x), y * (boxHeight + boxHeaderExtraHeight) + 5 + spacing + (y * spacing) + ((boxHeight + boxHeaderExtraHeight) / 3) + y_offset , boxWidth, (boxHeight + boxHeaderExtraHeight) /1 );
                        sf.LineAlignment = StringAlignment.Near;
                        graphics.DrawString(getNextItem(section, y), fontBold, brushWhite, r, sf);
                        
                    }
                  else
                    {
                        if (highlited_item == getNextItem(section, y) && the_chapter == current_chapter)
                        {
                            
                            //Current Item
                            graphics.FillRectangle(brushShaddow, x * boxWidth + spacing + (spacing * x), y * boxHeight + spacing + boxHeaderExtraHeight + (y * spacing) + y_offset, boxWidth, boxHeight);
                            

                            graphics.DrawRectangle(pen, x * boxWidth + spacing + (spacing * x), y * boxHeight + spacing + boxHeaderExtraHeight + (y * spacing) + y_offset, boxWidth, boxHeight);

                            Rectangle r = new Rectangle(x * boxWidth + spacing + (spacing * x), y * boxHeight + spacing + boxHeaderExtraHeight + (y * spacing) + y_offset, boxWidth, boxHeight);

                            StringFormat sf = new StringFormat();
                            sf.LineAlignment = StringAlignment.Center;
                            sf.Alignment = StringAlignment.Center;
                            graphics.DrawString(getNextItem(section, y), fontBold, brushWhite, r, sf);

                        }
                        else
                        {

                            if (current_box + current_chapter_id < up_to_box)

                            {
                                graphics.FillRectangle(brushPast, x * boxWidth + spacing + (spacing * x), y * boxHeight + spacing + boxHeaderExtraHeight + (y * spacing) + y_offset, boxWidth, boxHeight);
                            }
                            else
                            {
                                graphics.FillRectangle(brushWhite, x * boxWidth + spacing + (spacing * x), y * boxHeight + spacing + boxHeaderExtraHeight + (y * spacing) + y_offset, boxWidth, boxHeight);
                            }


                            //graphics.FillRectangle(brushWhite, x * boxWidth + spacing + (spacing * x), y * boxHeight + spacing + boxHeaderExtraHeight + (y * spacing) + y_offset, boxWidth, boxHeight);

                            graphics.DrawRectangle(pen, x * boxWidth + spacing + (spacing * x), y * boxHeight + spacing + boxHeaderExtraHeight + (y * spacing) + y_offset, boxWidth, boxHeight);

                            Rectangle r = new Rectangle(x * boxWidth + spacing + (spacing * x), y * boxHeight + spacing + boxHeaderExtraHeight + (y * spacing) + y_offset, boxWidth, boxHeight);

                            StringFormat sf = new StringFormat();
                            sf.LineAlignment = StringAlignment.Center;
                            sf.Alignment = StringAlignment.Center;
                            graphics.DrawString(getNextItem(section, y), fontBold, brushBlack, r, sf);

                        }

                    }

        //y++;
                current_box++;

                }
                var filename = "images\\" + prepend_filename + " - " + cleanFilename(current_chapter + " - " + highlited_item + ".png");

                bitmap.Save(filename);

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