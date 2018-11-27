using ImageMagick;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;

namespace qr_generator
{
    class Program
    {
        public static List<string> imageName = new List<string>();
        static void Main(string[] args)
        {
            if (!File.Exists(AppDomain.CurrentDomain.BaseDirectory + "ticket.png"))
            {
                Console.WriteLine("No ticket.png found!");
            }
            else if (!File.Exists(AppDomain.CurrentDomain.BaseDirectory + "ticket.xlsx"))
            {
                Console.WriteLine("No ticket.xlsx found!");
            }
            else
            {
                if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + "ticketImg"))
                {
                    Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + "ticketImg");
                }
                if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + "ticketPdf"))
                {
                    Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + "ticketPdf");
                }
                //ReadExcel();

                //Oversea
                //Building 1
                //Consulting 1
                //Infrastructure 1
                //Business Services

                //1
                //12
                //838
                //2222
                //73471

                //1
                //12
                //123
                //1000
                //9999

                //1
                //20
                //121
                //Head Table
                imageName.Add(DrawImage("1", "Oversea", "1", "1"));
                imageName.Add(DrawImage("12", "Building 1", "1", "1"));
                imageName.Add(DrawImage("123", "Consulting 1", "1", "1"));
                imageName.Add(DrawImage("1234", "Infrastructure 1", "1", "1"));
                imageName.Add(DrawImage("12345", "Business Services", "1", "1"));


                ConvertToPdf();
                Console.WriteLine("End");
                Console.WriteLine("Press enter to exit");
            }
            //Console.ReadKey();
            
        }

        public static string DrawImage(string staffNo, string groupName, string tableNo, string ticketNo)
        {
            Console.WriteLine("Generating images...");
            try
            {
                //creating a image object
                System.Drawing.Image bitmap = (System.Drawing.Image)Bitmap.FromFile(AppDomain.CurrentDomain.BaseDirectory + "ticket.png"); // set image 

                //draw the image object using a Graphics object
                Graphics graphicsImage = Graphics.FromImage(bitmap);

                //Set the alignment based on the coordinates   
                StringFormat stringformat = new StringFormat();
                stringformat.Alignment = StringAlignment.Far;
                stringformat.LineAlignment = StringAlignment.Far;
                StringFormat stringformat2 = new StringFormat();
                stringformat2.Alignment = StringAlignment.Center;
                stringformat2.LineAlignment = StringAlignment.Center;
                //Set the font color/format/size etc..  
                Color StringColor = System.Drawing.ColorTranslator.FromHtml("#000");//direct color adding

                int x = 0;
                switch (staffNo.Length)
                {
                    case 1: x = 1763; break;
                    case 2: x = 1785; break; //+12
                    case 3: x = 1806; break; //+21  +9
                    case 4: x = 1836; break; //35 + 9
                    case 5: x = 1880; break; //35 + 9
                }

                graphicsImage.DrawString(staffNo, new Font("arial", 10,
                FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 740),
                stringformat);

                x = 1729;
                int add = 12;
                for (int i = 0; i < groupName.Length - 1; i++)
                {
                    x += add;
                    add += 1;
                }

                graphicsImage.DrawString(groupName, new Font("arial", 10,
                FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 820),
                stringformat);

                switch (tableNo.Length)
                {
                    case 1: x = 1730; break;
                    case 2: x = 1785; break; //+12
                    case 3: x = 1806; break; //+21  +9
                    case 4: x = 1836; break; //35 + 9
                }
                graphicsImage.DrawString(tableNo, new Font("arial", 10,
                FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 901),
                stringformat);


                switch (ticketNo.Length)
                {
                    case 1: x = 1763; break;
                    case 2: x = 1785; break; //+12
                    case 3: x = 1806; break; //+21  +9
                    case 4: x = 1836; break; //35 + 9
                }
                graphicsImage.DrawString(ticketNo, new Font("arial", 10,
                FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 981),
                stringformat);

                bitmap.Save(AppDomain.CurrentDomain.BaseDirectory  + @"ticketImg/" + staffNo + ".png");
                //bitmap.Save(Response.OutputStream, ImageFormat.Png);
            }
            catch (Exception ex)
            {
                Console.WriteLine("StaffNo. " + staffNo + " " + ex.Message); 
            }

            return staffNo + ".png";
        
        }

        public static void ConvertToPdf()
        {
            Console.WriteLine("Converting to pdf...");
            using (MagickImageCollection collection = new MagickImageCollection())
            {
                for (int i = 0; i < imageName.Count; i++)
                {
                    collection.Add(new MagickImage(AppDomain.CurrentDomain.BaseDirectory + @"ticketImg/" + imageName[i]));
                    // Create pdf file with two pages
                    collection.Write(AppDomain.CurrentDomain.BaseDirectory + @"ticketPdf/" + imageName[i].Replace(".png", ".pdf"));
                }
            }
        }

        public static void ReadExcel()
        {
            Console.WriteLine("Start...");
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "ticket.xlsx", FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = hssfwb.GetSheetAt(0);
            for (int row = 3; row <= sheet.LastRowNum; row++) // start from row 4
            {
                if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                {
                    IRow irow = sheet.GetRow(row);
                    ICell groupCell = irow.GetCell(0);
                    ICell staffCell = irow.GetCell(2);
                    ICell ticketNoCell = irow.GetCell(10);
                    ICell tableNoCell = irow.GetCell(11);

                    if (groupCell != null && staffCell != null && ticketNoCell != null && tableNoCell != null)
                    {
                        string group = groupCell.StringCellValue;

                        string staff = "";
                        try
                        {
                            staff = staffCell.NumericCellValue.ToString();
                        }
                        catch
                        {
                            staff = staffCell.StringCellValue;
                        }

                        string tickeNo = "";
                        try
                        {
                            tickeNo = ticketNoCell.NumericCellValue.ToString();
                        }
                        catch
                        {
                            tickeNo = ticketNoCell.StringCellValue;
                        }
                        string tableNo = "";
                        try
                        {
                            tableNo = tableNoCell.NumericCellValue.ToString();
                        }
                        catch
                        {
                            tableNo = tableNoCell.StringCellValue;
                        }

                    }
                }
            }
        }
    }
}
