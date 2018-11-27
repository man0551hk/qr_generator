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
            ReadExcel();
        }

        public static string DrawImage(string staffNo, string groupName, string tableNo, string ticketNo)
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


            graphicsImage.DrawString("26056", new Font("arial", 10,
            FontStyle.Regular), new SolidBrush(StringColor), new Point(1860, 740),
            stringformat);

            graphicsImage.DrawString("Business Services", new Font("arial", 10,
            FontStyle.Regular), new SolidBrush(StringColor), new Point(2050, 820),
            stringformat);

            graphicsImage.DrawString("40", new Font("arial", 10,
            FontStyle.Regular), new SolidBrush(StringColor), new Point(1800, 901),
            stringformat);

            graphicsImage.DrawString("1933", new Font("arial", 10,
            FontStyle.Regular), new SolidBrush(StringColor), new Point(1850, 981),
            stringformat);

            bitmap.Save(AppDomain.CurrentDomain.BaseDirectory + staffNo + ".png");
            //bitmap.Save(Response.OutputStream, ImageFormat.Png);

            return staffNo + ".png";
        
        }

        public static void ConvertToPdf()
        {
            using (MagickImageCollection collection = new MagickImageCollection())
            {
                for (int i = 0; i < imageName.Count; i++)
                {
                    collection.Add(new MagickImage(AppDomain.CurrentDomain.BaseDirectory + imageName[i]));
                    // Create pdf file with two pages
                    collection.Write(imageName[i].Replace(".png", ".pdf"));
                }
            }
        }

        public static void ReadExcel()
        {
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

                        Console.WriteLine(group);
                        Console.WriteLine(staff);
                        Console.WriteLine(tickeNo);
                        Console.WriteLine(tableNo);
                    }
                }
            }
        }
    }
}
