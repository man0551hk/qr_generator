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
            if (!File.Exists(AppDomain.CurrentDomain.BaseDirectory + "ticket_17x7.png"))
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
                imageName.Add(DrawImage("12", "Building 1", "12", "12"));
                imageName.Add(DrawImage("123", "Consulting 1", "123", "123"));
                imageName.Add(DrawImage("1234", "Infrastructure 1", "1234", "1234"));
                imageName.Add(DrawImage("12345", "Business Services", "12345", "12345"));


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
                    System.Drawing.Image bitmap = (System.Drawing.Image)Bitmap.FromFile(AppDomain.CurrentDomain.BaseDirectory + "ticket_17x7.png"); // set image 

                    //draw the image object using a Graphics object
                    Graphics graphicsImage = Graphics.FromImage(bitmap);

                    MessagingToolkit.QRCode.Codec.QRCodeEncoder encoder = new MessagingToolkit.QRCode.Codec.QRCodeEncoder();
                    encoder.QRCodeScale = 24;
                    Bitmap qrBMP = encoder.Encode(ticketNo);

                    graphicsImage.DrawImage(qrBMP, 1663, 300, 300, 300);

                    if (Convert.ToInt32(ticketNo) >= 1955 && Convert.ToInt32(ticketNo) <= 2044)
                    {
                    }
                    else {
                        //Set the alignment based on the coordinates   
                        StringFormat stringformat = new StringFormat();
                        stringformat.Alignment = StringAlignment.Far;
                        stringformat.LineAlignment = StringAlignment.Far;
                        Color StringColor = System.Drawing.ColorTranslator.FromHtml("#000");//direct color adding

                        int x = 0;
                        switch (staffNo.Length)
                        {
                            case 1: x = 1763; break;
                            case 2: x = 1785; break; //+12
                            case 3: x = 1806; break; //+21  +9
                            case 4: x = 1832; break; //35 + 9
                            case 5: x = 1850; break; //35 + 9
                        }

                        graphicsImage.DrawString(staffNo, new Font("arial", 10,
                        FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 680),
                        stringformat);

                        x = 1773;
                        switch (groupName.Length)
                        {
                            case 7: x = 1862; break;
                            case 10: x = 1883; break;
                            case 12: x = 1937; break;
                            case 16: x = 1982; break;
                            case 17: x = 2042; break;
                        }

                        graphicsImage.DrawString(groupName, new Font("arial", 10,
                        FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 760),
                        stringformat);

                        switch (tableNo.Length)
                        {
                            case 1: x = 1767; break;
                            case 2: x = 1788; break; //+12
                            case 3: x = 1813; break; //+21  +9
                            case 4: x = 1834; break; //35 + 9
                            case 5: x = 1861; break; //35 + 9
                        }
                        graphicsImage.DrawString(tableNo, new Font("arial", 10,
                        FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 841),
                        stringformat);


                        switch (ticketNo.Length)
                        {
                            case 1: x = 1780; break;
                            case 2: x = 1802; break; //+12
                            case 3: x = 1824; break; //+21  +9
                            case 4: x = 1850; break; //35 + 9
                            case 5: x = 1870; break; //35 + 9
                        }
                        graphicsImage.DrawString(ticketNo, new Font("arial", 10,
                        FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 921),
                        stringformat);
                    }
                    bitmap.Save(AppDomain.CurrentDomain.BaseDirectory + @"ticketImg/" + staffNo + ".png");
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
            for (int i = 0; i < imageName.Count; i++)
            {
                using (MagickImageCollection collection = new MagickImageCollection())
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
