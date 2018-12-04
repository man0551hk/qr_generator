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
                //imageName.Add(DrawImage("1", "Oversea", "1", "1", "Crystal", "Sin-Yu", "Cheung", "1"));
                //imageName.Add(DrawImage("12", "Building 1", "12", "12", "Shawn", "Shaokun", "Chen", "0"));
                //imageName.Add(DrawImage("123", "Consulting 1", "123", "123", "Crystal", "Sin-Yu", "Cheung", "1"));
                //imageName.Add(DrawImage("1234", "Infrastructure 1", "1234", "1234", "Crystal", "Sin-Yu", "Cheung", "1"));
                //imageName.Add(DrawImage("12345", "Business Services", "12345", "12345", "Lawrence", "Wai-Yiu", "Kan", "1"));

                //for (int i = 1955; i <= 2044; i++)
                //{
                //    imageName.Add(DrawImage("", "", "", i.ToString(), "", "", "", ""));
                //}

                ConvertToPdf();
                Console.WriteLine("End");
                Console.WriteLine("Press enter to exit");
            }
            Console.ReadKey();
            
        }

        public static string DrawImage(string staffNo, string groupName, string tableNo, string ticketNo, string englishname, string fistname, string lastname, string parnter)
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

                    graphicsImage.DrawImage(qrBMP, 1730, 300, 180, 180);
                    Color StringColor = System.Drawing.ColorTranslator.FromHtml("#000");//direct color adding
                    if (Convert.ToInt32(ticketNo) >= 1955 && Convert.ToInt32(ticketNo) <= 2044)
                    {
                        int x = 1746;
                        graphicsImage.DrawString(ticketNo, new Font("arial", 10,
                        FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 869));
                    }
                    else {
                        int x = 1570;
                      

                        graphicsImage.DrawString(lastname + " " + fistname, new Font("arial", 10, FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 510));
                        graphicsImage.DrawString(englishname + " + Partner", new Font("arial", 10, FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 560));

                        //Set the alignment based on the coordinates   
                        x = 1728;
                        graphicsImage.DrawString(staffNo, new Font("arial", 10, FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 628));


                        x = 1700;
                        graphicsImage.DrawString(groupName, new Font("arial", 10,
                        FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 708));

                        x = 1740;
                        graphicsImage.DrawString(tableNo, new Font("arial", 10,
                        FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 789));

                        x = 1746;
                        graphicsImage.DrawString(ticketNo, new Font("arial", 10,
                        FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 869));
                    }
                    //bitmap.Save(AppDomain.CurrentDomain.BaseDirectory + @"ticketImg/" + staffNo + ".png");
                    bitmap.Save(AppDomain.CurrentDomain.BaseDirectory + @"ticketImg/ticket_" + ticketNo + ".png");
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

            DirectoryInfo d = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + @"ticketImg");//Assuming Test is your Folder
            FileInfo[] Files = d.GetFiles("*.png"); //Getting Text files

            foreach (FileInfo file in Files)
            {
                Console.WriteLine(file.FullName);
                using (MagickImageCollection collection = new MagickImageCollection())
                {
                    collection.Add(new MagickImage(file.FullName));
                    // Create pdf file with two pages
                    collection.Write(AppDomain.CurrentDomain.BaseDirectory + @"ticketPdf/" + file.Name.Replace(".png", ".pdf"));
                }
            }

            //for (int i = 0; i < imageName.Count; i++)
            //{
            //    using (MagickImageCollection collection = new MagickImageCollection())
            //    {
            //        collection.Add(new MagickImage(AppDomain.CurrentDomain.BaseDirectory + @"ticketImg/" + imageName[i]));
            //        // Create pdf file with two pages
            //        collection.Write(AppDomain.CurrentDomain.BaseDirectory + @"ticketPdf/" + imageName[i].Replace(".png", ".pdf"));
            //    }
            //}
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
                    ICell englishNameCell = irow.GetCell(4);
                    ICell firstNameCell = irow.GetCell(5);
                    ICell lastNameCell = irow.GetCell(6);

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

                        string englishNAme = englishNameCell.StringCellValue;
                        string firstName = firstNameCell.StringCellValue;
                        string lastName = lastNameCell.StringCellValue;

                        DrawImage(staff, group, tableNo, tickeNo, englishNAme, firstName, lastName, "");
                    }
                }
            }
        }
    }
}
