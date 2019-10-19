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
            //if (!File.Exists(AppDomain.CurrentDomain.BaseDirectory + "ticket_17x7.png"))
            //{
            //    Console.WriteLine("No ticket.png found!");
            //}
            //else if (!File.Exists(AppDomain.CurrentDomain.BaseDirectory + "ticket.xlsx"))
            //{
            //    Console.WriteLine("No ticket.xlsx found!");
            //}
            //else
            //{
                if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + "ticketImg"))
                {
                    Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + "ticketImg");
                }
                if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + "ticketPdf"))
                {
                    Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + "ticketPdf");
                }
                ReadExcel();

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
                //imageName.Add(DrawImage("2829", "2229", "Infrastructure 1 (TungChung New Town Ext)", "Head Table", "1001", "KWOK Ka-Yue,", "Michael + Partner"));
                //imageName.Add(DrawImage("12", "Building 1", "12", "12", "Shawn", "Shaokun", "Chen", "0"));
                //imageName.Add(DrawImage("123", "Consulting 1", "123", "123", "Crystal", "Sin-Yu", "Cheung", "1"));
                //imageName.Add(DrawImage("1234", "Infrastructure 1", "1234", "1234", "Crystal", "Sin-Yu", "Cheung", "1"));
                //imageName.Add(DrawImage("12345", "Business Services", "12345", "12345", "Lawrence", "Wai-Yiu", "Kan", "1"));

                //int thisQr = 3773;
                //for (int i = 1955; i <= 2044; i++)
                //{
                //    DrawImage(thisQr.ToString(), "", "", "", i.ToString(), "", "");
                //    thisQr += 1;
                //}

                //ConvertToPdf();
                Console.WriteLine("End");
                Console.WriteLine("Press enter to exit");
            //}
            //Console.ReadKey();
            
        }

        public static string DrawImage(string qrcode, string staffNo, string groupName, string tableNo, string ticketNo, string firstLine, string secondLine)
        {
            Console.WriteLine("Generating images...");

            try
            {
                string filename = AppDomain.CurrentDomain.BaseDirectory + "ticket.jpg";

                //creating a image object
                System.Drawing.Image bitmap = (System.Drawing.Image)Bitmap.FromFile(filename); // set image 

                //draw the image object using a Graphics object
                Graphics graphicsImage = Graphics.FromImage(bitmap);

                MessagingToolkit.QRCode.Codec.QRCodeEncoder encoder = new MessagingToolkit.QRCode.Codec.QRCodeEncoder();
                encoder.QRCodeScale = 24;
                Bitmap qrBMP = encoder.Encode(qrcode);
                graphicsImage.DrawImage(qrBMP, 1000, 180, 350, 350);
                Color StringColor = System.Drawing.ColorTranslator.FromHtml("#000");

                if (staffNo == "(leave it blank)")
                {
                    staffNo = "";
                }
                else if (staffNo == "N/A - Guest")
                {
                    staffNo = "";
                }
                if (groupName == "(leave it blank)")
                {
                    groupName = "";
                }
                else if (groupName == "N/A - Guest")
                {
                    groupName = "";
                }


                int x = 300;
                graphicsImage.DrawString(firstLine, new Font("arial", 10, FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 120));
                x = 375;
                graphicsImage.DrawString(staffNo, new Font("arial", 10, FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 258));
                x = 451;
                graphicsImage.DrawString(groupName, new Font("arial", 10, FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 398));
                x = 385;
                graphicsImage.DrawString(tableNo, new Font("arial", 10, FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 536));
                //if (Convert.ToInt32(ticketNo) >= 1955 && Convert.ToInt32(ticketNo) <= 2044)
                //{

                //    int x = 1627;
                //    graphicsImage.DrawString(ticketNo, new Font("arial", 10,
                //    FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 700));
                //}
                //else {
                //    
                      

                //    
                //    graphicsImage.DrawString(secondLine, new Font("arial", 10, FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 400));

                //    //Set the alignment based on the coordinates   
                //    x = 1610;
                //   


                //    x = 1572;
                //    if (groupName.Contains("("))
                //    {
                            
                //        string[] groups = groupName.Split('(');
                //        graphicsImage.DrawString(groups[0].TrimEnd(), new Font("arial", 7,
                //        FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 545));

                //        graphicsImage.DrawString("(" + groups[1].TrimStart(), new Font("arial", 7,
                //        FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 577));
                //    }
                //    else
                //    {
                //        graphicsImage.DrawString(groupName, new Font("arial", 10,
                //        FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 536));
                //    }

                //    x = 1618;
                //    graphicsImage.DrawString(tableNo, new Font("arial", 10,
                //    FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 620));

                //    x = 1627;
                //    graphicsImage.DrawString(ticketNo, new Font("arial", 10,
                //    FontStyle.Regular), new SolidBrush(StringColor), new Point(x, 700));
                //}
                //bitmap.Save(AppDomain.CurrentDomain.BaseDirectory + @"ticketImg/" + staffNo + ".png");
                bitmap.Save(AppDomain.CurrentDomain.BaseDirectory + @"ticketImg/ticket_" + ticketNo + ".jpg");
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

            
            DrawImage("3921", "256094", "H516", "1", "chan1", "CHAN, IVAN LONG-YAT", "");
            //int ticketNo = 1092;
            //for (int row = 5013; row <= 5043; row++)
            //{
            //    DrawImage(row.ToString(), "", "", "", ticketNo.ToString(), "", "");
            //    ticketNo += 1;
            //}

            for (int row = 0; row <= sheet.LastRowNum; row++) // start from row 4
            {
                if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                {
                    IRow irow = sheet.GetRow(row);
     
                    ICell nameCell = irow.GetCell(0);
                    ICell staffCell = irow.GetCell(1);
                    ICell groupCell = irow.GetCell(2);
                    ICell tableNoCell = irow.GetCell(3);
                    ICell qrcodeCell = irow.GetCell(4);



                    if (nameCell != null && staffCell != null && groupCell != null && tableNoCell != null && qrcodeCell != null)
                    {
                        string group = groupCell.StringCellValue;
                        string name = nameCell.StringCellValue;
                       
                        string qrcode = "";
                        try
                        {
                            qrcode = qrcodeCell.NumericCellValue.ToString();
                        }
                        catch
                        {
                            qrcode = qrcodeCell.StringCellValue;
                        }

                        string tickeNo = (row + 1).ToString();
                        //try
                        //{
                        //    tickeNo = ticketNoCell.NumericCellValue.ToString();
                        //}
                        //catch
                        //{
                        //    tickeNo = ticketNoCell.StringCellValue;
                        //}
                        string tableNo = "";
                        try
                        {
                            tableNo = tableNoCell.NumericCellValue.ToString();
                        }
                        catch
                        {
                            tableNo = tableNoCell.StringCellValue;
                        }

                        string staffNo = "";
                        try
                        {
                            staffNo = staffCell.NumericCellValue.ToString();
                        }
                        catch
                        {
                            staffNo = staffCell.StringCellValue;
                        }

                        //if ((Convert.ToInt32(tickeNo) >= 1442 && Convert.ToInt32(tickeNo) <= 1600) || Convert.ToInt32(tickeNo) >= 1721 && Convert.ToInt32(tickeNo) <= 1734)
                        //if ((Convert.ToInt32(tickeNo) >= 1574 && Convert.ToInt32(tickeNo) <= 1600) )
                        //if (Convert.ToInt32(tickeNo) == 1595)
                        //{
                        DrawImage(qrcode, staffNo, group, tableNo, tickeNo, name, "");
                        //}
                    }
                }
            }
        }
    }
}
