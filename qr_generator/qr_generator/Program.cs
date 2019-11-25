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
using System.Drawing.Text;
using System.Windows.Forms;

namespace qr_generator
{
    class Program
    {
        public static List<string> imageName = new List<string>();
        public static PrivateFontCollection pfc = new PrivateFontCollection();
        public static List<string> strList = new List<string>();
        public static List<string> strList2 = new List<string>();
        static void Main(string[] args)
        {
            pfc.AddFontFile(AppDomain.CurrentDomain.BaseDirectory + "Palatino.ttf");
            if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + "ticketImg"))
            {
                Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + "ticketImg");
            }
            if (!Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + "ticketPdf"))
            {
                Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + "ticketPdf");
            }
            AddList();
            AddList2();
            ReadExcel();
            Console.WriteLine("End");
            Console.WriteLine("Press enter to exit");            
        }

        //DrawImage(title, lastName, firstName, perferredName, empNumber);
        public static string DrawImage(string title, string lastName, string firstName, string perferredName, string empNumber)
        {
            Console.WriteLine("Generating images...");

            try
            {
                string filename = AppDomain.CurrentDomain.BaseDirectory + "ticket.jpg";

                //creating a image object
                System.Drawing.Image bitmap = (System.Drawing.Image)Bitmap.FromFile(filename); // set image 

                ////draw the image object using a Graphics object
                Graphics graphicsImage = Graphics.FromImage(bitmap);
                graphicsImage.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

                Pen blackPen = new Pen(Color.White, 40);
                Rectangle rect = new Rectangle(1458, 251, 210, 210);
                graphicsImage.DrawRectangle(blackPen, rect);


                MessagingToolkit.QRCode.Codec.QRCodeEncoder encoder = new MessagingToolkit.QRCode.Codec.QRCodeEncoder();
                encoder.QRCodeScale = 24;
                Bitmap qrBMP = encoder.Encode(empNumber);
                graphicsImage.DrawImage(qrBMP, 1453, 245, 220, 220);
                Color StringColor = System.Drawing.ColorTranslator.FromHtml("#fff");

                int middle = 1558;
                int fontSize = 55;
                foreach (string str in strList)
                {
                    if (str == empNumber) 
                    {
                        fontSize = 45;
                        break;
                    }
                }
                foreach (string str in strList2)
                {
                    if (str == empNumber)
                    {
                        fontSize = 35;
                        break;
                    }
                }
                string firstLine = title + " " + lastName + " " + firstName;
                Font font = new Font(pfc.Families[0], fontSize, FontStyle.Regular, GraphicsUnit.Pixel);
                int halfWidth = TextRenderer.MeasureText(firstLine, font).Width / 2;
                graphicsImage.DrawString(firstLine, font, new SolidBrush(StringColor), new Point(middle - halfWidth + 17, 487));

                halfWidth = TextRenderer.MeasureText(perferredName, font).Width /2 ;
                graphicsImage.DrawString(perferredName, font, new SolidBrush(StringColor), new Point(middle - halfWidth + 17, 547));

                font = new Font(pfc.Families[0], 55, FontStyle.Regular, GraphicsUnit.Pixel);
                halfWidth = TextRenderer.MeasureText(empNumber, font).Width / 2;
                graphicsImage.DrawString(empNumber, font, new SolidBrush(StringColor), new Point(middle - halfWidth + 5 ,720));

                bitmap.Save(AppDomain.CurrentDomain.BaseDirectory + @"ticketImg/ticket_" + empNumber + ".jpg");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }


            return empNumber + ".png";
        
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


            for (int row = 1; row <= sheet.LastRowNum; row++) // start from row 4
            {
                if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                {
                    IRow irow = sheet.GetRow(row);
     
                    ICell titleCell = irow.GetCell(0);
                    ICell lastNameCell = irow.GetCell(1);
                    ICell firstNameCell = irow.GetCell(2);
                    ICell perferredNameCell = irow.GetCell(3);
                    ICell empNumberCell = irow.GetCell(4);

                    if (titleCell != null && lastNameCell != null && firstNameCell != null && perferredNameCell != null && empNumberCell != null)
                    {
                        string title = titleCell.StringCellValue;
                        string lastName = lastNameCell.StringCellValue;
                        string firstName = firstNameCell.StringCellValue;
                        string perferredName = perferredNameCell.StringCellValue;
                        string empNumber = "";
                        try
                        {
                            empNumber = empNumberCell.StringCellValue;
                        }
                        catch (Exception ex)
                        {
                            empNumber = Convert.ToInt32(empNumberCell.NumericCellValue).ToString();
                        }
                        if (empNumber == "")
                        {
                            File.AppendAllText(AppDomain.CurrentDomain.BaseDirectory + "log.txt", "Row " + row.ToString() + System.Environment.NewLine);
                        }
                        else 
                        {
                            DrawImage(title, lastName, firstName, perferredName, empNumber);
                        }
                        Console.WriteLine(title + " " + lastName + " " + firstName + " " + perferredName + " " + empNumber);
                        //
                    }
                    else {
                        File.AppendAllText(AppDomain.CurrentDomain.BaseDirectory + "log.txt", "Row " + row.ToString() + System.Environment.NewLine);
                    }
                    //break;
                }
            }
        }
    
        public static void AddList()
        {
            strList.Add("1803");
            strList.Add("13721");
            strList.Add("17820");
            strList.Add("28616");
            strList.Add("30270");
            strList.Add("33112");
            strList.Add("36290");
            strList.Add("39494");
            strList.Add("42078");
            strList.Add("42993");
            strList.Add("45308");
            strList.Add("45658");
            strList.Add("50455");
            strList.Add("51615");
            strList.Add("52521");
            strList.Add("52666");
            strList.Add("52956");
            strList.Add("53643");
            strList.Add("54021");
            strList.Add("58151");
            strList.Add("58453");
            strList.Add("59089");
            strList.Add("60883");
            strList.Add("60971");
            strList.Add("68336");
            strList.Add("70281");
            strList.Add("71770");
            strList.Add("72476");
            strList.Add("72681");
            strList.Add("72798");
            strList.Add("72824");
            strList.Add("73649");
            strList.Add("74043");
            strList.Add("74065");
            strList.Add("75792");
            strList.Add("76491");
            strList.Add("77193");
            strList.Add("77743");
            strList.Add("78088");
            
        }

        public static void AddList2()
        {
            strList.Add("45308");
            strList.Add("45658");
            strList.Add("51615");
            strList.Add("52521");
            strList.Add("52956");
            strList.Add("53643");
            strList.Add("54021");
            strList.Add("58453");
            strList.Add("59089");
            strList.Add("60336");
            strList.Add("60883");
            strList.Add("60971");
            strList.Add("68336");
            strList.Add("70281");
            strList.Add("71770");
            strList.Add("72681");
            strList.Add("72824");
            strList.Add("73649");
            strList.Add("74065");
            strList.Add("75792");
            strList.Add("76491");
        }
    }
}
