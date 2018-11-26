using ImageMagick;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;

namespace qr_generator
{
    class Program
    {
        static void Main(string[] args)
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

            bitmap.Save(AppDomain.CurrentDomain.BaseDirectory + "newTicket.png");
            //bitmap.Save(Response.OutputStream, ImageFormat.Png);



            using (MagickImageCollection collection = new MagickImageCollection())
            {
                // Add first page
                collection.Add(new MagickImage(AppDomain.CurrentDomain.BaseDirectory + "newTicket.png"));
               

                // Create pdf file with two pages
                collection.Write("newTicket.pdf");
            }
        }
    }
}
