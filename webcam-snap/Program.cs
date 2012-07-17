/******************************************************
              http://k1im.ru			  
			  Modified by Ilya Klementiev 16/07/2012
*******************************************************/


using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;

namespace webcamsnap
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            /* (TortoiseSVN) agrs on Post-commit:
                PATH DEPTH MESSAGEFILE REVISION ERROR CWD  */

            string revision = "";
            string tag = "image";

            if (args.Length > 0)
                tag = args[0];

            if (args.Length > 4 && tag == "svn")
                revision = "_" + args[4];
            
            string path = String.Format(@"c:\temp\{0}{1}_{2}.jpg", tag, revision, DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss"));

            //Console.WriteLine(path);

            Image image = Capture.GetImage();

            image.Save(path, ImageFormat.Jpeg);

            image.Dispose();
        }
    }
}
