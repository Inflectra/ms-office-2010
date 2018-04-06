using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SpiraProjectAddIn
{
    /// <summary>
    /// This class uses System.Forms.Axhost to convert an Image file to an image type that can be applied to the toolbar item
    /// </summary>
    public class ConvertImage : System.Windows.Forms.AxHost
    {
        private ConvertImage()
            : base(null)
        {
        }

        /// <summary>
        /// Converts the image to the type expected by the toolbar
        /// </summary>
        /// <param name="image"></param>
        /// <returns></returns>
        public static stdole.IPictureDisp Convert(System.Drawing.Image image)
        {
            return (stdole.IPictureDisp)System.
                Windows.Forms.AxHost
                .GetIPictureDispFromPicture(image);
        }
    }
}
