using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PageToPNG
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void PageToPNGBTN_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            try
            {
                Microsoft.Office.Interop.Word.Pane pane = doc.ActiveWindow.Panes[1];
                Selection selection = doc.ActiveWindow.Selection;
                var page = selection.Information[WdInformation.wdActiveEndPageNumber];
                var bits = pane.Pages[page].EnhMetaFileBits;
                using (MemoryStream ms = new MemoryStream(bits))
                {
                    //Metafile mf = new Metafile(ms);
                    //mf.Save("c:\\test.png", ImageFormat.Png);
                    SaveFileDialog sfdlg = new SaveFileDialog(); sfdlg.Filter = "PNG File (*.png) | *.png";
                    if (sfdlg.ShowDialog() == DialogResult.OK)
                    {
                        System.Drawing.Image image = System.Drawing.Image.FromStream(ms);
                        // Assumes myImage is the PNG you are converting
                        using (var b = new Bitmap(image.Width, image.Height))
                        {
                            b.SetResolution(image.HorizontalResolution, image.VerticalResolution);
                            using (var g = Graphics.FromImage(b))
                            {
                                g.Clear(Color.White);
                                g.DrawImageUnscaled(image, 0, 0);
                                b.Save(sfdlg.FileName, ImageFormat.Png);
                            }

                            // Now save b as a JPEG like you normally would
                        }
                    }
                }

            }
            catch (System.Exception ex)
            {
                // Initializes the variables to pass to the MessageBox.Show method.
                string message = ex.Message;
                string caption = "Exception";
                MessageBoxButtons buttons = MessageBoxButtons.OK;
                DialogResult result;

                // Displays the MessageBox.
                result = MessageBox.Show(message, caption, buttons);
            }
        }
    }
}
