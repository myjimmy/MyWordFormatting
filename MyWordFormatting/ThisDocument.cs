using System;
using Word = Microsoft.Office.Interop.Word;

namespace MyWordFormatting
{
    public partial class ThisDocument
    {
        const int WordTrue = -1;
        const int WordFalse = 0;

        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            applyBoldFont.Click += new System.EventHandler(applyBoldFont_Click);
            applyItalicFont.Click += new System.EventHandler(applyItalicFont_Click);
            applyUnderlineFont.Click += new System.EventHandler(applyUnderlineFont_Click);
            Startup += new System.EventHandler(ThisDocument_Startup);
            Shutdown += new System.EventHandler(ThisDocument_Shutdown);

        }

        #endregion

        private void applyBoldFont_Click(object sender, EventArgs e)
        {
            if (applyBoldFont.Checked == true)
            {
                fontText.Bold = WordTrue;
            }
            else
            {
                fontText.Bold = WordFalse;
            }
        }

        private void applyItalicFont_Click(object sender, EventArgs e)
        {
            if (applyItalicFont.Checked == true)
            {
                fontText.Italic = WordTrue;
            }
            else
            {
                fontText.Italic = WordFalse;
            }
        }

        private void applyUnderlineFont_Click(object sender, EventArgs e)
        {
            if (applyUnderlineFont.Checked == true)
            {
                fontText.Underline = Word.WdUnderline.wdUnderlineSingle;
            }
            else
            {
                fontText.Underline = Word.WdUnderline.wdUnderlineNone;
            }
        }
    }
}
