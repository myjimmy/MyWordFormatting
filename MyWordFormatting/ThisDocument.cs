using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
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
            this.applyBoldFont.Click += new System.EventHandler(this.applyBoldFont_Click);
            this.applyItalicFont.Click += new System.EventHandler(this.applyItalicFont_Click);
            this.applyUnderlineFont.Click += new System.EventHandler(this.applyUnderlineFont_Click);
            this.Startup += new System.EventHandler(this.ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(this.ThisDocument_Shutdown);

        }

        #endregion

        private void applyBoldFont_Click(object sender, EventArgs e)
        {
            if (this.applyBoldFont.Checked == true)
            {
                this.fontText.Bold = WordTrue;
            }
            else
            {
                this.fontText.Bold = WordFalse;
            }
        }

        private void applyItalicFont_Click(object sender, EventArgs e)
        {
            if (this.applyItalicFont.Checked == true)
            {
                this.fontText.Italic = WordTrue;
            }
            else
            {
                this.fontText.Italic = WordFalse;
            }
        }

        private void applyUnderlineFont_Click(object sender, EventArgs e)
        {
            if (this.applyUnderlineFont.Checked == true)
            {
                this.fontText.Underline = Word.WdUnderline.wdUnderlineSingle;
            }
            else
            {
                this.fontText.Underline = Word.WdUnderline.wdUnderlineNone;
            }
        }
    }
}
