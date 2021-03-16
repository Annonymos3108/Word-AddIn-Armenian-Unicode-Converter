using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace Armenian_Uniocode_Converter
{
    public partial class Band
    {
        private void Band_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp,
            object toFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchwholeWord = false;
            object matchwildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            
            wordApp.Selection.Find.Execute(ref toFindText, ref matchCase,
                                            ref matchwholeWord, ref matchwildCards,
                                            ref matchSoundLike, ref nmatchAllforms,
                                            ref forward, ref wrap, ref format,
                                            ref replaceWithText, ref replace,
                                            ref matchKashida, ref matchDiactitics,
                                            ref matchAlefHamza, ref matchControl);
        }

        private void Convert_to_Unicode(Microsoft.Office.Interop.Word.Application wordApp)
        {
            //markers
            //question mark // ?            
            FindAndReplace(wordApp, char.ConvertFromUtf32((int)177),
                                    char.ConvertFromUtf32((int)1374));

            //emphasis // '
            FindAndReplace(wordApp, char.ConvertFromUtf32((int)176),
                                    char.ConvertFromUtf32((int)1371));

            //exclamation mark // !
            FindAndReplace(wordApp, char.ConvertFromUtf32((int)175),
                                    char.ConvertFromUtf32((int)1372));

            //punctuation mark //՝
            FindAndReplace(wordApp, char.ConvertFromUtf32((int)170),
                                    char.ConvertFromUtf32((int)1373));

            // ' //apatarts
            FindAndReplace(wordApp, char.ConvertFromUtf32((int)254),
                                    char.ConvertFromUtf32((int)1370));

            // letter "Yev"
            FindAndReplace(wordApp, char.ConvertFromUtf32((int)168),
                                    char.ConvertFromUtf32((int)1415));

            //small letters
            int i = 1377;
            int j = 1198;
            int x = 0;
            do
            {
                x = i - j;
                FindAndReplace(wordApp, char.ConvertFromUtf32((int)x),
                                        char.ConvertFromUtf32((int)i));
                i++;
                j--;
            }
            while (i < 1415);

            //capital letters
            int k = 1329;
            int l = 1151;
            int c = 0;
            do
            {
                c = k - l;
                FindAndReplace(wordApp, char.ConvertFromUtf32((int)c),
                                        char.ConvertFromUtf32((int)k));
                k++;
                l--;
            } while (k < 1367);

            //Quotes // " " //<<  >>
            FindAndReplace(wordApp, char.ConvertFromUtf32((int)166), 
                                    char.ConvertFromUtf32((int)'»'));
            FindAndReplace(wordApp, char.ConvertFromUtf32((int)167), 
                                    char.ConvertFromUtf32((int)'«'));
        }

        private void btn_Armenian_Covert_to_Unicode_Click(object sender, 
            RibbonControlEventArgs e)
        {
            Convert_to_Unicode(Globals.ThisAddIn.Application);
            Globals.ThisAddIn.Application.ActiveDocument.Save();
            MessageBox.Show("Document converted successfully.\nBest Regards, Arthur Nersisyan",
                "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btn_Convert_to_unicode_and_save_to_new_file_Click(object sender,
            RibbonControlEventArgs e)
        {
            Word.Document refDoc = Globals.ThisAddIn.Application.ActiveDocument;
            refDoc.Save();
            string refName = refDoc.Name.Substring(0, refDoc.Name.Length - 5);
            string newDocFullName = refDoc.Path + "\\" + refName + "_unicode.docx";
            refDoc.SaveAs2(newDocFullName);              
            Word.Application newWordApp = Globals.ThisAddIn.Application;
            Convert_to_Unicode(newWordApp);
            newWordApp.ActiveDocument.Save();
            MessageBox.Show("Document converted and saved to a new file.\n" +
               "The file is located in the same directory as the original one.",
               "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btn_Save_to_PDF_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            string sfileName_Document = doc.Name.Substring(0, doc.Name.Length - 5);
            string sPath = doc.Path;
            string sFullpath_pdf = sPath + "\\" + sfileName_Document + ".pdf";
            doc.ExportAsFixedFormat(sFullpath_pdf, Word.WdExportFormat.wdExportFormatPDF,
                OpenAfterExport: true);
        }
    }
}