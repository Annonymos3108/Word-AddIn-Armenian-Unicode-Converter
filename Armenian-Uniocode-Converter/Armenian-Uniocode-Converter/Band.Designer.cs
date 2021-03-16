
namespace Armenian_Uniocode_Converter
{
    partial class Band : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Band()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.Unicode_Converter = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btn_Armenian_Covert_to_Unicode = this.Factory.CreateRibbonButton();
            this.btn_Convert_to_unicode_and_save_to_new_file = this.Factory.CreateRibbonButton();
            this.Save_to_pdf = this.Factory.CreateRibbonGroup();
            this.btn_Save_to_PDF = this.Factory.CreateRibbonButton();
            this.Unicode_Converter.SuspendLayout();
            this.group1.SuspendLayout();
            this.Save_to_pdf.SuspendLayout();
            this.SuspendLayout();
            // 
            // Unicode_Converter
            // 
            this.Unicode_Converter.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.Unicode_Converter.Groups.Add(this.group1);
            this.Unicode_Converter.Groups.Add(this.Save_to_pdf);
            this.Unicode_Converter.Label = "Armenian Unicode Converter";
            this.Unicode_Converter.Name = "Unicode_Converter";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btn_Armenian_Covert_to_Unicode);
            this.group1.Items.Add(this.btn_Convert_to_unicode_and_save_to_new_file);
            this.group1.Label = "Armenian";
            this.group1.Name = "group1";
            // 
            // btn_Armenian_Covert_to_Unicode
            // 
            this.btn_Armenian_Covert_to_Unicode.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_Armenian_Covert_to_Unicode.Label = "Convert to Unicode";
            this.btn_Armenian_Covert_to_Unicode.Name = "btn_Armenian_Covert_to_Unicode";
            this.btn_Armenian_Covert_to_Unicode.OfficeImageId = "FontsReplaceFonts";
            this.btn_Armenian_Covert_to_Unicode.ShowImage = true;
            this.btn_Armenian_Covert_to_Unicode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Armenian_Covert_to_Unicode_Click);
            // 
            // btn_Convert_to_unicode_and_save_to_new_file
            // 
            this.btn_Convert_to_unicode_and_save_to_new_file.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_Convert_to_unicode_and_save_to_new_file.Label = "Convert to unicode and save to а new file";
            this.btn_Convert_to_unicode_and_save_to_new_file.Name = "btn_Convert_to_unicode_and_save_to_new_file";
            this.btn_Convert_to_unicode_and_save_to_new_file.OfficeImageId = "FontsReplaceFonts";
            this.btn_Convert_to_unicode_and_save_to_new_file.ShowImage = true;
            this.btn_Convert_to_unicode_and_save_to_new_file.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Convert_to_unicode_and_save_to_new_file_Click);
            // 
            // Save_to_pdf
            // 
            this.Save_to_pdf.Items.Add(this.btn_Save_to_PDF);
            this.Save_to_pdf.Label = "Save to PDF";
            this.Save_to_pdf.Name = "Save_to_pdf";
            // 
            // btn_Save_to_PDF
            // 
            this.btn_Save_to_PDF.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_Save_to_PDF.Label = "Save to PDF";
            this.btn_Save_to_PDF.Name = "btn_Save_to_PDF";
            this.btn_Save_to_PDF.OfficeImageId = "FileSaveAsPdfOrXps";
            this.btn_Save_to_PDF.ShowImage = true;
            this.btn_Save_to_PDF.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Save_to_PDF_Click);
            // 
            // Band
            // 
            this.Name = "Band";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.Unicode_Converter);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Band_Load);
            this.Unicode_Converter.ResumeLayout(false);
            this.Unicode_Converter.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.Save_to_pdf.ResumeLayout(false);
            this.Save_to_pdf.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab Unicode_Converter;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Armenian_Covert_to_Unicode;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Convert_to_unicode_and_save_to_new_file;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Save_to_pdf;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Save_to_PDF;
    }

    partial class ThisRibbonCollection
    {
        internal Band Band
        {
            get { return this.GetRibbon<Band>(); }
        }
    }
}
