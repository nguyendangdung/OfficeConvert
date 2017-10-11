// <copyright file="WordConverter.cs" company="FishDawg LLC">
//     Copyright (c) 2010, FishDawg LLC
//     All rights reserved.
//     Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
//     * Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
//     * Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
//     * Neither the name of FishDawg LLC nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.
//     THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
// </copyright>

namespace FishDawg.OfficeConvert
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Globalization;
    using System.Text;
    using Microsoft.Office.Core;
    using Word = Microsoft.Office.Interop.Word;
    using System.IO;

    internal class WordConverter : Converter
    {
        #region Fields

        private Word.Application _application;
        private List<FormatInfo> _formats;

        #endregion

        #region Constructors

        public WordConverter(Options options)
            : base(options)
        {
            this._application = new Word.ApplicationClass();
        }

        #endregion

        #region Properties

        public override IList<FormatInfo> Formats
        {
            get
            {
                return this._formats;
            }
        }

        #endregion

        #region Methods

        public override void Initialize()
        {
            base.Initialize();

            if (this._application == null)
            {
                throw new ObjectDisposedException(this.GetType().Name);
            }

            this._formats = LoadFormats();
        }

        public override void Convert(string inputFilePath, FormatInfo format, string outputFilePath, string password)
        {
            base.Convert(inputFilePath, format, outputFilePath, password);

            MsoAutomationSecurity originalAutomationSecurity = this._application.AutomationSecurity;
            this._application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            this._application.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;

            string tempFilePath = Converter.GetTempFileName(".docx");
            if (this.Options.UseAddin && Path.GetExtension(inputFilePath) == ".odt")
            {
                // convert odt to docx
                this.ConvertWithOdfConverter(inputFilePath, tempFilePath);
                inputFilePath = tempFilePath;
            }

            object addToRecentFiles = false;
            object encoding = Type.Missing;

            object sourceFileName = inputFilePath;
            object confirmConversions = Type.Missing;
            object openReadonly = true;
            object openPassword = password ?? Type.Missing;
            object templatePassword = Type.Missing;
            object revert = Type.Missing;
            object newDocumentPassword = Type.Missing;
            object newTemplatePassword = Type.Missing;
            object openFormat = Type.Missing; // Word.WdOpenFormat.wdOpenFormatAuto
            object visible = Type.Missing;
            object openAndRepair = Type.Missing;
            object documentDirection = Type.Missing;
            object noEncodingDialog = Type.Missing;
            object xmlTransform = Type.Missing;

            Word.Document document = this._application.Documents.Open(ref sourceFileName, ref confirmConversions, ref openReadonly, ref addToRecentFiles, ref openPassword, ref templatePassword, ref revert, ref newDocumentPassword, ref newTemplatePassword, ref openFormat, ref encoding, ref visible, ref openAndRepair, ref documentDirection, ref noEncodingDialog, ref xmlTransform);

            object targetFileName = outputFilePath;
            object saveFormat = format.SaveFormat;
            object lockComments = Type.Missing;
            object readPassword = password ?? Type.Missing;
            object writePassowrd = Type.Missing;
            object readOnlyRecommended = Type.Missing;
            object embedTrueTypeFonts = Type.Missing;
            object saveNativePictureFormat = Type.Missing;
            object saveFormsData = Type.Missing;
            object saveAsAOCELetter = Type.Missing;
            object insertLineBreaks = Type.Missing;
            object allowSubstitutions = Type.Missing;
            object lineEnding = Type.Missing;
            object addBiDiMarks = Type.Missing;
            object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
            object originalFormat = Word.WdOriginalFormat.wdOriginalDocumentFormat;
            object routeDocument = true;


            string tempFilePath2 = Converter.GetTempFileName(".docx");
            if (this.Options.UseAddin && format.SaveFormat == (int)Word.WdSaveFormat.wdFormatOpenDocumentText)
            {
                // export to odt using addin
                saveFormat = Word.WdSaveFormat.wdFormatXMLDocument;
                object tempOpenXmlDocument = tempFilePath2;
                document.SaveAs(ref tempOpenXmlDocument, ref saveFormat, ref lockComments, ref readPassword, ref addToRecentFiles, ref writePassowrd, ref readOnlyRecommended, ref embedTrueTypeFonts, ref saveNativePictureFormat, ref saveFormsData, ref saveAsAOCELetter, ref encoding, ref insertLineBreaks, ref allowSubstitutions, ref lineEnding, ref addBiDiMarks);
                ((Word._Document)document).Close(ref saveChanges, ref originalFormat, ref routeDocument);

                this.ConvertWithOdfConverter(tempFilePath2, outputFilePath);
            }
            else
            {
                document.SaveAs(ref targetFileName, ref saveFormat, ref lockComments, ref readPassword, ref addToRecentFiles, ref writePassowrd, ref readOnlyRecommended, ref embedTrueTypeFonts, ref saveNativePictureFormat, ref saveFormsData, ref saveAsAOCELetter, ref encoding, ref insertLineBreaks, ref allowSubstitutions, ref lineEnding, ref addBiDiMarks);
                ((Word._Document)document).Close(ref saveChanges, ref originalFormat, ref routeDocument);
            }

            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }
            if (File.Exists(tempFilePath2))
            {
                File.Delete(tempFilePath2);
            }
            this._application.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsAll;
            this._application.AutomationSecurity = originalAutomationSecurity;
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            QuitApplication();
        }

        private List<FormatInfo> LoadFormats()
        {
            Debug.Assert(this._application != null);

            Word.FileConverters fileConverters = this._application.FileConverters;

            List<FormatInfo> formats = new List<FormatInfo>(11 + fileConverters.Count);
            formats.Add(new FormatInfo("docx", "Word 2007+ Document", "docx", Word.WdSaveFormat.wdFormatDocumentDefault));
            formats.Add(new FormatInfo("doc", "Word 97-2003 Document", "doc", Word.WdSaveFormat.wdFormatDocument97));
            formats.Add(new FormatInfo("pdf", "PDF Document", "pdf", Word.WdSaveFormat.wdFormatPDF));
            formats.Add(new FormatInfo("xps", "XPS Document", "xps", Word.WdSaveFormat.wdFormatXPS));
            formats.Add(new FormatInfo("mhtml", "Web Archive", "mht", Word.WdSaveFormat.wdFormatWebArchive));
            formats.Add(new FormatInfo("html", "Web Page", "htm", Word.WdSaveFormat.wdFormatHTML));
            formats.Add(new FormatInfo("fhtml", "Web Page, Filtered", "htm", Word.WdSaveFormat.wdFormatFilteredHTML));
            formats.Add(new FormatInfo("rtf", "Rich Text Format", "rtf", Word.WdSaveFormat.wdFormatRTF));
            formats.Add(new FormatInfo("txt", "Plain Text", "txt", Word.WdSaveFormat.wdFormatText));
            formats.Add(new FormatInfo("xml", "XML Document", "xml", Word.WdSaveFormat.wdFormatXMLDocument));
            formats.Add(new FormatInfo("odt", "OpenDocument Text", "odt", Word.WdSaveFormat.wdFormatOpenDocumentText));

            foreach (Word.FileConverter fileConverter in fileConverters)
            {
                if (fileConverter.CanSave && !string.IsNullOrEmpty(fileConverter.ClassName))
                {
                    string[] fileExtensions = fileConverter.Extensions.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    if (fileExtensions.Length > 0)
                    {
                        string name = fileConverter.ClassName.Trim().ToLower(CultureInfo.CurrentCulture);

                        bool hasNameConflict;
                        int customFormatNumber = 1;
                        do
                        {
                            hasNameConflict = false;

                            foreach (FormatInfo format in formats)
                            {
                                if (string.Compare(format.Name, name, StringComparison.OrdinalIgnoreCase) == 0)
                                {
                                    hasNameConflict = true;
                                    ++customFormatNumber;
                                    name = string.Format(CultureInfo.CurrentCulture, "{0}-{1}", fileConverter.ClassName.Trim().ToLower(CultureInfo.CurrentCulture), customFormatNumber);
                                    break;
                                }
                            }
                        }
                        while (hasNameConflict);

                        formats.Add(new FormatInfo(name, fileConverter.FormatName, fileExtensions[0], fileConverter.SaveFormat));
                    }
                }
            }

            return formats;
        }

        private void QuitApplication()
        {
            object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
            object originalFormat = Type.Missing;
            object routeDocument = Type.Missing;

            ((Word._Application)this._application).Quit(ref saveChanges, ref originalFormat, ref routeDocument);

            this._application = null;
        }

        #endregion
    }
}
