// <copyright file="PowerPointConverter.cs" company="FishDawg LLC">
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
	using PowerPoint = Microsoft.Office.Interop.PowerPoint;

	internal class PowerPointConverter : Converter
	{
		#region Fields

		private PowerPoint.Application _application;
		private List<FormatInfo> _formats;

		#endregion

		#region Constructors

		public PowerPointConverter()
		{
			this._application = new PowerPoint.ApplicationClass();
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

			string sourceFileName = inputFilePath;
			MsoTriState openReadonly = MsoTriState.msoTrue;
			MsoTriState untitled = MsoTriState.msoFalse;
			MsoTriState withWindow = MsoTriState.msoFalse;

			PowerPoint.Presentation presentation = this._application.Presentations.Open(sourceFileName, openReadonly, untitled, withWindow);

			string targetFileName = outputFilePath;
			PowerPoint.PpSaveAsFileType saveFormat = (PowerPoint.PpSaveAsFileType)format.SaveFormat;
			MsoTriState embedTrueTypeFonts = MsoTriState.msoTriStateMixed;

			presentation.SaveAs(targetFileName, saveFormat, embedTrueTypeFonts);

			presentation.Close();

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

			PowerPoint.FileConverters fileConverters = this._application.FileConverters;

			List<FormatInfo> formats = new List<FormatInfo>(11 + fileConverters.Count);
			formats.Add(new FormatInfo("pptx", "PowerPoint 2007+ Presentation", "pptx", PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation));
			formats.Add(new FormatInfo("ppt", "PowerPoint 97-2003 Presentation", "ppt", PowerPoint.PpSaveAsFileType.ppSaveAsPresentation));
			formats.Add(new FormatInfo("pdf", "PDF Document", "pdf", PowerPoint.PpSaveAsFileType.ppSaveAsPDF));
			formats.Add(new FormatInfo("xps", "XPS Document", "xps", PowerPoint.PpSaveAsFileType.ppSaveAsXPS));
			formats.Add(new FormatInfo("ppsx", "PowerPoint 2007+ Show", "ppsx", PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLShow));
			formats.Add(new FormatInfo("pps", "PowerPoint 97-2003 Show", "pps", PowerPoint.PpSaveAsFileType.ppSaveAsShow));
			formats.Add(new FormatInfo("mhtml", "Web Archive", "mht", PowerPoint.PpSaveAsFileType.ppSaveAsWebArchive));
			formats.Add(new FormatInfo("html", "Web Page", "htm", PowerPoint.PpSaveAsFileType.ppSaveAsHTMLv3));
			formats.Add(new FormatInfo("rtf", "Rich Text Format", "rtf", PowerPoint.PpSaveAsFileType.ppSaveAsRTF));
			formats.Add(new FormatInfo("xml", "XML Presentation", "xml", PowerPoint.PpSaveAsFileType.ppSaveAsXMLPresentation));
			formats.Add(new FormatInfo("odp", "OpenDocument Presentation", "odp", PowerPoint.PpSaveAsFileType.ppSaveAsOpenDocumentPresentation));

			foreach (PowerPoint.FileConverter fileConverter in fileConverters)
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
			this._application.Quit();
			this._application = null;
		}

		#endregion
	}
}
