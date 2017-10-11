// <copyright file="ExcelConverter.cs" company="FishDawg LLC">
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
	using Excel = Microsoft.Office.Interop.Excel;

	internal class ExcelConverter : Converter
	{
		#region Constants

		private const int ExcelPdfFormat = -20001;
		private const int ExcelXpsFormat = -20002;

		#endregion

		#region Fields

		private Excel.Application _application;
		private List<FormatInfo> _formats;

		#endregion

		#region Constructors

		public ExcelConverter()
		{
			this._application = new Excel.ApplicationClass();
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

			object writeResPassword = Type.Missing;
			object addToMru = false;

			string sourceFileName = inputFilePath;
			object updateLinks = Type.Missing;
			object openReadOnly = true;
			object openFormat = Type.Missing;
			object openPassword = password ?? Type.Missing;
			object ignoreReadOnlyRecommended = Type.Missing;
			object origin = Type.Missing;
			object delimiter = Type.Missing;
			object editable = Type.Missing;
			object notify = Type.Missing;
			object Converter = Type.Missing; // 0
			object LocalDataStoreSlot = Type.Missing;
			object corruptLoad = Type.Missing;

			Excel.Workbook workbook = this._application.Workbooks.Open(sourceFileName, updateLinks, openReadOnly, openFormat, openPassword, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, Converter, addToMru, LocalDataStoreSlot, corruptLoad);

			if (format.SaveFormat != ExcelPdfFormat && format.SaveFormat != ExcelXpsFormat)
			{
				object targetFileName = outputFilePath;
				object saveFormat = format.SaveFormat;
				object savePassword = password ?? Type.Missing;
				object readOnlyRecommended = Type.Missing;
				object createBackup = Type.Missing;
				Excel.XlSaveAsAccessMode accessMode = Excel.XlSaveAsAccessMode.xlNoChange;
				object conflictResolution = Type.Missing;
				object textCodepage = Type.Missing;
				object textVisualLayout = Type.Missing;
				object local = Type.Missing;

				workbook.SaveAs(targetFileName, saveFormat, savePassword, writeResPassword, readOnlyRecommended, createBackup, accessMode, conflictResolution, addToMru, textCodepage, textVisualLayout, local);
			}
			else
			{
				Excel.XlFixedFormatType type;
				switch (format.SaveFormat)
				{
					case ExcelXpsFormat:
						type = Excel.XlFixedFormatType.xlTypeXPS;
						break;
					default:
						type = Excel.XlFixedFormatType.xlTypePDF;
						break;
				}

				object targetFileName = outputFilePath;
				object quality = Type.Missing;
				object includeDocProperties = Type.Missing;
				object ignorePrintAreas = Type.Missing;
				object from = Type.Missing;
				object to = Type.Missing;
				object openAfterPublish = Type.Missing;
				object fixedFormatExtClassPtr = Type.Missing;

				workbook.ExportAsFixedFormat(type, targetFileName, quality, includeDocProperties, ignorePrintAreas, from, to, openAfterPublish, fixedFormatExtClassPtr);
			}

			object saveChanges = false;
			object filename = Type.Missing;
			object routeWorkbook = true;

			workbook.Close(saveChanges, filename, routeWorkbook);

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

			Excel.FileExportConverters fileConverters = this._application.FileExportConverters;

			List<FormatInfo> formats = new List<FormatInfo>(11 + fileConverters.Count);
			formats.Add(new FormatInfo("xlsx", "Excel 2007+ Workbook", "xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook));
			formats.Add(new FormatInfo("xls", "Excel 97-2003 Workbook", "xls", Excel.XlFileFormat.xlExcel8));
			formats.Add(new FormatInfo("xlsb", "Excel Binary Workbook", "xlsb", Excel.XlFileFormat.xlExcel12));
			formats.Add(new FormatInfo("pdf", "PDF Document", "pdf", ExcelPdfFormat));
			formats.Add(new FormatInfo("xps", "XPS Document", "xps", ExcelXpsFormat));
			formats.Add(new FormatInfo("mhtml", "Web Archive", "mht", Excel.XlFileFormat.xlWebArchive));
			formats.Add(new FormatInfo("html", "Web Page", "htm", Excel.XlFileFormat.xlHtml));
			formats.Add(new FormatInfo("txt", "Text (Tab Delimited)", "txt", Excel.XlFileFormat.xlCurrentPlatformText));
			formats.Add(new FormatInfo("csv", "CSV (Comma Delimited)", "csv", Excel.XlFileFormat.xlCSV));
			formats.Add(new FormatInfo("xml", "XML Spreadsheet", "xml", Excel.XlFileFormat.xlXMLSpreadsheet));
			formats.Add(new FormatInfo("ods", "OpenDocument Spreadsheet", "ods", Excel.XlFileFormat.xlOpenDocumentSpreadsheet));

			foreach (Excel.FileExportConverter fileConverter in fileConverters)
			{
				string[] fileExtensions = fileConverter.Extensions.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
				if (fileExtensions.Length > 0)
				{
					string name = fileExtensions[0].Trim().ToLower(CultureInfo.CurrentCulture);

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
								name = string.Format(CultureInfo.CurrentCulture, "{0}-{1}", fileExtensions[0].Trim().ToLower(CultureInfo.CurrentCulture), customFormatNumber);
								break;
							}
						}
					}
					while (hasNameConflict);

					formats.Add(new FormatInfo(name, fileConverter.Description, fileExtensions[0], fileConverter.FileFormat));
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
