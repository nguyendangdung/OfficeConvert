// <copyright file="Converter.cs" company="FishDawg LLC">
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
	using System.IO;
	using System.Text;

	internal abstract class Converter : IDisposable
	{
		#region Constructors

		protected Converter()
		{
		}

		#endregion

		#region Properties

		public abstract IList<FormatInfo> Formats
		{
			get;
		}

		#endregion

		#region Methods

		public static Converter Create(Options options)
		{
			if (options == null)
			{
				throw new ArgumentNullException("options", "The options cannot be null.");
			}

			if (options.InputFilePath == null)
			{
				throw new OptionException("No file specified.");
			}

			if (options.TypeName == null)
			{
				string fileExtension = Path.GetExtension(options.InputFilePath);
				return CreateFromFileExtension(fileExtension);
			}
			else
			{
				return CreateFromTypeName(options.TypeName);
			}
		}

		public virtual FormatInfo GetBestFormat(string formatName)
		{
			if (this.Formats == null || this.Formats.Count == 0)
			{
				throw new InvalidOperationException("The converter cannot be used before it has been initialized.");
			}

			if (formatName != null)
			{
				// Find the specified format
				foreach (FormatInfo format in this.Formats)
				{
					if (string.Compare(format.Name, formatName, StringComparison.OrdinalIgnoreCase) == 0)
					{
						return format;
					}
				}

				throw new OptionException(string.Format(CultureInfo.CurrentCulture, "Invalid format: {0}", formatName));
			}

			// Use the default format
			return this.Formats[0];
		}

		public virtual string GetBestOutputFilePath(string inputFilePath, FormatInfo format, string outputFilePath)
		{
			if (inputFilePath == null)
			{
				throw new ArgumentNullException("inputFilePath", "The input file path cannot be null.");
			}

			if (format == null)
			{
				throw new ArgumentNullException("format", "The format cannot be null.");
			}

			if (this.Formats == null || this.Formats.Count == 0)
			{
				throw new InvalidOperationException("The converter cannot be used before it has been initialized.");
			}

			if (outputFilePath == null)
			{
				return Path.ChangeExtension(inputFilePath, format.DefaultFileExtension);
			}
			else if (Directory.Exists(outputFilePath))
			{
				return Path.Combine(outputFilePath, Path.ChangeExtension(Path.GetFileName(inputFilePath), format.DefaultFileExtension));
			}

			return outputFilePath;
		}

		public virtual void Initialize()
		{
			// Do nothing
		}

		public virtual void Convert(string inputFilePath, FormatInfo format, string outputFilePath, string password)
		{
			if (inputFilePath == null)
			{
				throw new ArgumentNullException("inputFilePath", "The input file path cannot be null.");
			}

			if (format == null)
			{
				throw new ArgumentNullException("format", "The format cannot be null.");
			}

			if (outputFilePath == null)
			{
				throw new ArgumentNullException("outputFilePath", "The output file path cannot be null.");
			}

			if (this.Formats == null || this.Formats.Count == 0)
			{
				throw new InvalidOperationException("The converter cannot be used before it has been initialized.");
			}

			// Do nothing
		}

		public void Dispose()
		{
			this.Dispose(true);

			// Suppress finalization in case a derived class implements a finalizer
			GC.SuppressFinalize(this);
		}

		protected virtual void Dispose(bool disposing)
		{
			// Do nothing
		}

		private static Converter CreateFromFileExtension(string fileExtension)
		{
			Debug.Assert(fileExtension != null);

			if (string.Compare(fileExtension, ".doc", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(fileExtension, ".docx", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(fileExtension, ".docm", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(fileExtension, ".odt", StringComparison.OrdinalIgnoreCase) == 0)
			{
				return new WordConverter();
			}
			else if (string.Compare(fileExtension, ".xls", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(fileExtension, ".xlsx", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(fileExtension, ".xlsm", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(fileExtension, ".xlsb", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(fileExtension, ".ods", StringComparison.OrdinalIgnoreCase) == 0)
			{
				return new ExcelConverter();
			}
			else if (string.Compare(fileExtension, ".ppt", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(fileExtension, ".pptx", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(fileExtension, ".pptm", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(fileExtension, ".pps", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(fileExtension, ".ppsx", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(fileExtension, ".odp", StringComparison.OrdinalIgnoreCase) == 0)
			{
				return new PowerPointConverter();
			}

			throw new OptionException(string.Format(CultureInfo.CurrentCulture, "Unrecognized file extension: {0}", fileExtension));
		}

		private static Converter CreateFromTypeName(string typeName)
		{
			Debug.Assert(typeName != null);

			if (string.Compare(typeName, "word", StringComparison.OrdinalIgnoreCase) == 0)
			{
				return new WordConverter();
			}
			else if (string.Compare(typeName, "excel", StringComparison.OrdinalIgnoreCase) == 0)
			{
				return new ExcelConverter();
			}
			else if (string.Compare(typeName, "powerpoint", StringComparison.OrdinalIgnoreCase) == 0)
			{
				return new PowerPointConverter();
			}

			throw new OptionException(string.Format(CultureInfo.CurrentCulture, "Invalid type: {0}", typeName));
		}

		#endregion
	}
}
