// <copyright file="Program.cs" company="FishDawg LLC">
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
	using System.Globalization;
	using System.IO;
	using System.Reflection;
	using System.Text;

	internal class Program
	{
		#region Fields

		private static bool _isDebugEnabled;

		#endregion

		#region Constructors

		public Program()
		{
		}

		#endregion

		#region Methods

		private static int Main(string[] args)
		{
			try
			{
				DisplayName();

				Options options = Options.FromArguments(args);
				_isDebugEnabled = options.EnableDebug;

				if (options.ShowHelp)
				{
					DisplayHelp();
				}
				else if (options.ShowVersion)
				{
					DisplayVersion();
				}
				else
				{
					if (options.ShowFormats)
					{
						using (Converter converter = new WordConverter())
						{
							converter.Initialize();
							DisplayFormats("Word", converter.Formats);
						}
						using (Converter converter = new ExcelConverter())
						{
							converter.Initialize();
							DisplayFormats("Excel", converter.Formats);
						}
						using (Converter converter = new PowerPointConverter())
						{
							converter.Initialize();
							DisplayFormats("Powerpoint", converter.Formats);
						}
					}
					else
					{
						using (Converter converter = Converter.Create(options))
						{
							converter.Initialize();

							string inputFilePath = Path.GetFullPath(options.InputFilePath);
							FormatInfo format = converter.GetBestFormat(options.FormatName);
							string outputFilePath = Path.GetFullPath(converter.GetBestOutputFilePath(inputFilePath, format, options.OutputFilePath));

							DisplayFilePath(outputFilePath);
							converter.Convert(inputFilePath, format, outputFilePath, options.Password);
						}
					}
				}

				return 0;
			}
			catch (Exception ex)
			{
				DisplayError(ex);
				return 1;
			}
		}

		private static void DisplayName()
		{
			Assembly assembly = Assembly.GetExecutingAssembly();

			Version version = assembly.GetName().Version;
			Console.WriteLine(string.Format(CultureInfo.CurrentCulture, "Office Convert {0}.{1}", version.Major, version.Minor));

			object[] copyrightAttributes = assembly.GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
			if (copyrightAttributes.Length > 0)
			{
				Console.WriteLine(((AssemblyCopyrightAttribute)copyrightAttributes[0]).Copyright);
			}
		}

		private static void DisplayHelp()
		{
			Console.WriteLine();
			Console.WriteLine("Converts the format of a Microsoft Office document.");
			Console.WriteLine();
			Console.WriteLine("Syntax (Windows style):");
			Console.WriteLine("OfficeConvert.exe [/T type] [/F format] [/O destination] [/P password] source");
			Console.WriteLine("OfficeConvert.exe /L");
			Console.WriteLine();
			Console.WriteLine("  source                  Specifies the file to convert.");
			Console.WriteLine("  /T type                 Specifies the type of document to convert. Inferred");
			Console.WriteLine("                          from file extension if omitted.");
			Console.WriteLine("                            Options: word, excel, powerpoint");
			Console.WriteLine("  /F format               Specifies the format to convert the file to. Default");
			Console.WriteLine("                          used based on type of document if ommitted.");
			Console.WriteLine("  /O destination          Specifies the location to output the converted file.");
			Console.WriteLine("  /P password             Specifies the password used to open the file.");
			Console.WriteLine("  /L                      Lists all supported formats.");
			Console.WriteLine("  /?                      Displays this help.");
			Console.WriteLine();
			Console.WriteLine("Syntax (Postix style):");
			Console.WriteLine("OfficeConvert.exe [--type=TYPE] [--format=FORMAT] [--output=FILE]");
			Console.WriteLine("[--password=PASSWORD] [--version] FILE");
			Console.WriteLine("OfficeConvert.exe --list");
			Console.WriteLine();
			Console.WriteLine("  FILE                    Specifies the file to convert.");
			Console.WriteLine("  --type=TYPE             Specifies the type of document to convert. Inferred");
			Console.WriteLine("                          from file extension if omitted.");
			Console.WriteLine("                            Options: word, excel, powerpoint");
			Console.WriteLine("  --format=FORMAT         Specifies the format to convert the file to. Default");
			Console.WriteLine("                          used based on type of document if ommitted.");
			Console.WriteLine("  --output=FILE           Specifies the location to output the converted file.");
			Console.WriteLine("  --password=PASSWORD     Specifies the password used to open the file.");
			Console.WriteLine("  --list                  Lists all supported formats.");
			Console.WriteLine("  --version               Displays the version information.");
			Console.WriteLine("  --help                  Displays this help.");
			Console.WriteLine();
			Console.WriteLine("Support: officeconvert@fishdawg.com <http://www.fishdawg.com/>");
		}

		private static void DisplayVersion()
		{
			Console.WriteLine();
			Console.WriteLine("License RBSD: Revised BSD <http://www.opensource.org/licenses/bsd-license.php>");
			Console.WriteLine("This is free software: you are free to change and redistribute it.");
			Console.WriteLine("There is NO WARRANTY, to the extent permitted by law.");
		}

		private static void DisplayError(Exception exception)
		{
			string errorMessage = exception.Message.Replace("\r\n", "\n").Replace("\r", "\n");
			Console.WriteLine();
			Console.WriteLine(errorMessage);

			if (_isDebugEnabled)
			{
				Console.WriteLine();
				Console.WriteLine("Exception:");
				Console.WriteLine(exception);
			}
		}

		private static void DisplayFormats(string typeName, IList<FormatInfo> formats)
		{
			Console.WriteLine();
			Console.WriteLine(string.Format(CultureInfo.CurrentCulture, "{0} Formats:", typeName));
			foreach (FormatInfo format in formats)
			{
				Console.WriteLine(string.Format(CultureInfo.CurrentCulture, "{0}  {1} ({2})", format.Name.PadRight(8), format.Description, format.DefaultFileExtension));
			}
		}

		private static void DisplayFilePath(string filePath)
		{
			Console.WriteLine();
			Console.WriteLine(filePath);
		}

		#endregion
	}
}
