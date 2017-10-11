// <copyright file="Options.cs" company="FishDawg LLC">
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

	internal class Options
	{
		#region Fields

		private string _inputFilePath;
		private string _typeName;
		private string _formatName;
		private string _outputFilePath;
		private string _password;
		private bool _showFormats;
		private bool _showVersion;
		private bool _showHelp;
		private bool _enableDebug;

		#endregion

		#region Constructors

		public Options()
		{
		}
		
		#endregion

		#region Properties

		public virtual string InputFilePath
		{
			get
			{
				return this._inputFilePath;
			}
			set
			{
				this._inputFilePath = value;
			}
		}

		public virtual string TypeName
		{
			get
			{
				return this._typeName;
			}
			set
			{
				this._typeName = value;
			}
		}

		public virtual string FormatName
		{
			get
			{
				return this._formatName;
			}
			set
			{
				this._formatName = value;
			}
		}

		public virtual string OutputFilePath
		{
			get
			{
				return this._outputFilePath;
			}
			set
			{
				this._outputFilePath = value;
			}
		}

		public virtual string Password
		{
			get
			{
				return this._password;
			}
			set
			{
				this._password = value;
			}
		}

		public virtual bool ShowFormats
		{
			get
			{
				return this._showFormats;
			}
			set
			{
				this._showFormats = value;
			}
		}

		public virtual bool ShowVersion
		{
			get
			{
				return this._showVersion;
			}
			set
			{
				this._showVersion = value;
			}
		}

		public virtual bool ShowHelp
		{
			get
			{
				return this._showHelp;
			}
			set
			{
				this._showHelp = value;
			}
		}

		public virtual bool EnableDebug
		{
			get
			{
				return this._enableDebug;
			}
			set
			{
				this._enableDebug = value;
			}
		}

		#endregion

		#region Methods

		public static Options FromArguments(string[] arguments)
		{
			if (arguments == null)
			{
				throw new ArgumentNullException("arguments", "The array of arguments cannot be null.");
			}

			Options configuration = new Options();
			configuration.LoadCommandLineArguments(arguments);
			return configuration;
		}

		private void LoadCommandLineArguments(string[] arguments)
		{
			Debug.Assert(arguments != null);

			string incompleteArgumentName = null;
			foreach (string argument in arguments)
			{
				if (!string.IsNullOrEmpty(argument))
				{
					// Read the supplied argument name/value pair
					bool isPostixArgument;
					string argumentName;
					string argumentValue;
					if (incompleteArgumentName == null)
					{
						if (argument.StartsWith("--", StringComparison.OrdinalIgnoreCase))
						{
							isPostixArgument = true;
							int separatorIndex = argument.IndexOf("=", 2, StringComparison.OrdinalIgnoreCase);
							if (separatorIndex != -1)
							{
								argumentName = argument.Substring(2, separatorIndex - 2);
								argumentValue = argument.Substring(separatorIndex + 1);
							}
							else
							{
								argumentName = argument.Substring(2);
								argumentValue = null;
							}
						}
						else if (argument.StartsWith("/", StringComparison.OrdinalIgnoreCase) || argument.StartsWith("-", StringComparison.OrdinalIgnoreCase))
						{
							isPostixArgument = false;
							argumentName = argument.Substring(1);
							argumentValue = null;
						}
						else
						{
							isPostixArgument = false;
							argumentName = null;
							argumentValue = argument;
						}
					}
					else
					{
						isPostixArgument = false;
						argumentName = incompleteArgumentName;
						argumentValue = argument;
						incompleteArgumentName = null;
					}

					// Interpret the parsed argument name-value pair
					if (!isPostixArgument)
					{
						if (argumentName == null)
						{
							if (this._inputFilePath == null)
							{
								this._inputFilePath = argumentValue;
							}
							else
							{
								throw new OptionException(string.Format(CultureInfo.CurrentCulture, "Unexpected argument: {0}", argument));
							}
						}
						else if (string.Compare(argumentName, "t", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(argumentName, "type", StringComparison.OrdinalIgnoreCase) == 0)
						{
							if (!string.IsNullOrEmpty(argumentValue))
							{
								if (this._typeName == null)
								{
									this._typeName = argumentValue;
								}
							}
							else
							{
								incompleteArgumentName = argumentName;
							}
						}
						else if (string.Compare(argumentName, "f", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(argumentName, "format", StringComparison.OrdinalIgnoreCase) == 0)
						{
							if (!string.IsNullOrEmpty(argumentValue))
							{
								if (this._formatName == null)
								{
									this._formatName = argumentValue;
								}
							}
							else
							{
								incompleteArgumentName = argumentName;
							}
						}
						else if (string.Compare(argumentName, "o", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(argumentName, "output", StringComparison.OrdinalIgnoreCase) == 0)
						{
							if (!string.IsNullOrEmpty(argumentValue))
							{
								if (this._outputFilePath == null)
								{
									this._outputFilePath = argumentValue;
								}
							}
							else
							{
								incompleteArgumentName = argumentName;
							}
						}
						else if (string.Compare(argumentName, "p", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(argumentName, "password", StringComparison.OrdinalIgnoreCase) == 0)
						{
							if (!string.IsNullOrEmpty(argumentValue))
							{
								if (this._password == null)
								{
									this._password = argumentValue;
								}
							}
							else
							{
								incompleteArgumentName = argumentName;
							}
						}
						else if (string.Compare(argumentName, "l", StringComparison.OrdinalIgnoreCase) == 0 || string.Compare(argumentName, "list", StringComparison.OrdinalIgnoreCase) == 0)
						{
							this._showFormats = true;
						}
						else if (string.Compare(argumentName, "?", StringComparison.OrdinalIgnoreCase) == 0)
						{
							this._showHelp = true;
						}
						else if (string.Compare(argumentName, "debug", StringComparison.OrdinalIgnoreCase) == 0)
						{
							this._enableDebug = true;
						}
						else
						{
							throw new OptionException(string.Format(CultureInfo.CurrentCulture, "Invalid argument: {0}", argument));
						}
					}
					else
					{
						if (string.Compare(argumentName, "type", StringComparison.OrdinalIgnoreCase) == 0)
						{
							if (!string.IsNullOrEmpty(argumentValue))
							{
								if (this._typeName == null)
								{
									this._typeName = argumentValue;
								}
							}
							else
							{
								throw new OptionException(string.Format(CultureInfo.CurrentCulture, "Incomplete argument: {0}", argumentName));
							}
						}
						else if (string.Compare(argumentName, "format", StringComparison.OrdinalIgnoreCase) == 0)
						{
							if (!string.IsNullOrEmpty(argumentValue))
							{
								if (this._formatName == null)
								{
									this._formatName = argumentValue;
								}
							}
							else
							{
								throw new OptionException(string.Format(CultureInfo.CurrentCulture, "Incomplete argument: {0}", argumentName));
							}
						}
						else if (string.Compare(argumentName, "output", StringComparison.OrdinalIgnoreCase) == 0)
						{
							if (!string.IsNullOrEmpty(argumentValue))
							{
								if (this._outputFilePath == null)
								{
									this._outputFilePath = argumentValue;
								}
							}
							else
							{
								throw new OptionException(string.Format(CultureInfo.CurrentCulture, "Incomplete argument: {0}", argumentName));
							}
						}
						else if (string.Compare(argumentName, "password", StringComparison.OrdinalIgnoreCase) == 0)
						{
							if (!string.IsNullOrEmpty(argumentValue))
							{
								if (this._password == null)
								{
									this._password = argumentValue;
								}
							}
							else
							{
								throw new OptionException(string.Format(CultureInfo.CurrentCulture, "Incomplete argument: {0}", argumentName));
							}
						}
						else if (string.Compare(argumentName, "list", StringComparison.OrdinalIgnoreCase) == 0)
						{
							this._showFormats = true;
						}
						else if (string.Compare(argumentName, "version", StringComparison.OrdinalIgnoreCase) == 0)
						{
							this._showVersion = true;
						}
						else if (string.Compare(argumentName, "help", StringComparison.OrdinalIgnoreCase) == 0)
						{
							this._showHelp = true;
						}
						else if (string.Compare(argumentName, "debug", StringComparison.OrdinalIgnoreCase) == 0)
						{
							this._enableDebug = true;
						}
						else
						{
							throw new OptionException(string.Format(CultureInfo.CurrentCulture, "Invalid argument: {0}", argument));
						}
					}
				}
			}

			if (incompleteArgumentName != null)
			{
				throw new OptionException(string.Format(CultureInfo.CurrentCulture, "Incomplete argument: {0}", incompleteArgumentName));
			}
		}

		#endregion
	}
}
