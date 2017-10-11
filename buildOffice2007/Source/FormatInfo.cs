﻿// <copyright file="FormatInfo.cs" company="FishDawg LLC">
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
	using System.Text;
	using Excel = Microsoft.Office.Interop.Excel;
	using PowerPoint = Microsoft.Office.Interop.PowerPoint;
	using Word = Microsoft.Office.Interop.Word;

	internal sealed class FormatInfo
	{
		#region Fields

		private readonly string _name;
		private readonly string _description;
		private readonly string _defaultFileExtension;
		private readonly int _saveFormat;

		#endregion

		#region Constructors

		public FormatInfo(string name, string description, string defaultFileExtension, int saveFormat)
		{
			_name = name;
			_description = description;
			_defaultFileExtension = defaultFileExtension;
			_saveFormat = saveFormat;
		}

		public FormatInfo(string name, string description, string defaultFileExtension, Word.WdSaveFormat saveFormat)
		{
			_name = name;
			_description = description;
			_defaultFileExtension = defaultFileExtension;
			_saveFormat = (int)saveFormat;
		}

		public FormatInfo(string name, string description, string defaultFileExtension, Excel.XlFileFormat saveFormat)
		{
			_name = name;
			_description = description;
			_defaultFileExtension = defaultFileExtension;
			_saveFormat = (int)saveFormat;
		}

		public FormatInfo(string name, string description, string defaultFileExtension, PowerPoint.PpSaveAsFileType saveFormat)
		{
			_name = name;
			_description = description;
			_defaultFileExtension = defaultFileExtension;
			_saveFormat = (int)saveFormat;
		}

		#endregion

		#region Properties

		public string Name
		{
			get
			{
				return _name;
			}
		}

		public string Description
		{
			get
			{
				return _description;
			}
		}

		public string DefaultFileExtension
		{
			get
			{
				return _defaultFileExtension;
			}
		}

		public int SaveFormat
		{
			get
			{
				return _saveFormat;
			}
		}

		#endregion
	}
}
