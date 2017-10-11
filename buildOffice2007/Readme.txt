Office Convert 1.1
Copyright (c) 2010 FishDawg LLC

Converts the format of a Microsoft Office document.

Syntax (Windows style):
OfficeConvert.exe [/T type] [/F format] [/O destination] [/P password] source
OfficeConvert.exe /L

  source                  Specifies the file to convert.
  /T type                 Specifies the type of document to convert. Inferred
                          from file extension if omitted.
                            Options: word, excel, powerpoint
  /F format               Specifies the format to convert the file to. Default
                          used based on type of document if ommitted.
  /O destination          Specifies the location to output the converted file.
  /P password             Specifies the password used to open the file.
  /L                      Lists all supported formats.
  /?                      Displays this help.

Syntax (Postix style):
OfficeConvert.exe [--type=TYPE] [--format=FORMAT] [--output=FILE]
[--password=PASSWORD] [--version] FILE
OfficeConvert.exe --list

  FILE                    Specifies the file to convert.
  --type=TYPE             Specifies the type of document to convert. Inferred
                          from file extension if omitted.
                            Options: word, excel, powerpoint
  --format=FORMAT         Specifies the format to convert the file to. Default
                          used based on type of document if ommitted.
  --output=FILE           Specifies the location to output the converted file.
  --password=PASSWORD     Specifies the password used to open the file.
  --list                  Lists all supported formats.
  --version               Displays the version information.
  --help                  Displays this help.

Support: officeconvert@fishdawg.com <http://www.fishdawg.com/>
