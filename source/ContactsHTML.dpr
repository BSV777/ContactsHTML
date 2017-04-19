program ContactsHTML;

{$APPTYPE CONSOLE}

uses
  Forms,
  SysUtils,
  Parser in 'Parser.pas',
  ooCalc in 'ooCalc.pas';

var
  InputFile, OutputDir, LdifFile : string;
  i : byte;

{$R *.res}

begin
  Application.Initialize;
  if (ParamCount = 0) or ((ParamCount = 1) and (ParamStr(1) = '/?')) then ConsoleWriteLn('Параметры запуска: ContactsHTML.exe [-с файл_contacts.ods] [-o каталог_для_HTML] [-l файл_LDIF]');
  InputFile := ExtractFilePath(Application.EXEName) + 'contacts.ods';
  OutputDir := ExtractFilePath(Application.EXEName);
  LdifFile := '';
  if ParamCount > 0 then
    begin
      for i := 0 to ParamCount do
        begin
          if (ParamStr(i) = '-c') and (ParamCount > i) and FileExists(pchar(ParamStr(i + 1))) then InputFile := ParamStr(i + 1);
          if (ParamStr(i) = '-o') and (ParamCount > i) and DirectoryExists(pchar(ParamStr(i + 1))) then OutputDir := ParamStr(i + 1);
          if (ParamStr(i) = '-l') and (ParamCount > i) and FileExists(pchar(ParamStr(i + 1))) then LdifFile := ParamStr(i + 1);
        end;
    end;
  CalcToHTML(InputFile, OutputDir, LdifFile);
  Application.Run;
end.
