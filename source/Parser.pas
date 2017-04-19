unit Parser;

interface

uses
  Forms, SysUtils, Variants, Classes, ooCalc, ExtCtrls, StrUtils, Windows;

procedure ConsoleWriteLn(const S: string);
procedure CalcToHTML(InputFile, OutputDir, LdifFile: string);
procedure LdifParser(LdifFile: string);
procedure SaveNotGoodLogins;
procedure SaveNewLDAP;
function EncodeBase64(Value: String): String;
function DecodeBase64(Value: String): String;
function CleanStr(Value: String): String;
function FindUserID(FullUserName: String): integer;

type
  LdapStr = record
  DN                          : string;  // CN=USERNAME,CN=Users,DC=Company,DC=ru
  objectClass                 : string;  // user
  cn                          : string;  // USERNAME
  sn                          : string;  // Фамилия
  title                       : string;  // Должность
  description                 : string;  // Описание
  physicalDeliveryOfficeName  : string;  // Комната
  telephoneNumber             : string;  // Телефон
  givenName                   : string;  // Имя
  initials	                  : string;  // Иниц
  distinguishedName	          : string;  // CN=USERNAME,CN=Users,DC=Company,DC=ru
  instanceType                : string;  //
  whenCreated                 : string;  //
  whenChanged                 : string;  //
  displayName	                : string;  // Полное имя
  uSNCreated                  : string;  //
  uSNChanged                  : string;  //
  department	                : string;  // Отдел
  company	                    : string;  // Организация
  name	                      : string;  // USERNAME
  userAccountControl          : string;  //
  codePage                    : string;  //
  countryCode                 : string;  //
  accountExpires              : string;  //
  sAMAccountName	            : string;  // USERNAME
  userPrincipalName	          : string;  // USERNAME@Company.ru
  ipPhone	                    : string;  // IP тел
  objectCategory	            : string;  // CN=Person,CN=Schema,CN=Configuration,DC=Company,DC=ru
  mail	                      : string;  // USERNAME@Company.ru
  homePhone	                  : string;  // Тел дом
  mobile	                    : string;  // Тел моб
  logonHours                  : string;  //
  adminCount                  : string;  //
  primaryGroupID              : string;  //
  showInAdvancedViewOnly      : string;  //
  servicePrincipalName        : string;  //
  oldLdapUsername             : string;  //
  oldLdapGood                 : boolean;
  end;

const
  OneSecond = 1 / (24 * 60 * 60);

var
  Calc : TopofCalc;
  LdapArray : array of LdapStr;
  StartTime : TDateTime;

implementation

procedure ConsoleWriteLn(const S: string);
var
  NewStr: string;
begin
  SetLength(NewStr, Length(S));
  CharToOem(PChar(S), PChar(NewStr));
  WriteLn(NewStr);
end;

procedure CalcToHTML(InputFile, OutputDir, LdifFile: string);
var
  i, n, d : byte;
  s, t : string;
  e, ld : boolean;
  m, p, UserID : integer;
  HTMLfile, AllUsers, AllUsers2, EmailList, AllEmailList: TextFile;
  tel, Tempfile, line, FullUserName, Otdel : string;
  RepF : TReplaceFlags;
begin
  Otdel := '';
  SetLength(LdapArray, 1);
  RepF := [rfReplaceAll, rfIgnoreCase];
  if not FileExists(InputFile) then
    begin
      ConsoleWriteLn('Файл ' + InputFile + ' не найден.');
      Exit;
    end;
  Tempfile := ExtractFilePath(Application.EXEName) + 'temp' + FormatDateTime('ddmmyyhhmm', Now) +'.ods';
  ConsoleWriteLn('Исходный файл ' + InputFile);
  ConsoleWriteLn('копируется во временный файл ' + Tempfile);
  CopyFile(pchar(InputFile), pchar(Tempfile), True);
  if not FileExists(Tempfile) then
    begin
      ConsoleWriteLn('Файл ' + Tempfile + ' не найден.');
      Exit;
    end;
  ConsoleWriteLn('Открывается файл ' + Tempfile);
  Calc := TopofCalc.OpenTable(Tempfile, False);
  if Calc.ProgLoaded then
    begin
      ConsoleWriteLn('Файл открыт');
      if not DirectoryExists(OutputDir + 'HTML') then MkDir(OutputDir + 'HTML');
      n := Calc.GetCountSheets;
      ConsoleWriteLn('В файле содержится ' + IntToStr(n) + ' листов');
      ConsoleWriteLn('Обрабатываются только листы, у которых в первой ячейке первый символ №');
      AssignFile(AllUsers, OutputDir + 'HTML\' + 'AllUsers.htm');
      Rewrite(AllUsers);
      AssignFile(AllEmailList, OutputDir + 'HTML\' + 'AllEmais.txt');
      Rewrite(AllEmailList);
      WriteLn(AllUsers, '<HTML>');
      WriteLn(AllUsers, '<HEAD>');
      WriteLn(AllUsers, '<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1251">');
      WriteLn(AllUsers, '<TITLE>Адресная книга </TITLE>');
      WriteLn(AllUsers, '</HEAD>');
      WriteLn(AllUsers, '<BODY LANG="ru-RU" DIR="LTR">');
      WriteLn(AllUsers, '<TABLE WIDTH="900" BORDER=1 BORDERCOLOR="#000000" CELLPADDING=2 CELLSPACING=0>');
      WriteLn(AllUsers, '<colgroup>');
      WriteLn(AllUsers, '  <col width="5"></col>');
      WriteLn(AllUsers, '  <col width="160"></col>');
      WriteLn(AllUsers, '  <col width="200"></col>');
      WriteLn(AllUsers, '  <col width="200"></col>');
      WriteLn(AllUsers, '  <col width="50"></col>');
      WriteLn(AllUsers, '  <col width="50"></col>');
      WriteLn(AllUsers, '  <col width="100"></col>');
      WriteLn(AllUsers, '</colgroup>');
      WriteLn(AllUsers, '<tbody>');
      if (LdifFile <> '') and FileExists(LdifFile) then
        begin
          LdifParser(LdifFile);
          AssignFile(AllUsers2, OutputDir + 'HTML\' + 'AllUsers2.htm');
          Rewrite(AllUsers2);
          WriteLn(AllUsers2, '<HTML>');
          WriteLn(AllUsers2, '<HEAD>');
          WriteLn(AllUsers2, '<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1251">');
          WriteLn(AllUsers2, '<TITLE>Адресная книга</TITLE>');
          WriteLn(AllUsers2, '</HEAD>');
          WriteLn(AllUsers2, '<BODY LANG="ru-RU" DIR="LTR">');
          WriteLn(AllUsers2, '<TABLE WIDTH="900" BORDER=1 BORDERCOLOR="#000000" CELLPADDING=2 CELLSPACING=0>');
          WriteLn(AllUsers2, '<colgroup>');
          WriteLn(AllUsers2, '  <col width="5"></col>');
          WriteLn(AllUsers2, '  <col width="160"></col>');
          WriteLn(AllUsers2, '  <col width="200"></col>');
          WriteLn(AllUsers2, '  <col width="200"></col>');
          WriteLn(AllUsers2, '  <col width="50"></col>');
          WriteLn(AllUsers2, '  <col width="50"></col>');
          WriteLn(AllUsers2, '  <col width="100"></col>');
          WriteLn(AllUsers2, '  <col width="100"></col>');
          WriteLn(AllUsers2, '  <col width="100"></col>');
          WriteLn(AllUsers2, '  <col width="100"></col>');
          WriteLn(AllUsers2, '</colgroup>');
          WriteLn(AllUsers2, '<tbody>');
          ld := True;
        end else
        begin
          if LdifFile <> '' then ConsoleWriteLn('Файл ' + LdifFile + ' не найден.');
          ld := False;
        end;
      for i := 1 to n do
        begin
          Calc.ActivateSheetByIndex(i);
          t := CleanStr(Calc.GetCellText(1, 1));
          if (Length(t) > 0) and (t[1] = '№') then
            begin
              s := Calc.GetActiveSheetName;
              ConsoleWriteLn(s);
              AssignFile(HTMLfile, OutputDir + 'HTML\' + s + '.htm');
              Rewrite(HTMLfile);
              AssignFile(EmailList, OutputDir + 'HTML\' + s + '.txt');
              Rewrite(EmailList);
              WriteLn(HTMLfile, '<TABLE WIDTH="900" BORDER=1 BORDERCOLOR="#000000" CELLPADDING=2 CELLSPACING=0>');
              WriteLn(HTMLfile, '<colgroup>');
              WriteLn(HTMLfile, '  <col width="5"></col>');
              WriteLn(HTMLfile, '  <col width="160"></col>');
              WriteLn(HTMLfile, '  <col width="200"></col>');
              WriteLn(HTMLfile, '  <col width="200"></col>');
              WriteLn(HTMLfile, '  <col width="50"></col>');
              WriteLn(HTMLfile, '  <col width="50"></col>');
              WriteLn(HTMLfile, '  <col width="100"></col>');
              WriteLn(HTMLfile, '</colgroup>');
              WriteLn(HTMLfile, '<tbody>');
              line := '  <TR><TD COLSPAN=7 style="font-family:Arial;font-size:16px;font-weight:bold;text-align:center;" BGCOLOR="#00D000"><P>' + StringReplace(s, '_', ' ', RepF) + '</P></TD></TR>';
              WriteLn(AllUsers, line);
              if ld then line := '  <TR><TD COLSPAN=10 style="font-family:Arial;font-size:16px;font-weight:bold;text-align:center;" BGCOLOR="#00D000"><P>' + StringReplace(s, '_', ' ', RepF) + '</P></TD></TR>';
              if ld then WriteLn(AllUsers2, line);
              line := '<TR>';
              WriteLn(HTMLfile, line);
              WriteLn(AllUsers, line);
              if ld then WriteLn(AllUsers2, line);
              line := '  <TD style="font-family:Arial;font-size:12px;text-align:center;font-style:italic;"><P>№<BR>п/п</P></TD>';
              WriteLn(HTMLfile, line);
              WriteLn(AllUsers, line);
              if ld then WriteLn(AllUsers2, line);
              line := '  <TD style="font-family:Arial;font-size:12px;text-align:center;font-style:italic;"><P>' + CleanStr(Calc.GetCellText(1, 2)) + '</P></TD>';
              WriteLn(HTMLfile, line);
              WriteLn(AllUsers, line);
              if ld then WriteLn(AllUsers2, line);
              line := '  <TD style="font-family:Arial;font-size:12px;text-align:center;font-style:italic;"><P>' + CleanStr(Calc.GetCellText(1, 3)) + '</P></TD>';
              WriteLn(HTMLfile, line);
              WriteLn(AllUsers, line);
              if ld then WriteLn(AllUsers2, line);
              line := '  <TD nowrap="" style="font-family:Arial;font-size:12px;text-align:center;font-style:italic;"><P>' + CleanStr(Calc.GetCellText(1, 4)) + '</P></TD>';
              WriteLn(HTMLfile, line);
              WriteLn(AllUsers, line);
              if ld then WriteLn(AllUsers2, line);
              line := '  <TD style="font-family:Arial;font-size:12px;text-align:center;font-style:italic;"><P>' + CleanStr(Calc.GetCellText(1, 7)) + '</P></TD>';
              WriteLn(HTMLfile, line);
              WriteLn(AllUsers, line);
              if ld then WriteLn(AllUsers2, line);
              line := '  <TD style="font-family:Arial;font-size:12px;text-align:center;font-style:italic;"><P>' + CleanStr(Calc.GetCellText(1, 8)) + '</P></TD>';
              WriteLn(HTMLfile, line);
              WriteLn(AllUsers, line);
              if ld then WriteLn(AllUsers2, line);
              line := '  <TD style="font-family:Arial;font-size:12px;text-align:center;font-style:italic;"><P>' + CleanStr(Calc.GetCellText(1, 9)) + '</P></TD>';
              WriteLn(HTMLfile, line);
              WriteLn(AllUsers, line);
              if ld then WriteLn(AllUsers2, line);
              if ld then line := '  <TD style="font-family:Arial;font-size:12px;text-align:center;font-style:italic;"><P>Имя пользователя в LDAP</P></TD>';
              if ld then WriteLn(AllUsers2, line);
              if ld then line := '  <TD style="font-family:Arial;font-size:12px;text-align:center;font-style:italic;"><P>Имя компьютера в LDAP</P></TD>';
              if ld then WriteLn(AllUsers2, line);
              if ld then line := '  <TD style="font-family:Arial;font-size:12px;text-align:center;font-style:italic;"><P>e-mail в LDAP</P></TD>';
              if ld then WriteLn(AllUsers2, line);
              line := '</TR>';
              WriteLn(HTMLfile, line);
              WriteLn(AllUsers, line);
              if ld then WriteLn(AllUsers2, line);
              m := 1;
              e := True;
              while e do
                begin
                  e := False;
                  for d := 1 to 3 do
                    begin
                      t := CleanStr(Calc.GetCellText(m, d));
                      if t <> '' then e := True;
                    end;
                  for d := 1 to 3 do
                    begin
                      t := CleanStr(Calc.GetCellText(m + 1, d));
                      if t <> '' then e := True;
                    end;
                  for d := 1 to 3 do
                    begin
                      t := CleanStr(Calc.GetCellText(m + 2, d));
                      if t <> '' then e := True;
                    end;
                  m := m + 1;
                  t := CleanStr(Calc.GetCellText(m, 1));
                  if (t <> '') and (t[1] in ['0'..'9']) then
                    begin
                      line := '<TR>';
                      WriteLn(HTMLfile, line);
                      WriteLn(AllUsers, line);
                      if ld then WriteLn(AllUsers2, line);
                      line := '  <TD style="font-family:Arial;font-size:12px;text-align:center;"><P>' + CleanStr(Calc.GetCellText(m, 1)) + '</P></TD>';
                      WriteLn(HTMLfile, line);
                      WriteLn(AllUsers, line);
                      if ld then WriteLn(AllUsers2, line);
                      t := CleanStr(Calc.GetCellText(m, 2));
                      FullUserName := CleanStr(t);
                      p := Pos(' ', t);
                      if p <> 0 then
                        begin
                          line := '  <TD style="font-family:Arial;font-size:13px;text-align:left;"><P>' + AnsiUpperCase(LeftStr(t, p - 1)) + '<BR />' + CleanStr(RightStr(t, Length(t) - p)) + '</P></TD>';
                          FullUserName := AnsiUpperCase(LeftStr(t, p - 1)) + ' ' + CleanStr(RightStr(t, Length(t) - p));
                        end else
                        begin
                          if t <> '' then
                            begin
                              line := '  <TD style="font-family:Arial;font-size:13px;text-align:left;"><P>t</P></TD>';
                              FullUserName := t;
                            end else
                            begin
                              line := '  <TD style="font-family:Arial;font-size:13px;text-align:left;"><P>&nbsp</P></TD>';
                              FullUserName := '';
                            end;
                        end;
                      UserID := FindUserID(FullUserName);
                      if UserID <> 0 then
                        begin
                          LdapArray[UserID].displayName := FullUserName;
                          LdapArray[UserID].sn := AnsiUpperCase(LeftStr(t, p - 1));
                          LdapArray[UserID].givenName := CleanStr(RightStr(t, Length(t) - p));
                          LdapArray[UserID].initials := AnsiUpperCase(Copy(LdapArray[UserID].givenName, 1, 1) + '.' + Copy(LdapArray[UserID].givenName, Pos(' ', LdapArray[UserID].givenName) + 1, 1) + '.');
                          LdapArray[UserID].company := StringReplace(s, '_', ' ', RepF);
                        end;
                      WriteLn(HTMLfile, line);
                      WriteLn(AllUsers, line);
                      if ld then WriteLn(AllUsers2, line);
                      t := CleanStr(Calc.GetCellText(m, 3));
                      if t <> '' then
                        begin
                          line := '  <TD style="font-family:Arial;font-size:12px;text-align:left;"><P>' + t + '</P></TD>';
                          if UserID <> 0 then LdapArray[UserID].title := t;
                        end else line := '  <TD style="font-family:Arial;font-size:12px;text-align:left;"><P>&nbsp</P></TD>';
                      WriteLn(HTMLfile, line);
                      WriteLn(AllUsers, line);
                      if ld then WriteLn(AllUsers2, line);
                      tel := CleanStr(Calc.GetCellText(m, 4));
                      if UserID <> 0 then LdapArray[UserID].telephoneNumber := tel;
                      t := CleanStr(Calc.GetCellText(m, 5));
                      if (tel <> '') and (t <> '') then
                        begin
                          if UserID <> 0 then LdapArray[UserID].telephoneNumber := tel + ', ' + t;
                          tel := tel + '<BR>' + t;
                        end;
                      t := CleanStr(Calc.GetCellText(m, 6));
                      if UserID <> 0 then LdapArray[UserID].mobile := t;
                      if (tel <> '') and (t <> '') then tel := tel + '<BR>' + t;
                      if tel = '' then tel := '&nbsp';
                      line := '  <TD style="font-family:Arial;font-size:12px;text-align:center;"><P>' + tel + '</P></TD>';
                      WriteLn(HTMLfile, line);
                      WriteLn(AllUsers, line);
                      if ld then WriteLn(AllUsers2, line);
                      t := CleanStr(Calc.GetCellText(m, 7));
                      if UserID <> 0 then LdapArray[UserID].ipPhone := t;
                      if t <> '' then line := '  <TD style="font-family:Arial;font-size:12px;text-align:center;"><P>' + t + '</P></TD>' else
                        line := '  <TD style="font-family:Arial;font-size:12px;text-align:center;"><P>&nbsp</P></TD>';
                      WriteLn(HTMLfile, line);
                      WriteLn(AllUsers, line);
                      if ld then WriteLn(AllUsers2, line);
                      if CleanStr(Calc.GetCellText(m, 8)) <> '' then
                        begin
                          line := '  <TD style="font-family:Arial;font-size:12px;text-align:center;"><P><A HREF="mailto:' + CleanStr(Calc.GetCellText(m, 8)) + '">' + CleanStr(Calc.GetCellText(m, 8)) + '</A></P></TD>';
                          WriteLn(EmailList, CleanStr(Calc.GetCellText(m, 8)) + ',');
                          WriteLn(AllEmailList, CleanStr(Calc.GetCellText(m, 8)) + ',');
                        end else
                        line := '  <TD style="font-family:Arial;font-size:12px;text-align:center;"><P>&nbsp</P></TD>';
                      WriteLn(HTMLfile, line);
                      WriteLn(AllUsers, line);
                      if ld then WriteLn(AllUsers2, line);
                      t := CleanStr(Calc.GetCellText(m, 9));
                      if UserID <> 0 then LdapArray[UserID].physicalDeliveryOfficeName := t;
                      if t <> '' then line := '  <TD style="font-family:Arial;font-size:10px;text-align:center;"><P>' + t + '</P></TD>' else
                        line := '  <TD style="font-family:Arial;font-size:10px;text-align:center;"><P>&nbsp</P></TD>';
                      WriteLn(HTMLfile, line);
                      WriteLn(AllUsers, line);
                      if ld then WriteLn(AllUsers2, line);
                      if ld then
                        begin
                          if UserID <> 0 then line := '  <TD style="font-family:Arial;font-size:12px;text-align:center;"><P>' + LdapArray[UserID].cn + '</P></TD>' else
                            line := '  <TD style="font-family:Arial;font-size:12px;text-align:center;"><P>&nbsp</P></TD>';
                        end;
                      if ld then WriteLn(AllUsers2, line);
                      if ld then line := '  <TD style="font-family:Arial;font-size:12px;text-align:center;"><P>&nbsp</P></TD>';
                      if ld then WriteLn(AllUsers2, line);
                      if ld then line := '  <TD style="font-family:Arial;font-size:12px;text-align:center;"><P>&nbsp</P></TD>';
                      if ld then WriteLn(AllUsers2, line);
                      line := '</TR>';
                      WriteLn(HTMLfile, line);
                      WriteLn(AllUsers, line);
                      if ld then WriteLn(AllUsers2, line);
                      if UserID <> 0 then LdapArray[UserID].department := Otdel;
                      if UserID <> 0 then LdapArray[UserID].description := LdapArray[UserID].company + ' - ' + LeftStr(LdapArray[UserID].department, 32) + ' - ' + LeftStr(LdapArray[UserID].title, 32);
                    end else
                    begin
                      t := CleanStr(Calc.GetCellText(m, 1));
                      Otdel := t;
                      line := '  <TR><TD COLSPAN=7 style="font-family:Arial;font-size:14px;font-weight:bold;text-align:center;" BGCOLOR="#' + Format('%6x', [Calc.GetCellColor(m, 1)]) + '"><P>' + t + '</P></TD></TR>';
                      WriteLn(HTMLfile, line);
                      WriteLn(AllUsers, line);
                      if ld then line := '  <TR><TD COLSPAN=10 style="font-family:Arial;font-size:14px;font-weight:bold;text-align:center;" BGCOLOR="#' + Format('%6x', [Calc.GetCellColor(m, 1)]) + '"><P>' + t + '</P></TD></TR>';
                      if ld then WriteLn(AllUsers2, line);
                    end;
                end;
              WriteLn(HTMLfile, '</tbody>');
              WriteLn(HTMLfile, '</TABLE>');
              CloseFile(HTMLfile);
              CloseFile(EmailList);
            end else
            begin
              s := Calc.GetActiveSheetName;
              ConsoleWriteLn('Пропускаем лист ' + s);
            end;
        end;
      WriteLn(AllUsers, '</tbody>');
      WriteLn(AllUsers, '</TABLE>');
      WriteLn(AllUsers, '</BODY>');
      WriteLn(AllUsers, '</HTML>');
      CloseFile(AllUsers);
      CloseFile(AllEmailList);
      if ld then
        begin
          WriteLn(AllUsers2, '</tbody>');
          WriteLn(AllUsers2, '</TABLE>');
          WriteLn(AllUsers2, '</BODY>');
          WriteLn(AllUsers2, '</HTML>');
          CloseFile(AllUsers2);
        end;
      ConsoleWriteLn('HTML-страницы сохранены в папке ' + OutputDir + 'HTML\');
      try
        Calc.Destroy;
      except
      end;
      StartTime := Now;
      while (Now - StartTime) < OneSecond do Application.ProcessMessages;
      DeleteFile(pchar(Tempfile));
      SaveNotGoodLogins;
      SaveNewLDAP;
    end else ConsoleWriteLn('Ошибка открытия файла');
end;

function EncodeBase64(Value: String): String;
const
 b64alphabet: PChar = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
  pad: PChar = '====';

  function EncodeChunk(const Chunk: String): String;
  var
    W: LongWord;
    i, n: Byte;
  begin
    n := Length(Chunk); W := 0;
    for i := 0 to n - 1 do
      W := W + Ord(Chunk[i + 1]) shl ((2 - i) * 8);
    Result := b64alphabet[(W shr 18) and $3f] +
              b64alphabet[(W shr 12) and $3f] +
              b64alphabet[(W shr 06) and $3f] +
              b64alphabet[(W shr 00) and $3f];
    if n <> 3 then
      Result := Copy(Result, 0, n + 1) + Copy(pad, 0, 3 - n);   //add padding when out len isn't 24 bits
  end;

begin
  Result := '';
  while Length(Value) > 0 do
  begin
    Result := Result + EncodeChunk(Copy(Value, 0, 3));
    Delete(Value, 1, 3);
  end;
end;


function DecodeBase64(Value: String): String;
const b64alphabet: PChar = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
  function DecodeChunk(const Chunk: String): String;
  var
    W: LongWord;
    i: Byte;
  begin
    W := 0; Result := '';
    for i := 1 to 4 do
      if Pos(Chunk[i], b64alphabet) <> 0 then
        W := W + Word((Pos(Chunk[i], b64alphabet) - 1)) shl ((4 - i) * 6);
    for i := 1 to 3 do
      Result := Result + Chr(W shr ((3 - i) * 8) and $ff);
  end;
begin
  Result := '';
  if Length(Value) mod 4 <> 0 then Exit;
  while Length(Value) > 0 do
  begin
    Result := Result + DecodeChunk(Copy(Value, 0, 4));
    Delete(Value, 1, 4);
  end;
end;

function CleanStr(Value: String): String;
var
  RepF : TReplaceFlags;
  s : string;
begin
  RepF := [rfReplaceAll, rfIgnoreCase];
  s := StringReplace(Value, Chr(39), ' ', RepF);
  s := StringReplace(s, Chr(10), ' ', RepF);
  s := StringReplace(s, '"', ' ', RepF);
  while Pos('  ', s) <> 0 do s := StringReplace(s, '  ', ' ', RepF);
  Result := Trim(s);
end;

procedure LdifParser(LdifFile: string);
var
  LDline, LDline1, LDline2, Domain1, Domain2 : string;
  LDIF : TextFile;
  RepF : TReplaceFlags;
  i : integer;
begin
  i := 1;
  Domain1 := 'ru';
  Domain2 := 'Company';
//  Domain1 := 'local';
//  Domain2 := 'test';
  RepF := [rfReplaceAll, rfIgnoreCase];
  AssignFile(LDIF, LdifFile);
  Reset(LDIF);
  while not EOF(LDIF) do
    begin
      ReadLn(LDIF, LDline1);
      if (Pos('dn: uid=', LDline1) <> 0) and (Pos(',ou=People,dc=Company,dc=ru', LDline1) <> 0) then
        begin
          LDline := StringReplace(LDline1, 'dn: uid=', '', RepF);
          LDline := StringReplace(LDline, ',ou=People,dc=Company,dc=ru', '', RepF);
          i := High(LdapArray) + 1;
          SetLength(LdapArray, i + 1);
          LdapArray[i].cn := LDline;
          LdapArray[i].DN                          := 'CN=' + LdapArray[i].cn + ',CN=Users,DC=' + Domain2 + ',DC=' + Domain1;
          LdapArray[i].objectClass                 := 'user';
          LdapArray[i].sn                          := '';  // Фамилия
          LdapArray[i].title                       := '';  // Должность
          LdapArray[i].description                 := '';  // Описание
          LdapArray[i].physicalDeliveryOfficeName  := '';  // Комната
          LdapArray[i].telephoneNumber             := '';  // Телефон
          LdapArray[i].givenName                   := '';  // Имя
          LdapArray[i].initials	                   := '';  // Инициалы
          LdapArray[i].distinguishedName	         := 'CN=' + LdapArray[i].cn + ',CN=Users,DC=' + Domain2 + ',DC=' + Domain1;
          LdapArray[i].instanceType                := '';
          LdapArray[i].whenCreated                 := '';
          LdapArray[i].whenChanged                 := '';
          LdapArray[i].displayName                 := '';
          LdapArray[i].uSNCreated                  := '';
          LdapArray[i].uSNChanged                  := '';
          LdapArray[i].department	                 := '';  // Отдел
          LdapArray[i].company	                   := '';  // Организация
          LdapArray[i].name	                       := LdapArray[i].cn;
          LdapArray[i].userAccountControl          := '';
          LdapArray[i].codePage                    := '';
          LdapArray[i].countryCode                 := '';
          LdapArray[i].accountExpires              := '';
          LdapArray[i].sAMAccountName	             := LdapArray[i].cn;
          LdapArray[i].userPrincipalName	         := LdapArray[i].cn + '@' + Domain2 + '.' + Domain1;
          LdapArray[i].ipPhone	                   := '';  // IP тел
          LdapArray[i].objectCategory	             := 'CN=Person,CN=Schema,CN=Configuration,DC=' + Domain2 + ',DC=' + Domain1;
          LdapArray[i].mail                        := LdapArray[i].cn + '@' + Domain2 + '.' + Domain1;
          LdapArray[i].homePhone	                 := '';  // Тел дом
          LdapArray[i].mobile	                     := '';  // Тел моб
          LdapArray[i].logonHours                  := '';
          LdapArray[i].adminCount                  := '';
          LdapArray[i].primaryGroupID              := '';
          LdapArray[i].showInAdvancedViewOnly      := '';
          LdapArray[i].servicePrincipalName        := '';
          LdapArray[i].oldLdapUsername             := '';
          LdapArray[i].oldLdapGood                 := False;
        end;
      if (Pos('cn::', LDline1) <> 0) then
        begin
          if Length(LDline1) = 76 then
            begin
              ReadLn(LDIF, LDline2);
              LDline := LDline1 + LDline2;
            end else LDline := LDline1;
          LDline := StringReplace(LDline, 'cn::', '', RepF);
          LDline := StringReplace(LDline, ' ', '', RepF);
          LDline := CleanStr(UTF8Decode(DecodeBase64(LDline)));
          LdapArray[i].oldLdapUsername := LDline;
        end;
    end;
  CloseFile(LDIF);
end;

function FindUserID(FullUserName: String): integer;
var
 i : integer;
begin
  Result := 0;
  if FullUserName = '' then exit;
  for i := 1 to High(LdapArray) do
    begin
      if AnsiLowerCase(LdapArray[i].oldLdapUsername) = AnsiLowerCase(FullUserName) then
        begin
          Result := i;
          LdapArray[i].oldLdapGood := True;
          Exit;
        end;
    end;
end;

procedure SaveNotGoodLogins;
var
 i : integer;
 NotFoundUsers : TextFile;
begin
  AssignFile(NotFoundUsers, ExtractFilePath(Application.EXEName) + 'NotFoundUsers.txt');
  Rewrite(NotFoundUsers);
  for i := 1 to High(LdapArray) do if not LdapArray[i].oldLdapGood then WriteLn(NotFoundUsers, LdapArray[i].cn + '  ' + LdapArray[i].oldLdapUsername);
  CloseFile(NotFoundUsers);
end;

procedure SaveNewLDAP;
var
 i : integer;
 NewLDAP : TextFile;
 s : string;
begin
  AssignFile(NewLDAP, ExtractFilePath(Application.EXEName) + 'NewWinAD.csv');
  Rewrite(NewLDAP);
  s := 'DN,objectClass,cn,sn,title,description,physicalDeliveryOfficeName,telephoneNumber,givenName,initials,distinguishedName,instanceType,whenCreated,';
  s:= s + 'whenChanged,displayName,uSNCreated,uSNChanged,department,company,name,userAccountControl,codePage,countryCode,accountExpires,sAMAccountName,';
  s:= s + 'userPrincipalName,ipPhone,objectCategory,mail,homePhone,mobile,logonHours,adminCount,primaryGroupID,showInAdvancedViewOnly,servicePrincipalName';
  WriteLn(NewLDAP, s);
  for i := 1 to High(LdapArray) do
    if LdapArray[i].oldLdapGood then
      begin
        s := '"' + LdapArray[i].DN + '",';
        s := s + LdapArray[i].objectClass + ',';
        s := s + LdapArray[i].cn + ',';
        s := s + LdapArray[i].sn + ',';
        if LdapArray[i].title <> '' then s := s + '"' + LeftStr(LdapArray[i].title, 64) + '",' else s := s + ',';
        if LdapArray[i].description <> '' then s := s + '"' + LdapArray[i].description + '",' else s := s + ',';
        if LdapArray[i].physicalDeliveryOfficeName <> '' then s := s + '"' + LdapArray[i].physicalDeliveryOfficeName + '",' else s := s + ',';
        if LdapArray[i].telephoneNumber <> '' then s := s + '"' + LdapArray[i].telephoneNumber + '",' else s := s + ',';
        if LdapArray[i].givenName <> '' then s := s + '"' + LdapArray[i].givenName + '",' else s := s + ',';
        if LdapArray[i].initials <> '' then s := s + '"' + LdapArray[i].initials + '",' else s := s + ',';
        if LdapArray[i].distinguishedName <> '' then s := s + '"' + LdapArray[i].distinguishedName + '",' else s := s + ',';
        s := s + LdapArray[i].instanceType + ',';
        s := s + LdapArray[i].whenCreated + ',';
        s := s + LdapArray[i].whenChanged + ',';
        if LdapArray[i].displayName <> '' then s := s + '"' + LdapArray[i].displayName + '",' else s := s + ',';
        s := s + LdapArray[i].uSNCreated + ',';
        s := s + LdapArray[i].uSNChanged + ',';
        if LdapArray[i].department <> '' then s := s + '"' + LeftStr(LdapArray[i].department, 64) + '",' else s := s + ',';
        if LdapArray[i].company <> '' then s := s + '"' + LdapArray[i].company + '",' else s := s + ',';
        s := s + LdapArray[i].name + ',';
        s := s + LdapArray[i].userAccountControl + ',';
        s := s + LdapArray[i].codePage + ',';
        s := s + LdapArray[i].countryCode + ',';
        s := s + LdapArray[i].accountExpires + ',';
        s := s + LdapArray[i].sAMAccountName + ',';
        s := s + LdapArray[i].userPrincipalName + ',';
        s := s + LdapArray[i].ipPhone + ',';
        if LdapArray[i].objectCategory <> '' then s := s + '"' + LdapArray[i].objectCategory + '",' else s := s + ',';
        s := s + LdapArray[i].mail + ',';
        s := s + LdapArray[i].homePhone + ',';
        s := s + LdapArray[i].mobile + ',';
        s := s + LdapArray[i].logonHours + ',';
        s := s + LdapArray[i].adminCount + ',';
        s := s + LdapArray[i].primaryGroupID + ',';
        s := s + LdapArray[i].showInAdvancedViewOnly + ',';
        s := s + LdapArray[i].servicePrincipalName;
        WriteLn(NewLDAP, s);
    end;
  CloseFile(NewLDAP);
  ConsoleWriteLn('Создан файл NewWinAD.csv');
  ConsoleWriteLn('Загрузка в Active Directory командой: csvde.exe -i -f NewWinAD.csv');
end;


end.
