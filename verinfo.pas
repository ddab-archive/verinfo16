{ ##
  @FILE                     verinfo.pas
  @COMMENTS                 Version Information Component (16 bit) source code
                            (development split from 32 bit version after v1.0).
  @PROJECT_NAME             Version Information Component
  @PROJECT_DESC             Component that reads version information from files.
  @AUTHOR                   Peter Johnson, LLANARTH, Ceredigion, Wales, UK
  @EMAIL                    peter.johnson@openlink.org
  @WEBSITE                  http://www.delphidabbler.com/
  @COPYRIGHT                © Peter D Johnson, 1998-2003.
  @LEGAL_NOTICE             This components and source code is placed in the
                            public domain. It may be freely copied and
                            circulated on a not for profit basis providing that
                            the code is unmodified and this notice and
                            information about the author and his copyright
                            remains attached to the source code.
  @CREDITS                  In producing this component some techniques were
                            used which were learned from FVersion by PJ Veger,
                            Best, The Netherlands (Feb/96). In particular the
                            method of accessing language and char-set tables was
                            taken from PJ Veger's code.
  @HISTORY(
    @REVISION(
      @VERSION              1.0
      @DATE                 25/04/1998
      @COMMENTS             Original version - 16 bit only.
    )
    @REVISION(
      @VERSION              1.0.1
      @DATE                 09/04/1999
      @COMMENTS             Changed palette where component installs to "PJ
                            Stuff" from "Own".
    )
    @REVISION(
      @VERSION              1.0.2
      @COMMENTS             + Changed palette where component installs from
                              "Own" to "DelphiDabbler".
      @DATE                 10/07/2003
    )
  )
}


unit VerInfo;

interface

uses
  SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,
  Forms, Dialogs, Ver;

type
  {Record holding version numbers}
  TVersionNumber = record
    V1, V2, V3, V4 : Word;
  end;

  {The component class}
  TPJVersionInfo = class(TComponent)
  private
    FPInfoBuffer : PChar;           {pointer to info buffer}
    FFixedInfo : Tvs_FixedFileInfo; {storage for fixed file info}
    FTransStr : string;             {translation info encoded in string}
    FLanguageCode : Word;           {code number for the language}
    FCharSetCode : Word;            {code number for the char-set}
    {property storage}
    FFileName : string;
    FHaveInfo : Boolean;
    {general private methods}
    procedure GetInfoBuffer(Len : Word);
      {Creates an info buffer of required size}
    {property access methods}
    procedure SetFileName(AName : string);
    function GetFileVersionNumber : TVersionNumber;
    function GetProductVersionNumber : TVersionNumber;
    function GetFileOS : LongInt;
    function GetFileType : LongInt;
    function GetFileSubType : LongInt;
    function GetFileFlagsMask : LongInt;
    function GetFileFlags : LongInt;
    function GetLanguage : string;
    function GetCharSet : string;
  protected
    {protected methods}
    procedure ClearProperties; virtual;
      {forces properties to return cleared values}
    procedure ReadVersionInfo; virtual;
      {reads version info from file}
    {property access method}
    function GetStringFileInfo(Index : integer) : string; virtual;
  public
    {class constructor & destructor}
    constructor Create(AOwner : TComponent); override;
    destructor Destroy; override;
    property HaveInfo : Boolean read FHaveInfo;
      {Property true if file version info for the file per FileName property has
      been successfully read}
    property FileVersionNumber : TVersionNumber read GetFileVersionNumber;
      {Version number of file, in numeric format, from fixed file info}
    property ProductVersionNumber : TVersionNumber read GetProductVersionNumber;
      {Version number of product, in numeric format, from fixed file info}
    property FileOS : LongInt read GetFileOS;
      {Code describing operating system to be used by file}
    property FileType : LongInt read GetFileType;
      {Code descibing type of file}
    property FileSubType : LongInt read GetFileSubType;
      {Code describing sub-type of file - only used for certain values of
      FileType property}
    property FileFlagsMask : LongInt read GetFileFlagsMask;
      {Code describing which FileFlags are valid}
    property FileFlags : LongInt read GetFileFlags;
      {Flags describing file state}
    property Comments : string  index 0 read GetStringFileInfo;
      {String file info property giving user defined comments}
    property CompanyName : string index 1 read GetStringFileInfo;
      {String file info property giving name of company}
    property FileDescription : string index 2 read GetStringFileInfo;
      {String file info property giving description of file}
    property FileVersion : string index 3 read GetStringFileInfo;
      {String file info property giving version number of file in string format}
    property InternalName : string index 4 read GetStringFileInfo;
      {String file info property giving internal name of file}
    property LegalCopyright : string index 5 read GetStringFileInfo;
      {String file info property giving copyright message}
    property LegalTrademarks : string index 6 read GetStringFileInfo;
      {String file info property giving trademark info}
    property OriginalFileName : string index 7 read GetStringFileInfo;
      {String file info property giving original name of file}
    property PrivateBuild : string index 8 read GetStringFileInfo;
      {String file info property giving information about a private build of
      file}
    property ProductName : string index 9 read GetStringFileInfo;
      {String file info property giving name of product}
    property ProductVersion : string index 10 read GetStringFileInfo;
      {String file info property giving version number of product in string
      format}
    property SpecialBuild : string index 11 read GetStringFileInfo;
      {String file info property giving information about a special build of
      file}
    property Language : string read GetLanguage;
      {Name of language in use}
    property CharSet : string read GetCharSet;
      {Name of character set in use}
  published
    property FileName : string read FFileName write SetFileName;
      {Name of file to which version information relates}
  end;

procedure Register;

implementation

{ --- Public methods for class --- }

constructor TPJVersionInfo.Create(AOwner : TComponent);
  {Class constructor}
begin
  inherited Create(AOwner);
  {Default is no file name - refers to executable file for application}
  FileName := '';
end;

destructor TPJVersionInfo.Destroy;
  {Class destructor}
begin
  {Ensure that info buffer is freed if allocated}
  if FPInfoBuffer <> nil then
    StrDispose(FPInfoBuffer);
  inherited Destroy;
end;

procedure TPJVersionInfo.ClearProperties;
  {Set that file version info not read - this effectively clears properties
  since each property read access method checks this flag before returning
  result}
begin
  FHaveInfo := False;
end;

procedure TPJVersionInfo.ReadVersionInfo;
  {Read version info from file}
type
  TPFFI = ^Tvs_FixedFileInfo;     {pointer to Tvs_FixedFileInfo structure}
  TLangCharSet = record           {record to hold language & char set codes}
    Lang, CharSet : Word;
  end;
  TPLangCharSet = ^TLangCharSet;  {pointer to TLangCharSet}
var
  PFileName : PChar;              {0 terminated string for file name}
  Len : Word;                     {length of structures returned from API calls}
  Ptr : Pointer;                  {pointer to version info structures}
  PFixed : TPFFI;                 {pointer to fixed file info structure}
  PLCSet : TPLangCharSet;         {pointer to language & char set array element}
  InfoSize : integer;             {size of info buffer}
  InfoHandle : LongInt;           {handle to file version info resource}
begin
  {Record default value of HaveInfo property - no info read}
  FHaveInfo := False;
  {Allocate storage & copy file name to 0 terminated string}
  PFileName := StrAlloc(Length(FFileName) + 1);
  StrPCopy(PFileName, FFileName);
  {Record required size of version info buffer}
  InfoSize := GetFileVersionInfoSize(PFileName, InfoHandle);
  {Check that there was no error}
  if InfoSize > 0 then
  begin
    {Found info size OK}
    {Ensure we have a sufficiently large buffer allocated}
    GetInfoBuffer(InfoSize);
    {Read file version info into storage and check this has happened
    successfully}
    if GetFileVersionInfo(PFileName, InfoHandle, InfoSize, FPInfoBuffer) then
    begin
      {Success - we've read file version info to storage OK}
      FHaveInfo := True;
      {Get fixed file info & copy to own storage}
      VerQueryValue(FPInfoBuffer, '\', Ptr, Len);
      FFixedInfo := TPFFI(Ptr)^;
      {Get first translation table info}
      {get the info}
      VerQueryValue(FPInfoBuffer, '\VarFileInfo\Translation', Ptr, Len);
      {make our pointer point to it}
      PLCSet := TPLangCharSet(Ptr);
      {record first entry in langauge/char-set table - we ignore all (any)
      others}
      FLanguageCode := PLCSet^.Lang;
      FCharSetCode := PLCSet^.CharSet;
      {build a string holding language/char-set info}
      FTransStr := Format('%4.4x%4.4x',[FLanguageCode, FCharSetCode]);
      {Get string file information}
    end;
  end;
  {dispose of storage for file name}
  StrDispose(PFileName);
end;

procedure TPJVersionInfo.GetInfoBuffer(Len : Word);
  {Allocate an info buffer of required size, ensuring that any previous buffer
  is first cleared}
begin
  if FPInfoBuffer <> nil then
    StrDispose(FPInfoBuffer);
  FPInfoBuffer := StrAlloc(Len);
end;

function TPJVersionInfo.GetStringFileInfo(Index : integer) : string;
  {Read access method for all string file info properties - returns appropriate
  string for the given property}
var
  CommandBuf : array[0..255] of char;   {buffer to build API call command str}
  Ptr : Pointer;                        {pointer to result of API call}
  Len : Word;                           {length of structure returned from API}
const
  CNames : array[0..11] of string[20] = {names of all string file info strings}
    ('Comments', 'CompanyName', 'FileDescription', 'FileVersion',
    'InternalName', 'LegalCopyright', 'LegalTrademarks', 'OriginalFileName',
    'PrivateBuild', 'ProductName', 'ProductVersion', 'SpecialBuild');
begin
  {Set default failure result to empty string}
  Result := '';
  {Check if we have valid information recorded in info buffer - exit if not}
  if not FHaveInfo then Exit;
  {Build API call command string for reading string file info for required
  language and character set}
  StrPCopy(CommandBuf, '\StringFileInfo\' + FTransStr + '\' + CNames[Index]);
  {Call API to get required string and return it if successful}
  if VerQueryValue(FPInfoBuffer, CommandBuf, Ptr, Len) then
    Result := StrPas(PChar(Ptr));
end;

procedure TPJVersionInfo.SetFileName(AName : string);
  {Write access method for FileName property - action at design time is
  different to run time}
begin
  if csDesigning in ComponentState then
    {We are designing, simply record the required name}
    FFileName := AName
  else
  begin
    {It's run-time}
    {use Application exec file name if name is ''}
    if AName = '' then
      FFileName := Application.ExeName
    else
      FFileName := AName;
    {Clear all properties and read file version info for new file}
    ClearProperties;
    ReadVersionInfo;
  end;
end;

function TPJVersionInfo.GetFileVersionNumber : TVersionNumber;
  {Read access method for FileVersionNumber property}
begin
  if FHaveInfo then
  begin
    {We've got some file version info - fill structure with required info}
    Result.V1 := HiWord(FFixedInfo.dwFileVersionMS);
    Result.V2 := LoWord(FFixedInfo.dwFileVersionMS);
    Result.V3 := HiWord(FFixedInfo.dwFileVersionLS);
    Result.V4 := LoWord(FFixedInfo.dwFileVersionLS);
  end
  else
  begin
    {We've not got any file version info - set structure to zeros}
    Result.V1 := 0;
    Result.V2 := 0;
    Result.V3 := 0;
    Result.V4 := 0;
  end;
end;

function TPJVersionInfo.GetProductVersionNumber : TVersionNumber;
  {Read access method for ProductVersionNumber property}
begin
  if FHaveInfo then
  begin
    {We've got some file version info - fill structure with required info}
    Result.V1 := HiWord(FFixedInfo.dwProductVersionMS);
    Result.V2 := LoWord(FFixedInfo.dwProductVersionMS);
    Result.V3 := HiWord(FFixedInfo.dwProductVersionLS);
    Result.V4 := LoWord(FFixedInfo.dwProductVersionLS);
  end
  else
  begin
    {We've not got any file version info - set structure to zeros}
    Result.V1 := 0;
    Result.V2 := 0;
    Result.V3 := 0;
    Result.V4 := 0;
  end;
end;

function TPJVersionInfo.GetFileOS : LongInt;
  {Read access method for FileOS property - return required value if we have
  some file version info and 0 if we haven't}
begin
  if FHaveInfo then
    Result := FFixedInfo.dwFileOS
  else
    Result := 0;
end;

function TPJVersionInfo.GetFileType : LongInt;
  {Read access method for FileType property - return required value if we have
  some file version info and 0 if we haven't}
begin
  if FHaveInfo then
    Result := FFixedInfo.dwFileType
  else
    Result := 0;
end;

function TPJVersionInfo.GetFileSubType : LongInt;
  {Read access method for FileSubType property - return required value if we
  have some file version info and 0 if we haven't}
begin
  if FHaveInfo then
    Result := FFixedInfo.dwFileSubType
  else
    Result := 0;
end;

function TPJVersionInfo.GetFileFlagsMask : LongInt;
  {Read access method for FileFlagsMask property - return required value if we
  have some file version info and 0 if we haven't}
begin
  if FHaveInfo then
    Result := FFixedInfo.dwFileFlagsMask
  else
    Result := 0;
end;

function TPJVersionInfo.GetFileFlags : LongInt;
  {Read access method for FileFlags property - return required value if we have
  some file version info and 0 if we haven't}
begin
  if FHaveInfo then
    Result := FFixedInfo.dwFileFlags
  else
    Result := 0;
end;

function TPJVersionInfo.GetLanguage : string;
  {Read access method for Language property - return string describing language
  if we have some version info and empty string if we haven't}
var
  Buf : array[0..255] of char;  {buffer for storing langauge string from API
                                call}
begin
  if HaveInfo then
  begin
    {We've got some file version info - get language name from API}
    VerLanguageName(FLanguageCode, Buf, 255);
    Result := StrPas(Buf);
  end
  else
    {No info}
    Result := '';
end;

function TPJVersionInfo.GetCharSet : string;
  {Read access method for CharSet property - return string describing character
  setif we have some version info and empty string if we haven't}
const
  CCharSets : array[0..11] of record
    Code : Word;
    Str : string[40];
  end = (         {structure of code numbers and char-set names}
    (Code: 0;    Str: '7-bit ASCII'),
    (Code: 932;  Str: 'Windows, Japan (Shift - JIS X-0208)'),
    (Code: 949;  Str: 'Windows, Korea (Shift - KSC 5601)'),
    (Code: 950;	Str: 'Windows, Taiwan (GB5)'),
    (Code: 1200;	Str: 'Unicode'),
    (Code: 1250;	Str: 'Windows, Latin-2 (Eastern European)'),
    (Code: 1251;	Str: 'Windows, Cyrillic'),
    (Code: 1252;	Str: 'Windows, Multilingual'),
    (Code: 1253;	Str: 'Windows, Greek'),
    (Code: 1254;	Str: 'Windows, Turkish'),
    (Code: 1255;	Str: 'Windows, Hebrew'),
    (Code: 1256;	Str: 'Windows, Arabic')
  );
var
  I : integer;    {loop control}
begin
  if HaveInfo then
  begin
    {We've got some file version info}
    {record default result - unknown}
    Result := 'Unknown';
    {scan table of codes looking for correct entry, if any}
    for I := 0 to 11 do
      if FCharSetCode = CCharSets[I].Code then
      begin
        {found one - record its name}
        Result := CCharSets[I].Str;
        Exit;
      end;
  end
  else
    {No info}
    Result := '';
end;

{ --- Component registration routine --- }

procedure Register;
begin
  RegisterComponents('DelphiDabbler', [TPJVersionInfo]);
end;

end.
