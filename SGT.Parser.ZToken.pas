{******************************************************************
*  (c)copyrights Corona Ltd. Donetsk 1999
*  Project: Zeos Library
*  Module: Functions for lexic and syntax process (version 1.2)
*  Author: Sergey Seroukhov   E-Mail: voland@cm.dongu.donetsk.ua
*  Date: 01/12/98
*
*  List of changes:
*  1.1  Add new functions for param string split
*  1.2  Add new function to SQL queries processinf
******************************************************************}

unit SGT.Parser.ZToken;

interface
uses Classes, SysUtils;

const
// Special symbols
  _TAB_   = #9;
  _CR_    = #13;
  _NL_    = #10;
  _DELIM_ = ' :;.,+-<>/*%^=()[]|&~@#\`{}'+_TAB_;
  _SPACE_ = ' ';

type
// Lexem types
  TTokenType = (ttUnknown, ttDelim, ttDigit, ttAlpha, ttString, ttCommand);

{******************* Check types functions ********************}

// Check if delimiter
function IsDelim(Value: Char): Boolean;
// Check if white spaces
function IsWhite(Value: Char): Boolean;
// Check if digit
function IsDigit(Value: Char): Boolean;
// Check if alpha
function IsAlpha(Value: Char): Boolean;
// Check if end of line
function IsEOL(Value: Char): Boolean;


{****************** Functions for lexical analize *************}

// Extract lexem
function ExtractToken(var Buffer, Token: String): TTokenType;
// Extract high level lexem
function ExtractTokenEx(var Buffer, Token: String): TTokenType;
// Extract high high leevl lexem
function ExtractHighToken(var Buffer: String; Cmds: TStringList;
                          var Token: String; var CmdNo: Integer): TTokenType;
// Putback lexem to buffer
procedure PutbackToken(var Buffer: String; Value: String);
// Delete begin and end quotes
procedure DeleteQuotes(var Buffer: String);
// Convert string to C-escape string format
function ConvStr(Value: String): String;
// Convert string from C-escape string format
function UnconvStr(Value: String): String;

{**************** Functions for params string processing *********************}

// Extract parameter value by it index
function ExtractParamByNo(Buffer: String; KeyNo: Integer): String;
// Extract parameter value by it name
function ExtractParam(Buffer, Key: String): String;
// Split params string
procedure SplitParams(Buffer: String; ParamNames, ParamValues: TStringList);

{*************** Functions for SQL queries processing ******************}

// Define table names in SELECT query
procedure ExtractSelectTables(Query: String; Tables, Aliases: TStringList);

implementation

// Check if delemiter
function IsDelim(Value: Char): Boolean;
var Temp: String;
begin
  Temp := Value;
  if (Pos(Temp,_DELIM_)<>0) or (Value=_TAB_) or (Value=_NL_) then
    Result := true
  else
    Result := false;
end;

// Check if white space
function IsWhite(Value: Char): Boolean;
begin
  if (Value=_SPACE_) or (Value=_TAB_) then Result := true
  else Result := false;
end;

// Check if digit
function IsDigit(Value: Char): Boolean;
begin
  if ((Value >= '0') and (Value <= '9')) Or (Value='$') then Result := true
  else Result := false;
end;

// Check if alpha
function IsAlpha(Value: Char): Boolean;
begin
  if (not IsDelim(Value)) and (not IsDigit(Value)) then
    Result := true
  else Result := false;
end;

// Check if quotes
function IsQuote(Value: Char): Boolean;
begin
  if (Value = #34) or (Value = #39) then Result := true
  else Result := false;
end;

// Check if end of line
function IsEOL(Value: Char): Boolean;
begin
  if (Value = _NL_) or (Value = _CR_) then Result := true
  else Result := false;
end;

// Convert string to C-escape string format
function ConvStr(Value: String): String;
var I: Integer;
begin

  Result := '';
  for I := 1 to Length(Value) do begin
    case Value[I] of
      #34: Result := Result + '\' + #34;
      #39: Result := Result + '\' + #39;
      _TAB_: Result := Result + '\t';
      _NL_: begin end;
      _CR_: Result := Result + '\n';
      '\': Result := Result + '\\';
      else Result := Result + Value[I];
    end;
  end;
end;

// Convert string from C-escape string format
function UnconvStr(Value: String): String;
var I: Integer;
begin
  Result := '';
  I := 1;
  while I<=Length(Value) do begin
    if Value[I]='\' then begin
      if I=Length(Value) then break;
      Inc(I);
      case Value[I] of
        'n': Result := Result + _CR_;
        't': Result := Result + _TAB_
        else Result := Result + Value[I];
      end;
    end else Result := Result + Value[I];
    Inc(I);
  end;
end;

// Extract lowerlevel token
function ExtractToken(var Buffer, Token: String): TTokenType;
label ExitProc;
var
  p: Integer;
  Quote: String;
begin
  p := 1;
  Result := ttUnknown;
  Token := '';
  if Buffer='' then exit;

  while IsWhite(Buffer[p]) do begin
    Inc(p);
    if Length(Buffer)<p then goto ExitProc;
  end;

  if IsDelim(Buffer[p]) then begin
    Result := ttDelim;
    Token := Buffer[p];
    Inc(p);
    goto ExitProc;
  end;

  if IsEOL(Buffer[p]) then begin
    Token := #13;
    Result := ttDelim;
    Inc(p);
    if IsEOL(Buffer[p]) then Inc(p);
    goto ExitProc;
  end;

  if IsQuote(Buffer[p]) then begin
    Quote := Buffer[p];
    Result := ttString;
    Token := Quote;
    Inc(p);
    while(p <= Length(Buffer)) do begin
      Token := Token + Buffer[p];
      Inc(p);
      if(Buffer[p-1]=Quote) then
       if (Buffer[p-2]<>'\') then break
       else Delete(Token,length(token)-1,1);
    end;
  end else begin
    if IsDigit(Buffer[p]) then Result := ttDigit
    else Result := ttAlpha;
    while(p <= Length(Buffer)) do begin
      Token := Token + Buffer[p];
      Inc(p);
      if IsDelim(Buffer[p]) or IsEOL(Buffer[p]) or IsQuote(Buffer[p]) then
        break;
    end;
  end;

ExitProc:
  Buffer := Copy(Buffer, p, Length(Buffer)-p+1);
end;

// Putback lexem to buffer
procedure PutbackToken(var Buffer: String; Value: String);
begin
  Buffer := Value + _SPACE_ + Buffer;
end;

// Delete begin and end quotes
procedure DeleteQuotes(var Buffer: String);
begin
  if Buffer='' then exit;
  if IsQuote(Buffer[1]) then
    if Buffer[1]=Buffer[Length(Buffer)] then
      Buffer := Copy(Buffer, 2, Length(Buffer)-2)
    else
      Buffer := Copy(Buffer, 2, Length(Buffer)-1);
end;

// Extract param value by it index
function ExtractParamByNo(Buffer: String; KeyNo: Integer): String;
var
  N: Integer;
  Token: String;
  TokenType: TTokenType;
begin
  N := -1;
  Result := '';
  while (Buffer<>'') and (N<KeyNo) do begin
    TokenType := ExtractToken(Buffer, Token);
    if TokenType in [ttAlpha, ttDigit] then Inc(N);
    if (Token='=') and (n<KeyNo) then ExtractToken(Buffer, Token);
  end;

  if n<>KeyNo then exit;

  ExtractToken(Buffer, Token);
  if Token<>'=' then exit;

  TokenType := ExtractToken(Buffer, Token);
  if TokenType = ttString then DeleteQuotes(Token);
  Result := Token;
end;

// Extract param value by it name
function ExtractParam(Buffer, Key: String): String;
var
  Token: String;
  TokenType: TTokenType;
begin
  while Buffer<>'' do begin
    ExtractToken(Buffer, Token);
    if Token=Key then break;
    if Token='=' then ExtractToken(Buffer, Token);
  end;

  if Buffer='' then exit;

  ExtractToken(Buffer, Token);
  if Token<>'=' then exit;

  TokenType := ExtractToken(Buffer, Token);
  if TokenType = ttString then DeleteQuotes(Token);
  Result := Token;
end;

// Split params string
procedure SplitParams(Buffer: String; ParamNames, ParamValues: TStringList);
var
  Token: String;
  TokenType: TTokenType;
begin
  if Assigned(ParamNames) then ParamNames.Clear;
  if Assigned(ParamValues) then ParamValues.Clear;

  while Buffer<>'' do begin
    TokenType := ExtractToken(Buffer, Token);
    if TokenType in [ttUnknown, ttDelim] then continue;

    if (TokenType = ttString) then DeleteQuotes(Token);
    if Assigned(ParamNames) then ParamNames.Add(Token);

    ExtractToken(Buffer, Token);
    if Token<>'=' then begin
      PutbackToken(Buffer, Token);
      if Assigned(ParamValues) then ParamValues.Add('');
    end else begin
      TokenType := ExtractToken(Buffer, Token);
      if (TokenType = ttString) then DeleteQuotes(Token);

      if TokenType in [ttDelim, ttUnknown] then begin
        if Assigned(ParamValues) then ParamValues.Add('');
      end else
        if Assigned(ParamValues) then ParamValues.Add(Token);
    end;
  end;
end;

// Define table names in SELECT query
procedure ExtractSelectTables(Query: String; Tables, Aliases: TStringList);
var
  Token, From, Table, Alias: String;
  TokenType: TTokenType;
  NextTable, Find: Boolean;
  I: Integer;
begin
  Tables.Clear;
  Aliases.Clear;

// Cut string from FROM to GROUP/ORDER/EOL
  repeat
    ExtractToken(Query, Token);
  until (Query='') or (UpperCase(Token)='FROM');
  if Query='' then exit;
  From := '';
  repeat
    TokenType := ExtractToken(Query, Token);
    if (TokenType=ttAlpha) and ((UpperCase(Token)='GROUP')
       or (UpperCase(Token)='ORDER')) then break;
    if not IsEOL(Token[1]) then
      From := From + ' ' + Token;
  until (Query='');

// Extract table names
  NextTable := true;
  while From<>'' do begin
    ExtractToken(From, Token);
    if NextTable then begin
      NextTable := false;
      Table := Token;

// Define table alias
      ExtractToken(From, Token);
      if (UpperCase(Token)='AS') or (Token='=') then ExtractToken(From, Alias)
      else Alias := Table;

      Find := false;
      for I:=0 to Tables.Count-1 do
        if (Tables[I]=Table) and (Aliases[I]=Alias) then begin
          Find := true;
          break;
        end;

      if not Find then begin
        Tables.Add(Table);
        Aliases.Add(Alias);
      end;
      continue;
    end;
    if (UpperCase(Token)='JOIN') or (Token=',') then
      NextTable := true;
  end;
end;

// Extract high level lexem
function ExtractHighToken(var Buffer: String; Cmds:TStringList;
                          var Token: String; var CmdNo: Integer): TTokenType;
var
  TempToken: String;
  I: Integer;
  TokenType: TTokenType;
begin
  TokenType := ExtractToken(Buffer, Token);
  CmdNo := -1;

// Extract float numbers
  if (TokenType=ttDigit) and (Buffer[1]='.') then begin
    ExtractToken(Buffer, TempToken);
    Token := Token + TempToken;
    if IsDigit(Buffer[1]) then begin
      ExtractToken(Buffer, TempToken);
      Token := Token + TempToken;
    end;
  end;

// Define command index
  if (TokenType=ttAlpha) and Assigned(Cmds) then begin
    for I:=0 to Cmds.Count-1 do
      if Cmds[I]=Token then begin
        CmdNo := I;
        TokenType := ttCommand;
        break;
      end;
  end;

  Result := TokenType;
end;

// Extract hight level lexem
function ExtractTokenEx(var Buffer, Token: String): TTokenType;
var
  P: Integer;
  TokenType: TTokenType;
begin
  repeat
    TokenType := ExtractToken(Buffer, Token);
  until (Token<>#10)and(Token<>#13);
  if Token='[' then begin
    TokenType := ttAlpha;
    P := Pos(']',Buffer); Token := '';
    if P>0 then begin
      Token := Copy(Buffer,1,P-1);
      Buffer := Copy(Buffer,P+1,Length(Buffer)-P);
    end;
  end;
  Result := TokenType;
end;

end.