// BASED ON OLD Parser Zeos library.
{ ******************************************************************
  *  (c)copyrights Corona Ltd. Donetsk 1999
  *  Project: Zeos Library
  *  Module: Formula parser component
  *  Author: Sergey Seroukhov   E-Mail: voland@cm.dongu.donetsk.ua
  *  Date: 26/03/99
  *
  *  List of changes:
  *  27/03/99 - Class convert to component, add vars
  *  16/04/99 - Add some functions, operators LIKE, XOR
  ******************************************************************* }

unit SGT.Parser.ZParser;

interface

uses SysUtils, Classes, SGT.Parser.ZToken, SGT.Parser.ZMatch, Variants,
  Rtti, Generics.Collections;

type

  TParseItemType = (ptFunction, ptVariable, ptDelim, ptString, ptInteger,
    ptFloat, ptBoolean, ptNull);

  TParseItem = record
    ItemValue: Variant;
    ItemType: TParseItemType;
  end;

  TParser = class;

  EParseException = class(Exception);

  TOnGetVar = procedure(sender: TParser; VarName: string; var Value: Variant)
    of object;

  TGetVar = reference to procedure (sender: TParser; VarName: string; var Value: Variant);

  TFunction = reference to function(sender: TParser; const Args: Array of Variant): Variant;

  TParser = class
  private
    class var FFunctions: TDictionary<String, TFunction>;
    class function GetGlobalFunction(const FuncName: String): TFunction;
    class constructor Create;
    class destructor Destroy;
  public
    class procedure AddFunction(const FuncName: String;
      FunctionCode: TFunction);
  private
    FParseItems: TList<TParseItem>;
    FLocalFunctions: TDictionary<String, TFunction>;
    FErrCheck: Integer;
    FEquation: String;
    FParseStack: TStack<Variant>;
    FVariables: TList<String>;
    FOnGetVar: TOnGetVar;
    FGetVar: TGetVar;
    function ExtractTokenEx(var Buffer, Token: String): TParseItemType;
    function OpLevel(Operat: String): Integer;
    function Parse(Level: Integer; var Buffer: String): Integer;
    procedure SetEquation(Value: String);
    procedure CheckTypes(var Value1, Value2: Variant; IsForOp: Boolean = FALSE);
    function ConvType(Value: Variant): Variant;
    function CheckFunc(var Buffer: String): Boolean;
    function EvalFunction(const ParseItem: TParseItem): Variant;
    procedure Push(const Value: Variant);
    function Pop: Variant;
    function GetVariables: TArray<String>;
    function GetFunction(FuncName: String): TFunction;
  protected
    function DoGetVar(VarName: String): Variant; Virtual;
  public
    constructor Create(GetVarFunction: TGetVar = nil);
    destructor Destroy; override;
    function Evalute: Variant; overload;
    function Evalute<T>: T; overload;
    procedure AddLocalFunction(const Name: String; AFunction: TFunction);
    procedure Clear;
    property Equation: String read FEquation write SetEquation;
    property OnGetVar: TOnGetVar read FOnGetVar write FOnGetVar;

    property Variables: TArray<String> read GetVariables;
  end;

procedure CheckParamCount(const Args: Array of Variant; ArgCount: byte;
  const FuncName: String); overload;
procedure CheckParamCount(const Args: Array of Variant; ArgCount: Array of byte;
  const FuncName: String); overload;

resourcestring
  rsInvalidParamsNo = 'Invalid param count in function %s';
  rsSyntaxError = 'Syntax error.';
  rsIndexOutOfRange = 'Index out of range.';
  rsTypesmismatch = 'Type mismatch.';
  rsFunctionNotFound = 'Function %s don''t. exists.';
  rsDuplicateFunction = 'Duplicate funcion name (%s).';

implementation

function VarIsNull(const V: Variant): Boolean;
begin
  result := Variants.VarIsNull(V) Or VarIsEmpty(V);
end;

{ TParserFuncion }

function TParser.GetFunction(FuncName: string): TFunction;
begin
  FuncName := LowerCase(FuncName);
  if not FLocalFunctions.TryGetValue(FuncName, result) then
    FFunctions.TryGetValue(FuncName, result);
end;

class function TParser.GetGlobalFunction(const FuncName: String): TFunction;
begin
  FFunctions.TryGetValue(LowerCase(FuncName), result);
end;

function TParser.GetVariables: TArray<String>;
begin
  result := FVariables.ToArray;
end;

{ ************** User functions implementation ************** }

procedure CheckParamCount(const Args: Array of Variant; ArgCount: Array of byte;
  const FuncName: String); overload;
var
  LOk: Boolean;
  Lb, LArg: byte;
begin
  LArg := Length(Args);
  for Lb in ArgCount do
    if Lb = LArg then
      exit;

  raise EParseException.CreateFmt(rsInvalidParamsNo, [FuncName]);
end;

procedure CheckParamCount(const Args: Array of Variant; ArgCount: byte;
  const FuncName: String); overload;
begin
  CheckParamCount(Args, [ArgCount], FuncName);
end;

{ ******************* TParser implementation **************** }

constructor TParser.Create;
begin
  inherited Create;
  FParseItems := TList<TParseItem>.Create;
  FParseItems.Capacity := 10;

  FParseStack := TStack<Variant>.Create;
  FParseStack.Capacity := 10;

  FVariables := TList<String>.Create;
  FErrCheck := 0;
  FLocalFunctions := TDictionary<String, TFunction>.Create;
  FGetVar := GetVarFunction;
end;

// Class destructor
destructor TParser.Destroy;
begin
  FParseItems.Free;
  FParseStack.Free;
  FLocalFunctions.Free;
  FVariables.Free;
  inherited Destroy;
end;

// Extract highlevel lexem
function TParser.ExtractTokenEx(var Buffer, Token: String): TParseItemType;
var
  P: Integer;
  TokenType: TTokenType;
  Temp: string;
begin
  repeat
    TokenType := ExtractToken(Buffer, Token);
  until (Token <> #10) and (Token <> #13);

  if Token = '[' then
  begin
    TokenType := ttAlpha;
    P := Pos(']', Buffer);
    Token := '';
    if P > 0 then
    begin
      Token := Copy(Buffer, 1, P - 1);
      Buffer := Copy(Buffer, P + 1, Length(Buffer) - P);
    end;
  end;

  if Token = '#' then
  begin
    P := Pos('#', Buffer);
    Token := '';
    if P > 0 then
    begin
      Token := Copy(Buffer, 1, P - 1);
      Token := FloatToStr(StrToDateTime(Token));
      if FormatSettings.DecimalSeparator <> '.' then
        Token := StringReplace(Token, FormatSettings.DecimalSeparator, '.', []);
      if Pos('.', Token) = 0 then
        Token := Token + '.0';
      Buffer := Copy(Buffer, P + 1, Length(Buffer) - P);
      result := ptFloat;
    end;
    exit;
  end;

  if (Buffer <> '') and (Token = '>') and (Buffer[1] = '=') then
  begin
    ExtractToken(Buffer, Temp);
    Token := Token + Temp;
  end;
  if (Buffer <> '') and (Token = '<') and ((Buffer[1] = '=') or (Buffer[1] = '>')) then
  begin
    ExtractToken(Buffer, Temp);
    Token := Token + Temp;
  end;

  Temp := UpperCase(Token);
  if (Temp = 'AND') or (Temp = 'NOT') or (Temp = 'OR') or (Temp = 'XOR') or
    (Temp = 'IN') or (Temp = 'LIKE') then
  begin
    Token := Temp;
    result := ptDelim;
    exit;
  end;
  if (Temp = 'TRUE') or (Temp = 'FALSE') then
  begin
    Token := Temp;
    result := ptBoolean;
    exit;
  end;
  if Temp = 'NULL' then
  begin
    Token := Temp;
    result := ptNull;
    exit;
  end;
  if Temp = 'IS' then
  begin
    Token := '=';
    result := ptDelim;
    exit;
  end;
  case TokenType of
    ttString:
      if Pos(FormatSettings.DateSeparator, Token) in [2, 3] then
      begin
        try
          Temp := FloatToStr(StrToDateTime(Copy(Token, 2, Length(Token) - 2)));
          if FormatSettings.DecimalSeparator <> '.' then
            Token := StringReplace(Temp,
              FormatSettings.DecimalSeparator, '.', [])
          else
            Token := Temp;
          result := ptFloat;
        except
          result := ptString;
        end;
      end
      else
        result := ptString;
    ttAlpha:
      result := ptVariable;
    ttDelim:
      result := ptDelim;
    ttDigit:
      begin
        if (Buffer <> '') and (Buffer[1] = '.') then
        begin
          ExtractToken(Buffer, Temp);
          Token := Token + '.';
          if (Buffer <> '') and (Buffer[1] >= '0') and (Buffer[1] <= '9') then
          begin
            ExtractToken(Buffer, Temp);
            Token := Token + Temp;
          end;
          result := ptFloat;
        end
        else
          result := ptInteger;
      end;
  end;
end;

// Get priority level of operation
function TParser.OpLevel(Operat: String): Integer;
var
  Temp: String;
begin
  result := 7;
  Temp := UpperCase(Operat);
  if (Temp = 'AND') or (Temp = 'OR') or (Temp = 'XOR') then
    result := 1;
  if (Temp = 'NOT') then
    result := 2;
  if (Temp = '<') or (Temp = '>') or (Temp = '=') or (Temp = '>=') or (Temp = '<=') or (Temp = '<>') then
    result := 3;
  if (Temp[1] = '+') or (Temp[1] = '-') or (Temp = 'LIKE') then
    result := 4;
  if (Temp[1] = '/') or (Temp[1] = '*') or (Temp[1] = '%') then
    result := 5;
  if (Temp[1] = '^') then
    result := 6;
end;

// Internal convert equation from infix form to postfix
function TParser.Parse(Level: Integer; var Buffer: String): Integer;
var
  ParseType: TParseItemType;
  Token, FuncName: String;
  Temp: Char;
  NewLevel: Integer;
  IValue: Variant;

  procedure AddParseItem(Value: Variant; _Type: TParseItemType);
  var
    Item: TParseItem;
  begin
    Item.ItemValue := Value;
    Item.ItemType := _Type;
    FParseItems.Add(Item);
  end;

  procedure ExtractParams;
  var
    Params, SaveCount: Integer;
  begin
    SaveCount := FParseItems.Count;
    Params := 0;
    repeat
      FErrCheck := 0;
      Parse(0, Buffer);
      ExtractTokenEx(Buffer, Token);
      if Token = '' then
        raise EParseException.Create(rsSyntaxError);
      case Token[1] of
        ',':
          begin
            Inc(Params);
            continue;
          end;
        ')':
          begin
            if SaveCount < { FParseCount } FParseItems.Count then
              Inc(Params);
            { FParseItems[FParseCount].ItemValue := ConvType(Params);
              FParseItems[FParseCount].ItemType := ptInteger;
              Inc(FParseCount); }
            AddParseItem(ConvType(Params), ptInteger);
            break;
          end;
      else
        raise EParseException.Create(rsSyntaxError);
      end;
    until Buffer = '';
  end;

begin
  result := 0;
  while Buffer <> '' do
  begin
    ParseType := ExtractTokenEx(Buffer, Token);

    if Token = '' then
      exit;
    if (Token = ')') or (Token = ',') then
    begin
      PutbackToken(Buffer, Token);
      exit;
    end;

    if Token = '(' then
    begin
      FErrCheck := 0;
      Parse(0, Buffer);
      ExtractTokenEx(Buffer, Token);
      if Token <> ')' then
        raise EParseException.Create(rsSyntaxError);
      FErrCheck := 1;
      continue;
    end;

    if ParseType = ptDelim then
    begin
      NewLevel := OpLevel(Token);
      if (FErrCheck = 2) and (Token <> 'NOT') then
        raise EParseException.Create(rsSyntaxError);
      if FErrCheck = 0 then
        if (Token <> 'NOT') and (Token <> '+') and (Token <> '-') then
          raise EParseException.Create(rsSyntaxError)
        else if Token <> 'NOT' then
          NewLevel := 6;

      if (Token <> 'NOT') and (NewLevel <= Level) then
      begin
        PutbackToken(Buffer, Token);
        result := NewLevel;
        exit;
      end
      else if (Token = 'NOT') and (NewLevel < Level) then
      begin
        PutbackToken(Buffer, Token);
        result := NewLevel;
        exit;
      end;

      if (FErrCheck = 0) and (Token = '+') then
        continue;
      if (FErrCheck = 0) and (Token = '-') then
        Token := '~';
      FErrCheck := 2;
      if Token <> 'IN' then
      begin
        while (Buffer <> '') and (Buffer[1] <> ')') and
          (Parse(NewLevel, Buffer) > NewLevel) do;
      end
      else if not CheckFunc(Buffer) then
        raise EParseException.Create(rsSyntaxError)
      else
      begin
        FuncName := Token;
        ExtractParams;
        Token := FuncName;
      end;
      AddParseItem(Token, ptDelim);
      result := NewLevel;
      continue;
    end;

    if FErrCheck = 1 then
      raise EParseException.Create(rsSyntaxError);
    FErrCheck := 1;
    case ParseType of
      ptVariable:
        begin
          // FParseItems[FParseCount].ItemValue := Token;
          IValue := Token;
          if CheckFunc(Buffer) then
            ParseType := ptFunction
        end;
      ptNull:
        IValue := NULL;
      ptInteger: // FParseItems[FParseCount].ItemValue := StrToInt(Token);
        IValue := StrToInt(Token);
      ptFloat:
        begin
          Temp := FormatSettings.DecimalSeparator;
          FormatSettings.DecimalSeparator := '.';
          IValue := StrToFloat(Token);
          FormatSettings.DecimalSeparator := Temp;
        end;
      ptString:
        begin
          DeleteQuotes(Token);
          IValue := Token;
        end;
      ptBoolean:
        IValue := Token = 'TRUE';
    end;
    if ParseType = ptFunction then
    begin
      FuncName := UpperCase(Token);
      ExtractParams;
      IValue := FuncName;
    end
    else if (ParseType = ptVariable) And not FVariables.Contains
      (LowerCase(Token)) then
      FVariables.Add(LowerCase(Token));

    AddParseItem(ConvType(IValue), ParseType);
  end;
end;

// Split equation to stack
// Value - equation buffer
procedure TParser.SetEquation(Value: String);
begin
  // FParseCount := 0;
  Clear;
  FErrCheck := 0;
  FEquation := Value;
  while Value <> '' do
    Parse(0, Value);

  FVariables.TrimExcess;
end;

// Convert types of two variant values
function TParser.ConvType(Value: Variant): Variant;
begin
  case VarType(Value) of
    varByte, varSmallint, varInteger:
      result := VarAsType(Value, varInteger);
    varSingle, varDouble, varCurrency:
      result := VarAsType(Value, varDouble);
    varOleStr, varString, varVariant, varUString:
      result := VarAsType(Value, varString);
    varDate:
      result := Double(varToDateTime(Value));
    varBoolean:
      result := Value;
    varNull:
      result := NULL;
  else
    raise EParseException.Create(rsTypesmismatch);
  end;
end;

class constructor TParser.Create;
begin
  FFunctions := TDictionary<String, TFunction>.Create;
end;

function StrToFloatEx(Value: String): Double;
var
  Temp: Char;
begin
  result := 0;
  if Value <> '' then
  begin
    Temp := FormatSettings.DecimalSeparator;
    FormatSettings.DecimalSeparator := '.';
    try
      result := StrToFloat(Value);
    except
      result := 0;
    end;
    FormatSettings.DecimalSeparator := Temp;
  end;
end;

// Convert types of two variant values
procedure TParser.CheckTypes(var Value1, Value2: Variant;
  IsForOp: Boolean = FALSE);
begin
  case VarType(Value1) of
    varInt64, varSmallint, varByte, varWord, varLongWord, varUInt64, varInteger:
      case VarType(Value2) of
        varString, varOleStr:
          Value2 := StrToFloatEx(Value2);
        varNull:
          if IsForOp then
            Value2 := 0;
        varDouble:
          begin
          end;
      else
        Value2 := VarAsType(Value2, varInteger);
      end;
    varString, varUString, varOleStr:
      if not VarIsNull(Value2) then
        Value2 := VarAsType(Value2, varString)
      else if IsForOp then
        Value2 := '';
    varSingle, varCurrency, varDouble:
      case VarType(Value2) of
        varUString, varString, varOleStr:
          Value2 := StrToFloatEx(Value2);
        varNull:
          if IsForOp then
            Value2 := 0.0;
      else
        Value2 := VarAsType(Value2, varDouble);
      end;
    varBoolean:
      case VarType(Value2) of
        varByte, varInteger, varDouble:
          Value2 := Value2 <> 0;
        varString, varOleStr, varUString:
          Value2 := StrToFloatEx(Value2) <> 0;
        varBoolean:
        else
          raise EParseException.Create(rsTypesmismatch);
      end;
    varEmpty, varNull:
      if not VarIsNull(Value2) then
      begin
        if IsForOp then
          CheckTypes(Value2, Value1, TRUE)
        else
          raise EParseException.Create(rsTypesmismatch)
      end
      else
        exit;
  else
    raise EParseException.Create(rsTypesmismatch);
  end;
end;

function TParser.EvalFunction(const ParseItem: TParseItem): Variant;
var
  Func: TFunction;
  Args: Array of Variant;
  i, Count: Integer;

begin
  Func := GetFunction(ParseItem.ItemValue);

  if Assigned(Func) then
  begin
    Count := FParseStack.Pop;
    SetLength(Args, Count);
    for i := Count - 1 downto 0 do
      Args[i] := FParseStack.Pop;

    result := Func(Self, Args);
  end
  else
    raise EParseException.CreateFmt(rsFunctionNotFound, [ParseItem.ItemValue]);
end;

// Calculate an equation

function TParser.Evalute<T>: T;
var
  res: Variant;
begin
  res := Evalute;
  if Variants.VarIsNull(res) Or VarIsEmpty(res) then
    result := default (T)
  else
    result := TValue.FromVariant(res).AsType<T>;
end;

function TParser.Evalute: Variant;
var
  i, Params: Integer;
  Value1, Value2: Variant;
  Op: String;
  S: String;
  l: TStringList;

  function CheckForNull: Boolean;
  begin
    result := VarIsNull(Value1) or VarIsNull(Value2)
  end;

begin
  // FStackCount := 0;
  FParseStack.Clear;
  FParseStack.Capacity := 10;
  for i := 0 to FParseItems.Count - 1 do
  begin
    case FParseItems[i].ItemType of
      ptFunction:
        Push(EvalFunction(FParseItems[i]));
      ptNull:
        Push(NULL);
      ptVariable:
        Push(DoGetVar(FParseItems[i].ItemValue));
      ptFloat, ptInteger, ptString, ptBoolean:
        Push(FParseItems[i].ItemValue);
      ptDelim:
        begin
          Op := VarAsType(FParseItems[i].ItemValue, varString);
          if Op[1] in ['+', '-', '*', '/', '%'] then
          begin
            Value2 := Pop;
            Value1 := Pop;
            CheckTypes(Value1, Value2, TRUE);
            case Op[1] of
              '+':
                Push(Value1 + Value2);
              '-':
                Push(Value1 - Value2);
              '*':
                Push(Value1 * Value2);
              '/':
                Push(Value1 / Value2);
              '%':
                Push(Value1 mod Value2);
            end;
            continue;
          end;

          if (Op = '=') or (Op = '<') or (Op = '>') then
          begin
            Value2 := Pop;
            Value1 := Pop;
            if not CheckForNull then
              CheckTypes(Value1, Value2);
            if (VarType(Value1) = varString) or (VarType(Value1) = varOleStr) Or
              (VarType(Value1) = varUString) then
              Value1 := AnsiUpperCase(Value1);
            if (VarType(Value2) = varString) or (VarType(Value2) = varOleStr) Or
              (VarType(Value1) = varUString) then
              Value2 := AnsiUpperCase(Value2);
            case Op[1] of
              '=':
                if VarIsNull(Value1) or VarIsNull(Value2) then
                  Push(VarIsNull(Value1) and VarIsNull(Value2))
                else
                  Push(Value1 = Value2);
              '<':
                if CheckForNull then
                  Push(FALSE)
                else
                  Push(Value1 < Value2);
              '>':
                if CheckForNull then
                  Push(FALSE)
                else
                  Push(Value1 > Value2);
            end;
            continue;
          end;

          if (Op = '>=') or (Op = '<=') or (Op = '<>') then
          begin
            Value2 := Pop;
            Value1 := Pop;
            if CheckForNull then
              if Op = '<>' then
                Push(NOT(VarIsNull(Value1) and VarIsNull(Value2)))
              else
                Push(FALSE)
            else
            begin
              CheckTypes(Value1, Value2);
              if (VarType(Value1) = varString) or (VarType(Value1) = varOleStr)
                Or (VarType(Value1) = varUString) then
                Value1 := AnsiUpperCase(Value1);
              if (VarType(Value2) = varString) or (VarType(Value2) = varOleStr)
                Or (VarType(Value1) = varUString) then
                Value2 := AnsiUpperCase(Value2);
              if Op = '>=' then
                Push(Value1 >= Value2);
              if Op = '<=' then
                Push(Value1 <= Value2);
              if Op = '<>' then
                Push(Value1 <> Value2);
            end;
            continue;
          end;

          if (Op = 'IN') then
          begin
            l := TStringList.Create;
            try
              for Params := 1 to Integer(Pop) do
              begin
                Value1 := Pop;
                if VarIsNull(Value1) then
                  Value1 := 'NULL';
                l.Add(AnsiUpperCase(Value1));
              end;
              Value1 := Pop;
              if VarIsNull(Value1) then
                Value1 := 'NULL';
              Push(l.IndexOf(AnsiUpperCase(VarToStr(Value1))) <> -1);
            finally
              l.Free;
            end;
            continue;
          end;

          if (Op = 'AND') or (Op = 'OR') or (Op = 'XOR') then
          begin
            Value1 := Pop;
            Value2 := Pop;
            if Op = 'AND' then
              Push(Value1 and Value2);
            if Op = 'OR' then
              Push(Value1 or Value2);
            if Op = 'XOR' then
              Push((not Value1 and Value2) or (Value1 and not Value2));
            continue;
          end;

          if Op = '~' then
          begin
            Value1 := Pop;
            Push(-Value1);
            continue;
          end;

          if Op = 'NOT' then
          begin
            Value1 := Pop;
            Value2 := TRUE;
            CheckTypes(Value2, Value1);
            Push(not Value1);
            continue;
          end;

          if (Op = '^') then
          begin
            Value2 := VarAsType(Pop, varDouble);
            Value1 := VarAsType(Pop, varDouble);
            Push(exp(Value2 * ln(Value1)));
            continue;
          end;

          if (Op = 'LIKE') then
          begin
            Value2 := VarAsType(Pop, varString);
            Value1 := Pop;
            if VarIsNull(Value1) then
              Push(FALSE)
            else
            begin
              if varIsArray(Value1) then
              begin
                SetLength(S, VarArrayHighBound(Value1, 1) + 1);
                Move(varArrayLock(Value1)^, PChar(S)^, Length(S));
                VarArrayUnLock(Value1);
                Push(IsMatch(Value2, S));
                {
                  Value2:=UpperCase(Copy(Value2,2,Length(value2)-2));
                  Push(Pos(Value2,UpperCase(S))<>0) }
              end
              else // Value1 := VarAsType(Value1, varString);
                Push(IsMatch(Value2, Value1));
            end;
            continue;
          end;
          raise EParseException.CreateFmt('Operación invalida (%s)', [Op]);
        end;
    end;
  end;
  result := Pop;

  if FParseStack.Count > 0 then
    raise EParseException.Create('Error de evaluación');
end;

// Push value to stack
procedure TParser.Push(const Value: Variant);
begin
  FParseStack.Push(Value)
end;

// Pop value from stack
function TParser.Pop: Variant;
begin
  result := FParseStack.Pop
end;

// Clear all variables and equation
procedure TParser.Clear;
begin
  FParseStack.Clear;
  FParseItems.Clear;
  FParseStack.Capacity := 10;
  FParseItems.Capacity := 10;
  FVariables.Clear;
  FVariables.Capacity := 5;
  FEquation := '';
end;

// Define function
class procedure TParser.AddFunction(const FuncName: String;
  FunctionCode: TFunction);
begin
  FFunctions.Add(LowerCase(FuncName), FunctionCode);
end;

procedure TParser.AddLocalFunction(const Name: String; AFunction: TFunction);
begin
  FLocalFunctions.AddOrSetValue(LowerCase(Name), AFunction);
end;

function TParser.CheckFunc(var Buffer: String): Boolean;
var
  i: Integer;
  Token: String;
begin
  i := 1;
  while (i <= Length(Buffer)) and (Buffer[i] in [' ', #9, #10, #13]) do
    Inc(i);
  if (i <= Length(Buffer)) and (Buffer[i] = '(') then
  begin
    result := TRUE;
    ExtractToken(Buffer, Token);
  end
  else
    result := FALSE;
end;

class destructor TParser.Destroy;
begin
  FFunctions.Free;
end;

function TParser.DoGetVar(VarName: String): Variant;
begin
  if Assigned(OnGetVar) then
  OnGetVar(Self, VarName, Result)
  else
    if Assigned(FGetVar) then
      FGetVar(Self, VarName, Result);

end;

end.
