// BASED ON OLD Parser Zeos library.
{******************************************************************
*  (c)copyrights Corona Ltd. Donetsk 1999
*  Project: Zeos Library
*  Module: Formula parser component
*  Author: Sergey Seroukhov   E-Mail: voland@cm.dongu.donetsk.ua
*  Date: 26/03/99
*
*  List of changes:
*  27/03/99 - Class convert to component, add vars
*  16/04/99 - Add some functions, operators LIKE, XOR
*******************************************************************}


unit SGT.Parser.ZParser;

interface
uses SysUtils, Classes, Sgt.Parser.ZToken, SGt.Parser.ZMatch, Variants,
    Rtti, SGT.ListIntf, Generics.Collections;

type

TParseItemType=(ptFunction, ptVariable, ptDelim, ptString, ptInteger, ptFloat,
                ptBoolean,  ptNull);

TParseItem = record
  ItemValue: Variant;
  ItemType: TParseItemType;
end;

TParser = class;

EParseException = class(Exception);

TOnGetVar = procedure (sender: TParser; VarName: string; var Value: Variant) of object;

TFunction = function (Sender: TParser; const Args: Array of Variant): Variant;

TParser = class
private
  class var FFunctions: TDictionary<String, TFunction>;
  class function GetGlobalFunction(const FuncName: String): TFunction;
  class constructor Create;
  class destructor Destroy;
public
  class procedure AddFunction(const FuncName: String; FunctionCode: TFunction);
private
  FParseItems: TList<TParseItem>;
  FLocalFunctions: TDictionary<String, TFunction>;
  FErrCheck: Integer;
  FEquation: String;
  FParseStack: TStack<Variant>;
  FVariables: TList<String>;
  FOnGetVar: TOnGetVar;
  function ExtractTokenEx(var Buffer, Token: String): TParseItemType;
  function OpLevel(Operat: String): Integer;
  function Parse(Level: Integer; var Buffer: String): Integer;
  procedure SetEquation(Value: String);
  procedure CheckTypes(var Value1,Value2: Variant; IsForOp:Boolean=FALSE);
  function ConvType(Value: Variant): Variant;
  function CheckFunc(var Buffer: String): Boolean;
  function EvalFunction(const ParseItem:TParseItem): Variant;
  procedure Push(const Value: Variant);
  function Pop: Variant;
  function GetVariables: TArray<String>;
  function GetFunction(FuncName: String): TFunction;
protected
  function DoGetVar(VarName: String): Variant; Virtual;
public
  constructor Create;
  destructor Destroy; override;
  function Evalute: Variant;
  procedure AddLocalFunction(const Name: String; AFunction: TFunction);
  procedure Clear;
  property Equation: String read FEquation write SetEquation;
  property OnGetVar: TOnGetVar read FOnGetVar write FOnGetVar;
  property Variables: TArray<String> read GetVariables;
end;

procedure CheckParamCount(const Args: Array of Variant; ArgCount: byte;  const FuncName: String); overload;
procedure CheckParamCount(const Args: Array of Variant; ArgCount: Array of byte;  const FuncName: String); overload;

implementation

ResourceString
   rsInvalidParamsNo ='Invalid param count in function %s';
   rsSyntaxError     ='Syntax error.';
   rsIndexOutOfRange ='Index out of range.';
   rsTypesmismatch   ='Type mismatch.';
   rsFunctionNotFound='Function %s don''t. exists.' ;
   rsDuplicateFunction='Duplicate funcion name (%s).';

function VarIsNull(const V: variant): boolean;
begin

  result:= Variants.VarIsNull(V) Or VarIsEmpty(V);
end;

{ TParserFuncion }

function TParser.GetFunction( FuncName:string): TFunction;
begin
  FuncName:= LowerCase(FuncName);
  if not FLocalFunctions.TryGetValue(FuncName, result) then
    FFunctions.TryGetValue(FuncName, result);
end;

class function TParser.GetGlobalFunction(const FuncName: String): TFunction;
begin
  FFunctions.TryGetValue(LowerCase(FuncName), result);
end;

function TParser.GetVariables: TArray<String>;
begin
  result:= FVariables.ToArray;
end;

{************** User functions implementation **************}

procedure CheckParamCount(const Args: Array of Variant; ArgCount: Array of Byte;  const FuncName: String); overload;
var
  LOk: Boolean;
  Lb, LArg: Byte;
begin
  Larg:= Length(Args);
  for lb in ArgCount do if lb= Larg then exit;

  raise EParseException.CreateFmt(rsInvalidParamsNo,[FuncName]);
end;


procedure CheckParamCount(const Args: Array of Variant; ArgCount: Byte;  const FuncName: String); overload;
begin
  CheckParamCount(Args,[ArgCount], FuncName);
end;


// Get current date and time
function Now_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 0, 'Now');
  Result := Now
end;

function FormatNum_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 2, 'FormatNum');
  if Result<>2 then
    EParseException.CreateFmt(rsInvalidParamsNo,['FormatNum']);
  Result := FormatFloat(Args[0], Args[1]);
end;

function Fcur_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 1, 'FCur');
  Result := FormatFloat(',##0.00 €',Args[0]);
end;

function IsNull_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 1, 'IsNull');
  Result := VarIsNull(Args[0]) Or VarIsEmpty(Args[0]);
end;

function FNum_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 1, 'FNum');
  Result := Format('%n',[Args[0]]);
end;

function StrPos_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 2, 'StrPos');
  Result := Pos(Args[0],Args[1]);
end;

function Replace_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 3, 'Replace');
  Result := StringReplace(Args[0],Args[1],Args[2],[rfIgnoreCase])
end;

function ReplaceAll_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 3, 'ReplaceAll');
  Result := StringReplace(Args[0],Args[1],Args[2],[rfIgnoreCase, rfReplaceAll])
end;


function StrCopy_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 3, 'StrCopy');
  Result := Copy(Args[0], Integer(Args[1]), Integer(Args[2]));
end;

function StrLen_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 1, 'StrLen');
  Result := Length(Args[0]);
end;

function IIf_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 3, 'Iif');
  if Boolean(VarAsType(Args[0], varBoolean)) then result:= Args[1]
  else result:= Args[2];
end;



function Upper_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 1, 'Upper');
  result:=AnsiUpperCase(Args[0]);
end;

function Lower_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 1, 'Lower');
  result:=AnsiLowerCase(Args[0]);
end;



function Date_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 1, 'Date');
  Result := Date //FormatDateTime('yyyy-mm-dd hh:nn:ss', Now());
end;

function FormatDate_Execute(Sender: TParser; const Args: Array of Variant): Variant;
begin
  CheckParamCount(Args, 2, 'FomatDate');
  result:= FormatDateTime(Args[0], Args[1])
end;

function GetItem_Execute(Sender: TParser; const Args: Array of Variant): Variant;
var
  LList: IListInterface;
  LValue: TValue;
begin
  CheckParamCount(Args, 2, 'GetListItem');
  if VarType(Args[0])= varInteger then
    LValue:= TObject(Integer(Args[0]))
  else
    LValue:= TValue.FromVariant(Args[0]);
  LList:= TListInterface.Get(Lvalue);
  if LList=nil then
    raise Exception.Create('Error GetItem function. Parameter not is a list');

  try
    LValue:= LList.GetItem(Args[1]);
    result:= LValue.AsVariant;
  finally
    LList.DetachList;
  end;
end;

function GetListCount_Execute(Sender: TParser; const Args: Array of Variant): Variant;
var
  LList: IListInterface;
  LValue: TValue;
begin
  CheckParamCount(Args, 1, 'GetListSize');
  if VarType(Args[0])= varInteger then
    LValue:= TObject(Integer(Args[0]))
  else
    LValue:= TValue.FromVariant(Args[0]);
  LList:= TListInterface.Get(Lvalue);
  if LList=nil then
    raise Exception.Create('Error GetListSize function. Parameter not is a list');
  try
    result:= LList.Size;
  finally
    LList.DetachList;
  end;
end;

{******************* TParser implementation ****************}

constructor TParser.Create;
begin
  inherited Create;
  FParseItems:= TList<TParseItem>.Create;
  FParseItems.Capacity:=10;

  FParseStack:= TStack<Variant>.Create;
  FparseStack.Capacity:= 10;

  FVariables:= TList<String>.Create;
  FErrCheck := 0;
  FLocalFunctions:= TDictionary<String, TFunction>.Create;
end;

// Class destructor
destructor TParser.Destroy;
begin
  FParseItems.Free;
  FParseStack.Free;
  FLocalFunctions.Free;
  FVariables.Free;
  inherited destroy;
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
  until (Token<>#10)and(Token<>#13);

  if Token='[' then begin
    TokenType := ttAlpha;
    P := Pos(']',Buffer); Token := '';
    if P>0 then begin
      Token := Copy(Buffer,1,P-1);
      Buffer := Copy(Buffer,P+1,Length(Buffer)-P);
    end;
  end;

  if Token='#' then begin
    p:=Pos('#',Buffer); Token:='';
    if p>0 then begin
      Token:=Copy(Buffer,1,p-1);
      Token:=FloatToStr(StrToDateTime(Token));
      if FormatSettings.DecimalSeparator<>'.' then
        Token:=StringReplace(Token,FormatSettings.DecimalSeparator,'.',[]);
      if Pos('.',Token)=0 then Token:=Token+'.0';
      Buffer := Copy(Buffer,P+1,Length(Buffer)-P);
      Result:=ptFloat;
    end;
    exit;
  end;

  if (Buffer<>'')and(Token='>')and(Buffer[1]='=') then begin
    ExtractToken(Buffer, Temp);
    Token := Token + Temp;
  end;
  if (Buffer<>'')and(Token='<')and((Buffer[1]='=')or(Buffer[1]='>')) then begin
    ExtractToken(Buffer, Temp);
    Token := Token + Temp;
  end;

  Temp := UpperCase(Token);
  if (Temp='AND') or (Temp='NOT') or (Temp='OR') or (Temp='XOR') or (Temp='IN') or
     (Temp='LIKE') then begin
    Token := Temp;
    Result := ptDelim;
    exit;
  end;
  if (Temp='TRUE') or (Temp='FALSE') then begin
    Token := Temp;
    Result := ptBoolean;
    exit;
  end;
  if Temp='NULL' then begin
    Token := Temp;
    Result := ptNull;
    exit;
  end;
  if Temp='IS' then begin
    Token:='=';
    Result:=ptDelim;
    exit;
  end;
  case TokenType of
    ttString: if Pos(FormatSettings.DateSeparator,Token) in [2,3] then begin
                try
                  Temp:= FloatToStr(StrToDateTime(Copy(Token,2,Length(Token)-2)));
                  if FormatSettings.DecimalSeparator<>'.' then
                    Token:= StringReplace(Temp,FormatSettings.DecimalSeparator,'.',[])
                  else Token:= Temp;
                  Result:= ptFloat;
                except
                  Result:= ptString;
                end;
              end else Result:= ptString;
    ttAlpha: Result := ptVariable;
    ttDelim: Result := ptDelim;
    ttDigit: begin
        if (Buffer<>'') and (Buffer[1]='.') then begin
          ExtractToken(Buffer, Temp);
          Token := Token + '.';
          if (Buffer<>'')and(Buffer[1]>='0')and(Buffer[1]<='9') then begin
            ExtractToken(Buffer,Temp);
            Token := Token + Temp;
          end;
          Result := ptFloat;
        end else Result := ptInteger;
      end;
  end;
end;

// Get priority level of operation
function TParser.OpLevel(Operat: String): Integer;
var Temp: String;
begin
  Result := 7;
  Temp := UpperCase(Operat);
  if (Temp='AND') or (Temp='OR') or (Temp='XOR') then Result := 1;
  if (Temp='NOT') then Result := 2;
  if (Temp='<')or(Temp='>')or(Temp='=')or(Temp='>=')or(Temp='<=')
    or(Temp='<>') then Result := 3;
  if (Temp[1]='+') or (Temp[1]='-') or (Temp='LIKE') then Result := 4;
  if (Temp[1]='/') or (Temp[1]='*') or (Temp[1]='%') then Result := 5;
  if (Temp[1]='^') then Result := 6;
end;

// Internal convert equation from infix form to postfix
function TParser.Parse(Level: Integer; var Buffer: String): Integer;
var
  ParseType: TParseItemType;
  Token, FuncName: String;
  Temp: Char;
  NewLevel: Integer;
  IValue: Variant;

  procedure AddParseItem(Value:Variant; _Type: TParseItemType);
  var Item: TParseItem;
  begin
    Item.ItemValue:= Value;
    Item.ItemType:= _Type;
    FParseItems.Add(Item);
  end;

  procedure ExtractParams;
  var Params, SaveCount: Integer;
  begin
    SaveCount:= FParseItems.Count;
    Params := 0;
    repeat
      FErrCheck := 0;
      Parse(0,Buffer);
      ExtractTokenEx(Buffer, Token);
      if Token ='' then
        raise EParseException.Create(rsSyntaxerror);
      case Token[1] of
        ',': begin
            Inc(Params);
            continue;
          end;
        ')': begin
            if SaveCount<{FParseCount}FParseItems.Count then Inc(Params);
            {FParseItems[FParseCount].ItemValue := ConvType(Params);
            FParseItems[FParseCount].ItemType := ptInteger;
            Inc(FParseCount);}
            AddParseItem(ConvType(Params),ptInteger);
            break;
          end;
        else
          raise EParseException.Create(rsSyntaxerror);
      end;
    until Buffer='';
  end;

begin
  Result := 0;
  while Buffer<>'' do begin
    ParseType := ExtractTokenEx(Buffer, Token);
    if Token='' then exit;
    if (Token=')') or (Token=',') then begin
      PutbackToken(Buffer, Token);
      exit;
    end;
    if Token='(' then begin
      FErrCheck := 0;
      Parse(0,Buffer);
      ExtractTokenEx(Buffer, Token);
      if Token<>')' then
        raise EParseException.Create(rsSyntaxError);
      FErrCheck := 1;
      continue;
    end;

    if ParseType=ptDelim then begin
      NewLevel := OpLevel(Token);
      if (FErrCheck=2)and(Token<>'NOT') then
        raise EParseException.Create(rsSyntaxError);
      if FErrCheck=0 then
        if (Token<>'NOT')and(Token<>'+')and(Token<>'-') then
          raise EParseException.Create(rsSyntaxError)
        else if Token<>'NOT' then NewLevel := 6;

      if (Token<>'NOT')and(NewLevel<=Level) then begin
        PutbackToken(Buffer, Token);
        Result := NewLevel;
        exit;
      end else if (Token='NOT')and(NewLevel<Level) then begin
        PutbackToken(Buffer, Token);
        Result := NewLevel;
        exit;
      end;

      if (FErrCheck=0) and (Token='+') then continue;
      if (FErrCheck=0) and (Token='-') then Token := '~';
      FErrCheck := 2;
      if Token<>'IN' then begin
        while (Buffer<>'')and(Buffer[1]<>')')and(Parse(NewLevel, Buffer)>NewLevel) do;
      end else
        if not CheckFunc(buffer) then raise EParseException.Create(rsSyntaxError)
        else begin
          FuncName:=Token;
          ExtractParams;
          Token:=FuncName;
        end;
      AddParseItem(Token,ptDelim);
      Result := NewLevel;
      continue;
    end;

    if FErrCheck=1 then
      raise EParseException.Create(rsSyntaxError);
    FErrCheck := 1;
    case ParseType of
      ptVariable: begin
          //FParseItems[FParseCount].ItemValue := Token;
          IValue:=Token;
          if CheckFunc(Buffer) then ParseType := ptFunction
        end;
      ptNull:   IValue:=NULL;
      ptInteger: //FParseItems[FParseCount].ItemValue := StrToInt(Token);
                IValue:=StrToInt(Token);
      ptFloat: begin
          Temp := FormatSettings.DecimalSeparator;
          FormatSettings.DecimalSeparator := '.';
          IValue:=StrToFloat(Token);
          FormatSettings.DecimalSeparator := Temp;
        end;
      ptString: begin
          DeleteQuotes(Token);
          IValue:=Token;
        end;
      ptBoolean: IValue:=Token='TRUE';
    end;
    if ParseType = ptFunction then begin
      FuncName := UpperCase(Token);
      ExtractParams;
      IValue:=FuncName;
    end
    else
    if (ParseType = ptVariable) And not FVariables.Contains(LowerCase(token)) then
      FVariables.Add(LowerCase(token));

    AddParseItem(ConvType(IValue),ParseType);
  end;
end;

// Split equation to stack
// Value - equation buffer
procedure TParser.SetEquation(Value: String);
begin
  //FParseCount := 0;
  Clear;
  FErrCheck := 0;
  FEquation := Value;
  while Value<>'' do Parse(0, Value);

  FVariables.TrimExcess;
end;

// Convert types of two variant values
function TParser.ConvType(Value: Variant): Variant;
begin
  case VarType(Value) of
    varByte, varSmallint, varInteger:
      Result := VarAsType(Value, varInteger);
    varSingle, varDouble, varCurrency:
      Result := VarAsType(Value, varDouble);
    varOleStr, varString, varVariant, varUString:
      Result := VarAsType(Value, varString);
    varDate: Result:=Double(varToDateTime(Value));
    varBoolean:
      Result := Value;
    varNull: Result:=NULL;
  else
      raise EParseException.Create(rsTypesmismatch);
  end;
end;

class constructor TParser.Create;
begin
  FFunctions:= TDictionary<String, TFunction>.Create;
  AddFunction('now',Now_Execute);
  AddFunction('iif',Iif_Execute);
  AddFunction('strcopy',StrCopy_Execute);
  AddFunction('upper',Upper_Execute);
  AddFunction('lower',lower_Execute);
  AddFunction('date',Date_Execute);
  AddFunction('formatdate',FormatDate_Execute);
  AddFunction('replace',Replace_Execute);
  AddFunction('RepalceAll',ReplaceAll_Execute);
  AddFunction('strpos', StrPos_Execute);
  AddFunction('FormatNum',FormatNum_Execute);
  AddFunction('fcur',FCur_Execute);
  AddFunction('fnum',FNum_Execute);
  AddFunction('isnull',IsNull_Execute);
  AddFunction('strlen',StrLen_Execute);
  AddFunction('getlistitem',GetItem_Execute);
  AddFunction('getlistsize',GetListCount_Execute);
end;

function StrToFloatEx(Value: String): Double;
var Temp: Char;
begin
  Result := 0;
  if Value<>'' then begin
    Temp := FormatSettings.DecimalSeparator;
    FormatSettings.DecimalSeparator := '.';
    try Result := StrToFloat(Value);
    except Result := 0; end;
    FormatSettings.DecimalSeparator := Temp;
  end;
end;

// Convert types of two variant values
procedure TParser.CheckTypes(var Value1,Value2: Variant; isForOp:Boolean=FALSE);
begin
  case VarType(Value1) of
    varInt64,
    varSmallInt,
    varByte,
    varWord,
    varLongWord,
    varUInt64,
    varInteger:
      case varType(Value2) of
         varString,varOleStr: Value2 := StrToFloatEx(Value2);
         varNull: if isForOp then Value2:=0;
         varDouble: begin end;
        else Value2 := VarAsType(Value2, varInteger);
      end;
    varString,
    varUString,
    varOleStr: if not varIsNull(Value2) then Value2 := VarAsType(Value2, varString)
               else if IsForOp then  Value2:='';
    varSingle,
    varCurrency,
    varDouble:
      case varType(Value2) of
         varUString, varString,varOleStr: Value2 := StrToFloatEx(Value2);
         varNull: if isForOp then Value2:=0.0;
        else Value2 := VarAsType(Value2, varDouble);
      end;
    varBoolean:
      case VarType(Value2) of
        varByte,varInteger, varDouble:
          Value2 := Value2<>0;
        varString,varOleStr, varUString:
          Value2 := StrToFloatEx(Value2)<>0;
        varBoolean:
        else
          raise EParseException.Create(rsTypesmismatch);
      end;
    varEmpty,
    varNull:if not VarIsNull(Value2) then begin
             if IsForOp then CheckTypes(Value2,Value1,TRUE)
             else raise EParseException.Create(rsTypesmismatch)
            end  else Exit;
    else
      raise EParseException.Create(rsTypesmismatch);
  end;
end;

function TParser.EvalFunction(const ParseItem: TParseItem):Variant;
var Func: TFunction;
    Args: Array of Variant;
    i,Count: Integer;

begin
  Func:= GetFunction(ParseItem.ItemValue);

  if Assigned(Func) then
  begin
    Count:= FParseStack.Pop;
    SetLength(Args, Count);
    for i:= count -1 downto 0 do
      Args[i]:= FparseStack.Pop;

    Result:= Func(Self, Args);
  end
  else
    raise EParseException.CreateFmt(rsFunctionNotFound,[ParseItem.ItemValue]);
end;

// Calculate an equation
function TParser.Evalute: Variant;
var
  I,params: Integer;
  Value1, Value2: Variant;
  Op: String;
  S:String;
  l:TStringList;

  function CheckForNull:Boolean;
  begin
    result:=varIsNull(Value1) or varIsNull(Value2)
  end;

begin
  //FStackCount := 0;
  FParseStack.Clear;
  FParseStack.Capacity:=10;
  for I := 0 to FParseItems.Count-1 do begin
    case FParseItems[I].ItemType of
      ptFunction: Push(EvalFunction(FParseItems[i]));
      ptNull: Push(NULL);
      ptVariable: Push(DoGetVar(FParseItems[i].ItemValue));
      ptFloat, ptInteger, ptString, ptBoolean: Push(FParseItems[I].ItemValue);
      ptDelim: begin
          Op := VarAsType(FParseItems[I].ItemValue, varString);
          if Op[1] in ['+','-','*','/','%'] then begin
            Value2 := Pop; Value1 := Pop;
            CheckTypes(Value1, Value2,TRUE);
            case Op[1] of
              '+': Push(Value1 + Value2);
              '-': Push(Value1 - Value2);
              '*': Push(Value1 * Value2);
              '/': Push(Value1 / Value2);
              '%': Push(Value1 mod Value2);
            end;
            continue;
          end;

          if (Op='=')or(Op='<')or(Op='>') then begin
            Value2 := Pop; Value1 := Pop;
            if not CheckForNull then CheckTypes(Value1, Value2);
            if (varType(value1)=varString) or (varType(value1)=varOleStr) Or (varType(value1)=varUString)then
                value1:=AnsiUpperCase(value1);
            if (varType(value2)=varString) or (varType(value2)=varOleStr) Or (varType(value1)=varUString) then
                value2:=AnsiUpperCase(value2);
            case Op[1] of
              '=': if varIsNull(value1) or varIsNull(value2) then
                      Push(varIsNull(value1) and varIsNull(value2))
                   else Push(Value1 = Value2);
              '<': if CheckForNull then Push(FALSE) else Push(Value1 < Value2);
              '>': if CheckForNull then Push(FALSE) else Push(Value1 > Value2);
            end;
            continue;
          end;

          if (Op='>=')or(Op='<=')or(Op='<>') then begin
            Value2 := Pop; Value1 := Pop;
            if CheckForNull then
              if Op='<>' then Push(NOT (VarIsNull(value1) and VarIsNull(Value2)))
              else Push(FALSE)
            else begin
              CheckTypes(Value1, Value2);
              if (varType(value1)=varString) or (varType(value1)=varOleStr) Or (varType(value1)=varUString) then
                value1:=AnsiUpperCase(value1);
              if (varType(value2)=varString) or (varType(value2)=varOleStr) Or (varType(value1)=varUString) then
                value2:=AnsiUpperCase(value2);
              if Op='>=' then Push(Value1 >= Value2);
              if Op='<=' then Push(Value1 <= Value2);
              if Op='<>' then Push(Value1 <> Value2);
            end;
            continue;
          end;

          if (Op='IN') then begin
             l:=TStringlist.create;
             try
               for Params:=1 to Integer(Pop) do begin
                 Value1:=Pop; if varIsNull(Value1) then Value1:='NULL';
                 l.Add(AnsiUpperCase(Value1));
               end;
               Value1:=Pop; if varIsNull(Value1) then Value1:='NULL';
               Push(l.IndexOf(AnsiUpperCase(VarToStr(Value1)))<>-1);
             finally
               l.free;
             end;
             continue;
          end;

          if (Op='AND')or(Op='OR')or(Op='XOR') then begin
            Value1 := Pop; Value2 := Pop;
            if Op='AND' then Push(Value1 and Value2);
            if Op='OR' then Push(Value1 or Value2);
            if Op='XOR' then
              Push((not Value1 and Value2) or (Value1 and not Value2));
            continue;
          end;

          if Op='~' then begin
            Value1 := Pop;
            Push(-Value1);
            continue;
          end;

          if Op='NOT' then begin
            Value1 := Pop;
            Value2 := True;
            CheckTypes(Value2, Value1);
            Push(not Value1);
            continue;
          end;

          if (Op='^') then begin
            Value2 := VarAsType(Pop,varDouble);
            Value1 := VarAsType(Pop,varDouble);
            Push(exp(Value2*ln(Value1)));
            continue;
          end;

          if (Op='LIKE') then begin
            Value2 := VarAsType(Pop, varString);
            Value1:=Pop;
            if VarIsNull(Value1) then Push(FALSE)
            else begin
              if varIsArray(Value1) then begin
                SetLength(S,VarArrayHighBound(Value1,1)+1);
                Move(varArrayLock(value1)^,PChar(S)^,Length(S));
                VarArrayUnLock(Value1);
                Push(IsMatch(Value2,S));
                {
                Value2:=UpperCase(Copy(Value2,2,Length(value2)-2));
                Push(Pos(Value2,UpperCase(S))<>0)}
              end else //Value1 := VarAsType(Value1, varString);
                Push(IsMatch(Value2,Value1));
            end;
            continue;
          end;
          raise  EParseException.CreateFmt('Operación invalida (%s)',[Op]);
        end;
    end;
  end;
  Result := Pop;

  if FParseStack.Count>0 then
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
 result:= FParseStack.Pop
end;

// Clear all variables and equation
procedure TParser.Clear;
begin
  FParseStack.Clear;
  FParseItems.Clear;
  FparseStack.Capacity:=10;
  FparseItems.Capacity:=10;
  FVariables.Clear;
  FVariables.Capacity:=5;
  FEquation := '';
end;

// Define function
class procedure TParser.AddFunction(const FuncName: String;
  FunctionCode: TFunction);
begin
  FFunctions.Add(Lowercase(FuncName), FunctionCode);
end;

procedure TParser.AddLocalFunction(const Name: String; AFunction: TFunction);
begin
  FLocalFunctions.AddOrSetValue(LowerCase(Name), AFunction);
end;

function TParser.CheckFunc(var Buffer: String): Boolean;
var
  I: Integer;
  Token: String;
begin
  I := 1;
  while (I<=Length(Buffer)) and (Buffer[I] in [' ',#9,#10,#13]) do
    Inc(I);
  if (i<=length(Buffer)) and (Buffer[I]='(') then begin
    Result := true;
    ExtractToken(Buffer, Token);
  end else Result := false;
end;

class destructor TParser.Destroy;
begin
  FFunctions.Free;
end;

function TParser.DoGetVar(VarName: String): Variant;
begin
  if Assigned(OnGetVar) then
    OnGetVar(Self,VarName,Result);
end;


end.


