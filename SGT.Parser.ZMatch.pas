{******************************************************************
*  (c)copyrights Corona Ltd. Donetsk 1999
*  Project: Zeos Library
*  Module: Function for string patterns process
*  Author: Sergey Seroukhov   E-Mail: voland@cm.dongu.donetsk.ua
*  Date: 17/04/99
*
*  List of changes:
******************************************************************}

unit SGT.Parser.ZMatch;
{
  File: match.pas
  Author: Kevin Boylan
  Modifer: Sergey Seroukhov

  This code is meant to allow wildcard pattern matches.
  It is VERY useful for matching filename wildcard patterns.
  It allows unix grep-like pattern comparisons, for instance:

	?	   	Matches any single characer
	*	   	Matches any contiguous characters
	[abc]  	Matches a or b or c at that position
	[^abc]	Matches anything but a or b or c at that position
	[!abc]	Ditto
	[a-e]  	Matches a through e at that position

	'ma?ch.*'	-Would match match.exe, mavch.dat, march.on, etc
	'this [e-n]s a [!zy]est' - Would match 'this is a test',
                               but would not match 'this as a yest'

  This is a Delphi VCL translation from C code that was downloaded from CIS.
  C code was written by J. Kerceval and released to public domain 02/20/1991.
  This code is ofcourse also public domain. I would appreciate it if you would
  let me know if you find any bugs.  I would also appreciate any notes sent my
  way letting me know if you find it useful.
}

interface
uses SysUtils;

// Check if Text equal to pattern
function IsMatch(Pattern, Text: String): Boolean;

implementation

const
// Match defines
  MATCH_PATTERN	  = 6;
  MATCH_LITERAL	  = 5;
  MATCH_RANGE	  = 4;
  MATCH_ABORT	  = 3;
  MATCH_END	  = 2;
  MATCH_VALID	  = 1;
// pattern defines
  PATTERN_VALID	  = 0;
  PATTERN_ESC	  = -1;
  PATTERN_RANGE	  = -2;
  PATTERN_CLOSE	  = -3;
  PATTERN_EMPTY	  = -4;
// character defines
  MATCH_CHAR_SINGLE	     = '?';
  MATCH_CHAR_KLEENE_CLOSURE  = '%';
  MATCH_CHAR_RANGE_OPEN	     = '[';
  MATCH_CHAR_RANGE	     = '-';
  MATCH_CHAR_RANGE_CLOSE     = ']';
  MATCH_CHAR_CARAT_NEGATE    = '^';
  MATCH_CHAR_EXCLAMATION_NEGATE	= '!';

function Matche(Pattern, Text: String): Integer; forward;
function MatchAfterStar(Pattern, Text: String): Integer; forward;
//function IsPattern(Pattern: String): Boolean; forward;

function IsMatch(Pattern, Text: String): Boolean;
begin
  Result := Matche(Pattern, Text)=1;
end;

function Matche(Pattern, Text: String): Integer;
var
  range_start, range_end, p, t, plen, tlen: Integer;
  invert, member_match, loop: Boolean;
begin
  p := 1;
  t := 1;
  pattern := LowerCase(pattern);
  text := LowerCase(Text);
  plen := Length( pattern );
  tlen := Length( text );
  Result := 0;
  while ( (Result = 0) and (p <= plen) ) do begin
    if (t > tlen) then begin
      if (pattern[p] = MATCH_CHAR_KLEENE_CLOSURE) and (p+1 > plen) then
        Result := MATCH_VALID
      else
        Result := MATCH_ABORT;
      exit;
    end else
      case (pattern[p]) of
        MATCH_CHAR_KLEENE_CLOSURE:
          Result := MatchAfterStar( Copy(pattern,p,plen),Copy(text,t,tlen) );
        MATCH_CHAR_RANGE_OPEN:
          begin
            Inc(p);
            invert := False;
            if (pattern[p] = MATCH_CHAR_EXCLAMATION_NEGATE) or
              (pattern[p] = MATCH_CHAR_CARAT_NEGATE) then begin
              invert := True;
              Inc(p);
            end;
            if (pattern[p] = MATCH_CHAR_RANGE_CLOSE) then begin
              Result := MATCH_PATTERN;
              exit;
            end;
            member_match := False;
            loop := True;
            while ( (loop) and (pattern[p] <> MATCH_CHAR_RANGE_CLOSE) ) do begin
              range_start := p;
              range_end := p;
              Inc(p);
              if (p > plen) then begin
                Result := MATCH_PATTERN;
                exit;
              end;
              if (pattern[p] = MATCH_CHAR_RANGE) then begin
                Inc(p);
                range_end := p;
              if (p > plen) or (pattern[range_end] = MATCH_CHAR_RANGE_CLOSE) then begin
                Result := MATCH_PATTERN;
                exit;
              end;
              Inc(p);
            end;
            if (p > plen) then begin
              Result := MATCH_PATTERN;
              exit;
            end;
            if (range_start < range_end) then begin
              if (text[t] >= pattern[range_start]) and
                (text[t] <= pattern[range_end]) then begin
                member_match := True;
                loop := False;
              end;
            end else begin
              if (text[t] >= pattern[range_end]) and
                (text[t] <= pattern[range_start]) then begin
                member_match := True;
                loop := False;
              end;
            end; { if (range_start < range_end) }
          end; { while (loop) }
          if (invert and member_match) or (not(invert or member_match)) then begin
            Result := MATCH_RANGE;
            exit;
          end;
          if (member_match) then
            while ((p <= plen) and (pattern[p] <> MATCH_CHAR_RANGE_CLOSE)) do
              Inc(p);
            if (p > plen) then begin
              Result := MATCH_PATTERN;
              exit;
            end;
          end; { MATCH_CHAR_RANGE_OPEN: }
        else
          if (pattern[p] <> MATCH_CHAR_SINGLE) then
            if (pattern[p] <> text[t]) then
              Result := MATCH_LITERAL;
      end; { Case pattern[p] }
    Inc(p);
    Inc(t);
  end; { While ( (Result = 0) and (p < plen) ) }
  if (Result = 0) then
    if (t <= tlen) then
      Result := MATCH_END
    else
      Result := MATCH_VALID;
end;

function MatchAfterStar(Pattern, Text: String): Integer;
var p, t, plen, tlen: Integer;
begin
  Result := 0;
  p := 1;
  t := 1;
  plen := Length(pattern);
  tlen := Length(text);
  if (plen=0)or(tlen=0) then begin
    Result := 3;
    Exit;
  end;
  while (( t <= tlen ) and (p < plen)) and (pattern[p] = MATCH_CHAR_SINGLE) or
    (pattern[p] = MATCH_CHAR_KLEENE_CLOSURE) do begin
    if (pattern[p] = MATCH_CHAR_SINGLE) then
      Inc(t);
    Inc(p);
  end;
  if (t > tlen) then begin
    Result := MATCH_ABORT;
    exit;
  end;
  if (p > plen) then begin
    Result := MATCH_VALID;
    exit;
  end;
  repeat
    if (pattern[p] = text[t]) or (pattern[p] = MATCH_CHAR_RANGE_OPEN) then begin
      pattern := Copy(pattern,p,plen);
      text := Copy(text,t,tlen);
      plen := Length(pattern);
      tlen := Length(text);
      p := 1;
      t := 1;
      Result := matche( pattern , text );
    end;
    if (t > tlen) then begin
      Result := MATCH_ABORT;
      exit;
    end;
    Inc(t);
  Until ( Result <> 0 );
end;
{
function IsPattern(Pattern: String): Boolean;
var Index:Integer;
begin
  Result:=false;
  for Index:=1 to length(pattern) do
    if Pos(pattern[Index],'[]?*')>0 then begin
      Result:=true;
      exit;
    end;
end;
}
end.


