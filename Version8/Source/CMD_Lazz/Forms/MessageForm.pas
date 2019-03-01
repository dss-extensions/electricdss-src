unit MessageForm;

{
  ----------------------------------------------------------
  Copyright (c) 2008-2015, Electric Power Research Institute, Inc.
  All rights reserved.
  ----------------------------------------------------------
}

{$IFDEF FPC}{$MODE Delphi}{$ENDIF}

interface

uses
    LCLIntf,
    LCLType,
    SysUtils,
    Classes,
    Graphics,
    Controls,
    Forms,
    Dialogs,
    RichMemo;

type
    TMessageForm1 = class(TForm)
        Editor: TRichMemo;
    PRIVATE
    { Private declarations }
    PUBLIC
    { Public declarations }
    end;

var
    MessageForm1, ResultForm, SummaryForm: TMessageForm1;

implementation

{$R *.lfm}


initialization
    MessageForm1 := NIL;
    ResultForm := NIL;
    SummaryForm := NIL;
end.
