unit TOPExport;

{$IFDEF FPC}{$MODE Delphi}{$ENDIF}

{
  ----------------------------------------------------------
  Copyright (c) 2008-2015, Electric Power Research Institute, Inc.
  All rights reserved.
  ----------------------------------------------------------
}

{Supports Creation of a STO file for interfacing to TOP and the
 invoking of TOP.}

interface

uses
    Classes,
    ArrayDef;

type
    time_t = Longint;

    ToutfileHdr = packed record
        Size: Word;
        Signature: array[0..15] of ANSICHAR;
        VersionMajor,
        VersionMinor: Word;
        FBase,
        VBase: Double;
        tStart,
        tFinish: time_t;
        StartTime,
        StopT,
        DeltaT: Double;
        Nsteps: Longword;
        NVoltages,
        NCurrents,
        VoltNameSize,
        CurrNameSize: Word;
        IdxVoltNames,
        IdxCurrentNames,
        IdxBaseData,
        IdxData: Longint;
        Title1,
        Title2,
        Title3,
        Title4,
        Title5: array[0..79] of ANSICHAR;  // Fixed length 80-byte string  space
    end;

    TOutFile32 = class(Tobject)
        Header: ToutfileHdr;
        Fname: String;  {Default is RLCWOUT.STO'}
        Fout: file;

    PRIVATE

    PUBLIC
          {constructor Create(Owner: TObject);}
        procedure Open;
        procedure Close;
        procedure WriteHeader(const t_start, t_stop, h: Double; const NV, NI, NameSize: Integer; const Title: Ansistring);
        procedure WriteNames(var Vnames, Cnames: TStringList);
        procedure WriteData(const t: Double; const V, Curr: pDoubleArray);
        procedure OpenR;  {Open for Read Only}
        procedure ReadHeader; {Opposite of WriteHeader}
        procedure GetVoltage(T, V: pDoubleArray; Idx, MaxPts: Integer); {Read a single node voltage from disk}
        procedure SendToTop;

        property FileName: String READ Fname WRITE Fname;

    end;

var
    TOPTransferFile: TOutFile32;
    TOP_Object: Variant;  // For Top Automation

implementation

uses
    SysUtils, {Dialogs,} DSSGlobals,
    CmdForms; // TEMc

var
    TOP_Inited: Boolean;

procedure StartTop;

begin
//  TOP_Object := CreateOleObject('TOP2000.MAIN');
//  TOP_Inited := TRUE;
end;

procedure TOutFile32.SendToTop;
begin
    DSSInfoMessageDlg('TOP Export (COM Interface) is not supported on Linux');
  (*
  TRY
     If NOT TOP_Inited Then StartTop;


   TRY
     TOP_Object.OpenFile(TOPTransferFile.FName);

   Except {Top has become disconnected}
     // Oops.  Connection to TOP is not valid;
       Try
          StartTop;
          TOP_Object.OpenFile(TOPTransferFile.FName);
       Except
        ShowMessage('Export to TOP failed.  Connection lost?');
       End;
   End;
  Except

        On E:Exception Do ShowMessage('Error Connecting to TOP: '+E.Message);
  End;
*)
end;


procedure TOutFile32.Open;
begin
    AssignFile(Fout, Fname);
    ReWrite(Fout, 1);  {Open untyped file with a recordsize of 1 byte}
end;


procedure TOutFile32.Close;
begin

    CloseFile(Fout);  {Close the output file}

end;

procedure TOutFile32.WriteHeader(const t_start, t_stop, h: Double; const NV, NI, NameSize: Integer; const Title: Ansistring);

var
    NumWrite: Integer;


begin

    with Header do
    begin

        Size := SizeOf(TOutFileHdr);
        Signature := 'SuperTran V1.00'#0;
        VersionMajor := 1;
        VersionMinor := 1;
        FBase := DefaultBaseFreq;
        VBase := 1.0;
        tStart := 0;
        TFinish := 0;
        StartTime := t_start;
        StopT := t_stop;
        DeltaT := h;
        Nsteps := Trunc(t_stop / h) + 1;
        NVoltages := NV;
        NCurrents := NI;


        VoltNameSize := NameSize;
        CurrNameSize := NameSize;
        IdxVoltNames := Header.Size;
        IdxCurrentNames := IdxVoltNames + NVoltages * VoltNameSize;
        IDXData := IDXCurrentNames + NCurrents * CurrNameSize;
        IdxBaseData := 0;

        sysutils.StrCopy(Title1, pAnsichar(Title));
        Title2[0] := #0;
        Title3[0] := #0;
        Title4[0] := #0;
        Title5[0] := #0;


    end;

     { Zap the header to disk }
    BlockWrite(Fout, Header, SizeOf(Header), NumWrite);

end;

procedure TOutFile32.WriteNames(var Vnames, Cnames: TStringList);

var
    NumWrite: Integer;
    i: Integer;
    Buf: array[0..120] of AnsiChar;  //120 char buffer to hold names  + null terminator

begin

    if Header.NVoltages > 0 then
        for i := 0 to Vnames.Count - 1 do
        begin
            Sysutils.StrCopy(Buf, pAnsichar(Ansistring(Vnames.Strings[i])));    // Assign string to a buffer
            BlockWrite(Fout, Buf, Header.VoltNameSize, NumWrite);    // Strings is default property of TStrings
        end;

    if Header.NCurrents > 0 then
        for i := 0 to Cnames.Count - 1 do
        begin
            Sysutils.StrCopy(Buf, pAnsichar(Ansistring(Cnames.Strings[i])));    // Assign string to a buffer
            BlockWrite(Fout, Buf, Header.CurrNameSize, NumWrite);
        end;

end;

procedure TOutFile32.WriteData(const t: Double; const V, Curr: pDoubleArray);

var
    NumWrite: Integer;

begin

    BlockWrite(Fout, t, SizeOf(Double), NumWrite);
    if Header.NVoltages > 0 then
        BlockWrite(Fout, V^[1], SizeOf(Double) * Header.NVoltages, NumWrite);
    if Header.NCurrents > 0 then
        BlockWrite(Fout, Curr^[1], SizeOf(Double) * Header.NCurrents, NumWrite);

end;

procedure TOutFile32.OpenR;  {Open for Read Only}

begin
    AssignFile(Fout, Fname);
    Reset(Fout, 1);
end;

procedure TOutFile32.ReadHeader; {Opposite of WriteHeader}

var
    NumRead: Integer;

begin
    BlockRead(Fout, Header, SizeOf(Header), NumRead);
end;

procedure TOutFile32.GetVoltage(T, V: pDoubleArray; Idx, MaxPts: Integer); {Read a voltage from disk}

{Gets a specified voltage from an STO file for plotting.  Idx specifies the index into the voltage array}
var
    Vtemp, Ctemp: pDoubleArray;
    i: Integer;
    NumRead: Integer;

begin
    {Assumes V is Allocated to hold result}

    i := 0;
    Seek(Fout, Header.IdxData);

    GetMem(Vtemp, Sizeof(Double) * Header.NVoltages);
    GetMem(Ctemp, Sizeof(Double) * Header.NCurrents);

    while (not Eof(Fout)) and (i < MaxPts) do
    begin
        Inc(i);
        BlockRead(Fout, T^[i], SizeOf(Double), NumRead);
        BlockRead(Fout, Vtemp^[1], SizeOf(Double) * Header.Nvoltages, NumRead);
        BlockRead(Fout, Ctemp^[1], SizeOf(Double) * Header.NCurrents, NumRead);
        V^[i] := Vtemp^[Idx];
    end;
    FreeMem(Vtemp, Sizeof(Double) * Header.NVoltages);
    FreeMem(Ctemp, Sizeof(Double) * Header.NCurrents);

end;

initialization

    TOP_Inited := FALSE;
    TOPTransferFile := TOutFile32.Create;
    TOPTransferFile.Fname := 'DSSTransfer.STO';

//   CoInitialize(Nil);
end.
