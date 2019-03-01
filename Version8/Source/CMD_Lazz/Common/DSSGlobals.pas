unit DSSGlobals;

{$IFDEF FPC}{$MODE Delphi}{$ENDIF}

{
  ----------------------------------------------------------
  Copyright (c) 2008-2015, Electric Power Research Institute, Inc.
  All rights reserved.
  ----------------------------------------------------------
}


{ Change Log
 8-14-99  SolutionAbort Added

 10-12-99 AutoAdd constants added;
 4-17-00  Added IsShuntCapacitor routine, Updated constants
 10-08-02 Moved Control Panel Instantiation and show to here
 11-6-02  Removed load user DLL because it was causing a conflict
}

{$WARN UNIT_PLATFORM OFF}

interface

uses
    Classes,
    DSSClassDefs,
    DSSObject,
    DSSClass,
    ParserDel,
    Hashlist,
    PointerList,
    UComplex,
    Circuit,
    IniRegSave, // Graphics,  TEMc

     {Some units which have global vars defined here}
    Spectrum,
    LoadShape,
    TempShape,
    PriceShape,
    XYCurve,
    GrowthShape,
    Monitor,
    EnergyMeter,
    Sensor,
    TCC_Curve,
    WireData,
    CNData,
    TSData,
    LineSpacing,
    Storage,
    PVSystem,
    InvControl,
    ExpControl;

const
    CRLF = #13#10;

    PI = 3.14159265359;

    TwoPi = 2.0 * PI;

    RadiansToDegrees = 57.29577951;

    EPSILON = 1.0e-12;   // Default tiny floating point
    EPSILON2 = 1.0e-3;   // Default for Real number mismatch testing

    POWERFLOW = 1;  // Load model types for solution
    ADMITTANCE = 2;

      // For YPrim matrices
    ALL_YPRIM = 0;
    SERIES = 1;
    SHUNT = 2;

      {Control Modes}
    CONTROLSOFF = -1;
    EVENTDRIVEN = 1;
    TIMEDRIVEN = 2;
    CTRLSTATIC = 0;

      {Randomization Constants}
    GAUSSIAN = 1;
    UNIFORM = 2;
    LOGNORMAL = 3;

      {Autoadd Constants}
    GENADD = 1;
    CAPADD = 2;

      {ERRORS}
    SOLUTION_ABORT = 99;

      {For General Sequential Time Simulations}
    USEDAILY = 0;
    USEYEARLY = 1;
    USEDUTY = 2;
    USENONE = -1;

      {Earth Model}
    SIMPLECARSON = 1;
    FULLCARSON = 2;
    DERI = 3;

      {Profile Plot Constants}
    PROFILE3PH = 9999; // some big number > likely no. of phases
    PROFILEALL = 9998;
    PROFILEALLPRI = 9997;
    PROFILELLALL = 9996;
    PROFILELLPRI = 9995;
    PROFILELL = 9994;


var

    DLLFirstTime: Boolean = TRUE;
    DLLDebugFile: TextFile;
    ProgramName: String;
    DSS_Registry: TIniRegSave; // Registry   (See Executive)

    IsDLL,
    NoFormsAllowed: Boolean;

    ActiveCircuit: TDSSCircuit;
    ActiveDSSClass: TDSSClass;
    LastClassReferenced: Integer;  // index of class of last thing edited
    ActiveDSSObject: TDSSObject;
    NumCircuits: Integer;
    MaxCircuits: Integer;
    MaxBusLimit: Integer; // Set in Validation
    MaxAllocationIterations: Integer;
    Circuits: TPointerList;
    DSSObjs: TPointerList;

    AuxParser: TParser;  // Auxiliary parser for use by anybody for reparsing values

//{****} DebugTrace:TextFile;


    ErrorPending: Boolean;
    CmdResult,
    ErrorNumber: Integer;
    LastErrorMessage: String;

    DefaultEarthModel: Integer;
    ActiveEarthModel: Integer;

    LastFileCompiled: String;
    LastCommandWasCompile: Boolean;

    CALPHA: Complex;  {120-degree shift constant}
    SQRT2: Double;
    SQRT3: Double;
    InvSQRT3: Double;
    InvSQRT3x1000: Double;
    SolutionAbort: Boolean;
    InShowResults: Boolean;
    Redirect_Abort: Boolean;
    In_Redirect: Boolean;
    DIFilesAreOpen: Boolean;
    AutoShowExport: Boolean;
    SolutionWasAttempted: Boolean;

    GlobalHelpString: String;
    GlobalPropertyValue: String;
    GlobalResult: String;
    LastResultFile: String;
    VersionString: String;

    LogQueries: Boolean;
    QueryFirstTime: Boolean;
    QueryLogFileName: String;
    QueryLogFile: TextFile;

    DefaultEditor: String;     // normally, Notepad
    DefaultFontSize: Integer;
    DefaultFontName: String;
    DefaultFontStyles: Integer; // TFontStyles; TEMc
    DSSFileName: String;     // Name of current exe or DLL
    DSSDirectory: String;     // where the current exe resides
    StartupDirectory: String;     // Where we started
    DataDirectory: String;     // used to be DSSDataDirectory
    OutputDirectory: String;     // output files go here, same as DataDirectory if writable
    CircuitName_: String;     // Name of Circuit with a "_" appended

    DefaultBaseFreq: Double;
    DaisySize: Double;

   // Some commonly used classes   so we can find them easily
    LoadShapeClass: TLoadShape;
    TShapeClass: TTshape;
    PriceShapeClass: TPriceShape;
    XYCurveClass: TXYCurve;
    GrowthShapeClass: TGrowthShape;
    SpectrumClass: TSpectrum;
    SolutionClass: TDSSClass;
    EnergyMeterClass: TEnergyMeter;
   // FeederClass        :TFeeder;
    MonitorClass: TDSSMonitor;
    SensorClass: TSensor;
    TCC_CurveClass: TTCC_Curve;
    WireDataClass: TWireData;
    CNDataClass: TCNData;
    TSDataClass: TTSData;
    LineSpacingClass: TLineSpacing;
    StorageClass: TStorage;
    PVSystemClass: TPVSystem;
    InvControlClass: TInvControl;
    ExpControlClass: TExpControl;

    EventStrings: TStringList;
    SavedFileList: TStringList;

    DSSClassList: TPointerList; // pointers to the base class types
    ClassNames: THashList;

    UpdateRegistry: Boolean;  // update on program exit
    CPU_Freq: Int64;          // Used to store the CPU frequency
    CPU_Cores: Integer;


procedure DoErrorMsg(const S, Emsg, ProbCause: String; ErrNum: Integer);
procedure DoSimpleMsg(const S: String; ErrNum: Integer);

procedure ClearAllCircuits;

procedure SetObject(const param: String);
function SetActiveBus(const BusName: String): Integer;
procedure SetDataPath(const PathName: String);

procedure SetLastResultFile(const Fname: String);

procedure MakeNewCircuit(const Name: String);

procedure AppendGlobalResult(const s: String);
procedure AppendGlobalResultCRLF(const S: String);  // Separate by CRLF

procedure ResetQueryLogFile;
procedure WriteQueryLogFile(const Prop, S: String);

procedure WriteDLLDebugFile(const S: String);

procedure ReadDSS_Registry;
procedure WriteDSS_Registry;

function IsDSSDLL(Fname: String): Boolean;

function GetOutputDirectory: String;

procedure MyReallocMem(var p: Pointer; newsize: Integer);
function MyAllocMem(nbytes: Cardinal): Pointer;


implementation


uses {Forms,   Controls,}
    SysUtils,
     {LCLIntf, LCLType,} dynlibs,
    resource,
    versiontypes,
    versionresource,
    CmdForms,
    Solution,
    Executive;

     {Intrinsic Ckt Elements}

type

    THandle = Integer;

    TDSSRegister = function(var ClassName: Pchar): Integer;  // Returns base class 1 or 2 are defined
   // Users can only define circuit elements at present

var

    LastUserDLLHandle: THandle;
    DSSRegisterProc: TDSSRegister;   // of last library loaded

function GetDefaultDataDirectory: String;
begin
{$IFDEF UNIX}
    Result := GetEnvironmentVariable('HOME') + '/Documents';
{$ENDIF}
{$IFDEF WINDOWS}
    Result := GetEnvironmentVariable('HOMEDRIVE') + GetEnvironmentVariable('HOMEPATH') + '\Documents';
{$ENDIF}
end;

function GetDefaultScratchDirectory: String;
//Var
//  ThePath:Array[0..MAX_PATH] of char;
begin
//  FillChar(ThePath, SizeOF(ThePath), #0);
//  SHGetFolderPath (0, CSIDL_LOCAL_APPDATA, 0, 0, ThePath);
//  Result := ThePath;
  {$IFDEF UNIX}
    Result := '/tmp';
  {$ENDIF}
  {$IFDEF WINDOWS}
    Result := GetEnvironmentVariable('LOCALAPPDATA');
  {$ENDIF}
end;

function GetOutputDirectory: String;
begin
    Result := OutputDirectory;
end;

{--------------------------------------------------------------}
function IsDSSDLL(Fname: String): Boolean;

begin
    Result := FALSE;

    // Ignore if "DSSLIB.DLL"
    if CompareText(ExtractFileName(Fname), 'dsslib.dll') = 0 then
        Exit;

    LastUserDLLHandle := LoadLibrary(Pchar(Fname));
    if LastUserDLLHandle <> 0 then
    begin

   // Assign the address of the DSSRegister proc to DSSRegisterProc variable
        @DSSRegisterProc := GetProcAddress(LastUserDLLHandle, 'DSSRegister');
        if @DSSRegisterProc <> NIL then
            Result := TRUE
        else
            FreeLibrary(LastUserDLLHandle);

    end;

end;

//----------------------------------------------------------------------------
procedure DoErrorMsg(const S, Emsg, ProbCause: String; ErrNum: Integer);

var
    Msg: String;
    Retval: Integer;
begin

    Msg := Format('Error %d Reported From OpenDSS Intrinsic Function: ', [Errnum]) + CRLF + S + CRLF + CRLF + 'Error Description: ' + CRLF + Emsg + CRLF + CRLF + 'Probable Cause: ' + CRLF + ProbCause;

    if not NoFormsAllowed then
    begin

        if In_Redirect then
        begin
            RetVal := DSSMessageDlg(Msg, FALSE);
            if RetVal = -1 then
                Redirect_Abort := TRUE;
        end
        else
            DSSMessageDlg(Msg, TRUE);

    end;

    LastErrorMessage := Msg;
    ErrorNumber := ErrNum;
    AppendGlobalResultCRLF(Msg);
end;

//----------------------------------------------------------------------------
procedure AppendGlobalResultCRLF(const S: String);

begin
    if Length(GlobalResult) > 0 then
        GlobalResult := GlobalResult + CRLF + S
    else
        GlobalResult := S;
end;

//----------------------------------------------------------------------------
procedure DoSimpleMsg(const S: String; ErrNum: Integer);

var
    Retval: Integer;
begin

    if not NoFormsAllowed then
    begin
        if In_Redirect then
        begin
            RetVal := DSSMessageDlg(Format('(%d) OpenDSS %s%s', [Errnum, CRLF, S]), FALSE);
            if RetVal = -1 then
                Redirect_Abort := TRUE;
        end
        else
            DSSInfoMessageDlg(Format('(%d) OpenDSS %s%s', [Errnum, CRLF, S]));
    end;

    LastErrorMessage := S;
    ErrorNumber := ErrNum;
    AppendGlobalResultCRLF(S);
end;


//----------------------------------------------------------------------------
procedure SetObject(const param: String);

{Set object active by name}

var
    dotpos: Integer;
    ObjName, ObjClass: String;

begin

      // Split off Obj class and name
    dotpos := Pos('.', Param);
    case dotpos of
        0:
            ObjName := Copy(Param, 1, Length(Param));  // assume it is all name; class defaults
    else
    begin
        ObjClass := Copy(Param, 1, dotpos - 1);
        ObjName := Copy(Param, dotpos + 1, Length(Param));
    end;
    end;

    if Length(ObjClass) > 0 then
        SetObjectClass(ObjClass);

    ActiveDSSClass := DSSClassList.Get(LastClassReferenced);
    if ActiveDSSClass <> NIL then
    begin
        if not ActiveDSSClass.SetActive(Objname) then
        begin // scroll through list of objects untill a match
            DoSimpleMsg('Error! Object "' + ObjName + '" not found.' + CRLF + parser.CmdString, 904);
        end
        else
            with ActiveCircuit do
            begin
                case ActiveDSSObject.DSSObjType of
                    DSS_OBJECT: ;  // do nothing for general DSS object

                else
                begin   // for circuit types, set ActiveCircuit Element, too
                    ActiveCktElement := ActiveDSSClass.GetActiveObj;
                end;
                end;
            end;
    end
    else
        DoSimpleMsg('Error! Active object type/class is not set.', 905);

end;

//----------------------------------------------------------------------------
function SetActiveBus(const BusName: String): Integer;


begin

   // Now find the bus and set active
    Result := 0;

    with ActiveCircuit do
    begin
        if BusList.ListSize = 0 then
            Exit;   // Buslist not yet built
        ActiveBusIndex := BusList.Find(BusName);
        if ActiveBusIndex = 0 then
        begin
            Result := 1;
            AppendGlobalResult('SetActiveBus: Bus ' + BusName + ' Not Found.');
        end;
    end;

end;

procedure ClearAllCircuits;

begin

    ActiveCircuit := Circuits.First;
    while ActiveCircuit <> NIL do
    begin
        ActiveCircuit.Free;
        ActiveCircuit := Circuits.Next;
    end;
    Circuits.Free;
    Circuits := TPointerList.Create(2);   // Make a new list of circuits
    NumCircuits := 0;

    // Revert on key global flags to Original States
    DefaultEarthModel := DERI;
    LogQueries := FALSE;
    MaxAllocationIterations := 2;

end;


procedure MakeNewCircuit(const Name: String);

//Var
//   handle :Integer;
var
    S: String;

begin


    if NumCircuits <= MaxCircuits - 1 then
    begin
        ActiveCircuit := TDSSCircuit.Create(Name);
        ActiveDSSObject := ActiveSolutionObj;
         {*Handle := *} Circuits.Add(ActiveCircuit);
        Inc(NumCircuits);
        S := Parser.Remainder;    // Pass remainder of string on to vsource.
         {Create a default Circuit}
        SolutionABort := FALSE;
         {Voltage source named "source" connected to SourceBus}
        DSSExecutive.Command := 'New object=vsource.source Bus1=SourceBus ' + S;  // Load up the parser as if it were read in
    end
    else
    begin
        DoErrorMsg('MakeNewCircuit',
            'Cannot create new circuit.',
            'Max. Circuits Exceeded.' + CRLF +
            '(Max no. of circuits=' + inttostr(Maxcircuits) + ')', 906);
    end;
end;


procedure AppendGlobalResult(const S: String);

// Append a string to Global result, separated by commas

begin
    if Length(GlobalResult) = 0 then
        GlobalResult := S
    else
        GlobalResult := GlobalResult + ', ' + S;
end;


function GetDSSVersion: String;
(* Unlike most of AboutText (below), this takes significant activity at run-    *)
 (* time to extract version/release/build numbers from resource information      *)
 (* appended to the binary.                                                      *)

var
    Stream: TResourceStream;
    vr: TVersionResource;
    fi: TVersionFixedInfo;

begin
    RESULT := 'Unknown.';
    try

 (* This raises an exception if version info has not been incorporated into the  *)
 (* binary (Lazarus Project -> Project Options -> Version Info -> Version        *)
 (* numbering).                                                                  *)

        Stream := TResourceStream.CreateFromID(HINSTANCE, 1, Pchar(RT_VERSION));
        try
            vr := TVersionResource.Create;
            try
                vr.SetCustomRawDataStream(Stream);
                fi := vr.FixedInfo;
                RESULT := 'Version ' + IntToStr(fi.FileVersion[0]) + '.' + IntToStr(fi.FileVersion[1]) +
                    ' release ' + IntToStr(fi.FileVersion[2]) + ' build ' + IntToStr(fi.FileVersion[3]) + LineEnding;
                vr.SetCustomRawDataStream(NIL)
            finally
                vr.Free
            end
        finally
            Stream.Free
        end
    except
    end
end;


procedure WriteDLLDebugFile(const S: String);

begin

    AssignFile(DLLDebugFile, OutputDirectory + 'DSSDLLDebug.TXT');
    if DLLFirstTime then
    begin
        Rewrite(DLLDebugFile);
        DLLFirstTime := FALSE;
    end
    else
        Append(DLLDebugFile);
    Writeln(DLLDebugFile, S);
    CloseFile(DLLDebugFile);

end;

function IsDirectoryWritable(const Dir: String): Boolean;
var
    TempFile: array[0..MAX_PATH] of Char;
begin
    if GetTempFileName(Pchar(Dir), 'DA', 0, TempFile) <> 0 then
        Result := DeleteFile(TempFile)
    else
        Result := FALSE;
end;

procedure SetDataPath(const PathName: String);
var
    ScratchPath: String;
// Pathname may be null
begin
    if (Length(PathName) > 0) and not DirectoryExists(PathName) then
    begin
  // Try to create the directory
        if not CreateDir(PathName) then
        begin
            DosimpleMsg('Cannot create ' + PathName + ' directory.', 907);
            Exit;
        end;
    end;

    DataDirectory := PathName;

  // Put a \ on the end if not supplied. Allow a null specification.
    if Length(DataDirectory) > 0 then
    begin
        ChDir(DataDirectory);   // Change to specified directory
{$IFDEF WINDOWS}
        if DataDirectory[Length(DataDirectory)] <> '\' then
            DataDirectory := DataDirectory + '\';
{$ENDIF}
{$IFDEF UNIX}
        if DataDirectory[Length(DataDirectory)] <> '/' then
            DataDirectory := DataDirectory + '/';
{$ENDIF}
    end;

  // see if DataDirectory is writable. If not, set OutputDirectory to the user's appdata
    if IsDirectoryWritable(DataDirectory) then
    begin
        OutputDirectory := DataDirectory;
    end
    else
    begin
{$IFDEF WINDOWS}
        ScratchPath := GetDefaultScratchDirectory + '\' + ProgramName + '\';
{$ENDIF}
{$IFDEF UNIX}
        ScratchPath := GetDefaultScratchDirectory + '/' + ProgramName + '/';
{$ENDIF}
        if not DirectoryExists(ScratchPath) then
            CreateDir(ScratchPath);
        OutputDirectory := ScratchPath;
    end;
end;

procedure ReadDSS_Registry;
var
    TestDataDirectory: String;
begin
    DSS_Registry.Section := 'MainSect';
  {$IFDEF Darwin}
    DefaultEditor := DSS_Registry.ReadString('Editor', 'open -t');
    DefaultFontSize := StrToInt(DSS_Registry.ReadString('ScriptFontSize', '12'));
    DefaultFontName := DSS_Registry.ReadString('ScriptFontName', 'Geneva');
  {$ENDIF}
  {$IFDEF Linux}
    DefaultEditor := DSS_Registry.ReadString('Editor', 'xdg-open');
    DefaultFontSize := StrToInt(DSS_Registry.ReadString('ScriptFontSize', '10'));
    DefaultFontName := DSS_Registry.ReadString('ScriptFontName', 'Arial');
  {$ENDIF}
  {$IFDEF Windows}
    DefaultEditor := DSS_Registry.ReadString('Editor', 'Notepad.exe');
    DefaultFontSize := StrToInt(DSS_Registry.ReadString('ScriptFontSize', '8'));
    DefaultFontName := DSS_Registry.ReadString('ScriptFontName', 'MS Sans Serif');
  {$ENDIF}
    DefaultFontStyles := 1; // []; TEMc
//  If DSS_Registry.ReadBool('ScriptFontBold', TRUE)    Then DefaultFontStyles := DefaultFontStyles + [fsbold];
//  If DSS_Registry.ReadBool('ScriptFontItalic', FALSE) Then DefaultFontStyles := DefaultFontStyles + [fsItalic];
    DefaultBaseFreq := StrToInt(DSS_Registry.ReadString('BaseFrequency', '60'));
    LastFileCompiled := DSS_Registry.ReadString('LastFile', '');
    TestDataDirectory := DSS_Registry.ReadString('DataPath', DataDirectory);
    if SysUtils.DirectoryExists(TestDataDirectory) then
        SetDataPath(TestDataDirectory)
    else
        SetDataPath(DataDirectory);
end;


procedure WriteDSS_Registry;
begin
    if UpdateRegistry then
    begin
        DSS_Registry.Section := 'MainSect';
        DSS_Registry.WriteString('Editor', DefaultEditor);
        DSS_Registry.WriteString('ScriptFontSize', Format('%d', [DefaultFontSize]));
        DSS_Registry.WriteString('ScriptFontName', Format('%s', [DefaultFontName]));
        DSS_Registry.WriteBool('ScriptFontBold', FALSE); // (fsBold in DefaultFontStyles));  TEMc
        DSS_Registry.WriteBool('ScriptFontItalic', FALSE); // (fsItalic in DefaultFontStyles));  TEMc
        DSS_Registry.WriteString('BaseFrequency', Format('%d', [Round(DefaultBaseFreq)]));
        DSS_Registry.WriteString('LastFile', LastFileCompiled);
        DSS_Registry.WriteString('DataPath', DataDirectory);
    end;
end;

procedure ResetQueryLogFile;
begin
    QueryFirstTime := TRUE;
end;


procedure WriteQueryLogfile(const Prop, S: String);

{Log file is written after a query command if LogQueries is true.}

begin

    try
        QueryLogFileName := OutputDirectory + 'QueryLog.CSV';
        AssignFile(QueryLogFile, QueryLogFileName);
        if QueryFirstTime then
        begin
            Rewrite(QueryLogFile);  // clear the file
            Writeln(QueryLogFile, 'Time(h), Property, Result');
            QueryFirstTime := FALSE;
        end
        else
            Append(QueryLogFile);

        Writeln(QueryLogFile, Format('%.10g, %s, %s', [ActiveCircuit.Solution.DynaVars.dblHour, Prop, S]));
        CloseFile(QueryLogFile);
    except
        On E: Exception do
            DoSimpleMsg('Error writing Query Log file: ' + E.Message, 908);
    end;

end;

procedure SetLastResultFile(const Fname: String);

begin
    LastResultfile := Fname;
    ParserVars.Add('@lastfile', Fname);
end;

function MyAllocMem(nbytes: Cardinal): Pointer;
begin
    Result := AllocMem(Nbytes);
    WriteDLLDebugFile(Format('Allocating %d bytes @ %p', [nbytes, Result]));
end;

procedure MyReallocMem(var p: Pointer; newsize: Integer);

begin
    WriteDLLDebugFile(Format('Reallocating @ %p, new size= %d', [p, newsize]));
    ReallocMem(p, newsize);
end;

initialization

   // LCL can't show any forms until after Application.Initialize
    NoFormsAllowed := TRUE;

   {Various Constants and Switches}

    CALPHA := Cmplx(-0.5, -0.866025); // -120 degrees phase shift
    SQRT2 := Sqrt(2.0);
    SQRT3 := Sqrt(3.0);
    InvSQRT3 := 1.0 / SQRT3;
    InvSQRT3x1000 := InvSQRT3 * 1000.0;
    CmdResult := 0;
    DIFilesAreOpen := FALSE;
    ErrorNumber := 0;
    ErrorPending := FALSE;
    GlobalHelpString := '';
    GlobalPropertyValue := '';
    LastResultFile := '';
    In_Redirect := FALSE;
    InShowResults := FALSE;
    IsDLL := FALSE;
    LastCommandWasCompile := FALSE;
    LastErrorMessage := '';
    MaxCircuits := 1;  //  This version only allows one circuit at a time
    MaxAllocationIterations := 2;
    SolutionAbort := FALSE;
    AutoShowExport := FALSE;
    SolutionWasAttempted := FALSE;

    DefaultBaseFreq := 60.0;
    DaisySize := 1.0;
    DefaultEarthModel := DERI;
    ActiveEarthModel := DefaultEarthModel;

   {Initialize filenames and directories}

    ProgramName := 'OpenDSS';
    DSSFileName := GetDSSExeFile;
    DSSDirectory := ExtractFilePath(DSSFileName);
   // want to know if this was built for 64-bit, not whether running on 64 bits
   // (i.e. we could have a 32-bit build running on 64 bits; not interested in that
{$IFDEF CPU64}
    VersionString := 'Version ' + GetDSSVersion + ' (64-bit build)';
{$ELSE ! CPU32}
    VersionString := 'Version ' + GetDSSVersion + ' (32-bit build)';
{$ENDIF}
{$IFDEF WINDOWS}
    StartupDirectory := GetCurrentDir + '\';
    SetDataPath(GetDefaultDataDirectory + '\' + ProgramName + '\');
    DSS_Registry := TIniRegSave.Create(DataDirectory + 'opendss.ini');
{$ENDIF}
{$IFDEF UNIX}
    StartupDirectory := GetCurrentDir + '/';
    SetDataPath(GetDefaultDataDirectory + '/' + ProgramName + '/');
    DSS_Registry := TIniRegSave.Create(DataDirectory + 'opendss.ini');
{$ENDIF}

    AuxParser := TParser.Create;
{$IFDEF Darwin}
    DefaultEditor := 'open -t';
    DefaultFontSize := 12;
    DefaultFontName := 'Geneva';
{$ENDIF}
{$IFDEF Linux}
    DefaultEditor := 'xdg-open';
    DefaultFontSize := 10;
    DefaultFontName := 'Arial';
{$ENDIF}
{$IFDEF Windows}
    DefaultEditor := 'NotePad';
    DefaultFontSize := 8;
    DefaultFontName := 'MS Sans Serif';
{$ENDIF}

    EventStrings := TStringList.Create;
    SavedFileList := TStringList.Create;

    LogQueries := FALSE;
    QueryLogFileName := '';
    UpdateRegistry := TRUE;
    CPU_Freq := 1000; // until we can query it
//   QueryPerformanceFrequency(CPU_Freq);
    CPU_Cores := CPUCount;


   //WriteDLLDebugFile('DSSGlobals');

finalization

  // Dosimplemsg('Enter DSSGlobals Unit Finalization.');
    Auxparser.Free;

    EventStrings.Free;
    SavedFileList.Free;

    with DSSExecutive do
        if RecorderOn then
            Recorderon := FALSE;

    DSSExecutive.Free;  {Writes to Registry}
    DSS_Registry.Free;  {Close Registry}


end.
