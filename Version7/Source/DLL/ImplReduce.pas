unit ImplReduce;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
    ComObj,
    ActiveX,
    OpenDSSengine_TLB,
    StdVcl;

type
    TReduce = class(TAutoObject, IReduce)
    PROTECTED
        function Get_Zmag: Double; SAFECALL;
        procedure Set_Zmag(Value: Double); SAFECALL;
        function Get_KeepLoad: Wordbool; SAFECALL;
        procedure Set_KeepLoad(Value: Wordbool); SAFECALL;
        function Get_EditString: Widestring; SAFECALL;
        procedure Set_EditString(const Value: Widestring); SAFECALL;
        function Get_StartPDElement: Widestring; SAFECALL;
        procedure Set_StartPDElement(const Value: Widestring); SAFECALL;
        function Get_EnergyMeter: Widestring; SAFECALL;
        procedure SaveCircuit(const CktName: Widestring); SAFECALL;
        procedure Set_EnergyMeter(const Value: Widestring); SAFECALL;
        procedure DoDefault; SAFECALL;
        procedure DoShortLines; SAFECALL;
        procedure Do1phLaterals; SAFECALL;
        procedure DoBranchRemove; SAFECALL;
        procedure DoDangling; SAFECALL;
        procedure DoLoopBreak; SAFECALL;
        procedure DoParallelLines; SAFECALL;
        procedure DoSwitches; SAFECALL;

    end;

implementation

uses
    Circuit,
    DSSGlobals,
    ComServ,
    Executive,
    EnergyMeter,
    ReduceAlgs,
    PDElement;

var
    ReduceEditString: String;
    EnergyMeterName: String;
    FirstPDelement: String;  // Full name

function TReduce.Get_Zmag: Double;
begin
    if Assigned(ActiveCircuit) then
        Result := ActiveCircuit.ReductionZmag
end;

procedure TReduce.Set_Zmag(Value: Double);
begin
    if Assigned(ActiveCircuit) then
        ActiveCircuit.ReductionZmag := Value;
end;

function TReduce.Get_KeepLoad: Wordbool;
begin
    if Assigned(ActiveCircuit) then
        Result := ActiveCircuit.ReduceLateralsKeepLoad;
end;

procedure TReduce.Set_KeepLoad(Value: Wordbool);
begin
    if Assigned(ActiveCircuit) then
        ActiveCircuit.ReduceLateralsKeepLoad := Value;
end;

function TReduce.Get_EditString: Widestring;
begin
    Result := ReduceEditString;
end;

procedure TReduce.Set_EditString(const Value: Widestring);
begin
    ReduceEditString := Value;
end;

function TReduce.Get_StartPDElement: Widestring;
begin
    Result := FirstPDelement;
end;

procedure TReduce.Set_StartPDElement(const Value: Widestring);
begin
    FirstPDelement := Value;
end;

function TReduce.Get_EnergyMeter: Widestring;
begin
    Result := EnergyMeterName;
end;

procedure TReduce.SaveCircuit(const CktName: Widestring);
begin
    DSSExecutive.Command := 'Save Circuit Dir=' + CktName;
end;

procedure TReduce.Set_EnergyMeter(const Value: Widestring);
begin
    EnergyMeterName := Value;
end;

procedure TReduce.DoDefault;
begin
    if EnergyMeterClass.SetActive(EnergyMeterName) then
        ActiveEnergyMeterObj := EnergyMeterClass.ElementList.Active;
    if Assigned(ActiveEnergyMeterObj) then
        with ActiveEnergyMeterObj do
        begin
            if not assigned(BranchList) then
                MakeMeterZoneLists;
            DoReduceDefault(BranchList);
        end;
end;

procedure TReduce.DoShortLines;
begin
    if EnergyMeterClass.SetActive(EnergyMeterName) then
        ActiveEnergyMeterObj := EnergyMeterClass.ElementList.Active;
    if Assigned(ActiveEnergyMeterObj) then
        with ActiveEnergyMeterObj do
        begin
            if not assigned(BranchList) then
                MakeMeterZoneLists;
            DoReduceShortLines(BranchList);
        end;
end;

procedure TReduce.Do1phLaterals;
begin
    if EnergyMeterClass.SetActive(EnergyMeterName) then
        ActiveEnergyMeterObj := EnergyMeterClass.ElementList.Active;
    if Assigned(ActiveEnergyMeterObj) then
        with ActiveEnergyMeterObj do
        begin
            if not assigned(BranchList) then
                MakeMeterZoneLists;
            DoRemoveAll_1ph_Laterals(BranchList);
        end;
end;

procedure TReduce.DoBranchRemove;
begin
    if Assigned(ActiveCircuit) then
    begin
        if EnergyMeterClass.SetActive(EnergyMeterName) then
            ActiveEnergyMeterObj := EnergyMeterClass.ElementList.Active;
        if Assigned(ActiveEnergyMeterObj) then
            with ActiveEnergyMeterObj do
            begin
                if not assigned(BranchList) then
                    MakeMeterZoneLists;
                with ActiveCircuit do
                begin
                    if SetElementActive(FirstPDelement) >= 0 then // element was found  0-based array
                        DoRemoveBranches(BranchList, ActiveCktElement as TPDElement, ReduceLateralsKeepLoad, ReduceEditString);
                end;
            end;
    end;
end;

procedure TReduce.DoDangling;
begin
    if EnergyMeterClass.SetActive(EnergyMeterName) then
        ActiveEnergyMeterObj := EnergyMeterClass.ElementList.Active;
    if Assigned(ActiveEnergyMeterObj) then
        with ActiveEnergyMeterObj do
        begin
            if not assigned(BranchList) then
                MakeMeterZoneLists;
            DoReduceDangling(BranchList);
        end;
end;

procedure TReduce.DoLoopBreak;
begin
    if EnergyMeterClass.SetActive(EnergyMeterName) then
        ActiveEnergyMeterObj := EnergyMeterClass.ElementList.Active;
    if Assigned(ActiveEnergyMeterObj) then
        with ActiveEnergyMeterObj do
        begin
            if not assigned(BranchList) then
                MakeMeterZoneLists;
            DoBreakLoops(BranchList);
        end;
end;

procedure TReduce.DoParallelLines;
begin
    if EnergyMeterClass.SetActive(EnergyMeterName) then
        ActiveEnergyMeterObj := EnergyMeterClass.ElementList.Active;
    if Assigned(ActiveEnergyMeterObj) then
        with ActiveEnergyMeterObj do
        begin
            if not assigned(BranchList) then
                MakeMeterZoneLists;
            DoMergeParallelLines(BranchList);
        end;
end;

procedure TReduce.DoSwitches;
begin
    if EnergyMeterClass.SetActive(EnergyMeterName) then
        ActiveEnergyMeterObj := EnergyMeterClass.ElementList.Active;
    if Assigned(ActiveEnergyMeterObj) then
        with ActiveEnergyMeterObj do
        begin
            if not assigned(BranchList) then
                MakeMeterZoneLists;
            DoRemoveAll_1ph_Laterals(BranchList);
        end;
end;

initialization
    TAutoObjectFactory.Create(ComServer, TReduce, Class_Reduce,
        ciInternal, tmApartment);

    ReduceEditString := ''; // Init to null string
    EnergyMeterName := '';
    FirstPDelement := '';
end.
