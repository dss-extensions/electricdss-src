unit ExpControl;

{$IFDEF FPC}{$MODE Delphi}{$ENDIF}

{
  ----------------------------------------------------------
  Copyright (c) 2015-2016, University of Pittsburgh
  All rights reserved.
  ----------------------------------------------------------

  Notes: adapted and simplified from InvControl for adaptive controller research
}

interface

uses
    Command,
    ControlClass,
    ControlElem,
    CktElement,
    DSSClass,
    PVSystem,
    uComplex,
    utilities,
    Dynamics,
    PointerList,
    Classes;

type
    TExpControl = class(TControlClass)
    PROTECTED
        procedure DefineProperties;
        function MakeLike(const ExpControlName: String): Integer; OVERRIDE;
    PUBLIC
        constructor Create;
        destructor Destroy; OVERRIDE;

        function Edit: Integer; OVERRIDE;     // uses global parser
        function NewObject(const ObjName: String): Integer; OVERRIDE;
        procedure UpdateAll;
    end;

    TExpControlObj = class(TControlElem)
    PRIVATE
        ControlActionHandle: Integer;
        ControlledElement: array of TPVSystemObj;    // list of pointers to controlled PVSystem elements
        MonitoredElement: TDSSCktElement;  // First PVSystem element for now

            // PVSystemList information
        FListSize: Integer;
        FPVSystemNameList: TStringList;
        FPVSystemPointerList: PointerList.TPointerList;

            // working storage for each PV system under management
        FPriorVpu: array of Double;
        FPresentVpu: array of Double;
        FPendingChange: array of Integer;
        FVregs: array of Double;
        FPriorQ: array of Double;
        FTargetQ: array of Double;
        FWithinTol: array of Boolean;

            // temp storage for biggest PV system, not each one
        cBuffer: array of Complex;

            // user-supplied parameters (also PVSystemList and EventLog)
        FVregInit: Double;
        FSlope: Double;
        FVregTau: Double;
        FQbias: Double;
        FVregMin: Double;
        FVregMax: Double;
        FQmaxLead: Double;
        FQmaxLag: Double;
        FdeltaQ_factor: Double;
        FVoltageChangeTolerance: Double; // hard-wire now?
        FVarChangeTolerance: Double;     // hard-wire now?

        procedure Set_PendingChange(Value: Integer; DevIndex: Integer);
        function Get_PendingChange(DevIndex: Integer): Integer;
        function ReturnElementsList: String;
        procedure UpdateExpControl(i: Integer);
    PROTECTED
        procedure Set_Enabled(Value: Boolean); OVERRIDE;
    PUBLIC

        constructor Create(ParClass: TDSSClass; const ExpControlName: String);
        destructor Destroy; OVERRIDE;

        procedure MakePosSequence; OVERRIDE;  // Make a positive Sequence Model
        procedure RecalcElementData; OVERRIDE;
        procedure CalcYPrim; OVERRIDE;    // Always Zero for an ExpControl

            // Sample control quantities and set action times in Control Queue
        procedure Sample; OVERRIDE;

            // Do the action that is pending from last sample
        procedure DoPendingAction(const Code, ProxyHdl: Integer); OVERRIDE;

        procedure Reset; OVERRIDE;  // Reset to initial defined state

        procedure GetInjCurrents(Curr: pComplexArray); OVERRIDE;
        procedure GetCurrents(Curr: pComplexArray); OVERRIDE;

        procedure InitPropertyValues(ArrayOffset: Integer); OVERRIDE;
        procedure DumpProperties(var F: TextFile; Complete: Boolean); OVERRIDE;

        function MakePVSystemList: Boolean;
        function GetPropertyValue(Index: Integer): String; OVERRIDE;

        property PendingChange[DevIndex: Integer]: Integer READ Get_PendingChange WRITE Set_PendingChange;

    end;

var
    ActiveExpControlObj: TExpControlObj;

{--------------------------------------------------------------------------}
implementation

uses
    ParserDel,
    Sysutils,
    DSSClassDefs,
    DSSGlobals,
    Circuit,
    uCmatrix,
    MathUtil,
    Math;

const

    NumPropsThisClass = 11;

    NONE = 0;
    CHANGEVARLEVEL = 1;

{--------------------------------------------------------------------------}
constructor TExpControl.Create;  // Creates superstructure for all ExpControl objects
begin
    inherited Create;
    Class_name := 'ExpControl';
    DSSClassType := DSSClassType + EXP_CONTROL;
    DefineProperties;
    CommandList := TCommandList.Create(Slice(PropertyName^, NumProperties));
    CommandList.Abbrev := TRUE;
end;

{--------------------------------------------------------------------------}
destructor TExpControl.Destroy;

begin

    inherited Destroy;
end;

//- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
procedure TExpControl.DefineProperties;
begin

    Numproperties := NumPropsThisClass;
    CountProperties;   // Get inherited property count
    AllocatePropertyArrays;

     // Define Property names
    PropertyName[1] := 'PVSystemList';
    PropertyName[2] := 'Vreg';
    PropertyName[3] := 'Slope';
    PropertyName[4] := 'VregTau';
    PropertyName[5] := 'Qbias';
    PropertyName[6] := 'VregMin';
    PropertyName[7] := 'VregMax';
    PropertyName[8] := 'QmaxLead';
    PropertyName[9] := 'QmaxLag';
    PropertyName[10] := 'EventLog';
    PropertyName[11] := 'DeltaQ_factor';

    PropertyHelp[1] := 'Array list of PVSystems to be controlled.' + CRLF + CRLF +
        'If not specified, all PVSystems in the circuit are assumed to be controlled by this ExpControl.';
    PropertyHelp[2] := 'Per-unit voltage at which reactive power is zero; defaults to 1.0.' + CRLF + CRLF +
        'This may self-adjust when VregTau > 0, limited by VregMin and VregMax ' +
        'The equilibrium point of reactive power is also affected by Qbias';
    PropertyHelp[3] := 'Per-unit reactive power injection / per-unit voltage deviation from Vreg; defaults to 50.' + CRLF + CRLF +
        'Unlike InvControl, base reactive power is constant at the inverter kva rating.';
    PropertyHelp[4] := 'Time constant for adaptive Vreg. Defaults to 1200 seconds.' + CRLF + CRLF +
        'When the control injects or absorbs reactive power due to a voltage deviation from the Q=0 crossing of vvc_curve1, ' +
        'the Q=0 crossing will move toward the actual terminal voltage with this time constant. ' +
        'Over time, the effect is to gradually bring inverter reactive power to zero as the grid voltage changes due to non-solar effects. ' +
        'If zero, then Vreg stays fixed';
    PropertyHelp[5] := 'Equilibrium per-unit reactive power when V=Vreg; defaults to 0.' + CRLF + CRLF +
        'Enter > 0 for lagging (capacitive) bias, < 0 for leading (inductive) bias.';
    PropertyHelp[6] := 'Lower limit on adaptive Vreg; defaults to 0.95 per-unit';
    PropertyHelp[7] := 'Upper limit on adaptive Vreg; defaults to 1.05 per-unit';
    PropertyHelp[8] := 'Limit on leading (inductive) reactive power injection, in per-unit of base kva; defaults to 1.' + CRLF + CRLF +
        'Even if QmaxLead = 1, the reactive power injection is still ' +
        'limited by dynamic headroom when actual real power output exceeds 0%';
    PropertyHelp[9] := 'Limit on lagging (capacitive) reactive power injection, in per-unit of base kva; defaults to 1.' + CRLF + CRLF +
        'Even if QmaxLag = 1, the reactive power injection is still ' +
        'limited by dynamic headroom when actual real power output exceeds 0%';
    PropertyHelp[10] := '{Yes/True* | No/False} Default is No for ExpControl. Log control actions to Eventlog.';
    PropertyHelp[11] := 'Convergence parameter; Defaults to 0.7. ' + CRLF + CRLF +
        'Sets the maximum change (in per unit) from the prior var output level to the desired var output level during each control iteration. ' +
        'If numerical instability is noticed in solutions such as var sign changing from one control iteration to the next and voltages oscillating between two values with some separation, ' +
        'this is an indication of numerical instability (use the EventLog to diagnose). ' +
        'If the maximum control iterations are exceeded, and no numerical instability is seen in the EventLog of via monitors, then try increasing the value of this parameter to reduce the number ' +
        'of control iterations needed to achieve the control criteria, and move to the power flow solution.';

    ActiveProperty := NumPropsThisClass;
    inherited DefineProperties;  // Add defs of inherited properties to bottom of list
end;

{--------------------------------------------------------------------------}
function TExpControl.NewObject(const ObjName: String): Integer;
begin
    // Make a new ExpControl and add it to ExpControl class list
    with ActiveCircuit do
    begin
        ActiveCktElement := TExpControlObj.Create(Self, ObjName);
        Result := AddObjectToList(ActiveDSSObject);
    end;
end;

{--------------------------------------------------------------------------}
function TExpControl.Edit: Integer;
var
    ParamPointer: Integer;
    ParamName: String;
    Param: String;


begin
    ActiveExpControlObj := ElementList.Active;
    ActiveCircuit.ActiveCktElement := ActiveExpControlObj;
    Result := 0;
    with ActiveExpControlObj do
    begin
        ParamPointer := 0;
        ParamName := Parser.NextParam;
        Param := Parser.StrValue;
        while Length(Param) > 0 do
        begin
            if Length(ParamName) = 0 then
                Inc(ParamPointer)
            else
                ParamPointer := CommandList.GetCommand(ParamName);

            if (ParamPointer > 0) and (ParamPointer <= NumProperties) then
                PropertyValue[ParamPointer] := Param;

            case ParamPointer of
                0:
                    DoSimpleMsg('Unknown parameter "' + ParamName + '" for Object "' + Class_Name + '.' + Name + '"', 364);
                1:
                begin
                    InterpretTStringListArray(Param, FPVSystemNameList);
                    FPVSystemPointerList.Clear; // clear this for resetting on first sample
                    FListSize := FPVSystemNameList.count;
                end;
                2:
                    if Parser.DblValue > 0 then
                        FVregInit := Parser.DblValue;
                3:
                    if Parser.DblValue > 0 then
                        FSlope := Parser.DblValue;
                4:
                    if Parser.DblValue >= 0 then
                        FVregTau := Parser.DblValue; // zero means fixed Vreg
                5:
                    FQbias := Parser.DblValue;
                6:
                    if Parser.DblValue > 0 then
                        FVregMin := Parser.DblValue;
                7:
                    if Parser.DblValue > 0 then
                        FVregMax := Parser.DblValue;
                8:
                    if Parser.DblValue >= 0 then
                        FQmaxLead := Parser.DblValue;
                9:
                    if Parser.DblValue >= 0 then
                        FQmaxLag := Parser.DblValue;
                10:
                    ShowEventLog := InterpretYesNo(param);
                11:
                    FdeltaQ_factor := Parser.DblValue;
            else
        // Inherited parameters
                ClassEdit(ActiveExpControlObj, ParamPointer - NumPropsthisClass)
            end;
            ParamName := Parser.NextParam;
            Param := Parser.StrValue;
        end;
        RecalcElementData;
    end;
end;

function TExpControl.MakeLike(const ExpControlName: String): Integer;
var
    OtherExpControl: TExpControlObj;
    i, j: Integer;
begin
    Result := 0;
   {See if we can find this ExpControl name in the present collection}
    OtherExpControl := Find(ExpControlName);
    if OtherExpControl <> NIL then
        with ActiveExpControlObj do
        begin

            NPhases := OtherExpControl.Fnphases;
            NConds := OtherExpControl.Fnconds; // Force Reallocation of terminal stuff

            for i := 1 to FPVSystemPointerList.ListSize do
            begin
                ControlledElement[i] := OtherExpControl.ControlledElement[i];
                FWithinTol[i] := OtherExpControl.FWithinTol[i];
            end;

            FListSize := OtherExpControl.FListSize;
            FVoltageChangeTolerance := OtherExpControl.FVoltageChangeTolerance;
            FVarChangeTolerance := OtherExpControl.FVarChangeTolerance;
            FVregInit := OtherExpControl.FVregInit;
            FSlope := OtherExpControl.FSlope;
            FVregTau := OtherExpControl.FVregTau;
            FQbias := OtherExpControl.FQbias;
            FVregMin := OtherExpControl.FVregMin;
            FVregMax := OtherExpControl.FVregMax;
            FQmaxLead := OtherExpControl.FQmaxLead;
            FQmaxLag := OtherExpControl.FQmaxLag;
            FdeltaQ_factor := OtherExpControl.FdeltaQ_factor;
            for j := 1 to ParentClass.NumProperties do
                PropertyValue[j] := OtherExpControl.PropertyValue[j];

        end
    else
        DoSimpleMsg('Error in ExpControl MakeLike: "' + ExpControlName + '" Not Found.', 370);

end;

{==========================================================================}
{                    TExpControlObj                                        }
{==========================================================================}

constructor TExpControlObj.Create(ParClass: TDSSClass; const ExpControlName: String);

begin
    inherited Create(ParClass);
    Name := LowerCase(ExpControlName);
    DSSObjType := ParClass.DSSClassType;

    ElementName := '';

     {
       Control elements are zero current sources that attach to a terminal of a
       power-carrying device, but do not alter voltage or current flow.
       Define a default number of phases and conductors here and update in
       RecalcElementData routine if necessary. This allocates arrays for voltages
       and currents and gives more direct access to the values, if needed
     }
    NPhases := 3;  // Directly set conds and phases
    Fnconds := 3;
    Nterms := 1;  // this forces allocation of terminals and conductors
     // This general feature should not be used for ExpControl,
     // because it controls more than one PVSystem

    ShowEventLog := FALSE;

    ControlledElement := NIL;
    FPVSystemNameList := NIL;
    FPVSystemPointerList := NIL;
    cBuffer := NIL;
    FPriorVpu := NIL;
    FPresentVpu := NIL;
    FPendingChange := NIL;
    FPriorQ := NIL;
    FTargetQ := NIL;
    FWithinTol := NIL;

    FVoltageChangeTolerance := 0.0001;  // per-unit
    FVarChangeTolerance := 0.0001;  // per-unit

    FPVSystemNameList := TSTringList.Create;
    FPVSystemPointerList := PointerList.TPointerList.Create(20);  // Default size and increment

     // user parameters for dynamic Vreg
    FVregInit := 1.0;
    FSlope := 50.0;
    FVregTau := 1200.0;
    FVregs := NIL;
    FQbias := 0.0;
    FVregMin := 0.95;
    FVregMax := 1.05;
    FQmaxLead := 1.0;
    FQmaxLag := 1.0;
    FdeltaQ_factor := 0.7; // only on control iterations, not the final solution

     //generic for control
    FPendingChange := NIL;

    InitPropertyValues(0);
end;

destructor TExpControlObj.Destroy;
begin
    ElementName := '';
    Finalize(ControlledElement);
    Finalize(cBuffer);
    Finalize(FPriorVpu);
    Finalize(FPresentVpu);
    Finalize(FPendingChange);
    Finalize(FPriorQ);
    Finalize(FTargetQ);
    Finalize(FWithinTol);
    Finalize(FVregs);
    inherited Destroy;
end;

procedure TExpControlObj.RecalcElementData;
var
    i: Integer;
    maxord: Integer;
begin
    if FPVSystemPointerList.ListSize = 0 then
        MakePVSystemList;

    if FPVSystemPointerList.ListSize > 0 then
    begin
    {Setting the terminal of the ExpControl device to same as the 1st PVSystem element}
        MonitoredElement := TDSSCktElement(FPVSystemPointerList.Get(1));   // Set MonitoredElement to 1st PVSystem in lise
        Setbus(1, MonitoredElement.Firstbus);
    end;

    maxord := 0; // will be the size of cBuffer
    for i := 1 to FPVSystemPointerList.ListSize do
    begin
        // User ControlledElement[] as the pointer to the PVSystem elements
        ControlledElement[i] := TPVSystemObj(FPVSystemPointerList.Get(i));  // pointer to i-th PVSystem
        Nphases := ControlledElement[i].NPhases;  // TEMC TODO - what if these are different sizes (same concern exists with InvControl)
        Nconds := Nphases;
        if (ControlledElement[i] = NIL) then
            DoErrorMsg('ExpControl: "' + Self.Name + '"',
                'Controlled Element "' + FPVSystemNameList.Strings[i - 1] + '" Not Found.',
                ' PVSystem object must be defined previously.', 361);
        if ControlledElement[i].Yorder > maxord then
            maxord := ControlledElement[i].Yorder;
        ControlledElement[i].ActiveTerminalIdx := 1; // Make the 1 st terminal active
    end;
    if maxord > 0 then
        SetLength(cBuffer, SizeOF(Complex) * maxord);
end;

procedure TExpControlObj.MakePosSequence;
// ***  This assumes the PVSystem devices have already been converted to pos seq
begin
    if FPVSystemPointerList.ListSize = 0 then
        RecalcElementData;
  // TEMC - from here to inherited was copied from InvControl
    Nphases := 3;
    Nconds := 3;
    Setbus(1, MonitoredElement.GetBus(ElementTerminal));
    if FPVSystemPointerList.ListSize > 0 then
    begin
    {Setting the terminal of the ExpControl device to same as the 1st PVSystem element}
    { This sets it to a realistic value to avoid crashes later }
        MonitoredElement := TDSSCktElement(FPVSystemPointerList.Get(1));   // Set MonitoredElement to 1st PVSystem in lise
        Setbus(1, MonitoredElement.Firstbus);
        Nphases := MonitoredElement.NPhases;
        Nconds := Nphases;
    end;
    inherited;
end;

procedure TExpControlObj.CalcYPrim;
begin
end;

procedure TExpControlObj.GetCurrents(Curr: pComplexArray);
var
    i: Integer;
begin
// Control is a zero current source
    for i := 1 to Fnconds do
        Curr^[i] := CZERO;
end;

procedure TExpControlObj.GetInjCurrents(Curr: pComplexArray);
var
    i: Integer;
begin
// Control is a zero current source
    for i := 1 to Fnconds do
        Curr^[i] := CZERO;
end;

procedure TExpControlObj.DumpProperties(var F: TextFile; Complete: Boolean);
var
    i: Integer;
begin
    inherited DumpProperties(F, Complete);

    with ParentClass do
        for i := 1 to NumProperties do
        begin
            Writeln(F, '~ ', PropertyName^[i], '=', PropertyValue[i]);
        end;

    if Complete then
    begin
        Writeln(F);
    end;
end;

procedure TExpControlObj.DoPendingAction;
var
    i: Integer;
    Qset, DeltaQ: Double;
    Qmaxpu, Qpu: Double;
    Qbase: Double;
    PVSys: TPVSystemObj;
begin
    for i := 1 to FPVSystemPointerList.ListSize do
    begin
        PVSys := ControlledElement[i];   // Use local variable in loop
        if PendingChange[i] = CHANGEVARLEVEL then
        begin
            PVSys.VWmode := FALSE;
            PVSys.ActiveTerminalIdx := 1; // Set active terminal of PVSystem to terminal 1
            PVSys.Varmode := VARMODEKVAR;  // Set var mode to VARMODEKVAR to indicate we might change kvar
            FTargetQ[i] := 0.0;
            Qbase := PVSys.kVARating;
            Qpu := PVSys.Presentkvar / Qbase; // no change for now

            if (FWithinTol[i] = FALSE) then
            begin
        // look up Qpu from the slope crossing at Vreg, and add the bias
                Qpu := -FSlope * (FPresentVpu[i] - FVregs[i]) + FQbias;
                if ShowEventLog then
                    AppendtoEventLog('ExpControl.' + Self.Name + ',' + PVSys.Name,
                        Format(' Setting Qpu= %.5g at FVreg= %.5g, Vpu= %.5g', [Qpu, FVregs[i], FPresentVpu[i]]));
            end;

      // apply limits on Qpu, then define the target in kVAR
            Qmaxpu := Sqrt(1 - Sqr(PVSys.PresentkW / Qbase)); // dynamic headroom
            if Abs(Qpu) > Qmaxpu then
                Qpu := QmaxPu * Sign(Qpu);
            if Qpu < -FQmaxLead then
                Qpu := -FQmaxLead;
            if Qpu > FQmaxLag then
                Qpu := FQmaxLag;
            FTargetQ[i] := Qbase * Qpu;

      // only move the non-bias component by deltaQ_factor in this control iteration
            DeltaQ := FTargetQ[i] - FPriorQ[i];
            Qset := FPriorQ[i] + DeltaQ * FdeltaQ_factor;
 //     Qset := FQbias * Qbase;
            if PVSys.Presentkvar <> Qset then
                PVSys.Presentkvar := Qset;
            if ShowEventLog then
                AppendtoEventLog('ExpControl.' + Self.Name + ',' + PVSys.Name,
                    Format(' Setting PVSystem output kvar= %.5g',
                    [PVSys.Presentkvar]));
            FPriorQ[i] := Qset;
            FPriorVpu[i] := FPresentVpu[i];
            ActiveCircuit.Solution.LoadsNeedUpdating := TRUE;
      // Force recalc of power parms
            Set_PendingChange(NONE, i);
        end
    end;

end;

procedure TExpControlObj.Sample;
var
    i, j: Integer;
    basekV, Vpresent: Double;
    Verr, Qerr: Double;
    PVSys: TPVSystemObj;
begin
  // If list is not defined, go make one from all PVSystem in circuit
    if FPVSystemPointerList.ListSize = 0 then
        RecalcElementData;

    if (FListSize > 0) then
    begin
    // If an ExpControl controls more than one PV, control each one
    // separately based on the PVSystem's terminal voltages, etc.
        for i := 1 to FPVSystemPointerList.ListSize do
        begin
            PVSys := ControlledElement[i];   // Use local variable in loop
      // Calculate the present average voltage  magnitude
            PVSys.ComputeVTerminal;
            for j := 1 to PVSys.Yorder do
                cBuffer[j] := PVSys.Vterminal^[j];
            BasekV := ActiveCircuit.Buses^[PVSys.terminals^[1].busRef].kVBase;
            Vpresent := 0;
            for j := 1 to PVSys.NPhases do
                Vpresent := Vpresent + Cabs(cBuffer[j]);
            FPresentVpu[i] := (Vpresent / PVSys.NPhases) / (basekV * 1000.0);
      // both errors are in per-unit
            Verr := Abs(FPresentVpu[i] - FPriorVpu[i]);
            Qerr := Abs(PVSys.Presentkvar - FTargetQ[i]) / PVSys.kVARating;

      // process the sample
            if (PVSys.InverterON = FALSE) and (PVSys.VarFollowInverter = TRUE) then
            begin // not injecting
                if (FVregTau > 0.0) then
                    FVregs[i] := FPresentVpu[i]; // tracking grid voltage while not injecting
                continue;
            end;
            PVSys.VWmode := FALSE;
            if (FWithinTol[i] = FALSE) then
            begin
                if ((Verr > FVoltageChangeTolerance) or (Qerr > FVarChangeTolerance) or
                    (ActiveCircuit.Solution.ControlIteration = 1)) then
                begin
                    FWithinTol[i] := FALSE;
                    Set_PendingChange(CHANGEVARLEVEL, i);
                    with  ActiveCircuit.Solution.DynaVars do
                        ControlActionHandle := ActiveCircuit.ControlQueue.Push(intHour, t + TimeDelay, PendingChange[i], 0, Self);
                    if ShowEventLog then
                        AppendtoEventLog('ExpControl.' + Self.Name + ' ' + PVSys.Name, Format
                            (' outside Hit Tolerance, Verr= %.5g, Qerr=%.5g', [Verr, Qerr]));
                end
                else
                begin
                    if ((Verr <= FVoltageChangeTolerance) and (Qerr <= FVarChangeTolerance)) then
                        FWithinTol[i] := TRUE;
                    if ShowEventLog then
                        AppendtoEventLog('ExpControl.' + Self.Name + ' ' + PVSys.Name, Format
                            (' within Hit Tolerance, Verr= %.5g, Qerr=%.5g', [Verr, Qerr]));
                end;
            end;
        end;  {For}
    end; {If FlistSize}
end;

procedure TExpControlObj.InitPropertyValues(ArrayOffset: Integer);
begin
    PropertyValue[1] := '';      // PVSystem list
    PropertyValue[2] := '1';     // initial Vreg
    PropertyValue[3] := '50';    // slope
    PropertyValue[4] := '1200.0'; // VregTau
    PropertyValue[5] := '0';     // Q bias
    PropertyValue[6] := '0.95';  // Vreg min
    PropertyValue[7] := '1.05';  // Vreg max
    PropertyValue[8] := '1';     // Qmax leading
    PropertyValue[9] := '1';     // Qmax lagging
    PropertyValue[10] := 'no';    // write event log?
    PropertyValue[11] := '0.7';   // DeltaQ_factor
    inherited  InitPropertyValues(NumPropsThisClass);
end;

function TExpControlObj.MakePVSystemList: Boolean;
var
    PVSysClass: TDSSClass;
    PVSys: TPVsystemObj;
    i: Integer;
begin
    Result := FALSE;
    PVSysClass := GetDSSClassPtr('PVsystem');
    if FListSize > 0 then
    begin    // Name list is defined - Use it
        SetLength(ControlledElement, FListSize + 1);  // Use this as the main pointer to PVSystem Elements
        SetLength(FPriorVpu, FListSize + 1);
        SetLength(FPresentVpu, FListSize + 1);
        SetLength(FPendingChange, FListSize + 1);
        SetLength(FPriorQ, FListSize + 1);
        SetLength(FTargetQ, FListSize + 1);
        SetLength(FWithinTol, FListSize + 1);
        SetLength(FVregs, FListSize + 1);
        for i := 1 to FListSize do
        begin
            PVSys := PVSysClass.Find(FPVSystemNameList.Strings[i - 1]);
            if Assigned(PVSys) and PVSys.Enabled then
                FPVSystemPointerList.New := PVSys;
        end;
    end
    else
    begin
     {Search through the entire circuit for enabled pvsysten objects and add them to the list}
        for i := 1 to PVSysClass.ElementCount do
        begin
            PVSys := PVSysClass.ElementList.Get(i);
            if PVSys.Enabled then
                FPVSystemPointerList.New := PVSys;
            FPVSystemNameList.Add(PVSys.Name);
        end;
        FListSize := FPVSystemPointerList.ListSize;

        SetLength(ControlledElement, FListSize + 1);

        SetLength(FPriorVpu, FListSize + 1);
        SetLength(FPresentVpu, FListSize + 1);

        SetLength(FPendingChange, FListSize + 1);
        SetLength(FPriorQ, FListSize + 1);
        SetLength(FTargetQ, FListSize + 1);
        SetLength(FWithinTol, FListSize + 1);
        SetLength(FVregs, FListSize + 1);
    end;  {Else}

  //Initialize arrays
    for i := 1 to FlistSize do
    begin
//    PVSys := PVSysClass.Find(FPVSystemNameList.Strings[i-1]);
//    Set_NTerms(PVSys.NTerms); // TODO - what is this for?
        FPriorVpu[i] := 0.0;
        FPresentVpu[i] := 0.0;
        FPriorQ[i] := -1.0;
        FTargetQ[i] := 0.0;
        FWithinTol[i] := FALSE;
        FVregs[i] := FVregInit;
        FPendingChange[i] := NONE;
    end; {For}
    RecalcElementData;
    if FPVSystemPointerList.ListSize > 0 then
        Result := TRUE;
end;

procedure TExpControlObj.Reset;
begin
  // inherited;
end;

function TExpControlObj.GetPropertyValue(Index: Integer): String;
begin
    Result := '';
    case Index of
        1:
            Result := ReturnElementsList;
        2:
            Result := Format('%.6g', [FVregInit]);
        3:
            Result := Format('%.6g', [FSlope]);
        4:
            Result := Format('%.6g', [FVregTau]);
        5:
            Result := Format('%.6g', [FQbias]);
        6:
            Result := Format('%.6g', [FVregMin]);
        7:
            Result := Format('%.6g', [FVregMax]);
        8:
            Result := Format('%.6g', [FQmaxLead]);
        9:
            Result := Format('%.6g', [FQmaxLag]);
        11:
            Result := Format('%.6g', [FdeltaQ_factor]);
    // 10 skipped, EventLog always went to the default handler
    else  // take the generic handler
        Result := inherited GetPropertyValue(index);
    end;
end;

function TExpControlObj.ReturnElementsList: String;
var
    i: Integer;
begin
    if FListSize = 0 then
    begin
        Result := '';
        Exit;
    end;
    Result := '[' + FPVSystemNameList.Strings[0];
    for i := 1 to FListSize - 1 do
    begin
        Result := Result + ', ' + FPVSystemNameList.Strings[i];
    end;
    Result := Result + ']';  // terminate the array
end;

procedure TExpControlObj.Set_Enabled(Value: Boolean);
begin
    inherited;
  {Reset controlled PVSystems to original PF}
end;

procedure TExpControlObj.Set_PendingChange(Value: Integer; DevIndex: Integer);
begin
    FPendingChange[DevIndex] := Value;
    DblTraceParameter := Value;
end;

procedure TExpControlObj.UpdateExpControl(i: Integer);
var
    j: Integer;
    PVSys: TPVSystemObj;
    dt, Verr: Double; // for DYNAMICVREG
begin
    for j := 1 to FPVSystemPointerList.ListSize do
    begin
        PVSys := ControlledElement[j];
        FWithinTol[j] := FALSE;
        if FVregTau > 0.0 then
        begin
            dt := ActiveCircuit.Solution.Dynavars.h;
            Verr := FPresentVpu[j] - FVregs[j];
            FVregs[j] := FVregs[j] + Verr * (1 - Exp(-dt / FVregTau));
        end
        else
        begin
            Verr := 0.0;
        end;
        if FVregs[j] < FVregMin then
            FVregs[j] := FVregMin;
        if FVregs[j] > FVregMax then
            FVregs[j] := FVregMax;
        PVSys.Set_Variable(5, FVregs[j]);
        if ShowEventLog then
            AppendtoEventLog('ExpControl.' + Self.Name + ',' + PVSys.Name,
                Format(' Setting new Vreg= %.5g Vpu=%.5g Verr=%.5g',
                [FVregs[j], FPresentVpu[j], Verr]));
    end;
end;

function TExpControlObj.Get_PendingChange(DevIndex: Integer): Integer;
begin
    Result := FPendingChange[DevIndex];
end;

//Called at end of main power flow solution loop
procedure TExpControl.UpdateAll;
var
    i: Integer;
begin
    for i := 1 to ElementList.ListSize do
        with TExpControlObj(ElementList.Get(i)) do
            if Enabled then
                UpdateExpControl(i);
end;

initialization

finalization

end.
