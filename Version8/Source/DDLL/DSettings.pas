unit DSettings;

interface

function SettingsI(mode: Longint; arg: Longint): Longint; CDECL;
function SettingsF(mode: Longint; arg: Double): Double; CDECL;
function SettingsS(mode: Longint; arg: pAnsiChar): pAnsiChar; CDECL;
procedure SettingsV(mode: Longint; out arg: Variant); CDECL;

implementation

uses
    DSSGlobals,
    ExecHelper,
    Variants;

function SettingsI(mode: Longint; arg: Longint): Longint; CDECL;
begin
    Result := 0;       // Deafult return value
    case mode of
        0:
        begin  // Setting.Allowduplicates read
            Result := 0;
            if ActiveCircuit[ActiveActor] <> NIL then
                if ActiveCircuit[ActiveActor].DuplicatesAllowed then
                    Result := 1
                else
                    Result := 0;
        end;
        1:
        begin  // Setting.Allowduplicates read
            if ActiveCircuit[ActiveActor] <> NIL then
            begin
                if arg = 1 then
                    ActiveCircuit[ActiveActor].DuplicatesAllowed := TRUE
                else
                    ActiveCircuit[ActiveActor].DuplicatesAllowed := FALSE
            end;
        end;
        2:
        begin  // Settings.ZoneLock read
            if ActiveCircuit[ActiveActor] <> NIL then
            begin
                Result := 0;
                if ActiveCircuit[ActiveActor].ZonesLocked then
                    Result := 1;
            end
            else
                Result := 0;
        end;
        3:
        begin // Settings.ZoneLock Write
            if ActiveCircuit[ActiveActor] <> NIL then
            begin
                if arg = 1 then
                    ActiveCircuit[ActiveActor].ZonesLocked := TRUE
                else
                    ActiveCircuit[ActiveActor].ZonesLocked := FALSE
            end;
        end;
        4:
        begin // Settings.CktModel read
            if ActiveCircuit[ActiveActor] <> NIL then
            begin
                if ActiveCircuit[ActiveActor].PositiveSequence then
                    Result := 2
                else
                    Result := 1;
            end
            else
                Result := 0;
        end;
        5:
        begin  // Settings.CktModel Write
            if ActiveCircuit[ActiveActor] <> NIL then
                case arg of
                    2:
                        ActiveCircuit[ActiveActor].PositiveSequence := TRUE;
                else
                    ActiveCircuit[ActiveActor].PositiveSequence := FALSE;
                end;
        end;
        6:
        begin // Settings.Trapezoidal read
            if ActiveCircuit[ActiveActor] <> NIL then
            begin
                if ActiveCircuit[ActiveActor].TrapezoidalIntegration then
                    Result := 1;
            end
            else
                Result := 0;
        end;
        7:
        begin  // Settings.Trapezoidal Write
            if ActiveCircuit[ActiveActor] <> NIL then
            begin
                if arg = 1 then
                    ActiveCircuit[ActiveActor].TrapezoidalIntegration := TRUE
                else
                    ActiveCircuit[ActiveActor].TrapezoidalIntegration := FALSE;
            end;
        end
    else
        Result := -1;
    end;
end;

//****************************Floating point type properties**********************
function SettingsF(mode: Longint; arg: Double): Double; CDECL;
begin
    Result := 0.0; // Deafult return value
    case mode of
        0:
        begin  // Settings.AllocationFactors
            if ActiveCircuit[ActiveActor] <> NIL then
                DoSetAllocationFactors(arg);
        end;
        1:
        begin // Settings.NormVminpu read
            if ActiveCircuit[ActiveActor] <> NIL then
                Result := ActiveCircuit[ActiveActor].NormalMinVolts
            else
                Result := 0.0;
            ;
        end;
        2:
        begin  // Settings.NormVminpu write
            if ActiveCircuit[ActiveActor] <> NIL then
                ActiveCircuit[ActiveActor].NormalMinVolts := arg;
        end;
        3:
        begin  // Settings.NormVmaxpu read
            if ActiveCircuit[ActiveActor] <> NIL then
                Result := ActiveCircuit[ActiveActor].NormalMaxVolts
            else
                Result := 0.0;
            ;
        end;
        4:
        begin  // Settings.NormVmaxpu write
            if ActiveCircuit[ActiveActor] <> NIL then
                ActiveCircuit[ActiveActor].NormalMaxVolts := arg;
        end;
        5:
        begin  // Settings.EmergVminpu read
            if ActiveCircuit[ActiveActor] <> NIL then
                Result := ActiveCircuit[ActiveActor].EmergMinVolts
            else
                Result := 0.0;
            ;
        end;
        6:
        begin  // Settings.EmergVminpu write
            if ActiveCircuit[ActiveActor] <> NIL then
                ActiveCircuit[ActiveActor].EmergMinVolts := arg;
        end;
        7:
        begin  // Settings.EmergVmaxpu read
            if ActiveCircuit[ActiveActor] <> NIL then
                Result := ActiveCircuit[ActiveActor].EmergMaxVolts
            else
                Result := 0.0;
            ;
        end;
        8:
        begin  // Settings.EmergVmaxpu write
            if ActiveCircuit[ActiveActor] <> NIL then
                ActiveCircuit[ActiveActor].EmergMaxVolts := arg;
        end;
        9:
        begin  // Settings.UEWeight read
            if ActiveCircuit[ActiveActor] <> NIL then
            begin
                Result := ActiveCircuit[ActiveActor].UEWeight
            end
            else
                Result := 0.0;
        end;
        10:
        begin  // Settings.UEWeight Write
            if ActiveCircuit[ActiveActor] <> NIL then
            begin
                ActiveCircuit[ActiveActor].UEWeight := arg
            end;
        end;
        11:
        begin  // Settings.LossWeight read
            if ActiveCircuit[ActiveActor] <> NIL then
            begin
                Result := ActiveCircuit[ActiveActor].LossWeight;
            end
            else
                Result := 0.0;
        end;
        12:
        begin  // Settings.LossWeight write
            if ActiveCircuit[ActiveActor] <> NIL then
            begin
                ActiveCircuit[ActiveActor].LossWeight := arg
            end;
        end;
        13:
        begin  // Settings.PriceSignal read
            if ActiveCircuit[ActiveActor] <> NIL then
                Result := ActiveCircuit[ActiveActor].Pricesignal
            else
                Result := 0.0;
        end;
        14:
        begin  // Settings.PriceSignal write
            if ActiveCircuit[ActiveActor] <> NIL then
                ActiveCircuit[ActiveActor].PriceSignal := arg;
        end
    else
        Result := -1.0;
    end;
end;

//*******************************Strings type properties**************************
function SettingsS(mode: Longint; arg: pAnsiChar): pAnsiChar; CDECL;

var
    i: Integer;

begin
    Result := pAnsiChar(Ansistring(''));  // Deafult return value
    case mode of
        0:
        begin  // Settings.AutoBusLits read
            if ActiveCircuit[ActiveActor] <> NIL then
                with ActiveCircuit[ActiveActor].AutoAddBusList do
                begin
                    for i := 1 to ListSize do
                        AppendGlobalResult(Get(i));
                    Result := pAnsiChar(Ansistring(GlobalResult));
                end
            else
                Result := pAnsiChar(Ansistring(''));
        end;
        1:
        begin  // Settings.AutoBusLits write
            if ActiveCircuit[ActiveActor] <> NIL then
                DoAutoAddBusList(Widestring(arg));
        end;
        2:
        begin  // Settings.PriceCurve read
            if ActiveCircuit[ActiveActor] <> NIL then
                Result := pAnsiChar(Ansistring(ActiveCircuit[ActiveActor].PriceCurve))
            else
                Result := pAnsiChar(Ansistring(''));
        end;
        3:
        begin  // Settings.PriceCurve write
            if ActiveCircuit[ActiveActor] <> NIL then
                with ActiveCircuit[ActiveActor] do
                begin
                    PriceCurve := Widestring(arg);
                    PriceCurveObj := LoadShapeClass[ActiveActor].Find(Pricecurve);
                    if PriceCurveObj = NIL then
                        DoSimpleMsg('Price Curve: "' + Pricecurve + '" not found.', 5006);
                end;
        end;
    else
        Result := pAnsiChar(Ansistring('Error, parameter not recognized'));
    end;
end;

//*******************************Variant type properties******************************
procedure SettingsV(mode: Longint; out arg: Variant); CDECL;

var
    i, j, Count, Num: Integer;

begin
    case mode of
        0:
        begin  // Settings.UERegs read
            if ActiveCircuit[ActiveActor] <> NIL then
            begin
                arg := VarArrayCreate([0, ActiveCircuit[ActiveActor].NumUERegs - 1], varInteger);
                for i := 0 to ActiveCircuit[ActiveActor].NumUERegs - 1 do
                begin
                    arg[i] := ActiveCircuit[ActiveActor].UERegs^[i + 1]
                end;
            end
            else
                arg := VarArrayCreate([0, 0], varInteger);
        end;
        1:
        begin  // Settings.UERegs write
            if ActiveCircuit[ActiveActor] <> NIL then
            begin
                ReAllocMem(ActiveCircuit[ActiveActor].UERegs, Sizeof(ActiveCircuit[ActiveActor].UERegs^[1]) * (1 - VarArrayLowBound(arg, 1) + VarArrayHighBound(arg, 1)));
                j := 1;
                for i := VarArrayLowBound(arg, 1) to VarArrayHighBound(arg, 1) do
                begin
                    ActiveCircuit[ActiveActor].UERegs^[j] := arg[i];
                    Inc(j);
                end;
            end;
        end;
        2:
        begin  // Settings.LossRegs read
            if ActiveCircuit[ActiveActor] <> NIL then
            begin
                arg := VarArrayCreate([0, ActiveCircuit[ActiveActor].NumLossRegs - 1], varInteger);
                for i := 0 to ActiveCircuit[ActiveActor].NumLossRegs - 1 do
                begin
                    arg[i] := ActiveCircuit[ActiveActor].LossRegs^[i + 1]
                end;
            end
            else
                arg := VarArrayCreate([0, 0], varInteger);
        end;
        3:
        begin  // Settings.LossRegs write
            if ActiveCircuit[ActiveActor] <> NIL then
            begin
                ReAllocMem(ActiveCircuit[ActiveActor].LossRegs, Sizeof(ActiveCircuit[ActiveActor].LossRegs^[1]) * (1 - VarArrayLowBound(arg, 1) + VarArrayHighBound(arg, 1)));
                j := 1;
                for i := VarArrayLowBound(arg, 1) to VarArrayHighBound(arg, 1) do
                begin
                    ActiveCircuit[ActiveActor].LossRegs^[j] := arg[i];
                    Inc(j);
                end;
            end;
        end;
        4:
        begin  // Settings.VoltageBases read
            if ActiveCircuit[ActiveActor] <> NIL then
                with ActiveCircuit[ActiveActor] do
                begin
          {Count the number of voltagebases specified}
                    i := 0;
                    repeat
                        Inc(i);
                    until LegalVoltageBases^[i] = 0.0;
                    Count := i - 1;
                    arg := VarArrayCreate([0, Count - 1], varDouble);
                    for i := 0 to Count - 1 do
                        arg[i] := LegalVoltageBases^[i + 1];
                end
            else
                arg := VarArrayCreate([0, 0], varDouble);
        end;
        5:
        begin  // Settings.VoltageBases write
            Num := VarArrayHighBound(arg, 1) - VarArrayLowBound(arg, 1) + 1;
     {LegalVoltageBases is a zero-terminated array, so we have to allocate
      one more than the number of actual values}
            with ActiveCircuit[ActiveActor] do
            begin
                Reallocmem(LegalVoltageBases, Sizeof(LegalVoltageBases^[1]) * (Num + 1));
                j := 1;
                for i := VarArrayLowBound(arg, 1) to VarArrayHighBound(arg, 1) do
                begin
                    LegalVoltageBases^[j] := arg[i];
                    Inc(j)
                end;
                LegalVoltageBases^[Num + 1] := 0.0;
            end;
        end
    else
        arg[0] := 'Error, parameter not recognized'
    end;
end;


end.
