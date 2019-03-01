unit Circuit;

{$MODE Delphi}

{
   ----------------------------------------------------------
  Copyright (c) 2008-2015, Electric Power Research Institute, Inc.
  All rights reserved.
  ----------------------------------------------------------
}

{
 Change Log
   10-12-99 Added DuplicatesAllowed and ZonesLocked
   10-24-99 Added Losses Property
   12-15-99 Added Default load shapes and generator dispatch reference
   4-17=00  Add Loads List
   5-30-00  Added Positive Sequence Flag
   8-24-00  Added PriceCurve stuff   Updated 3-6-11
   8-1-01  Modified Compute Capacity to report up to loadmult=1
   9-25-15  Fixed broken repository
}

{$WARN UNIT_PLATFORM OFF}

interface

uses
    Classes,
    Solution,
    SysUtils,
    ArrayDef,
    HashList,
    PointerList,
    CktElement,
    DSSClass, {DSSObject,} Bus,
    LoadShape,
    PriceShape,
    ControlQueue,
    uComplex,
    AutoAdd,
    EnergyMeter,
    NamedObject,
    CktTree;

type
    TReductionStrategy = (rsDefault, rsStubs, rsTapEnds, rsMergeParallel, rsBreakLoop, rsDangling, rsSwitches);

    CktElementDef = record
        CktElementClass: Integer;
        devHandle: Integer;
    end;

    pCktElementDefArray = ^CktElementDefArray;
    CktElementDefArray = array[1..1] of CktElementDef;


     // for adding markers to Plot
    TBusMarker = class(TObject)
    // Must be defined before calling circuit plot
    PRIVATE

    PUBLIC
        BusName: String;
        AddMarkerColor: Integer; // Tcolor; TEMc
        AddMarkerCode,
        AddMarkerSize: Integer;

        constructor Create;
        destructor Destroy; OVERRIDE;
    end;

    TDSSCircuit = class(TNamedObject)

    PRIVATE
        NodeBuffer: pIntegerArray;
        NodeBufferMax: Integer;
        FBusNameRedefined: Boolean;
        FActiveCktElement: TDSSCktElement;
        FCaseName: String;

          // Temp arrays for when the bus swap takes place
        SavedBuses: pTBusArray;
        SavedBusNames: pStringArray;
        SavedNumBuses: Integer;

        FLoadMultiplier: Double;  // global multiplier for every load

        AbortBusProcess: Boolean;

        Branch_List: TCktTree; // topology from the first source, lazy evaluation
        BusAdjPC, BusAdjPD: TAdjArray; // bus adjacency lists of PD and PC elements


        procedure AddDeviceHandle(Handle: Integer);
        procedure AddABus;
        procedure AddANodeBus;
        function AddBus(const BusName: String; NNodes: Integer): Integer;
        procedure Set_ActiveCktElement(Value: TDSSCktElement);
        procedure Set_BusNameRedefined(Value: Boolean);
        function Get_Losses: Complex; //Total Circuit losses
        procedure Set_LoadMultiplier(Value: Double);
        procedure SaveBusInfo;
        procedure RestoreBusInfo;

        function SaveMasterFile: Boolean;
        function SaveDSSObjects: Boolean;
        function SaveFeeders: Boolean;
        function SaveBusCoords: Boolean;
        function SaveVoltageBases: Boolean;

        procedure ReallocDeviceList;
        procedure Set_CaseName(const Value: String);

        function Get_Name: String;


    PUBLIC

        ActiveBusIndex: Integer;
        Fundamental: Double;    // fundamental and default base frequency

        Control_BusNameRedefined: Boolean;  // Flag for use by control elements to detect redefinition of buses

        BusList,
        AutoAddBusList,
        DeviceList: THashList;
        DeviceRef: pCktElementDefArray;  //Type and handle of device

          // lists of pointers to different elements by class
        Faults,
        CktElements,
        PDElements,
        PCElements,
        DSSControls,
        Sources,
        MeterElements,
        Sensors,
        Monitors,
        EnergyMeters,
        Generators,
        StorageElements,
        PVSystems,
        Substations,
        Transformers,
        CapControls,
        RegControls,
        Lines,
        Loads,
        ShuntCapacitors,
        Reactors, // added for CIM XML export
        Feeders,
        SwtControls: PointerList.TPointerList;

        ControlQueue: TControlQueue;

        Solution: TSolutionObj;
        AutoAddObj: TAutoAdd;

          // For AutoAdd stuff
        UEWeight,
        LossWeight: Double;

        NumUEregs,
        NumLossRegs: Integer;
        Ueregs,
        LossRegs: pIntegerArray;

        CapacityStart,
        CapacityIncrement: Double;

        TrapezoidalIntegration,   // flag for trapezoidal integratio
        LogEvents: Boolean;

        LoadDurCurve: String;
        LoadDurCurveObj: TLoadShapeObj;
        PriceCurve: String;
        PriceCurveObj: TPriceShapeObj;

        NumDevices, NumBuses, NumNodes: Integer;
        MaxDevices, MaxBuses, MaxNodes: Integer;
        IncDevices, IncBuses, IncNodes: Integer;

          // Bus and Node stuff
        Buses: pTBusArray;
        MapNodeToBus: pTNodeBusArray;

          // Flags
        Issolved: Boolean;
        DuplicatesAllowed: Boolean;
        ZonesLocked: Boolean;
        MeterZonesComputed: Boolean;
        PositiveSequence: Boolean;  // Model is to be interpreted as Pos seq
        NeglectLoadY: Boolean;

          // Voltage limits
        NormalMinVolts,
        NormalMaxVolts,
        EmergMaxVolts,
        EmergMinVolts: Double;  //per unit voltage restraints for this circuit
        LegalVoltageBases: pDoubleArray;

          // Global circuit multipliers
        GeneratorDispatchReference,
        DefaultGrowthFactor,
        DefaultGrowthRate,
        GenMultiplier,   // global multiplier for every generator
        HarmMult: Double;
        DefaultHourMult: Complex;

        PriceSignal: Double; // price signal for entire circuit

          // EnergyMeter Totals
        RegisterTotals: TRegisterArray;

        DefaultDailyShapeObj,
        DefaultYearlyShapeObj: TLoadShapeObj;

        CurrentDirectory: String;

        ReductionStrategy: TReductionStrategy;
        ReductionMaxAngle, ReductionZmag: Double;
        ReductionStrategyString: String;

        PctNormalFactor: Double;

          {------Plot Marker Circuit Globals---------}
        NodeMarkerCode: Integer;
        NodeMarkerWidth: Integer;
        SwitchMarkerCode: Integer;

        TransMarkerSize: Integer;
        CapMarkerSize: Integer;
        RegMarkerSize: Integer;
        PVMarkerSize: Integer;
        StoreMarkerSize: Integer;
        FuseMarkerSize: Integer;
        RecloserMarkerSize: Integer;
        RelayMarkerSize: Integer;

        TransMarkerCode: Integer;
        CapMarkerCode: Integer;
        RegMarkerCode: Integer;
        PVMarkerCode: Integer;
        StoreMarkerCode: Integer;
        FuseMarkerCode: Integer;
        RecloserMarkerCode: Integer;
        RelayMarkerCode: Integer;

        MarkSwitches: Boolean;
        MarkTransformers: Boolean;
        MarkCapacitors: Boolean;
        MarkRegulators: Boolean;
        MarkPVSystems: Boolean;
        MarkStorage: Boolean;
        MarkFuses: Boolean;
        MarkReclosers: Boolean;
        MarkRelays: Boolean;

        BusMarkerList: TList;  // list of buses to mark

          {---------------------------------}

        ActiveLoadShapeClass: Integer;


        constructor Create(const aName: String);
        destructor Destroy; OVERRIDE;

        procedure AddCktElement(Handle: Integer);  // Adds last DSS object created to circuit
        procedure ClearBusMarkers;

        procedure TotalizeMeters;
        function ComputeCapacity: Boolean;

        function Save(Dir: String): Boolean;

        procedure ProcessBusDefs;
        procedure ReProcessBusDefs;
        procedure DoResetMeterZones;
        function SetElementActive(const FullObjectName: String): Integer;
        procedure InvalidateAllPCElements;

        procedure DebugDump(var F: TextFile);

          // Access to topology from the first source
        function GetTopology: TCktTree;
        procedure FreeTopology;
        function GetBusAdjacentPDLists: TAdjArray;
        function GetBusAdjacentPCLists: TAdjArray;

        property Name: String READ Get_Name;
        property CaseName: String READ FCaseName WRITE Set_CaseName;
        property ActiveCktElement: TDSSCktElement READ FActiveCktElement WRITE Set_ActiveCktElement;
        property Losses: Complex READ Get_Losses;  // Total Circuit PD Element losses
        property BusNameRedefined: Boolean READ FBusNameRedefined WRITE Set_BusNameRedefined;
        property LoadMultiplier: Double READ FLoadMultiplier WRITE Set_LoadMultiplier;

    end;

implementation

uses
    PDElement,
    CktElementClass,
    ParserDel,
    DSSClassDefs,
    DSSGlobals,
    Dynamics,
    Line,
    Vsource,
    Utilities,
    CmdForms;

//----------------------------------------------------------------------------
constructor TDSSCircuit.Create(const aName: String);

// Var Retval:Integer;

begin
    inherited Create('Circuit');

    IsSolved := FALSE;
     {*Retval   := *} SolutionClass.NewObject(Name);
    Solution := ActiveSolutionObj;

    LocalName := LowerCase(aName);

    CaseName := aName;  // Default case name to circuitname
                            // Sets CircuitName_

    Fundamental := DefaultBaseFreq;
    ActiveCktElement := NIL;
    ActiveBusIndex := 1;    // Always a bus

     // initial allocations increased from 100 to 1000 to speed things up

    MaxBuses := 1000;  // good sized allocation to start
    MaxDevices := 1000;
    MaxNodes := 3 * MaxBuses;
    IncDevices := 1000;
    IncBuses := 1000;
    IncNodes := 3000;

     // Allocate some nominal sizes
    BusList := THashList.Create(900);  // Bus name list Nominal size to start; gets reallocated
    DeviceList := THashList.Create(900);
    AutoAddBusList := THashList.Create(100);

    NumBuses := 0;  // Eventually allocate a single source
    NumDevices := 0;
    NumNodes := 0;

    Faults := TPointerList.Create(2);
    CktElements := TPointerList.Create(1000);
    PDElements := TPointerList.Create(1000);
    PCElements := TPointerList.Create(1000);
    DSSControls := TPointerList.Create(10);
    Sources := TPointerList.Create(10);
    MeterElements := TPointerList.Create(20);
    Monitors := TPointerList.Create(20);
    EnergyMeters := TPointerList.Create(5);
    Sensors := TPointerList.Create(5);
    Generators := TPointerList.Create(5);
    StorageElements := TPointerList.Create(5);
    PVSystems := TPointerList.Create(5);
    Feeders := TPointerList.Create(10);
    Substations := TPointerList.Create(5);
    Transformers := TPointerList.Create(10);
    CapControls := TPointerList.Create(10);
    SwtControls := TPointerList.Create(50);
    RegControls := TPointerList.Create(5);
    Lines := TPointerList.Create(1000);
    Loads := TPointerList.Create(1000);
    ShuntCapacitors := TPointerList.Create(20);
    Reactors := TPointerList.Create(5);

    Buses := Allocmem(Sizeof(Buses^[1]) * Maxbuses);
    MapNodeToBus := Allocmem(Sizeof(MapNodeToBus^[1]) * MaxNodes);
    DeviceRef := AllocMem(SizeOf(DeviceRef^[1]) * MaxDevices);

    ControlQueue := TControlQueue.Create;

    LegalVoltageBases := AllocMem(SizeOf(LegalVoltageBases^[1]) * 8);
     // Default Voltage Bases
    LegalVoltageBases^[1] := 0.208;
    LegalVoltageBases^[2] := 0.480;
    LegalVoltageBases^[3] := 12.47;
    LegalVoltageBases^[4] := 24.9;
    LegalVoltageBases^[5] := 34.5;
    LegalVoltageBases^[6] := 115.0;
    LegalVoltageBases^[7] := 230.0;
    LegalVoltageBases^[8] := 0.0;  // terminates array

    ActiveLoadShapeClass := USENONE; // Signify not set

    NodeBufferMax := 50;
    NodeBuffer := AllocMem(SizeOf(NodeBuffer^[1]) * NodeBufferMax); // A place to hold the nodes

     // Init global circuit load and harmonic source multipliers
    FLoadMultiplier := 1.0;
    GenMultiplier := 1.0;
    HarmMult := 1.0;

    PriceSignal := 25.0;   // $25/MWH

     // Factors for Autoadd stuff
    UEWeight := 1.0;  // Default to weighting UE same as losses
    LossWeight := 1.0;
    NumUEregs := 1;
    NumLossRegs := 1;
    UEregs := NIL;  // set to something so it wont break reallocmem
    LossRegs := NIL;
    Reallocmem(UEregs, sizeof(UEregs^[1]) * NumUEregs);
    Reallocmem(Lossregs, sizeof(Lossregs^[1]) * NumLossregs);
    UEregs^[1] := 10;   // Overload UE
    LossRegs^[1] := 13;   // Zone Losses

    CapacityStart := 0.9;     // for Capacity search
    CapacityIncrement := 0.005;

    LoadDurCurve := '';
    LoadDurCurveObj := NIL;
    PriceCurve := '';
    PriceCurveObj := NIL;

     // Flags
    DuplicatesAllowed := FALSE;
    ZonesLocked := FALSE;   // Meter zones recomputed after each change
    MeterZonesComputed := FALSE;
    PositiveSequence := FALSE;
    NeglectLoadY := FALSE;

    NormalMinVolts := 0.95;
    NormalMaxVolts := 1.05;
    EmergMaxVolts := 1.08;
    EmergMinVolts := 0.90;

    NodeMarkerCode := 16;
    NodeMarkerWidth := 1;
    MarkSwitches := FALSE;
    MarkTransformers := FALSE;
    MarkCapacitors := FALSE;
    MarkRegulators := FALSE;
    MarkPVSystems := FALSE;
    MarkStorage := FALSE;
    MarkFuses := FALSE;
    MarkReclosers := FALSE;

    SwitchMarkerCode := 5;
    TransMarkerCode := 35;
    CapMarkerCode := 38;
    RegMarkerCode := 17; //47;
    PVMarkerCode := 15;
    StoreMarkerCode := 9;
    FuseMarkerCode := 25;
    RecloserMarkerCode := 17;
    RelayMarkerCode := 17;

    TransMarkerSize := 1;
    CapMarkerSize := 3;
    RegMarkerSize := 5; //1;
    PVMarkerSize := 1;
    StoreMarkerSize := 1;
    FuseMarkerSize := 1;
    RecloserMarkerSize := 5;
    RelayMarkerSize := 5;

    BusMarkerList := TList.Create;
    BusMarkerList.Clear;

    TrapezoidalIntegration := FALSE;  // Default to Euler method
    LogEvents := FALSE;

    GeneratorDispatchReference := 0.0;
    DefaultGrowthRate := 1.025;
    DefaultGrowthFactor := 1.0;

    DefaultDailyShapeObj := LoadShapeClass.Find('default');
    DefaultYearlyShapeObj := LoadShapeClass.Find('default');

    CurrentDirectory := '';

    BusNameRedefined := TRUE;  // set to force rebuild of buslists, nodelists

    SavedBuses := NIL;
    SavedBusNames := NIL;

    ReductionStrategy := rsDefault;
    ReductionMaxAngle := 15.0;
    ReductionZmag := 0.02;

   {Misc objects}
    AutoAddObj := TAutoAdd.Create;

    Branch_List := NIL;
    BusAdjPC := NIL;
    BusAdjPD := NIL;


end;

//----------------------------------------------------------------------------
destructor TDSSCircuit.Destroy;
var
    i: Integer;
    pCktElem: TDSSCktElement;
    ElemName: String;

begin
    for i := 1 to NumDevices do
    begin
        try
            pCktElem := TDSSCktElement(CktElements.Get(i));
            ElemName := pCktElem.ParentClass.name + '.' + pCktElem.Name;
            pCktElem.Free;

        except
            ON E: Exception do
                DoSimpleMsg('Exception Freeing Circuit Element:' + ElemName + CRLF + E.Message, 423);
        end;
    end;

    for i := 1 to NumBuses do
        Buses^[i].Free;  // added 10-29-00

    Reallocmem(DeviceRef, 0);
    Reallocmem(Buses, 0);
    Reallocmem(MapNodeToBus, 0);
    Reallocmem(NodeBuffer, 0);
    Reallocmem(UEregs, 0);
    Reallocmem(Lossregs, 0);
    Reallocmem(LegalVoltageBases, 0);

    DeviceList.Free;
    BusList.Free;
    AutoAddBusList.Free;
    Solution.Free;
    PDElements.Free;
    PCElements.Free;
    DSSControls.Free;
    Sources.Free;
    Faults.Free;
    CktElements.Free;
    MeterElements.Free;
    Monitors.Free;
    EnergyMeters.Free;
    Sensors.Free;
    Generators.Free;
    StorageElements.Free;
    PVSystems.Free;
    Feeders.Free;
    Substations.Free;
    Transformers.Free;
    CapControls.Free;
    SwtControls.Free;
    RegControls.Free;
    Loads.Free;
    Lines.Free;
    ShuntCapacitors.Free;
    Reactors.Free;

    ControlQueue.Free;

    ClearBusMarkers;
    BusMarkerList.Free;

    AutoAddObj.Free;

    FreeTopology;

    inherited Destroy;
end;

//----------------------------------------------------------------------------
procedure TDSSCircuit.ProcessBusDefs;
var
    BusName: String;
    NNodes, NP, Ncond, i, j, iTerm, RetVal: Integer;
    NodesOK: Boolean;

begin
    with ActiveCktElement do
    begin
        np := NPhases;
        Ncond := NConds;

        Parser.Token := FirstBus;     // use parser functions to decode
        for iTerm := 1 to Nterms do
        begin
            NodesOK := TRUE;
           // Assume normal phase rotation  for default
            for i := 1 to np do
                NodeBuffer^[i] := i; // set up buffer with defaults

           // Default all other conductors to a ground connection
           // If user wants them ungrounded, must be specified explicitly!
            for i := np + 1 to NCond do
                NodeBuffer^[i] := 0;

           // Parser will override bus connection if any specified
            BusName := Parser.ParseAsBusName(NNodes, NodeBuffer);

           // Check for error in node specification
            for j := 1 to NNodes do
            begin
                if NodeBuffer^[j] < 0 then
                begin
                    retval := DSSMessageDlg('Error in Node specification for Element: "' + ParentClass.Name + '.' + Name + '"' + CRLF +
                        'Bus Spec: "' + Parser.Token + '"', FALSE);
                    NodesOK := FALSE;
                    if retval = -1 then
                    begin
                        AbortBusProcess := TRUE;
                        AppendGlobalresult('Aborted bus process.');
                        Exit
                    end;
                    Break;
                end;
            end;


           // Node -Terminal Connnections
           // Caution: Magic -- AddBus replaces values in nodeBuffer to correspond
           // with global node reference number.
            if NodesOK then
            begin
                ActiveTerminalIdx := iTerm;
                ActiveTerminal.BusRef := AddBus(BusName, Ncond);
                SetNodeRef(iTerm, NodeBuffer);  // for active circuit
            end;
            Parser.Token := NextBus;
        end;
    end;
end;


//----------------------------------------------------------------------------
procedure TDSSCircuit.AddABus;
begin
    if NumBuses > MaxBuses then
    begin
        Inc(MaxBuses, IncBuses);
        ReallocMem(Buses, SizeOf(Buses^[1]) * MaxBuses);
    end;
end;

//----------------------------------------------------------------------------
procedure TDSSCircuit.AddANodeBus;
begin
    if NumNodes > MaxNodes then
    begin
        Inc(MaxNodes, IncNodes);
        ReallocMem(MapNodeToBus, SizeOf(MapNodeToBus^[1]) * MaxNodes);
    end;
end;

//----------------------------------------------------------------------------
function TDSSCircuit.AddBus(const BusName: String; NNodes: Integer): Integer;

var
    NodeRef, i: Integer;
begin

// Trap error in bus name
    if Length(BusName) = 0 then
    begin  // Error in busname
        DoErrorMsg('TDSSCircuit.AddBus', 'BusName for Object "' + ActiveCktElement.Name + '" is null.',
            'Error in definition of object.', 424);
        for i := 1 to ActiveCktElement.NConds do
            NodeBuffer^[i] := 0;
        Result := 0;
        Exit;
    end;

    Result := BusList.Find(BusName);
    if Result = 0 then
    begin
        Result := BusList.Add(BusName);    // Result is index of bus
        Inc(NumBuses);
        AddABus;   // Allocates more memory if necessary
        Buses^[NumBuses] := TDSSBus.Create;
    end;

    {Define nodes belonging to the bus}
    {Replace Nodebuffer values with global reference number}
    with Buses^[Result] do
    begin
        for i := 1 to NNodes do
        begin
            NodeRef := Add(NodeBuffer^[i]);
            if NodeRef = NumNodes then
            begin  // This was a new node so Add a NodeToBus element ????
                AddANodeBus;   // Allocates more memory if necessary
                MapNodeToBus^[NumNodes].BusRef := Result;
                MapNodeToBus^[NumNodes].NodeNum := NodeBuffer^[i]
            end;
            NodeBuffer^[i] := NodeRef;  //  Swap out in preparation to setnoderef call
        end;
    end;
end;

//----------------------------------------------------------------------------
procedure TDSSCircuit.AddDeviceHandle(Handle: Integer);
begin
    if NumDevices > MaxDevices then
    begin
        MaxDevices := MaxDevices + IncDevices;
        ReallocMem(DeviceRef, Sizeof(DeviceRef^[1]) * MaxDevices);
    end;
    DeviceRef^[NumDevices].devHandle := Handle;    // Index into CktElements
    DeviceRef^[NumDevices].CktElementClass := LastClassReferenced;
end;


//----------------------------------------------------------------------------
function TDSSCircuit.SetElementActive(const FullObjectName: String): Integer;

// Fast way to set a cktelement active
var
    Devindex: Integer;
    DevClassIndex: Integer;
    DevType,
    DevName: String;

begin

    Result := 0;

    ParseObjectClassandName(FullObjectName, DevType, DevName);
    DevClassIndex := ClassNames.Find(DevType);
    if DevClassIndex = 0 then
        DevClassIndex := LastClassReferenced;
    Devindex := DeviceList.Find(DevName);
    while DevIndex > 0 do
    begin
        if DeviceRef^[Devindex].CktElementClass = DevClassIndex then   // we got a match
        begin
            ActiveDSSClass := DSSClassList.Get(DevClassIndex);
            LastClassReferenced := DevClassIndex;
            Result := DeviceRef^[Devindex].devHandle;
           // ActiveDSSClass.Active := Result;
          //  ActiveCktElement := ActiveDSSClass.GetActiveObj;
            ActiveCktElement := CktElements.Get(Result);
            Break;
        end;
        Devindex := Devicelist.FindNext;   // Could be duplicates
    end;

    CmdResult := Result;

end;

//----------------------------------------------------------------------------
procedure TDSSCircuit.Set_ActiveCktElement(Value: TDSSCktElement);
begin
    FActiveCktElement := Value;
    ActiveDSSObject := Value;
end;

//----------------------------------------------------------------------------
procedure TDSSCircuit.AddCktElement(Handle: Integer);


begin

   // Update lists that keep track of individual circuit elements
    Inc(NumDevices);

   // Resize DeviceList if no. of devices greatly exceeds allocation
    if Cardinal(NumDevices) > 2 * DeviceList.InitialAllocation then
        ReAllocDeviceList;
    DeviceList.Add(ActiveCktElement.Name);
    CktElements.Add(ActiveCktElement);

   {Build Lists of PC and PD elements}
    case (ActiveCktElement.DSSObjType and BaseClassMask) of
        PD_ELEMENT:
            PDElements.Add(ActiveCktElement);
        PC_ELEMENT:
            PCElements.Add(ActiveCktElement);
        CTRL_ELEMENT:
            DSSControls.Add(ActiveCktElement);
        METER_ELEMENT:
            MeterElements.Add(ActiveCktElement);
    else
       {Nothing}
    end;

   {Build  lists of Special elements and generic types}
    case (ActiveCktElement.DSSObjType and CLASSMASK) of
        MON_ELEMENT:
            Monitors.Add(ActiveCktElement);
        ENERGY_METER:
            EnergyMeters.Add(ActiveCktElement);
        SENSOR_ELEMENT:
            Sensors.Add(ActiveCktElement);
        GEN_ELEMENT:
            Generators.Add(ActiveCktElement);
        SOURCE:
            Sources.Add(ActiveCktElement);
        CAP_CONTROL:
            CapControls.Add(ActiveCktElement);
        SWT_CONTROL:
            SwtControls.Add(ActiveCktElement);
        REG_CONTROL:
            RegControls.Add(ActiveCktElement);
        LOAD_ELEMENT:
            Loads.Add(ActiveCktElement);
        CAP_ELEMENT:
            ShuntCapacitors.Add(ActiveCktElement);
        REACTOR_ELEMENT:
            Reactors.Add(ActiveCktElement);

       { Keep Lines, Transformer, and Lines and Faults in PDElements and separate lists
         so we can find them quickly.}
        XFMR_ELEMENT:
            Transformers.Add(ActiveCktElement);
        LINE_ELEMENT:
            Lines.Add(ActiveCktElement);
        FAULTOBJECT:
            Faults.Add(ActiveCktElement);
        FEEDER_ELEMENT:
            Feeders.Add(ActiveCktElement);

        STORAGE_ELEMENT:
            StorageElements.Add(ActiveCktElement);
        PVSYSTEM_ELEMENT:
            PVSystems.Add(ActiveCktElement);
    end;

  // AddDeviceHandle(Handle); // Keep Track of this device result is handle
    AddDeviceHandle(CktElements.ListSize); // Handle is global index into CktElements
    ActiveCktElement.Handle := CktElements.ListSize;

end;

//----------------------------------------------------------------------------
procedure TDSSCircuit.DoResetMeterZones;

begin

 { Do this only if meterzones unlocked .  Normally, Zones will remain unlocked
   so that all changes to the circuit will result in rebuilding the lists}
    if not MeterZonesComputed or not ZonesLocked then
    begin
        if LogEvents then
            LogThisEvent('Resetting Meter Zones');
        EnergyMeterClass.ResetMeterZonesAll;
        MeterZonesComputed := TRUE;
        if LogEvents then
            LogThisEvent('Done Resetting Meter Zones');
    end;

    FreeTopology;

end;

//----------------------------------------------------------------------------
procedure TDSSCircuit.SaveBusInfo;
var
    i: Integer;

begin

{Save existing bus definitions and names for info that needs to be restored}
    SavedBuses := Allocmem(Sizeof(SavedBuses^[1]) * NumBuses);
    SavedBusNames := Allocmem(Sizeof(SavedBusNames^[1]) * NumBuses);

    for i := 1 to NumBuses do
    begin
        SavedBuses^[i] := Buses^[i];
        SavedBusNames^[i] := BusList.get(i);
    end;
    SavedNumBuses := NumBuses;

end;

//----------------------------------------------------------------------------
procedure TDSSCircuit.RestoreBusInfo;

var
    i, j, idx, jdx: Integer;
    pBus: TDSSBus;

begin

// Restore  kV bases, other values to buses still in the list
    for i := 1 to SavedNumBuses do
    begin
        idx := BusList.Find(SavedBusNames^[i]);
        if idx <> 0 then
            with Buses^[idx] do
            begin
                pBus := SavedBuses^[i];
                kvBase := pBus.kVBase;
                x := pBus.x;
                Y := pBus.y;
                CoordDefined := pBus.CoordDefined;
                Keep := pBus.Keep;
               {Restore Voltages in new bus def that existed in old bus def}
                if assigned(pBus.VBus) then
                begin
                    for j := 1 to pBus.NumNodesThisBus do
                    begin
                        jdx := FindIdx(pBus.GetNum(j));  // Find index in new bus for j-th node  in old bus
                        if jdx > 0 then
                            Vbus^[jdx] := pBus.VBus^[j];
                    end;
                end;
            end;
        SavedBusNames^[i] := ''; // De-allocate string
    end;

    if Assigned(SavedBuses) then
        for i := 1 to SavedNumBuses do
            SavedBuses^[i].Free;  // gets rid of old bus voltages, too

    ReallocMem(SavedBuses, 0);
    ReallocMem(SavedBusNames, 0);

end;

//----------------------------------------------------------------------------
procedure TDSSCircuit.ReProcessBusDefs;

// Redo all Buslists, nodelists

var
    CktElementSave: TDSSCktElement;
    i: Integer;

begin
    if LogEvents then
        LogThisEvent('Reprocessing Bus Definitions');

    AbortBusProcess := FALSE;
    SaveBusInfo;  // So we don't have to keep re-doing this
     // Keeps present definitions of bus objects until new ones created

     // get rid of old bus lists
    BusList.Free;  // Clears hash list of Bus names for adding more
    BusList := THashList.Create(NumDevices);  // won't have many more buses than this

    NumBuses := 0;  // Leave allocations same, but start count over
    NumNodes := 0;

     // Now redo all enabled circuit elements
    CktElementSave := ActiveCktElement;
    ActiveCktElement := CktElements.First;
    while ActiveCktElement <> NIL do
    begin
        if ActiveCktElement.Enabled then
            ProcessBusDefs;
        if AbortBusProcess then
            Exit;
        ActiveCktElement := CktElements.Next;
    end;

    ActiveCktElement := CktElementSave;  // restore active circuit element

    for i := 1 to NumBuses do
        Buses^[i].AllocateBusVoltages;
    for i := 1 to NumBuses do
        Buses^[i].AllocateBusCurrents;

    RestoreBusInfo;     // frees old bus info, too
    DoResetMeterZones;  // Fix up meter zones to correspond

    BusNameRedefined := FALSE;  // Get ready for next time
end;

//----------------------------------------------------------------------------
procedure TDSSCircuit.Set_BusNameRedefined(Value: Boolean);
begin
    FBusNameRedefined := Value;

    if Value then
    begin
        Solution.SystemYChanged := TRUE;  // Force Rebuilding of SystemY if bus def has changed
        Control_BusNameRedefined := TRUE;  // So controls will know buses redefined
    end;
end;

//- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
function TDSSCircuit.Get_Losses: Complex;

var
    pdelem: TPDElement;
begin

{Return total losses in all PD Elements}

    pdelem := PDElements.First;
    Result := cZERO;
    while pdelem <> NIL do
    begin
        if pdelem.enabled then
        begin
              {Ignore Shunt Elements}
            if not pdElem.IsShunt then
                Caccum(Result, pdelem.losses);
        end;
        pdelem := PDElements.Next;
    end;

end;
//- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
procedure TDSSCircuit.DebugDump(var F: TextFile);

var
    i, j: Integer;

begin

    Writeln(F, 'NumBuses= ', NumBuses: 0);
    Writeln(F, 'NumNodes= ', NumNodes: 0);
    Writeln(F, 'NumDevices= ', NumDevices: 0);
    Writeln(F, 'BusList:');
    for i := 1 to NumBuses do
    begin
        Write(F, '  ', Pad(BusList.Get(i), 12));
        Write(F, ' (', Buses^[i].NumNodesThisBus: 0, ' Nodes)');
        for j := 1 to Buses^[i].NumNodesThisBus do
            Write(F, ' ', Buses^[i].Getnum(j): 0);
        Writeln(F);
    end;
    Writeln(F, 'DeviceList:');
    for i := 1 to NumDevices do
    begin
        Write(F, '  ', Pad(DeviceList.Get(i), 12));
        ActiveCktElement := CktElements.Get(i);
        if not ActiveCktElement.Enabled then
            Write(F, '  DISABLED');
        Writeln(F);
    end;
    Writeln(F, 'NodeToBus Array:');
    for i := 1 to NumNodes do
    begin
        j := MapNodeToBus^[i].BusRef;
        Write(F, '  ', i: 2, ' ', j: 2, ' (=', BusList.Get(j), '.', MapNodeToBus^[i].NodeNum: 0, ')');
        Writeln(F);
    end;


end;

//- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
procedure TDSSCircuit.InvalidateAllPCElements;

var
    p: TDSSCktElement;

begin

    p := PCElements.First;
    while (p <> NIL) do
    begin
        p.YprimInvalid := TRUE;
        p := PCElements.Next;
    end;

    Solution.SystemYChanged := TRUE;  // Force rebuild of matrix on next solution

end;


// - - ------------------------------------------------------
procedure TDSSCircuit.Set_LoadMultiplier(Value: Double);

begin

    if (Value <> FLoadMultiplier) then   // We may have to change the Y matrix if the load multiplier  has changed
        case Solution.LoadModel of
            ADMITTANCE:
                InvalidateAllPCElements
        else
            {nada}
        end;

    FLoadMultiplier := Value;

end;

procedure TDSSCircuit.TotalizeMeters;

{ Totalize all energymeters in the problem}

var
    pEM: TEnergyMeterObj;
    i: Integer;

begin
    for i := 1 to NumEMRegisters do
        RegisterTotals[i] := 0.;

    pEM := EnergyMeters.First;
    while pEM <> NIL do
        with PEM do
        begin

            for i := 1 to NumEMRegisters do
                RegisterTotals[i] := RegisterTotals[i] + Registers[i] * TotalsMask[i];

            pEM := EnergyMeters.Next;
        end;
end;

function TDSSCircuit.ComputeCapacity: Boolean;
var
    CapacityFound: Boolean;

    function SumSelectedRegisters(const mtrRegisters: TRegisterArray; Regs: pIntegerArray; count: Integer): Double;
    var
        i: Integer;
    begin
        Result := 0.0;
        for i := 1 to count do
        begin
            Result := Result + mtrRegisters[regs^[i]];
        end;
    end;

begin
    Result := FALSE;
    if (EnergyMeters.ListSize = 0) then
    begin
        DoSimpleMsg('Cannot compute system capacity with EnergyMeter objects!', 430);
        Exit;
    end;

    if (NumUeRegs = 0) then
    begin
        DoSimpleMsg('Cannot compute system capacity with no UE resisters defined.  Use SET UEREGS=(...) command.', 431);
        Exit;
    end;

    Solution.Mode := SNAPSHOT;
    LoadMultiplier := CapacityStart;
    CapacityFound := FALSE;

    repeat
        EnergyMeterClass.ResetAll;
        Solution.Solve;
        EnergyMeterClass.SampleAll;
        TotalizeMeters;

           // Check for non-zero in UEregs
        if SumSelectedRegisters(RegisterTotals, UEregs, NumUEregs) <> 0.0 then
            CapacityFound := TRUE;
           // LoadMultiplier is a property ...
        if not CapacityFound then
            LoadMultiplier := LoadMultiplier + CapacityIncrement;
    until (LoadMultiplier > 1.0) or CapacityFound;
    if LoadMultiplier > 1.0 then
        LoadMultiplier := 1.0;
    Result := TRUE;
end;

function TDSSCircuit.Save(Dir: String): Boolean;
{Save the present circuit - Enabled devices only}

var
    i: Integer;
    Success: Boolean;
    CurrDir, SaveDir: String;

begin
    Result := FALSE;

// Make a new subfolder in the present folder based on the circuit name and
// a unique sequence number
    SaveDir := GetCurrentDir;  // remember where to come back to
    Success := FALSE;
    if Length(Dir) = 0 then
    begin
        dir := Name;

        CurrDir := Dir;
        for i := 0 to 999 do  // Find a unique dir name
        begin
            if not DirectoryExists(CurrDir) then
            begin
                if CreateDir(CurrDir) then
                begin
                    SetCurrentDir(CurrDir);
                    Success := TRUE;
                    Break;
                end;
            end;
            CurrDir := dir + Format('%.3d', [i]);
        end;
    end
    else
    begin
        if not DirectoryExists(Dir) then
        begin
            CurrDir := dir;
            if CreateDir(CurrDir) then
            begin
                SetCurrentDir(CurrDir);
                Success := TRUE;
            end;
        end
        else
        begin  // Exists - overwrite
            CurrDir := Dir;
            SetCurrentDir(CurrDir);
            Success := TRUE;
        end;
    end;

    if not Success then
    begin
        DoSimpleMsg('Could not create a folder "' + Dir + '" for saving the circuit.', 432);
        Exit;
    end;

    SavedFileList.Clear;  {This list keeps track of all files saved}

    // Initialize so we will know when we have saved the circuit elements
    for i := 1 to CktElements.ListSize do
        TDSSCktElement(CktElements.Get(i)).HasBeenSaved := FALSE;

    // Initialize so we don't save a class twice
    for i := 1 to DSSClassList.ListSize do
        TDssClass(DSSClassList.Get(i)).Saved := FALSE;

    {Ignore Feeder Class -- gets saved with Energymeters}
   // FeederClass.Saved := TRUE;  // will think this class is already saved

    {Define voltage sources first}
    Success := WriteVsourceClassFile(GetDSSClassPtr('vsource'), TRUE);
    {Write library files so that they will be available to lines, loads, etc}
    {Use default filename=classname}
    if Success then
        Success := WriteClassFile(GetDssClassPtr('wiredata'), '', FALSE);
    if Success then
        Success := WriteClassFile(GetDssClassPtr('cndata'), '', FALSE);
    if Success then
        Success := WriteClassFile(GetDssClassPtr('tsdata'), '', FALSE);
    if Success then
        Success := WriteClassFile(GetDssClassPtr('linegeometry'), '', FALSE);
    // If Success Then Success :=  WriteClassFile(GetDssClassPtr('linecode'),'', FALSE);
    if Success then
        Success := WriteClassFile(GetDssClassPtr('linespacing'), '', FALSE);
    if Success then
        Success := WriteClassFile(GetDssClassPtr('linecode'), '', FALSE);
    if Success then
        Success := WriteClassFile(GetDssClassPtr('xfmrcode'), '', FALSE);
    if Success then
        Success := WriteClassFile(GetDssClassPtr('loadshape'), '', FALSE);
    if Success then
        Success := WriteClassFile(GetDssClassPtr('TShape'), '', FALSE);
    if Success then
        Success := WriteClassFile(GetDssClassPtr('priceshape'), '', FALSE);
    if Success then
        Success := WriteClassFile(GetDssClassPtr('growthshape'), '', FALSE);
    if Success then
        Success := WriteClassFile(GetDssClassPtr('XYcurve'), '', FALSE);
    if Success then
        Success := WriteClassFile(GetDssClassPtr('TCC_Curve'), '', FALSE);
    if Success then
        Success := WriteClassFile(GetDssClassPtr('Spectrum'), '', FALSE);
    if Success then
        Success := SaveFeeders; // Save feeders first
    if Success then
        Success := SaveDSSObjects;  // Save rest ot the objects
    if Success then
        Success := SaveVoltageBases;
    if Success then
        Success := SaveBusCoords;
    if Success then
        Success := SaveMasterFile;


    if Success then
        DoSimpleMsg('Circuit saved in directory: ' + GetCurrentDir, 433)
    else
        DoSimpleMsg('Error attempting to save circuit in ' + GetCurrentDir, 434);
    // Return to Original directory
    SetCurrentDir(SaveDir);

    Result := TRUE;

end;

function TDSSCircuit.SaveDSSObjects: Boolean;
var

    Dss_Class: TDSSClass;
    i: Integer;

begin
    Result := FALSE;

  // Write Files for all populated DSS Classes  Except Solution Class
    for i := 1 to DSSClassList.ListSize do
    begin
        Dss_Class := DSSClassList.Get(i);
        if (DSS_Class = SolutionClass) or Dss_Class.Saved then
            Continue;   // Cycle to next
            {use default filename=classname}
        if not WriteClassFile(Dss_Class, '', (DSS_Class is TCktElementClass)) then
            Exit;  // bail on error
        DSS_Class.Saved := TRUE;
    end;

    Result := TRUE;

end;

function TDSSCircuit.SaveVoltageBases: Boolean;
var
    F: TextFile;
    i: Integer;
begin

    Result := FALSE;
    try
        AssignFile(F, 'BusVoltageBases.DSS');
        Rewrite(F);

        for i := 1 to NumBuses do
            if Buses^[i].kVBase > 0.0 then
                Writeln(F, Format('SetkVBase Bus=%s  kvln=%.7g ', [BusList.Get(i), Buses^[i].kVBase]));

        CloseFile(F);
        Result := TRUE;
    except
        On E: Exception do
            DoSimpleMsg('Error Saving BusVoltageBases File: ' + E.Message, 43501);
    end;

end;

function TDSSCircuit.SaveMasterFile: Boolean;

var
    F: TextFile;
    i: Integer;

begin
    Result := FALSE;
    try
        AssignFile(F, 'Master.DSS');
        Rewrite(F);

        Writeln(F, 'Clear');
        Writeln(F, 'New Circuit.' + Name);
        Writeln(F);
        if PositiveSequence then
            Writeln(F, 'Set Cktmodel=Positive');
        if DuplicatesAllowed then
            Writeln(F, 'set allowdup=yes');
        Writeln(F);

      // Write Redirect for all populated DSS Classes  Except Solution Class
        for i := 1 to SavedFileList.Count do
        begin
            Writeln(F, 'Redirect ', SavedFileList.Strings[i - 1]);
        end;

        Writeln(F, 'MakeBusList');
        Writeln(F, 'Redirect BusVoltageBases.dss  ! set voltage bases');

        if FileExists('buscoords.dss') then
        begin
            Writeln(F, 'Buscoords buscoords.dss');
        end;

        CloseFile(F);
        Result := TRUE;
    except
        On E: Exception do
            DoSimpleMsg('Error Saving Master File: ' + E.Message, 435);
    end;

end;

function TDSSCircuit.SaveFeeders: Boolean;
var
    i: Integer;
    SaveDir, CurrDir: String;
    Meter: TEnergyMeterObj;
begin

    Result := TRUE;
{Write out all energy meter  zones to separate subdirectories}
    SaveDir := GetCurrentDir;
    for i := 1 to EnergyMeters.ListSize do
    begin
        Meter := EnergyMeters.Get(i); // Recast pointer
        CurrDir := Meter.Name;
        if DirectoryExists(CurrDir) then
        begin
            SetCurrentDir(CurrDir);
            Meter.SaveZone(CurrDir);
            SetCurrentDir(SaveDir);
        end
        else
        begin
            if CreateDir(CurrDir) then
            begin
                SetCurrentDir(CurrDir);
                Meter.SaveZone(CurrDir);
                SetCurrentDir(SaveDir);
            end
            else
            begin
                DoSimpleMsg('Cannot create directory: ' + CurrDir, 436);
                Result := FALSE;
                SetCurrentDir(SaveDir);  // back to whence we came
                Break;
            end;
        end;
    end;  {For}

end;

function TDSSCircuit.SaveBusCoords: Boolean;
var
    F: TextFile;
    i: Integer;
begin

    Result := FALSE;

    try
        AssignFile(F, 'BusCoords.dss');
        Rewrite(F);


        for i := 1 to NumBuses do
        begin
            if Buses^[i].CoordDefined then
                Writeln(F, CheckForBlanks(BusList.Get(i)), Format(', %-g, %-g', [Buses^[i].X, Buses^[i].Y]));
        end;

        Closefile(F);

        Result := TRUE;

    except
        On E: Exception do
            DoSimpleMsg('Error creating Buscoords.dss.', 437);
    end;

end;

procedure TDSSCircuit.ReallocDeviceList;

var
    TempList: THashList;
    i: Integer;

begin
{Reallocate the device list to improve the performance of searches}
    if LogEvents then
        LogThisEvent('Reallocating Device List');
    TempList := THashList.Create(2 * NumDevices);

    for i := 1 to DeviceList.ListSize do
    begin
        Templist.Add(DeviceList.Get(i));
    end;

    DeviceList.Free; // Throw away the old one.
    Devicelist := TempList;

end;

procedure TDSSCircuit.Set_CaseName(const Value: String);
begin
    FCaseName := Value;
    CircuitName_ := Value + '_';
end;

function TDSSCircuit.Get_Name: String;
begin
    Result := LocalName;
end;

function TDSSCircuit.GetBusAdjacentPDLists: TAdjArray;
begin
    if not Assigned(BusAdjPD) then
        BuildActiveBusAdjacencyLists(BusAdjPD, BusAdjPC);
    Result := BusAdjPD;
end;

function TDSSCircuit.GetBusAdjacentPCLists: TAdjArray;
begin
    if not Assigned(BusAdjPC) then
        BuildActiveBusAdjacencyLists(BusAdjPD, BusAdjPC);
    Result := BusAdjPC;
end;

function TDSSCircuit.GetTopology: TCktTree;
var
    i: Integer;
    elem: TDSSCktElement;
begin
    if not assigned(Branch_List) then
    begin
    {Initialize all Circuit Elements and Buses to not checked, then build a new tree}
        elem := CktElements.First;
        while assigned(elem) do
        begin
            elem.Checked := FALSE;
            for i := 1 to elem.Nterms do
                elem.Terminals^[i].Checked := FALSE;
            elem.IsIsolated := TRUE; // till proven otherwise
            elem := CktElements.Next;
        end;
        for i := 1 to NumBuses do
            Buses^[i].BusChecked := FALSE;
        Branch_List := GetIsolatedSubArea(Sources.First, TRUE);  // calls back to build adjacency lists
    end;
    Result := Branch_List;
end;

procedure TDSSCircuit.FreeTopology;
begin
    if Assigned(Branch_List) then
        Branch_List.Free;
    Branch_List := NIL;
    if Assigned(BusAdjPC) then
        FreeAndNilBusAdjacencyLists(BusAdjPD, BusAdjPC);
end;

procedure TDSSCircuit.ClearBusMarkers;
var
    i: Integer;
begin
    for i := 0 to BusMarkerList.count - 1 do
        TBusMarker(BusMarkerList.Items[i]).Free;
    BusMarkerList.Clear;
end;

{====================================================================}
{ TBusMarker }
{====================================================================}

constructor TBusMarker.Create;
begin
    inherited;
    BusName := '';
    AddMarkerColor := 1; // clBlack; TEMc
    AddMarkerCode := 4;
    AddMarkerSize := 1;
end;

destructor TBusMarker.Destroy;
begin
    BusName := '';
    inherited;
end;


end.
