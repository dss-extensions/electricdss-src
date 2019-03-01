unit Diakoptics;

interface

uses
    Circuit,
    Solution,
    DSSGlobals,
    SysUtils,
    DSSClassDefs;

function Solve_Diakoptics(): Integer;
procedure ADiakoptics_Tearing();
procedure ADiakopticsInit();
procedure Calc_C_Matrix(PLinks: PString; NLinks: Integer);
procedure Calc_ZCT();
procedure Calc_ZCC();

implementation

uses
    ExecHelper,
    Executive,
    ParserDel,
    YMatrix,
    KLUSolve,
    Ucomplex;

function Solve_Diakoptics(): Integer;
begin
  {Space left empty to implement the simplified Diakoptics algorithm}

    Result := 0;
end;


{*******************************************************************************
*              Calculates the Connections matrix ZCC in the                    *
*                      contours-contours domain                                *
*******************************************************************************}
procedure Calc_ZCC();
begin
    ActiveActor := 1;
    with ActiveCircuit[ActiveActor], ActiveCircuit[ActiveActor].Solution do
    begin
  {Space left empty to implement the simplified Diakoptics algorithm}
    end;
end;


{*******************************************************************************
*                   Calculates the Lateral matrix ZCT in                       *
*                         contours-trees domain                                *
*******************************************************************************}

  {Probably to be removed}
procedure Calc_ZCT();
var
    i, j,
    Ret,
    LIdx: Integer;
    VContours,
    VZCT: array of Complex;
    temp: String;
    myFile: TextFile;         // For debugging
begin
    ActiveActor := 1;
    with ActiveCircuit[ActiveActor], ActiveCircuit[ActiveActor].Solution do
    begin
        setlength(ZCT, length(Contours));
        for LIdx := 0 to High(ZCT) do
        begin
            setlength(ZCT[LIdx], length(Contours[LIdx]));
            setlength(VContours, length(Contours[LIdx]));
            setlength(VZCT, length(Contours[LIdx]));
            for i := 0 to High(Contours[LIdx]) do
                VContours[i] := Contours[LIdx][i];      // Moves the contour column
      // Solves the vector
            Ret := SolveSparseSet(hY, @VZCT[0], @VContours[0]);
      //  Stores the vector
            for i := 0 to High(VZCT) do
                ZCT[LIdx][i] := VZCT[i];
        end;
{    // For Debugging
    temp              :=  GetCurrentDir;
    AssignFile(myFile, temp  + '\ZCT.txt');
    ReWrite(myFile);
    for j := 0 to High(VZCT) do
    Begin
      for i := 0 to High(ZCT) do
        Write(myFile,(floattostr(ZCT[i][j].re) + '+' + floattostr(ZCT[i][j].im) + ' '));
      WriteLn(myFile,'');
    End;
    CloseFile(myFile);  }
    end;
end;

{*******************************************************************************
*                   Calculates the contours matrix based                       *
*             on the location in the graph of the link branches                *
*******************************************************************************}
procedure Calc_C_Matrix(PLinks: PString; NLinks: Integer);
var
    LIdx, k, l,
    j,
    i: Integer;
    Elem_Buses,
    Node_Names: array of String;
    temp: String;

    myFile: TextFile;         // For debugging
begin
    ActiveActor := 1;
    with ActiveCircuit[ActiveActor] do
    begin
        setlength(Contours, 0);
        setlength(Elem_Buses, 2);
        setlength(Node_Names, 0);
        for i := 1 to NumNodes do
        begin
            setlength(Node_Names, (length(Node_Names) + 1));
            with MapNodeToBus^[i] do
                Node_Names[High(Node_names)] := Format('%s.%-d', [lowercase(BusList.Get(Busref)), NodeNum]);
        end;

        for LIdx := 1 to (NLinks - 1) do
        begin
            inc(PLinks);                  // Pointing to the Next link branch (starting in 1)

            i := SetElementActive(String(PLinks^));
      // Gest the names of the buses fot this PDElement
            for i := 1 to ActiveCktElement.Nterms do
            begin
                Elem_Buses[i - 1] := ActiveCktElement.GetBus(i);
                j := ansipos('.', Elem_Buses[i - 1]);
                if j <> 0 then
                    Elem_Buses[i - 1] := copy(Elem_Buses[i - 1], 0, j);
            end;
      //  Marks the connection point in the contours matrix
            for l := 1 to ActiveCktElement.NPhases do
            begin
                setlength(Contours, (length(Contours) + 1));
                setlength(Contours[High(Contours)], length(Node_Names));
                for i := 0 to 1 do
                begin
                    temp := Elem_Buses[i] + inttostr(l);
                    for j := 0 to High(Node_Names) do
                    begin
                        k := ansipos(temp, Node_Names[j]);
                        if k <> 0 then
                        begin
                            if i = 0 then
                                Contours[High(Contours)][j].re := 1
                            else
                                Contours[High(Contours)][j].re := -1;
                            Contours[High(Contours)][j].im := 0;
                        end;
                    end;
                end;
            end;

        end;

    // For Debugging
{    temp              :=  GetCurrentDir;
    AssignFile(myFile, temp  + '\Contours.txt');
    ReWrite(myFile);
    for j := 0 to High(Node_Names) do
    Begin
      Write(myFile,Node_Names[j] + ' ');
      for i := 0 to High(Contours) do
        Write(myFile,(floattostr(Contours[i][j].re) + '+' + floattostr(Contours[i][j].im) + ' '));
      WriteLn(myFile,'');
    End;
    CloseFile(myFile);  }
    end;
end;

{*******************************************************************************
*           Tears the system using considering the number of                   *
*           available CPUs as reference                                        *
*******************************************************************************}
procedure ADiakoptics_Tearing();
var
    Prev_Mode,                              // Stores the previous solution mode
    Num_Ckts: Integer;                  // Stores the number of Sub-Circuits created
begin
    with ActiveCircuit[ActiveActor].Solution do
    begin
        Num_Ckts := ActiveCircuit[ActiveActor].Tear_Circuit();
        Prev_mode := Dynavars.SolutionMode;
        Dynavars.SolutionMode := 0;          // Shapshot mode
        solve(ActiveActor);
        ActiveCircuit[ActiveActor].Save_SubCircuits();
        Dynavars.SolutionMode := Prev_mode;  // Goes back to the previous solution mode
        ActiveCircuit[1].Num_SubCkts := Num_Ckts;
        GlobalResult := 'Sub-Circuits Created: ' + inttostr(Num_Ckts);
    end;
end;

{*******************************************************************************
*            Generates the subsystems, actors and memory space                 *
*                     For using the A-Diakoptics parallelism                   *
*******************************************************************************}
procedure ADiakopticsInit();
var
    DIdx,
    Diak_Actors: Integer;
    Dir, Proj_Dir,
    FileRoot: String;
    Links: array of String;                        // List of the Link Branches
begin
    ActiveActor := 1;
    ActiveCircuit[1].Num_SubCkts := CPU_Cores - 1;
    ADiakoptics_Tearing();
    Diak_Actors := ActiveCircuit[1].Num_SubCkts + 1;
  // Saves the Link Branch list locally
    setlength(Links, length(ActiveCircuit[1].Link_Branches));
    for DIdx := 0 to High(Links) do
        Links[DIdx] := ActiveCircuit[1].Link_Branches[DIdx];

  // Clears everything to craete the actors and compile the subsystems
    DSSExecutive.ClearAll;
    Fileroot := GetCurrentDir;    //  Gets the current directory

  // Compiles the interconnected Circuit for further calculations on actor 1
    ActiveActor := 1;
    Proj_Dir := 'compile "' + Fileroot + '\Torn_Circuit\master_interconnected.dss"';
    DssExecutive.Command := Proj_Dir;

  // Creates the other actors
    for DIdx := 2 to Diak_Actors do
    begin
        inc(NumOfActors);
        ActiveActor := NumOfActors;
        ActorCPU[ActiveActor] := ActiveActor - 1;
        DSSExecutive := TExecutive.Create;  // Make a DSS object
        Parser[ActiveActor] := TParser.Create;
        DSSExecutive.CreateDefaultDSSItems;
        if DIdx = 2 then
            Dir := ''
        else
            Dir := 'zone_' + inttostr(DIdx - 1) + '\';
        Proj_Dir := 'compile "' + Fileroot + '\Torn_Circuit\' + Dir + 'master.dss"';
        DssExecutive.Command := Proj_Dir;
    end;
  // Calculates the contours matrix
    Calc_C_Matrix(@Links[0], length(Links));

  // Opens the link branches in the interconnected Circuit and recalculates the YBus
    for DIdx := 1 to High(Links) do
    begin
        Proj_Dir := 'open ' + Links[DIdx] + ' term=1';
    end;
    Ymatrix.BuildYMatrix(SERIESONLY, FALSE, ActiveActor);

  // Calculates the connection matrix and the lateral matrices
    Calc_ZCT();
    Calc_ZCC();

    ActiveActor := 1;
    GlobalResult := 'Sub-Circuits Created: ' + inttostr(Diak_Actors - 1);
end;

end.
