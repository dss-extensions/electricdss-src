unit ReduceAlgs;

{
  ----------------------------------------------------------
  Copyright (c) 2008-2015, Electric Power Research Institute, Inc.
  All rights reserved.
  ----------------------------------------------------------
}

{$IFDEF FPC}{$MODE Delphi}{$ENDIF}

{Reduction Algorithms}

{Primarily called from EnergyMeter}

interface

uses
    CktTree;

procedure DoReduceDefault(var BranchList: TCktTree);
procedure DoReduceStubs(var BranchList: TCktTree);
procedure DoReduceDangling(var BranchList: TCktTree);
procedure DoReduceTapEnds(var BranchList: TCktTree);
procedure DoBreakLoops(var BranchList: TCktTree);
procedure DoMergeParallelLines(var BranchList: TCktTree);
procedure DoReduceSwitches(var Branchlist: TCktTree);


implementation

uses
    Line,
    Utilities,
    DSSGlobals,
    Load,
    uComplex,
    ParserDel,
    CktElement,
    sysutils;

procedure DoMergeParallelLines(var BranchList: TCktTree);
{Merge all lines in this zone that are marked in parallel}

var
    LineElement: TLineObj;

begin
    if BranchList <> NIL then
    begin
        BranchList.First;
        LineElement := BranchList.GoForward; // Always keep the first element
        while LineElement <> NIL do
        begin
            if BranchList.PresentBranch.IsParallel then
            begin
               {There will always be two lines in parallel.  The first operation will disable the second}
                if LineElement.Enabled then
                    LineElement.MergeWith(TLineObj(BranchList.PresentBranch.LoopLineObj), FALSE);  // Guaranteed to be a line
            end;
            LineElement := BranchList.GoForward;
        end;
    end;
end;

procedure DoBreakLoops(var BranchList: TCktTree);

{Break loops}
var
    LineElement: TLineObj;

begin
    if BranchList <> NIL then
    begin
        BranchList.First;
        LineElement := BranchList.GoForward; // Always keep the first element
        while LineElement <> NIL do
        begin
            if BranchList.PresentBranch.IsLoopedHere then
            begin
               {There will always be two lines in the loop.  The first operation will disable the second}
                if LineElement.Enabled then
                    TLineObj(BranchList.PresentBranch.LoopLineObj).Enabled := FALSE; // Disable the other
            end;
            LineElement := BranchList.GoForward;
        end;
    end;
end;


procedure DoReduceTapEnds(var BranchList: TCktTree);
(*Var
   pLineElem1, pLineElem2:TLineObj;
   ToBusRef:Integer;
   AngleTest:Double;
   ParentNode:TCktTreeNode;
*)
begin

end;


procedure DoReduceDangling(var BranchList: TCktTree);
var
    pLineElem1: TDSSCktElement;
    ToBusRef: Integer;
begin
    if BranchList <> NIL then
    begin
     {Let's throw away all dangling end branches}
        BranchList.First;
        pLineElem1 := BranchList.GoForward; // Always keep the first element

        while pLineElem1 <> NIL do
        begin

            if IsLineElement(pLineElem1) then
                with  BranchList.PresentBranch do
                begin

             {If it is at the end of a section and has no load,cap, reactor,
             or coordinate, just throw it away}
                    if IsDangling then
                    begin
                        ToBusRef := ToBusReference;  // only access this property once!
                        if ToBusRef > 0 then
                            with ActiveCircuit.Buses^[ToBusRef] do
                                if not (Keep) then
                                    pLineElem1.Enabled := FALSE;
                    end; {IF}
                end;  {If-With}
            pLineElem1 := BranchList.GoForward;
        end;
    end;

end;


procedure DoReduceStubs(var BranchList: TCktTree);
{Eliminate short stubs and merge with lines on either side}
var
    LineElement1, LineElement2: TLineObj;
    LoadElement: TLoadObj;
    ParentNode: TCktTreeNode;

begin
    if BranchList <> NIL then
    begin  {eliminate really short lines}
      {First, flag all elements that need to be merged}
        LineElement1 := BranchList.First;
        LineElement1 := BranchList.GoForward; // Always keep the first element
        while LineElement1 <> NIL do
        begin
            if IsLineElement(LineElement1) then
            begin
                if IsStubLine(LineElement1) then
                    LineElement1.Flag := TRUE   {Too small: Mark for merge with something}
                else
                    LineElement1.Flag := FALSE;
            end; {IF}
            LineElement1 := BranchList.GoForward;
        end; {WHILE}

        LineElement1 := BranchList.First;
        LineElement1 := BranchList.GoForward; // Always keep the first element
        while LineElement1 <> NIL do
        begin
            if LineElement1.Flag then  // Merge this element out
            begin
                with BranchList.PresentBranch do
                begin
                    if (NumChildBranches = 0) and (NumShuntObjects = 0) then
                        LineElement1.Enabled := FALSE     // just discard it
                    else
                    if (NumChildBranches = 0) or (NumChildBranches > 1) then
                 {Merge with Parent and move loads on parent to To node}
                    begin
                        ParentNode := ParentBranch;
                        if ParentNode <> NIL then
                        begin
                            if ParentNode.NumChildBranches = 1 then   // only works for in-line
                                if not ActiveCircuit.Buses^[ParentNode.ToBusReference].Keep then
                                begin
                             {Let's consider merging}
                                    LineElement2 := ParentNode.CktObject;
                                    if LineElement2.enabled then  // Check to make sure it hasn't been merged out
                                        if IsLineElement(LineElement2) then
                                            if LineElement2.MergeWith(LineElement1, TRUE) then {Move any loads to ToBus Reference of downline branch}
                                                if ParentNode.NumShuntObjects > 0 then
                                                begin
                                   {Redefine bus connection for PC elements hanging on the bus that is eliminated}
                                                    LoadElement := ParentNode.FirstShuntObject;
                                                    while LoadElement <> NIL do
                                                    begin
                                                        Parser.CmdString := 'bus1="' + ActiveCircuit.BusList.Get(ToBusReference) + '"';
                                                        LoadElement.Edit;
                                                        LoadElement := ParentNode.NextShuntObject;
                                                    end;  {While}
                                                end; {IF}
                                end; {IF}
                        end; {IF ParentNode}
                    end
                    else
                    begin{Merge with child}
                        if not ActiveCircuit.Buses^[ToBusReference].Keep then
                        begin
                       {Let's consider merging}
                            LineElement2 := FirstChildBranch.CktObject;
                            if IsLineElement(LineElement2) then
                                if LineElement2.MergeWith(LineElement1, TRUE) then
                                    if FirstChildBranch.NumShuntObjects > 0 then
                                    begin
                               {Redefine bus connection to upline bus}
                                        LoadElement := FirstChildBranch.FirstShuntObject;
                                        while LoadElement <> NIL do
                                        begin
                                            Parser.CmdString := 'bus1="' + ActiveCircuit.BusList.Get(FromBusReference) + '"';
                                            LoadElement.Edit;
                                            LoadElement := FirstChildBranch.NextShuntObject;
                                        end;  {While}
                                    end; {IF}
                        end; {IF not}
                    end; {ELSE}
                end;
            end;
            LineElement1 := BranchList.GoForward;
        end;

    end;
end;

procedure DoReduceSwitches(var Branchlist: TCktTree);

{Merge switches in with lines or delete if dangling}

var
    LineElement1, LineElement2: TLineObj;
begin

    if BranchList <> NIL then
    begin

        LineElement1 := BranchList.First;
        LineElement1 := BranchList.GoForward; // Always keep the first element
        while LineElement1 <> NIL do
        begin

            if LineElement1.Enabled then   // maybe we threw it away already
                if IsLineElement(LineElement1) then
                    if LineElement1.IsSwitch then
                        with BranchList.PresentBranch do
             {see if eligble for merging}
                            case NumChildBranches of
                                0: {Throw away if dangling}
                                    if NumShuntObjects = 0 then
                                        LineElement1.Enabled := FALSE;

                                1:
                                    if NumShuntObjects = 0 then
                                        if not ActiveCircuit.Buses^[ToBusReference].Keep then
                                        begin
                     {Let's consider merging}
                                            LineElement2 := FirstChildBranch.CktObject;
                                            if IsLineElement(LineElement2) then
                                                if not LineElement2.IsSwitch then
                                                    LineElement2.MergeWith(LineElement1, TRUE){Series Merge}
                                        end;
                            else {Nada}
                            end;

            LineElement1 := BranchList.GoForward;
        end;
    end;

end;

procedure DoReduceDefault(var BranchList: TCktTree);

var
    LineElement1, LineElement2: TLineObj;
begin
    if BranchList <> NIL then
    begin

     {Now merge remaining lines}

        LineElement1 := BranchList.First;
        LineElement1 := BranchList.GoForward; // Always keep the first element
        while LineElement1 <> NIL do
        begin

            if IsLineElement(LineElement1) then
                if not LineElement1.IsSwitch then
                    if LineElement1.Enabled then   // maybe we threw it away already
                        with BranchList.PresentBranch do
                        begin
                 {see if eligble for merging}
                            if NumChildBranches = 1 then
                                if NumShuntObjects = 0 then
                                    if not ActiveCircuit.Buses^[ToBusReference].Keep then
                                    begin
                     {Let's consider merging}
                                        LineElement2 := FirstChildBranch.CktObject;

                                        if IsLineElement(LineElement2) then
                                            if not LineElement2.IsSwitch then
                                                LineElement2.MergeWith(LineElement1, TRUE){Series Merge}
                                    end;

                        end;

            LineElement1 := BranchList.GoForward;
        end;
    end;

end;

end.
