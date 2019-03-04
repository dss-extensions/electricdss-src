unit ImplCmathLib;

{
  ----------------------------------------------------------
  Copyright (c) 2008-2015, Electric Power Research Institute, Inc.
  All rights reserved.
  ----------------------------------------------------------
}
{$WARN SYMBOL_PLATFORM OFF}

interface

uses
    ComObj,
    ActiveX,
    OpenDSSengine_TLB,
    StdVcl,
    Variants;

type
    TCmathLib = class(TAutoObject, ICmathLib)
    PROTECTED
        function Get_cmplx(RealPart, ImagPart: Double): Olevariant; SAFECALL;
        function Get_cabs(realpart, imagpart: Double): Double; SAFECALL;
        function Get_cdang(RealPart, ImagPart: Double): Double; SAFECALL;
        function Get_ctopolardeg(RealPart, ImagPart: Double): Olevariant; SAFECALL;
        function Get_pdegtocomplex(magnitude, angle: Double): Olevariant; SAFECALL;
        function Get_cmul(a1, b1, a2, b2: Double): Olevariant; SAFECALL;
        function Get_cdiv(a1, b1, a2, b2: Double): Olevariant; SAFECALL;

    end;

implementation

uses
    ComServ,
    Ucomplex;

function TCmathLib.Get_cmplx(RealPart, ImagPart: Double): Olevariant;
begin
    Result := VarArrayCreate([0, 1], varDouble);
    Result[0] := RealPart;
    Result[1] := ImagPart;
end;

function TCmathLib.Get_cabs(realpart, imagpart: Double): Double;
begin
    Result := cabs(cmplx(realpart, imagpart));
end;

function TCmathLib.Get_cdang(RealPart, ImagPart: Double): Double;
begin
    Result := cdang(cmplx(realpart, imagpart));
end;

function TCmathLib.Get_ctopolardeg(RealPart, ImagPart: Double): Olevariant;
var
    TempPolar: polar;
begin
    Result := VarArrayCreate([0, 1], varDouble);
    TempPolar := ctopolardeg(cmplx(RealPart, ImagPart));
    Result[0] := TempPolar.mag;
    Result[1] := TempPolar.ang;
end;

function TCmathLib.Get_pdegtocomplex(magnitude, angle: Double): Olevariant;
var
    cTemp: Complex;
begin
    Result := VarArrayCreate([0, 1], varDouble);
    cTemp := pdegtocomplex(magnitude, angle);
    Result[0] := cTemp.re;
    Result[1] := cTemp.im;
end;

function TCmathLib.Get_cmul(a1, b1, a2, b2: Double): Olevariant;
var
    cTemp: Complex;
begin
    Result := VarArrayCreate([0, 1], varDouble);
    cTemp := cmul(cmplx(a1, b1), cmplx(a2, b2));
    Result[0] := cTemp.re;
    Result[1] := cTemp.im;
end;

function TCmathLib.Get_cdiv(a1, b1, a2, b2: Double): Olevariant;
var
    cTemp: Complex;
begin
    Result := VarArrayCreate([0, 1], varDouble);
    cTemp := cdiv(cmplx(a1, b1), cmplx(a2, b2));
    Result[0] := cTemp.re;
    Result[1] := cTemp.im;
end;

initialization
    TAutoObjectFactory.Create(ComServer, TCmathLib, Class_CmathLib,
        ciInternal, tmApartment);
end.
