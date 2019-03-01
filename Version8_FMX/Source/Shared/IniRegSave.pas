unit IniRegSave;

{
  ----------------------------------------------------------
  Copyright (c) 2008-2015, Electric Power Research Institute, Inc.
  All rights reserved.
  ----------------------------------------------------------
}

{ Simple unit to create an Ini file equivalent in the registry.
  By default, creates a key under HKEY_CURRENT_USER

  Typically, you will want to move the key under 'Software' by creating as

  MyIniFile := TIniRegSave.Create('\Software\myprogramname');

  But it'll work anywhere.

}

interface

uses
    Registry;

type

    TIniRegSave = class(TObject)
        FSection: String;
        Fname: String;
        FIniFile: TRegIniFile;

    PRIVATE
        procedure Set_FSection(const Value: String);
    { Private declarations }

    PUBLIC
    { Public declarations }

        property Section: String READ FSection WRITE Set_FSection;

        procedure ClearSection;

        procedure WriteBool(const key: String; value: Boolean);
        procedure WriteInteger(const key: String; value: Integer);
        procedure WriteString(const key: String; value: String);

        function ReadBool(const key: String; default: Boolean): Boolean;
        function ReadInteger(const key: String; default: Integer): Integer;
        function ReadString(const key: String; const default: String): String;

        constructor Create(const Name: String);
        destructor Destroy; OVERRIDE;
    end;


implementation


constructor TIniRegSave.Create(const Name: String);
begin
    FName := Name;
    FIniFile := TRegIniFile.Create(Name);
    FSection := 'MainSect';
end;

destructor TIniRegSave.Destroy;
begin
    inherited;

end;

function TIniRegSave.ReadBool(const key: String; default: Boolean): Boolean;
begin
    Result := FiniFile.ReadBool(Fsection, key, default);
end;

function TIniRegSave.ReadInteger(const key: String; Default: Integer): Integer;
begin
    Result := FiniFile.ReadInteger(Fsection, key, default);
end;

function TIniRegSave.ReadString(const key: String; const Default: String): String;
begin
    Result := FiniFile.ReadString(Fsection, key, default);
end;

procedure TIniRegSave.Set_FSection(const Value: String);
begin
    FSection := Value;
end;

procedure TIniRegSave.WriteBool(const key: String; value: Boolean);
begin
    FiniFile.WriteBool(FSection, key, value);
end;

procedure TIniRegSave.WriteInteger(const key: String; value: Integer);
begin
    FiniFile.WriteInteger(FSection, key, value);
end;

procedure TIniRegSave.WriteString(const key: String; value: String);
begin
    FiniFile.WriteString(FSection, key, value);
end;

procedure TIniRegSave.ClearSection;
begin
    FiniFile.EraseSection(FSection);
end;

end.
