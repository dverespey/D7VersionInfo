unit VersionInfo;

interface

uses
  Windows, Messages, SysUtils, Classes;

type
  TVersionInfo = class(TComponent)
  private
    { Private declarations }
    //Only visible to this unit
    fErrMsg: string;
    fVF_Path: string;
  protected
    { Protected declarations }
    //Only visible to this unit and the sub-classes of it
  public
    { Public declarations }
    //Visible from anywhere within the application
    property ErrMsg: string
        read fErrMsg;
    constructor Create (AOwner: TComponent); override;
    destructor Destroy; override;
    function GetVersion: String;
  published
    { Published declarations }
    property FilePath: string
        read fVF_Path
        write fVF_Path;
  end;

procedure Register;

implementation

type
  TLangInfoBuffer = array [1..4] of SmallInt;

constructor TVersionInfo.Create(AOwner: TComponent);
begin
	inherited Create(AOwner);
  fVF_Path := '';
  fErrMsg := '';
end;

destructor TVersionInfo.Destroy;
begin
  fErrMsg := '';
  fVF_Path := '';
	inherited destroy;
end;

procedure Register;
begin
  RegisterComponents('NUMMI Tools', [TVersionInfo]);
end;

function TVersionInfo.GetVersion: String;
Var
  VInfoSize, DetSize: DWord;
  pVInfo, pDetail: Pointer;
  pLangInfo: ^TLangInfoBuffer;
  strLangId: string;
Begin
  fErrMsg := '';
  VInfoSize := GetFileVersionInfoSize (PChar (fVF_Path), DetSize);
  If VInfoSize > 0 Then
  Begin
    GetMem(pVInfo, VInfoSize);
    try
      try
        GetFileVersionInfo (PChar(fVF_Path), 0, VInfoSize, pVInfo);
        VerQueryValue(pVInfo, '\VarFileInfo\Translation', Pointer(pLangInfo), DetSize);
        strLangId := IntToHex (SmallInt (pLangInfo^ [1]), 4) +
             IntToHex (SmallInt (pLangInfo^ [2]), 4);
        strLangId := '\StringFileInfo\' + strLangId;
        VerQueryValue(pVInfo, PChar(strLangId + '\FileVersion'),
             pDetail, DetSize);
        fErrMsg := '';
        Result := PChar(pDetail);
      except
        on e:exception do
        Begin
          fErrMsg := e.Message;
          result := 'ERROR';
        end;
      end;    //Try...except
    finally
      FreeMem(pVInfo);
    End;    //Try...Finally
  End;
End;       //GetVersionInfo

end.
