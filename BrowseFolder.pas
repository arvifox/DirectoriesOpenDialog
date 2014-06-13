unit BrowseFolder;

interface

uses
  Windows, SysUtils, Classes, ShellAPI, ShlObj;

type
  TBrowseFolder = class(TComponent)
  private
    { Private declarations }
  protected
    { Protected declarations }
  public
    { Public declarations }
    StartDir: string;
    ResultDir: string;
    TitleName: string;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    function Execute(phwnd: HWND): boolean;
  published
    { Published declarations }

  end;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('Samples', [TBrowseFolder]);
end;

{ TBrowseFolder }

constructor TBrowseFolder.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TBrowseFolder.Destroy;
begin
  inherited Destroy;
end;

function BrowseCallbackProc(hwnd: HWND; uMsg: UINT; lParam: integer;
                            lpData: integer): integer; stdcall;
begin
  // get message
  if uMsg=BFFM_INITIALIZED then
    // send message to cheange dir
    SendMessage(hwnd,BFFM_SETSELECTION,1,lpData);
  result:=0;
end;

function TBrowseFolder.Execute(phwnd: HWND): boolean;
var 
  pp: string;
  lpItemID : PItemIDList;
  BrowseInfo : TBrowseInfo;
  DisplayName : array[0..MAX_PATH] of char;
  TempPath : array[0..MAX_PATH] of char;
begin
  result:=false;
  FillChar(BrowseInfo, sizeof(TBrowseInfo), #0);
  // init
  BrowseInfo.hwndOwner := phwnd;
  BrowseInfo.pszDisplayName := @DisplayName;
  BrowseInfo.lpszTitle := PChar(TitleName);
  BrowseInfo.ulFlags := BIF_RETURNONLYFSDIRS + BIF_NEWDIALOGSTYLE;
  BrowseInfo.lpfn:=@BrowseCallbackProc;
  pp:=StartDir;
  lstrcpy(TempPath,@pp[1]);
  BrowseInfo.lParam:=integer(@TempPath);
  // show window
  lpItemID := SHBrowseForFolder(BrowseInfo);
  if lpItemId <> nil then
  begin
    SHGetPathFromIDList(lpItemID, TempPath);
    SetLength(ResultDir,lstrlen(TempPath));
    lstrcpy(@ResultDir[1],TempPath);
    GlobalFreePtr(lpItemID); 
    result:=true;
  end;
end;

end.
