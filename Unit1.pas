unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, Winapi.imm,
  Vcl.Forms;

type
  TForm1 = class(TForm)
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
  public
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

const
  VK_LALT = $A4;
  VK_RALT = $A5;

type
  PKBDLLHOOKSTRUCT = ^KBDLLHOOKSTRUCT;

  KBDLLHOOKSTRUCT = record
    vkCode: DWORD;
    scanCode: DWORD;
    flags: DWORD;
    time: DWORD;
    dwExtraInfo: ULONG_PTR;
  end;

var
  Hook: HHOOK;
  LAltDown, RAltDown, AltDirty: Boolean;

// alt-ime-ahk IME_SET() 相当
procedure SetIme(Open: Boolean);
var
  GUIInfo: TGUIThreadInfo;
  TargetWnd, ImeWnd: THandle;
begin
  ZeroMemory(@GUIInfo, SizeOf(GUIInfo));
  GUIInfo.cbSize := SizeOf(GUIInfo);
  if GetGUIThreadInfo(0, GUIInfo) and (GUIInfo.hwndFocus <> 0) then
    TargetWnd := GUIInfo.hwndFocus
  else
    TargetWnd := GetForegroundWindow;

  ImeWnd := ImmGetDefaultIMEWnd(TargetWnd);
  if ImeWnd <> 0 then
    SendMessage(ImeWnd, WM_IME_CONTROL, $006, Ord(Open)); // IMC_SETOPENSTATUS
end;

function HookCallback(Code: Integer; WP: WPARAM; LP: LPARAM): LRESULT; stdcall;
var
  VK: Integer;
  Inputs: array [0 .. 1] of TInput;
begin
  Result := CallNextHookEx(Hook, Code, WP, LP);

  if Code < 0 then
    Exit;

  VK := PKBDLLHOOKSTRUCT(LP)^.vkCode;

  if (WP = WM_KEYDOWN) or (WP = WM_SYSKEYDOWN) then
  begin
    if VK = VK_LALT then
    begin
      LAltDown := True;
      AltDirty := False;
    end
    else if VK = VK_RALT then
    begin
      RAltDown := True;
      AltDirty := False;
    end
    else if LAltDown or RAltDown then
      AltDirty := True;
  end
  else if (WP = WM_KEYUP) or (WP = WM_SYSKEYUP) then
  begin
    if (VK = VK_LALT) and LAltDown then
    begin
      LAltDown := False;
      if not AltDirty then
      begin
        // alt-ime-ahk準拠: ダミーキーでメニュー起動を抑制してからIME切り替え
        ZeroMemory(@Inputs, SizeOf(Inputs));
        Inputs[0].Itype := INPUT_KEYBOARD;
        Inputs[0].ki.wVk := VK_F23;
        Inputs[1].Itype := INPUT_KEYBOARD;
        Inputs[1].ki.wVk := VK_F23;
        Inputs[1].ki.dwFlags := KEYEVENTF_KEYUP;
        SendInput(2, Inputs[0], SizeOf(TInput));
        SetIme(False);
      end;
    end
    else if (VK = VK_RALT) and RAltDown then
    begin
      RAltDown := False;
      if not AltDirty then
      begin
        ZeroMemory(@Inputs, SizeOf(Inputs));
        Inputs[0].Itype := INPUT_KEYBOARD;
        Inputs[0].ki.wVk := VK_F23;
        Inputs[1].Itype := INPUT_KEYBOARD;
        Inputs[1].ki.wVk := VK_F23;
        Inputs[1].ki.dwFlags := KEYEVENTF_KEYUP;
        SendInput(2, Inputs[0], SizeOf(TInput));
        SetIme(True);
      end;
    end;
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  ShowWindow(Application.Handle, SW_HIDE);
  Application.MainFormOnTaskbar := False;
  SetWindowLong(Application.Handle, GWL_EXSTYLE,
    GetWindowLong(Application.Handle, GWL_EXSTYLE) or WS_EX_TOOLWINDOW);
  Hide;

  Hook := SetWindowsHookEx(WH_KEYBOARD_LL, @HookCallback,
    GetModuleHandle(nil), 0);
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  if Hook <> 0 then
    UnhookWindowsHookEx(Hook);
end;

end.
