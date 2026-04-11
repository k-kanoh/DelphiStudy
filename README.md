# DelphiStudy

Delphi12 の学習、日曜プログラミング

## 機能

### 1. Alt単独押しでIME切り替え(alt-ime-ahk互換)
- **左Alt** 単独押し → IME OFF(英数入力)
- **右Alt** 単独押し → IME ON(日本語入力)
- Alt+他のキーの組み合わせ(Alt+Tab など)は通常通り動作する

### 2. ExcelでShift+EnterをAlt+Enterに変換
- **Excelがアクティブな場合のみ** Shift+Enter をAlt+Enter(セル内改行)に変換する
- Shiftを押しっぱなしにしての連続入力にも対応

## 参考

- [alt-ime-ahk](https://github.com/karakaram/alt-ime-ahk) - AutoHotkeyで同機能を実現したオリジナル実装。IME制御の方法とメニュー起動抑制の手法はこのプロジェクトに準拠している
- [ExcelのAlt+Enterをshift+Enterに割り当てるAHKスクリプト](https://www.naporitansushi.com/ahk-excel-newline/)

## 試行錯誤の記録

### IME切り替えの失敗した手法

**`ImmSetOpenStatus`**  
Windowsの古いIME APIで、古くからある手法。Windows 11の新しいMicrosoft IMEでは効かない。

**`SendInput`で`VK_DBE_HIRAGANA` / `VK_DBE_ALPHANUMERIC`を送る**  
仮想キーコードでIMEのon/offを直接送る方法。動作しなかった。

**TSF(Text Services Framework)経由**  
`ITfInputProcessorProfiles`や`ITfCompartment`を使う方法。Windows 11対応の正攻法だが、COMインターフェースの定義が複雑で実装が重くなる。

### IME切り替えの成功した手法

**`ImmGetDefaultIMEWnd` + `WM_IME_CONTROL`メッセージ**  
alt-ime-ahkが使っているのと同じ方法。フォーカスされたウィンドウのIMEウィンドウハンドルを取得して、`WM_IME_CONTROL`メッセージの`IMC_SETOPENSTATUS`(wParam = `$006`)を直接送る。Windows 11でも動作する。

```delphi
ImeWnd := ImmGetDefaultIMEWnd(TargetWnd);
SendMessage(ImeWnd, WM_IME_CONTROL, $006, Ord(Open));
```

ウィンドウハンドルは`GetForegroundWindow`だけでは不十分なケースがあるため、`GetGUIThreadInfo`でフォーカスされたコントロールのハンドルを取得している。

### Altキーのメニュー起動抑制の失敗した手法

**Altキーアップイベントだけをブロック(`Result := 1`)**  
キーアップだけ握りつぶすとシステムがAlt押しっぱなしと認識し、Escキーでウィンドウが切り替わるなど副作用が出た。

**Altキーダウン時に`vk07`(未定義の仮想キー)を挟む**  
alt-ime-ahkがメニュー起動抑制に使っている手法。Delphi版のLL hookではSendInputのイベントも拾ってしまうため、`vk07`がAltと組み合わさって半角カナ入力モードになる副作用が発生した。

### Altキーのメニュー起動抑制の成功した手法

**Altキーアップ時に`VK_F23`を挟む**  
Altキーが単独で離される直前に`VK_F23`のdown/upを`SendInput`で送る。これによりシステムが「Altは他のキーと組み合わせて使われた」と認識し、メニューバーがアクティブにならない。alt-ime-ahkは同目的で`vk07`を使っているが、Delphi版のLL hookでは`vk07`がAltと組み合わさって半角カナ入力モードになる副作用があったため、実用上ほぼ使われない`VK_F23`を代わりに採用した。

```delphi
// Altアップ前にVK_F23を挟む
Inputs[0].ki.wVk := VK_F23;           // F23 down
Inputs[1].ki.wVk := VK_F23;           // F23 up
Inputs[1].ki.dwFlags := KEYEVENTF_KEYUP;
SendInput(2, Inputs[0], SizeOf(TInput));
```

## 処理の説明

### グローバルキーボードフック(`WH_KEYBOARD_LL`)

`SetWindowsHookEx`でシステム全体のキーボードイベントを監視するフックを登録する。すべてのキー入力がこのフックを経由するため、どのアプリにフォーカスがあっても動作する。

フック内で`CallNextHookEx`を最初に呼んで他のフックにイベントを渡し、その後で処理を行う。イベントを握りつぶす場合のみ`Result := 1`を返して`Exit`する。

### `AltDirty`フラグ

Altキーが単独で押されたか、他のキーと組み合わせて押されたかを判定するフラグ。

- Altキーダウン時に`False`にリセット
- Altが押されている間に他のキーが押されたら`True`にセット
- Altキーアップ時に`AltDirty = False`の場合のみIME切り替えを実行

### メッセージループ

VCLを使わずWin32の`GetMessage`ループを直接回している。`WH_KEYBOARD_LL`フックはメッセージループが動いていることが必要なため。

```delphi
while GetMessage(Msg, 0, 0, 0) do
begin
  TranslateMessage(Msg);
  DispatchMessage(Msg);
end;
```

### フォアグラウンドプロセス判定(Excel用)

`GetForegroundWindow`でアクティブウィンドウのプロセスIDを取得し、`OpenProcess` + `GetModuleFileNameEx`(`Winapi.PsAPI`)でプロセス名を取得して`EXCEL.EXE`と比較する。全プロセス列挙より高速。

### Shift+Enter → Alt+Enter変換

AHKの記事を参考にShift+EnterをAlt+Enterに差し替える発想を得た。具体的な実装はこちらの工夫。

Shift+EnterのキーダウンをブロックしてAlt+Enterを`SendInput`で送る。単純にAlt+Enterを送るだけではShiftが押しっぱなしになりAlt+Shift+Enterになってしまうため、先にShiftを離す必要がある。またShiftを押しっぱなしにしての連続入力に対応するため、シーケンスの最後にShiftを押し直している。

```
Shift up → Alt down → Enter down → Alt up → Shift down
```

これら5つのイベントを`ZeroMemory`で初期化した配列に設定して`SendInput`で一括送信することで、イベント間に他の入力が割り込まない。

## 注意事項

- Windows Defenderにキーボードフックアプリとして検知されることがある。デバッグビルドは検知されやすく、Releaseビルドは通ることが多い
- `KBDLLHOOKSTRUCT`はDelphi 12の`Winapi.Windows`に定義されていないため、手動で定義している
- 終了はタスクマネージャーからプロセスをkillする

---

*このプロジェクトは [Claude Code](https://claude.ai/code) (claude-sonnet-4-6) との対話を通じて作成されました。*
