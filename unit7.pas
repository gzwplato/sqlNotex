// ***********************************************************************
// ***********************************************************************
// sqlNotex 1.x
// Author and copyright: Massimo Nardello, Modena (Italy) 2020.
// Free software released under GPL licence version 3 or later.

// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version. You can read the version 3
// of the Licence in http://www.gnu.org/licenses/gpl-3.0.txt
// or in the file Licence.txt included in the files of the
// source code of this software.

// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
// ***********************************************************************
// ***********************************************************************

unit Unit7;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, Menus,
  Clipbrd;

type

  { TfmInsertID }

  TfmInsertID = class(TForm)
    bnOK: TButton;
    bnCancel: TButton;
    bnPaste: TButton;
    edID: TEdit;
    lbIDKind: TLabel;
    pmMenu: TPopupMenu;
    procedure bnPasteClick(Sender: TObject);
    procedure edIDKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormShow(Sender: TObject);
  private

  public

  end;

var
  fmInsertID: TfmInsertID;

implementation

{$R *.lfm}

{ TfmInsertID }

procedure TfmInsertID.edIDKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = 13 then
  begin
    ModalResult := mrOK;
  end
  else
  if (((key < 48) or (key > 57)) and (key <> 8)) then
  begin
    key := 0;
  end;
end;

procedure TfmInsertID.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = 27 then
  begin
    ModalResult := mrCancel;
  end;
end;

procedure TfmInsertID.bnPasteClick(Sender: TObject);
begin
  edID.Clear;
  edID.PasteFromClipboard;
end;

procedure TfmInsertID.FormShow(Sender: TObject);
  var i: integer;
begin
  edID.SetFocus;
  edID.SelectAll;
  if TryStrToInt(Clipboard.AsText, i) = True then
  begin
    bnPaste.Enabled := True;
  end
  else
  begin
    bnPaste.Enabled := False;
  end;
end;

end.

