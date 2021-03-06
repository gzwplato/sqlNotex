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

unit Unit6;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls;

type

  { TfmRenameTags }

  TfmRenameTags = class(TForm)
    bnOK: TButton;
    bnCancel: TButton;
    edOldTag: TEdit;
    edNewTag: TEdit;
    lbOldTag: TLabel;
    lbNewTag: TLabel;
    procedure edNewTagKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState
      );
    procedure edOldTagKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState
      );
  private

  public

  end;

var
  fmRenameTags: TfmRenameTags;

implementation

{$R *.lfm}

{ TfmRenameTags }

procedure TfmRenameTags.edOldTagKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = 13 then
  begin
    edNewTag.SetFocus;
  end;
end;

procedure TfmRenameTags.edNewTagKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 13 then
  begin
    ModalResult := mrOK;
  end;
end;

end.

