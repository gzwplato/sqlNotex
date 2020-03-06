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

unit Unit5;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, Grids, ExtCtrls,
  StdCtrls, LazFileUtils;

type

  { TfmBookmarks }

  TfmBookmarks = class(TForm)
    bnClear: TButton;
    bnSet: TButton;
    bnCancel: TButton;
    bnGoTo: TButton;
    lbBookmarks: TLabel;
    pnBookmarks: TPanel;
    sgBookmarks: TStringGrid;
    shBookmarks: TShape;
    procedure bnCancelClick(Sender: TObject);
    procedure bnClearClick(Sender: TObject);
    procedure bnGoToClick(Sender: TObject);
    procedure bnSetClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure sgBookmarksDblClick(Sender: TObject);
  private

  public

  end;

var
  fmBookmarks: TfmBookmarks;

implementation

uses Unit1;

{$R *.lfm}

{ TfmBookmarks }

procedure TfmBookmarks.FormCreate(Sender: TObject);
begin
  sgBookmarks.SelectedColor := myColor;
  sgBookmarks.FocusRectVisible := False;
  if FileExistsUTF8(myHomeDir + 'bookmarks') then
  begin
    sgBookmarks.LoadFromFile(myHomeDir + 'bookmarks');
  end;
end;

procedure TfmBookmarks.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = 27 then
  begin
    ModalResult := mrCancel;
  end
  else
  if key = 13 then
  begin
    key := 0;
    bnGoToClick(nil);
  end;
end;

procedure TfmBookmarks.sgBookmarksDblClick(Sender: TObject);
begin
  bnGoToClick(nil);
end;

procedure TfmBookmarks.bnSetClick(Sender: TObject);
begin
  with sgBookmarks do
  begin
    if fmMain.zqNotesID.AsString <> '' then
    begin
      Cells[1, Row] := fmMain.zqNotebooksID.AsString;
      Cells[2, Row] := fmMain.zqSectionsID.AsString;
      Cells[3, Row] := fmMain.zqNotesID.AsString;
      Cells[4, Row] := fmMain.zqNotebooksTITLE.Value;
      Cells[5, Row] := fmMain.zqSectionsTITLE.Value;
      Cells[6, Row] := fmMain.zqNotesTITLE.Value;
    end;
  end;
end;

procedure TfmBookmarks.bnClearClick(Sender: TObject);
begin
  with sgBookmarks do
  begin
    Cells[1, Row] := '';
    Cells[2, Row] := '';
    Cells[3, Row] := '';
    Cells[4, Row] := '';
    Cells[5, Row] := '';
    Cells[6, Row] := '';
  end;
end;

procedure TfmBookmarks.bnGoToClick(Sender: TObject);
begin
  if sgBookmarks.Cells[1, sgBookmarks.Row] <> '' then
  begin;
    ModalResult := mrOK;
  end;
end;

procedure TfmBookmarks.bnCancelClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

procedure TfmBookmarks.FormDestroy(Sender: TObject);
begin
  try
    sgBookmarks.SaveOptions := [soContent];
    sgBookmarks.SaveToFile(myHomeDir + 'bookmarks');
  except
  end;
end;

end.

