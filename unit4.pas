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

unit Unit4;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ComCtrls,
  ExtCtrls, LazUTF8;

type

  { TfmOptions }

  TfmOptions = class(TForm)
    bnClose: TButton;
    bnStFontColor: TButton;
    bnStBackColor: TButton;
    bnDefault: TButton;
    bnStMarkColor: TButton;
    bnStHighlightColor: TButton;
    cbStFonts: TComboBox;
    cdColorDialog: TColorDialog;
    edBackup: TEdit;
    edServer: TEdit;
    edPath: TEdit;
    edPort: TEdit;
    edStSize: TEdit;
    edStSpaceAfter: TEdit;
    edStSizeTitles: TEdit;
    edStLineSpace: TEdit;
    lbBackup: TLabel;
    lbServer: TLabel;
    lbPath: TLabel;
    lbPort: TLabel;
    lbStSize: TLabel;
    lbStSpaceAfter: TLabel;
    lbStSizeTitles: TLabel;
    lbStLineSpace: TLabel;
    lnStFonts: TLabel;
    udStSize: TUpDown;
    udStSpaceAfter: TUpDown;
    udStSizeTitles: TUpDown;
    udStLineSpace: TUpDown;
    procedure bnDefaultClick(Sender: TObject);
    procedure bnStBackColorClick(Sender: TObject);
    procedure bnStFontColorClick(Sender: TObject);
    procedure bnCloseClick(Sender: TObject);
    procedure bnStHighlightColorClick(Sender: TObject);
    procedure bnStMarkColorClick(Sender: TObject);
    procedure cbStFontsChange(Sender: TObject);
    procedure edStLineSpaceChange(Sender: TObject);
    procedure edStSizeChange(Sender: TObject);
    procedure edStSizeTitlesChange(Sender: TObject);
    procedure edStSpaceAfterChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private

  public

  end;

var
  fmOptions: TfmOptions;

implementation

Uses Unit1;

{$R *.lfm}

{ TfmOptions }

procedure TfmOptions.FormCreate(Sender: TObject);
begin
  cbStFonts.Items := Screen.Fonts;
  cbStFonts.ItemIndex := 0;
end;

procedure TfmOptions.FormShow(Sender: TObject);
begin
  if fmMain.zcConnection.Connected = False then
  begin
    edServer.ReadOnly := False;
    edPort.ReadOnly := False;
    edPath.ReadOnly := False;
  end
  else
  begin
    edServer.ReadOnly := True;
    edPort.ReadOnly := True;
    edPath.ReadOnly := True;
  end;
  edServer.Text := fmMain.zcConnection.HostName;
  edPort.Text := IntToStr(fmMain.zcConnection.Port);
  edPath.Text := fmMain.zcConnection.Database;
  edBackup.Text := stBackupFile;
end;

procedure TfmOptions.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  with fmMain do
  begin
    if zcConnection.Connected = False then
    begin
      zcConnection.HostName := edServer.Text;
      zcConnection.Port := StrToInt(edPort.Text);
      zcConnection.Database := edPath.Text;
    end;
    stBackupFile := edBackup.Text;
    if UTF8LowerCase(edServer.Text) <> 'localhost' then
    begin
      fmMain.miToolsBackup.Enabled := False;
      fmMain.miToolsRestore.Enabled := False;
      fmMain.miToolsCompact.Enabled := False;
      fmMain.bnBackup.Enabled := False;
      fmMain.bnRestore.Enabled := False;
    end
    else
    if fmMain.zcConnection.Connected = False then
    begin
      fmMain.miToolsBackup.Enabled := True;
      fmMain.miToolsRestore.Enabled := True;
      fmMain.miToolsCompact.Enabled := True;
      fmMain.bnBackup.Enabled := True;
      fmMain.bnRestore.Enabled := True;
    end;
  end;
end;

procedure TfmOptions.cbStFontsChange(Sender: TObject);
  var
    stText: String;
begin
  if fmMain.dbText.Visible = True then
  begin
    fmMain.dbText.Font.Name := cbStFonts.Text;
    stText := fmMain.dbText.Text;
    fmMain.dbText.Clear;
    fmMain.dbText.Text := stText;
    fmMain.ListAndFormatTitle(True);
    fmMain.FormatMarkers(True);
    fmMain.sgTitles.Font.Name := cbStFonts.Text;
  end;
end;

procedure TfmOptions.edStSizeChange(Sender: TObject);
  var
    stText: String;
begin
  fmMain.dbText.Font.Size := StrToInt(edStSize.Text);
  stText := fmMain.dbText.Text;
  fmMain.dbText.Clear;
  fmMain.dbText.Text := stText;
  fmMain.ListAndFormatTitle(True);
  fmMain.FormatMarkers(True);
end;

procedure TfmOptions.edStSizeTitlesChange(Sender: TObject);
begin
  fmMain.sgTitles.Font.Size := StrToInt(edStSizeTitles.Text);
end;

procedure TfmOptions.edStSpaceAfterChange(Sender: TObject);
  var
    stText: String;
begin
    fmMain.SetAfterSpace(StrToInt(edStSpaceAfter.Text));
    stText := fmMain.dbText.Text;
    fmMain.dbText.Clear;
    fmMain.dbText.Text := stText;
    fmMain.MainFormat;
    fmMain.ListAndFormatTitle(True);
    fmMain.FormatMarkers(True);
end;

procedure TfmOptions.edStLineSpaceChange(Sender: TObject);
  var
    stText: String;
begin
    fmMain.SetLineSpace(StrToInt(edStLineSpace.Text));
    stText := fmMain.dbText.Text;
    fmMain.dbText.Clear;
    fmMain.dbText.Text := stText;
    fmMain.MainFormat;
    fmMain.ListAndFormatTitle(True);
    fmMain.FormatMarkers(True);
end;

procedure TfmOptions.bnStFontColorClick(Sender: TObject);
  var
    stText: String;
begin
  cdColorDialog.Color := fmMain.dbText.Font.Color;
  if cdColorDialog.Execute then
  begin
    fmMain.dbText.Font.Color := cdColorDialog.Color;
    stText := fmMain.dbText.Text;
    fmMain.dbText.Clear;
    fmMain.dbText.Text := stText;
    fmMain.ListAndFormatTitle(True);
    fmMain.FormatMarkers(True);
    fmMain.sgTitles.Font.Color := cdColorDialog.Color;
  end;
end;

procedure TfmOptions.bnStBackColorClick(Sender: TObject);
begin
  cdColorDialog.Color := fmMain.dbText.Color;
  if cdColorDialog.Execute then
  begin
    fmMain.dbText.Color := cdColorDialog.Color;
    fmMain.pnText.Color := cdColorDialog.Color;
    fmMain.sgTitles.Color := cdColorDialog.Color;
    fmMain.ListAndFormatTitle(True);
    fmMain.FormatMarkers(True);
  end;
end;

procedure TfmOptions.bnDefaultClick(Sender: TObject);
  var
    stText: String;
begin
  stText := fmMain.dbText.Text;
  fmMain.dbText.Clear;
  fmMain.dbText.Text := stText;
  fmMain.dbText.Color := clWhite;
  fmMain.dbText.Font.Color := clBlack;
  fmMain.pnText.Color := clWhite;
  fmMain.sgTitles.Font.Color := clBlack;
  fmMain.sgTitles.Color := clWhite;
  clMarkCol := clRed;
  clHighColor := clYellow;
  fmMain.ListAndFormatTitle(True);
  fmMain.FormatMarkers(True);
end;

procedure TfmOptions.bnStMarkColorClick(Sender: TObject);
  var
    stText: String;
begin
  cdColorDialog.Color := clMarkCol;
  if cdColorDialog.Execute then
  begin
    clMarkCol := cdColorDialog.Color;
    stText := fmMain.dbText.Text;
    fmMain.dbText.Clear;
    fmMain.dbText.Text := stText;
    fmMain.ListAndFormatTitle(True);
    fmMain.FormatMarkers(True);
  end;
end;

procedure TfmOptions.bnStHighlightColorClick(Sender: TObject);
  var
    stText: String;
begin
  cdColorDialog.Color := clHighColor;
  if cdColorDialog.Execute then
  begin
    clHighColor := cdColorDialog.Color;
    stText := fmMain.dbText.Text;
    fmMain.dbText.Clear;
    fmMain.dbText.Text := stText;
    fmMain.ListAndFormatTitle(True);
    fmMain.FormatMarkers(True);
  end;
end;

procedure TfmOptions.bnCloseClick(Sender: TObject);
begin
  Close;
end;

end.

