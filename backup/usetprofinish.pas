unit uSetProFinish;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  StdCtrls, ButtonPanel, CheckLst, ComCtrls, IniFiles, sqldb
  ;

type

  { TFormSetProFinish }

  TFormSetProFinish = class(TForm)
    ButtonPanel1: TButtonPanel;
    cgSections: TCheckGroup;
    cgZones: TCheckGroup;
    cgEditions: TCheckGroup;
    CheckGroup1: TCheckGroup;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    lwRules: TListView;
    Panel1: TPanel;
    Panel2: TPanel;
    rgproduct: TRadioGroup;
    ScrollBox1: TScrollBox;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    procedure ButtonPanel1Click(Sender: TObject);
    procedure cgEditionsItemClick(Sender: TObject; Index: integer);
    procedure cgSectionsItemClick(Sender: TObject; Index: integer);
    procedure cgZonesItemClick(Sender: TObject; Index: integer);
    procedure CheckGroup1ItemClick(Sender: TObject; Index: integer);
    procedure FormShow(Sender: TObject);
    procedure lwRulesSelectItem(Sender: TObject; Item: TListItem; Selected: Boolean);
    procedure OKButtonClick(Sender: TObject);
    procedure rgproductClick(Sender: TObject);
    procedure rgproductSelectionChanged(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
  private

  public
  end;

var
  FormSetProFinish: TFormSetProFinish;

implementation
Uses
  Umain;
{$R *.lfm}

{ TFormSetProFinish }



procedure TFormSetProFinish.lwRulesSelectItem(Sender: TObject; Item: TListItem; Selected: Boolean);
var
  SL: TStringList;
  i: Integer;
begin
  SL := TStringlist.Create;
  SL.Delimiter := ',';
  SL.StrictDelimiter := True;

  rgproduct.ItemIndex := rgproduct.Items.IndexOf(Item.SubItems[0]);

  SL.DelimitedText := Item.SubItems[1];
  For i := 0 to CheckGroup1.Items.Count - 1 do
    CheckGroup1.Checked[i] := (SL.IndexOf(CheckGroup1.Items[i]) >= 0);

  SL.DelimitedText := Item.SubItems[2];
  For i := 0 to cgSections.Items.Count - 1 do
    cgSections.Checked[i] := (SL.IndexOf(cgSections.Items[i]) >= 0);

  SL.DelimitedText := Item.SubItems[3];
  For i := 0 to cgZones.Items.Count - 1 do
    cgZones.Checked[i] := (SL.IndexOf(cgZones.Items[i]) >= 0);

  SL.DelimitedText := Item.SubItems[4];
  For i := 0 to cgEditions.Items.Count - 1 do
    cgEditions.Checked[i] := (SL.IndexOf(cgEditions.Items[i]) >= 0);

  SL.Free;
end;

procedure TFormSetProFinish.OKButtonClick(Sender: TObject);
var
  i: Integer;
  SL: TStringList;
  INI: TIniFile;
begin
  Try
    SL := TStringlist.Create;
    INI := TINIFile.Create(fMain.Config.CheckFinishfileName);
    //CopyFile('ProFinishToCC.ini', 'ProFinishToCC.bak');
    For i := 0 to SL.Count - 1 do
      INI.EraseSection(SL[i]);
    For i := 0 to FormSetProFinish.lwRules.Items.Count - 1 do
    Begin
      INI.WriteBool('Rule' + IntToStr(i), 'Use', FormSetProFinish.lwRules.Items[i].Checked);
      INI.WriteString('Rule' + IntToStr(i), 'Rulename', FormSetProFinish.lwRules.Items[i].Caption);
      INI.WriteString('Rule' + IntToStr(i), 'Aliasname', FormSetProFinish.lwRules.Items[i].SubItems[0]);
      INI.WriteString('Rule' + IntToStr(i), 'Days', FormSetProFinish.lwRules.Items[i].SubItems[1]);
      INI.WriteString('Rule' + IntToStr(i), 'Section', FormSetProFinish.lwRules.Items[i].SubItems[2]);
      INI.WriteString('Rule' + IntToStr(i), 'Zone', FormSetProFinish.lwRules.Items[i].SubItems[3]);
      INI.WriteString('Rule' + IntToStr(i), 'Edition', FormSetProFinish.lwRules.Items[i].SubItems[4]);
    end;


  finally
    SL.Free;
  end;
end;

procedure TFormSetProFinish.rgproductClick(Sender: TObject);
begin
  If rgproduct.ItemIndex >= 0 then
    If lwRules.SelCount > 0 then
      lwRules.Selected.SubItems[0] := rgproduct.Items[rgproduct.ItemIndex];
end;

procedure TFormSetProFinish.rgproductSelectionChanged(Sender: TObject);
begin

end;

procedure TFormSetProFinish.ToolButton1Click(Sender: TObject);
var
  LI: TListItem;
begin
  LI := lwRules.Items.Add;
  LI.Caption := 'New';
  LI.Checked := False;
  LI.SubItems.Add('');
  LI.SubItems.Add('');
  LI.SubItems.Add('');
  LI.SubItems.Add('');
  LI.SubItems.Add('');
end;

procedure TFormSetProFinish.ToolButton2Click(Sender: TObject);
var
  i: Integer;
begin
  If lwRules.SelCount > 0 then
    For i := 0 to lwRules.Items.Count - 1 do
      If lwRules.Items[i].Selected Then
      Begin
        lwRules.Items.Delete(i);
        Exit;
      end;
end;

procedure TFormSetProFinish.cgSectionsItemClick(Sender: TObject; Index: integer);
var
  TmpS: String;
  i: Integer;
begin
  If lwRules.SelCount > 0 then
  Begin
    TmpS := '';
    For i := 0 to cgSections.Items.Count - 1 do
      If cgSections.Checked[i] Then
        TmpS := Tmps + ',' + cgSections.Items[i];
    lwRules.Selected.SubItems[2] := Copy(TmpS, 2, Length(TmpS));
  end;
end;

procedure TFormSetProFinish.cgEditionsItemClick(Sender: TObject; Index: integer);
var
  TmpS: String;
  i: Integer;
begin
  If lwRules.SelCount > 0 then
    Begin
      TmpS := '';
      For i := 0 to cgEditions.Items.Count - 1 do
        If cgEditions.Checked[i] Then
          TmpS := Tmps + ',' + cgEditions.Items[i];
      lwRules.Selected.SubItems[4] := Copy(TmpS, 2, Length(TmpS));
    end;

end;

procedure TFormSetProFinish.ButtonPanel1Click(Sender: TObject);
begin

end;

procedure TFormSetProFinish.cgZonesItemClick(Sender: TObject; Index: integer);
var
  TmpS: String;
  i: Integer;
begin
  If lwRules.SelCount > 0 then
    Begin
      TmpS := '';
      For i := 0 to cgZones.Items.Count - 1 do
        If cgZones.Checked[i] Then
          TmpS := Tmps + ',' + cgZones.Items[i];
      lwRules.Selected.SubItems[3] := Copy(TmpS, 2, Length(TmpS));
    end;
end;

procedure TFormSetProFinish.CheckGroup1ItemClick(Sender: TObject; Index: integer);
var
  TmpS: String;
  i: Integer;
begin
  If lwRules.SelCount > 0 then
    Begin
      TmpS := '';
      For i := 0 to CheckGroup1.Items.Count - 1 do
        If CheckGroup1.Checked[i] Then
          TmpS := Tmps + ',' + CheckGroup1.Items[i];
      lwRules.Selected.SubItems[1] := Copy(TmpS, 2, Length(TmpS));
    end;
end;

procedure TFormSetProFinish.FormShow(Sender: TObject);
var
  Que: TSQLQuery;
  INI: TIniFile;
  SL: TStringList;
  i: Integer;
  LI: TListItem;
begin
  try
    Que := TSQLquery.Create(nil);
    Que.DataBase := fMain.MSSQLConnection;
    Que.Transaction := fMain.SQLTrans;
    Que.PacketRecords:= -1;

    Que.Close;
    Que.SQL.Text := 'Select Short_Name from PRO_NAME WITH (NOLOCK)';
    Que.SQL.Add('ORDER BY Short_NAME');
    Que.Open;
    rgproduct.Items.Clear;
    While not Que.EOF do
    begin
      rgproduct.Items.Add(Que.FieldByName('Short_NAME').AsString);
      Que.Next;
    end;


    INI := TINIFile.Create(fMain.Config.CheckFinishfileName);
    SL := TStringList.Create;
    INI.ReadSections(SL);

    lwRules.BeginUpdate;
    lwRules.Items.Clear;
    For i := 0 to SL.Count - 1 do
    begin
      LI := lwRules.Items.Add;
      LI.Caption := INI.ReadString(SL[i], 'Rulename', '');
      LI.Checked := INI.ReadBool(SL[i], 'Use', False);
      LI.SubItems.Add(INI.ReadString(SL[i], 'Aliasname', ''));
      LI.SubItems.Add(INI.ReadString(SL[i], 'Days', ''));
      LI.SubItems.Add(INI.ReadString(SL[i], 'Section', ''));
      LI.SubItems.Add(INI.ReadString(SL[i], 'Zone', ''));
      LI.SubItems.Add(INI.ReadString(SL[i], 'Edition', ''));
    end;
    For i := 0 to lwRules.ColumnCount - 1 do
      lwRules.Column[i].Width := -2;
    lwRules.EndUpdate;

    Que.Close;
    Que.SQL.Text := 'Select Name from SectionsName WITH (NOLOCK)';
    Que.SQL.Add('ORDER BY NAME');
    Que.Open;
    cgSections.Items.Clear;
    While not Que.EOF do
    begin
      cgSections.Items.Add(Que.FieldByName('Name').AsString);
      Que.Next;
    End;

    Que.Close;
    Que.SQL.Text := 'Select Name from ZonesName WITH (NOLOCK)';
    Que.SQL.Add('ORDER BY NAME');
    Que.Open;
    cgZones.Items.Clear;
    While not Que.EOF do
    begin
      FormSetProFinish.cgZones.Items.Add(Que.FieldByName('Name').AsString);
      Que.Next;
    End;

    Que.Close;
    Que.SQL.Text := 'Select Name from EditionsName WITH (NOLOCK)';
    Que.SQL.Add('ORDER BY NAME');
    Que.Open;
    cgEditions.Items.Clear;
    While not Que.EOF do
    begin
      cgEditions.Items.Add(Que.FieldByName('Name').AsString);
      Que.Next;
    End;


  except
    SL.Free;
    Que.Free;
    INI.Free;
  end;

end;

end.

