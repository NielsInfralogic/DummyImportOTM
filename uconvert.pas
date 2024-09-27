unit uConvert;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ComCtrls,
  ExtCtrls, Spin, EditBtn, ButtonPanel, CheckLst, DateTimePicker,
  TAChartListbox, INIfiles, mssqlconn, sqldb;

type

  { TfConvert }

  TfConvert = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    ButtonPanel1: TButtonPanel;
    cgProname: TCheckGroup;
    chbCrossAdSendToCC: TCheckBox;
    chbRulePartPagecount: TCheckBox;
    chbRuleForceCommon: TCheckBox;
    chbRuleForceUnique: TCheckBox;
    chbRuleCreate: TCheckBox;
    chbRuleAppendPagecount: TCheckBox;
    chbRuleSetOnlyPageCount: TCheckBox;
    chbRuleSearchInSecrem: TCheckBox;
    chbRuleSetUnique: TCheckBox;
    cgEdition: TCheckGroup;
    cgZone: TCheckGroup;
    cgSection: TCheckGroup;
    chbRuleTurnSection: TCheckBox;
    ChbRuleMasterToTurnSection: TCheckBox;
    ComboBoxRipSetup: TComboBox;
    deCrossAddDone: TDirectoryEdit;
    deCrossAddNoMatch: TDirectoryEdit;
    deCrossAddPath: TDirectoryEdit;
    eCrossAddFIleFilter: TEdit;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    EditTurnSectionDefault: TEdit;
    eRulePageType: TEdit;
    eRuleSecRem: TEdit;
    eRulesAlias: TEdit;
    eRulesEdi: TEdit;
    eRuleMasterToTurnSection: TEdit;
    eRulesSec: TEdit;
    eRulesZon: TEdit;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    GroupBox4: TGroupBox;
    GroupBox47: TGroupBox;
    GroupBox66: TGroupBox;
    GroupBox68: TGroupBox;
    GroupBox69: TGroupBox;
    GroupBox70: TGroupBox;
    Label1: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label150: TLabel;
    Label151: TLabel;
    Label152: TLabel;
    Label153: TLabel;
    Label156: TLabel;
    Label158: TLabel;
    Label162: TLabel;
    Label163: TLabel;
    Label164: TLabel;
    Label165: TLabel;
    Label166: TLabel;
    Label17: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label96: TLabel;
    Label97: TLabel;
    Label98: TLabel;
    ListView1: TListView;
    ListView2: TListView;
    ListView3: TListView;
    lwRuleName: TListView;
    PageControl1: TPageControl;
    Panel1: TGroupBox;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel50: TPanel;
    Panel51: TPanel;
    Panel52: TPanel;
    Panel53: TPanel;
    Panel57: TPanel;
    Panel58: TPanel;
    Panel6: TPanel;
    Panel62: TPanel;
    Panel64: TPanel;
    Panel65: TPanel;
    Panel66: TPanel;
    Panel67: TPanel;
    Panel68: TPanel;
    Panel9: TPanel;
    PanelTurnSection: TPanel;
    Panel7: TPanel;
    Panel73: TPanel;
    Panel8: TPanel;
    PanelPageType: TPanel;
    ScrollBox1: TScrollBox;
    ScrollBox2: TScrollBox;
    ScrollBox3: TScrollBox;
    ScrollBox4: TScrollBox;
    seRuleCommonEnd: TSpinEdit;
    seRuleCommonStart: TSpinEdit;
    seRuleUniqueEnd: TSpinEdit;
    seRuleUniqueStart: TSpinEdit;
    SpinTrimW: TSpinEdit;
    SpinTrimH: TSpinEdit;
    Splitter1: TSplitter;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    tbDeleteZone: TToolButton;
    tbDeleteZone1: TToolButton;
    tbDeleteZone2: TToolButton;
    tbRulesAdd: TToolButton;
    tbRulesDelete: TToolButton;
    ToolBar1: TToolBar;
    ToolBar2: TToolBar;
    ToolBar3: TToolBar;
    ToolBar5: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure ButtonPanel1Click(Sender: TObject);
    procedure chbRuleAppendPagecountEditingDone(Sender: TObject);
    procedure chbRuleCreateEditingDone(Sender: TObject);
    procedure chbRuleForceCommonEditingDone(Sender: TObject);
    procedure chbRuleForceUniqueEditingDone(Sender: TObject);
    procedure ChbRuleMasterToTurnSectionChange(Sender: TObject);
    procedure chbRulePartPagecountEditingDone(Sender: TObject);
    procedure chbRuleSearchInSecremChange(Sender: TObject);
    procedure chbRuleSearchInSecremEditingDone(Sender: TObject);
    procedure chbRuleSetOnlyPageCountEditingDone(Sender: TObject);
    procedure chbRuleSetUniqueEditingDone(Sender: TObject);
    procedure cgEditionItemClick(Sender: TObject; Index: integer);
    procedure cgZoneItemClick(Sender: TObject; Index: integer);
    procedure cgPronameItemClick(Sender: TObject; Index: integer);
    procedure cgSectionClick(Sender: TObject);
    procedure cgSectionItemClick(Sender: TObject; Index: integer);
    procedure chbRuleTurnSectionChange(Sender: TObject);
    procedure Edit7Change(Sender: TObject);
    procedure eRuleMasterToTurnSectionChange(Sender: TObject);
    procedure eRulePageTypeChange(Sender: TObject);
    procedure eRulesAliasChange(Sender: TObject);
    procedure eRuleSecRemChange(Sender: TObject);
    procedure eRulesEdiChange(Sender: TObject);
    procedure eRulesSecChange(Sender: TObject);
    procedure eRulesZonChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure GroupBox69Click(Sender: TObject);
    procedure Label9Click(Sender: TObject);
    procedure ListView1SelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure ListView2SelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure ListView3SelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure lwRuleNameSelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure OKButtonClick(Sender: TObject);
    procedure Panel68Click(Sender: TObject);
    procedure Panel8Click(Sender: TObject);
    procedure tbDeleteZone1Click(Sender: TObject);
    procedure tbDeleteZone2Click(Sender: TObject);
    procedure tbDeleteZoneClick(Sender: TObject);
    procedure tbRulesAddClick(Sender: TObject);
    procedure tbRulesDeleteClick(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure ToolButton3Click(Sender: TObject);
  private

  public

  end;

var
  fConvert: TfConvert;

implementation
Uses
  Umain;
{$R *.lfm}

{ TfConvert }


procedure TfConvert.FormShow(Sender: TObject);
var
  INI: TIniFile;
  slZone, slEdi, slSec, SL, SLL: TStringList;
  LI: TListItem;
  i, x, y: Integer;
  Que: TSQLQuery;
begin
  Try
    Que := TSQLquery.Create(nil);
    Que.DataBase := fMain.MSSQLConnection;
    Que.Transaction := fMain.SQLTrans;
    Que.PacketRecords:= -1;

    INI := TINIFile.Create(fMain.Config.INIfileName);

    deCrossAddDone.Text      := INI.ReadString('Path', 'Done', '');
    deCrossAddPath.Text      := INI.ReadString('Path', 'Input', '');
    eCrossAddFIleFilter.Text := INI.ReadString('Path', 'Filter', '');
    deCrossAddNoMatch.Text   := INI.ReadString('Path', 'No Match', '');
    INI.Free;

    INI := TINIFile.Create(fMain.Config.ConvertFileName);
    slZone := TStringlist.Create;
    Listview1.Items.Clear;
    INI.ReadSection('Zone', slZone);
    For i := 0 to slZone.Count - 1 do
    Begin
      LI := Listview1.Items.Add;
      LI.Caption := slZone[i];
      LI.SubItems.Add(INI.ReadString('Zone', slZone[i], ''));
    end;
    slEdi := TStringlist.Create;
    Listview2.Items.Clear;
    INI.ReadSection('Edition', slEdi);
    For i := 0 to slEdi.Count - 1 do
    Begin
      LI := Listview2.Items.Add;
      LI.Caption := slEdi[i];
      LI.SubItems.Add(INI.ReadString('Edition', slEdi[i], ''));
    end;
    slSec := TStringlist.Create;
    Listview3.Items.Clear;
    INI.ReadSection('Section', slSec);
    For i := 0 to slSec.Count - 1 do
    Begin
      LI := Listview3.Items.Add;
      LI.Caption := slSec[i];
      LI.SubItems.Add(INI.ReadString('Section', slSec[i], ''));
    end;
    INI.Free;
    slZone.Free;
    slEdi.Free;
    slSec.Free;



    Que.Close;
    Que.SQL.Text:='SELECT Long_Name, Short_Name, ID FROM Pro_Name order by Short_Name';
    Que.Open;
    cgProname.Items.Clear;
    While Not Que.EOF do
    Begin
      cgProname.Items.AddObject(Que.FieldByName('Short_Name').AsString, TObject(Que.FieldByName('ID').AsInteger));
      Que.Next;
    end;

    Que.Close;
    Que.SQL.Text:='SELECT Name, ID FROM EditionsNAME order by Name';
    Que.Open;
    cgEdition.Items.Clear;
    While Not Que.EOF do
    Begin
      cgEdition.Items.AddObject(Que.FieldByName('Name').AsString, TObject(Que.FieldByName('ID').AsInteger));
      Que.Next;
    end;

    Que.Close;
    Que.SQL.Text:='SELECT Name, ID FROM ZonesNAME order by Name';
    Que.Open;
    cgZone.Items.Clear;
    While Not Que.EOF do
    Begin
      cgZone.Items.AddObject(Que.FieldByName('Name').AsString, TObject(Que.FieldByName('ID').AsInteger));
      Que.Next;
    end;

    cgSection.Items.Clear;
    For i := 0 to 25 do
    Begin
      cgSection.Items.AddObject(Chr(i + 65), TObject(i));
    End;

    // ### NAN 2024-08-21
    ComboBoxRipSetup.Items.Clear;
    for i:=0 to Length(fMain.Config.Names.RipSetups)-1 do
    Begin
      ComboBoxRipSetup.Items.Add( fMain.Config.Names.RipSetups[i].Name);
    end;
    ComboBoxRipSetup.ItemIndex := ComboBoxRipSetup.Items.IndexOf(fMain.Config.Convert.DefaultTrimRipSetup);
    // ###


    INI := TINIFile.Create(fMain.Config.ConvertFileName);
    SL := TStringList.Create;
    SLL := TStringList.Create;
    SLL.Delimiter:= ';';
    SLL.StrictDelimiter:= True;
    lwRuleName.Items.Clear;
    INI.ReadSection('Rule', SL);
    For i := 0 to SL.Count - 1 do
    Begin
      SLL.DelimitedText := INI.ReadString('Rule', SL[i], '');
      Li := lwRuleName.Items.Add;

      While SLL.Count < lwRuleName.ColumnCount + 2  do
        SLL.Add(''); //Dette blot hvis der ikke er nok col

      LI.Checked := StrToBoolDef(SLL[0], false);
      LI.Caption := SLL[1];
      For x := 2 to SLL.Count - 1 do
        LI.SubItems.Add(SLL[x]);
    end;

    SL.Free;
    SLL.Free;
    INI.Free;


  except
  on E: Exception do
    begin
      fMain.StatusBar.Panels.Items[1].Text := E.Message;
      fMain.EventLog.Error(E.Message);
    end;
  end;

end;

procedure TfConvert.GroupBox69Click(Sender: TObject);
begin

end;

procedure TfConvert.Label9Click(Sender: TObject);
begin

end;


procedure TfConvert.eRulesAliasChange(Sender: TObject);
begin
  If lwRuleName.SelCount = 1 then
    lwRuleName.Selected.Caption:= Uppercase(eRulesAlias.Text);
end;

procedure TfConvert.eRuleSecRemChange(Sender: TObject);
begin
  If lwRuleName.SelCount = 1 then
    lwRuleName.Selected.SubItems[8]:= Uppercase(eRuleSecRem.Text);
end;

procedure TfConvert.eRulesEdiChange(Sender: TObject);
begin
  If lwRuleName.SelCount = 1 then
    lwRuleName.Selected.SubItems[1]:= Uppercase(eRulesEdi.Text);
end;

procedure TfConvert.eRulesSecChange(Sender: TObject);
begin
  If lwRuleName.SelCount = 1 then
    lwRuleName.Selected.SubItems[0] := Uppercase(eRulesSec.Text);
end;

procedure TfConvert.eRulesZonChange(Sender: TObject);
begin
  If lwRuleName.SelCount = 1 then
      lwRuleName.Selected.SubItems[2]:= Uppercase(eRulesZon.Text);
end;

procedure TfConvert.Button1Click(Sender: TObject);
begin
  Try
    If (Listview1.SelCount = 1) then
    Begin
      Listview1.Selected.Caption     := Uppercase(Edit1.Text);
      Listview1.Selected.SubItems[0] := Uppercase(Edit2.Text);
    end;
    Panel2.Visible:= False;

  except
  on E: Exception do
    begin
      fMain.StatusBar.Panels.Items[1].Text := E.Message;
      fMain.EventLog.Error(E.Message);
    end;
  end;
end;

procedure TfConvert.Button2Click(Sender: TObject);
begin
  Try
     If (Listview2.SelCount = 1) then
     Begin
       Listview2.Selected.Caption     := Uppercase(Edit3.Text);
       Listview2.Selected.SubItems[0] := Uppercase(Edit4.Text);
     end;
     Panel3.Visible:= False;

  except
  on E: Exception do
    begin
      fMain.StatusBar.Panels.Items[1].Text := E.Message;
      fMain.EventLog.Error(E.Message);
    end;
  end;
end;

procedure TfConvert.Button3Click(Sender: TObject);
begin
  Try
     If (Listview3.SelCount = 1) then
     Begin
       Listview3.Selected.Caption     := Uppercase(Edit5.Text);
       Listview3.Selected.SubItems[0] := Uppercase(Edit6.Text);
     end;
     Panel4.Visible:= False;

  except
  on E: Exception do
    begin
      fMain.StatusBar.Panels.Items[1].Text := E.Message;
      fMain.EventLog.Error(E.Message);
    end;
  end;
end;

procedure TfConvert.ButtonPanel1Click(Sender: TObject);
begin

end;

procedure TfConvert.chbRuleAppendPagecountEditingDone(Sender: TObject);
begin
  Try
      If lwRuleName.SelCount = 1 then
       lwRuleName.Selected.SubItems[19]:= BoolToStr(chbRuleAppendPagecount.Checked);
    except
    on E: Exception do
      begin
        fMain.StatusBar.Panels.Items[1].Text := E.Message;
        fMain.EventLog.Error(E.Message);
      end;
    end;
end;


procedure TfConvert.chbRuleCreateEditingDone(Sender: TObject);
begin
  Try
    If lwRuleName.SelCount = 1 then
       lwRuleName.Selected.SubItems[17]:= BoolToStr(chbRuleCreate.Checked);
  except
  on E: Exception do
    begin
      fMain.StatusBar.Panels.Items[1].Text := E.Message;
      fMain.EventLog.Error(E.Message);
    end;
  end;
end;

procedure TfConvert.chbRuleForceCommonEditingDone(Sender: TObject);
begin
  Try
    If lwRuleName.SelCount = 1 then
      lwRuleName.Selected.SubItems[12]:= BoolToStr(chbRuleForceCommon.Checked);

  except
  on E: Exception do
    begin
      fMain.StatusBar.Panels.Items[1].Text := E.Message;
      fMain.EventLog.Error(E.Message);
    end;
  end;

end;

procedure TfConvert.chbRuleForceUniqueEditingDone(Sender: TObject);
begin
  Try
  If lwRuleName.SelCount = 1 then
      lwRuleName.Selected.SubItems[9]:= BoolToStr(chbRuleForceUnique.Checked);
  except
  on E: Exception do
    begin
      fMain.StatusBar.Panels.Items[1].Text := E.Message;
      fMain.EventLog.Error(E.Message);
    end;
  end;

end;



procedure TfConvert.chbRulePartPagecountEditingDone(Sender: TObject);
begin
  Try
    If lwRuleName.SelCount = 1 then
      lwRuleName.Selected.SubItems[20]:= BoolToStr(chbRulePartPagecount.Checked);
     except
     on E: Exception do
       begin
         fMain.StatusBar.Panels.Items[1].Text := E.Message;
         fMain.EventLog.Error(E.Message);
       end;
     end;
end;

procedure TfConvert.chbRuleSearchInSecremChange(Sender: TObject);
begin
  GroupBox66.Enabled := not (chbRuleSearchInSecrem.Checked);
end;

procedure TfConvert.chbRuleSearchInSecremEditingDone(Sender: TObject);
begin
  Try
    If lwRuleName.SelCount = 1 then
     lwRuleName.Selected.SubItems[18]:= BoolToStr(chbRuleSearchInSecrem.Checked);
  except
  on E: Exception do
    begin
      fMain.StatusBar.Panels.Items[1].Text := E.Message;
      fMain.EventLog.Error(E.Message);
    end;
  end;

end;


procedure TfConvert.chbRuleSetOnlyPageCountEditingDone(Sender: TObject);
begin
  Try
    If lwRuleName.SelCount = 1 then
     lwRuleName.Selected.SubItems[16]:= BoolToStr(chbRuleSetOnlyPageCount.Checked);

  except
  on E: Exception do
    begin
      fMain.StatusBar.Panels.Items[1].Text := E.Message;
      fMain.EventLog.Error(E.Message);
    end;
  end;

end;


procedure TfConvert.chbRuleSetUniqueEditingDone(Sender: TObject);
begin
  Try
    If lwRuleName.SelCount = 1 then
     lwRuleName.Selected.SubItems[15]:= BoolToStr(chbRuleSetUnique.Checked);

  except
  on E: Exception do
    begin
      fMain.StatusBar.Panels.Items[1].Text := E.Message;
      fMain.EventLog.Error(E.Message);
    end;
  end;
end;

procedure TfConvert.cgEditionItemClick(Sender: TObject; Index: integer);
var
  SL: TStringList;
  i: Integer;
begin
  Try
    If lwRuleName.SelCount = 1 then
    Begin;
      SL := TStringlist.Create;
      For i := 0 to cgEdition.Items.Count - 1 do
        If cgEdition.Checked[i] then
          SL.Add(cgEdition.Items[i]);
      lwRuleName.Selected.SubItems[6] := SL.CommaText;
      SL.Free;
    end;


  except
  on E: Exception do
    begin
      fMain.StatusBar.Panels.Items[1].Text := E.Message;
      fMain.EventLog.Error(E.Message);
    end;
  end;
end;


procedure TfConvert.cgZoneItemClick(Sender: TObject; Index: integer);
var
  SL: TStringList;
  i: Integer;
begin
  Try
    If lwRuleName.SelCount = 1 then
    Begin;
      SL := TStringlist.Create;
      For i := 0 to cgZone.Items.Count - 1 do
        If cgZone.Checked[i] then
          SL.Add(cgZone.Items[i]);
      lwRuleName.Selected.SubItems[7] := SL.CommaText;
      SL.Free;
    end;

  except
  on E: Exception do
    begin
      fMain.StatusBar.Panels.Items[1].Text := E.Message;
      fMain.EventLog.Error(E.Message);
    end;
  end;
end;


procedure TfConvert.cgPronameItemClick(Sender: TObject; Index: integer);
var
  SL: TStringList;
  i: Integer;
begin
  Try
    If lwRuleName.SelCount = 1 then
    Begin;
      SL := TStringlist.Create;
      For i := 0 to cgProname.Items.Count - 1 do
        If cgProname.Checked[i] then
          SL.Add(cgProname.Items[i]);
          //SL.Add(IntToStr(Integer(cgProname.Items.Objects[i])));
      lwRuleName.Selected.SubItems[4] := SL.CommaText;
      SL.Free;
    end;

  except
  on E: Exception do
    begin
      fMain.StatusBar.Panels.Items[1].Text := E.Message;
      fMain.EventLog.Error(E.Message);
    end;
  end;
end;

procedure TfConvert.cgSectionClick(Sender: TObject);
begin

end;

procedure TfConvert.cgSectionItemClick(Sender: TObject; Index: integer);
var
  SL: TStringList;
  i: Integer;
begin
  Try
    If lwRuleName.SelCount = 1 then
    Begin;
      SL := TStringlist.Create;
      For i := 0 to cgSection.Items.Count - 1 do
        If cgSection.Checked[i] then
          SL.Add(cgSection.Items[i]);
      lwRuleName.Selected.SubItems[5] := SL.CommaText;
      SL.Free;
    end;

  except
  on E: Exception do
    begin
      fMain.StatusBar.Panels.Items[1].Text := E.Message;
      fMain.EventLog.Error(E.Message);
    end;
  end;
end;



procedure TfConvert.Edit7Change(Sender: TObject);
begin

end;

// ### NAN 2024-08-20
procedure TfConvert.chbRuleTurnSectionChange(Sender: TObject);
begin
  PanelTurnSection.Enabled := chbRuleTurnSection.Checked;
   If lwRuleName.SelCount = 1 then
      lwRuleName.Selected.SubItems[21]:= BoolToStr(chbRuleTurnSection.Checked);
end;

// ### NAN 2024-08-20
procedure TfConvert.ChbRuleMasterToTurnSectionChange(Sender: TObject);
begin
    If lwRuleName.SelCount = 1 then
      lwRuleName.Selected.SubItems[22]:= BoolToStr(ChbRuleMasterToTurnSection.Checked);
end;

// ### NAN 2024-08-20
procedure TfConvert.eRuleMasterToTurnSectionChange(Sender: TObject);
begin
  If lwRuleName.SelCount = 1 then
      lwRuleName.Selected.SubItems[23]:= eRuleMasterToTurnSection.Text
end;

procedure TfConvert.eRulePageTypeChange(Sender: TObject);
begin
   If lwRuleName.SelCount = 1 then
      lwRuleName.Selected.SubItems[24]:= eRulePageType.Text
end;

procedure TfConvert.ListView1SelectItem(Sender: TObject; Item: TListItem; Selected: Boolean);
begin
  Panel2.Visible:= Selected;
  Edit1.Text := Item.Caption;
  Edit2.Text := Item.SubItems[0];
end;

procedure TfConvert.ListView2SelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
begin
  Panel3.Visible:= Selected;
   Edit3.Text := Item.Caption;
   Edit4.Text := Item.SubItems[0];
end;

procedure TfConvert.ListView3SelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
begin
  Panel4.Visible:= Selected;
   Edit5.Text := Item.Caption;
   Edit6.Text := Item.SubItems[0];
end;

procedure TfConvert.lwRuleNameSelectItem(Sender: TObject; Item: TListItem;
  Selected: Boolean);
var
  sl: TStringList;
  i: Integer;
begin
  Try
    If selected then
    Begin
      eRulesAlias.Text := lwRuleName.Selected.Caption;
      eRulesSec.Text   := lwRuleName.Selected.SubItems[0];
      eRulesEdi.Text   := lwRuleName.Selected.SubItems[1];
      eRulesZon.Text   := lwRuleName.Selected.SubItems[2];
      eRuleSecRem.Text :=  lwRuleName.Selected.SubItems[8];
      chbRuleCreate.Checked           := StrToBoolDef(lwRuleName.Selected.SubItems[17], false);
      chbRuleSetUnique.Checked        := StrToBoolDef(lwRuleName.Selected.SubItems[15], false);
      chbRuleSetOnlyPageCount.Checked := StrToBoolDef(lwRuleName.Selected.SubItems[16], false);
      chbRuleForceUnique.Checked      := StrToBoolDef(lwRuleName.Selected.SubItems[9], false);
      chbRuleForceCommon.Checked      := StrToBoolDef(lwRuleName.Selected.SubItems[12], false);
      chbRuleSearchInSecrem.Checked   := StrToBoolDef(lwRuleName.Selected.SubItems[18], false);
      chbRuleAppendPagecount.Checked  := StrToBoolDef(lwRuleName.Selected.SubItems[19], false);
      chbRulePartPagecount.Checked  := StrToBoolDef(lwRuleName.Selected.SubItems[20], false);

      // ### NAN 2024-08-20 - nye felter for vendesektion
      chbRuleTurnSection.Checked  := StrToBoolDef(lwRuleName.Selected.SubItems[21], false);
      ChbRuleMasterToTurnSection.Checked := StrToBoolDef(lwRuleName.Selected.SubItems[22], false);
      eRuleMasterToTurnSection.Text :=  lwRuleName.Selected.SubItems[23];
      eRulePageType.Text  := lwRuleName.Selected.SubItems[24];

      PanelTurnSection.Enabled := chbRuleTurnSection.Checked;


      sl := TStringlist.Create;
      sl.Delimiter:= ',';

      sl.DelimitedText := lwRuleName.Selected.SubItems[4];
      Label1.Caption:= 'Selected:';
      For i := 0 to cgProname.Items.Count - 1 do
      Begin
        cgProname.Checked[i] := (sl.IndexOf(cgProname.Items[i]) >=0);
        If cgProname.Checked[i] then
          Label1.Caption := Label1.Caption + ' ' + cgProname.Items[i];
      end;

      sl.DelimitedText := lwRuleName.Selected.SubItems[5];
      Label3.Caption:= 'Selected:';
      For i := 0 to cgSection.Items.Count - 1 do
      Begin
        cgSection.Checked[i] := (sl.IndexOf(cgSection.Items[i]) >=0);
        If cgSection.Checked[i] then
          Label3.Caption := Label3.Caption + ' ' + cgSection.Items[i];
      end;

      sl.DelimitedText := lwRuleName.Selected.SubItems[6];
      Label4.Caption:= 'Selected:';
      For i := 0 to cgEdition.Items.Count - 1 do
      Begin
        cgEdition.Checked[i] := (sl.IndexOf(cgEdition.Items[i]) >=0);
        If cgEdition.Checked[i] then
          Label4.Caption := Label4.Caption + ' ' + cgEdition.Items[i];
      end;

      sl.DelimitedText := lwRuleName.Selected.SubItems[7];
      Label2.Caption:= 'Selected:';
      For i := 0 to cgZone.Items.Count - 1 do
      Begin
        cgZone.Checked[i] := (sl.IndexOf(cgZone.Items[i]) >=0);
        If cgZone.Checked[i] then
          Label2.Caption := Label2.Caption + ' ' + cgZone.Items[i];
      end;

      sl.Free;
    end;

  except
  on E: Exception do
    begin
      fMain.StatusBar.Panels.Items[1].Text := E.Message;
      fMain.EventLog.Error(E.Message);
    end;
  end;
end;

procedure TfConvert.OKButtonClick(Sender: TObject);
var
  INI: TIniFile;
  i, x: Integer;
  SL: TStringList;
begin
  Try
    INI := TINIFile.Create(fMain.Config.INIfileName);
    INI.WriteString('Path', 'Done', deCrossAddDone.Text);
    INI.WriteString('Path', 'Input', deCrossAddPath.Text);
    INI.WriteString('Path', 'Filter', eCrossAddFIleFilter.Text);
    INI.WriteString('Path', 'No match', deCrossAddNoMatch.Text);

    // ### NAN 2024-08-21 - nye defaults
    INI.WriteString('System', 'DefaultTurnSectionName', EditTurnSectionDefault.Text);
    INI.WriteString('System', 'DefaultTrimRipSetup', ComboBoxRipSetup.Text);
    INI.WriteInteger('System',  'DefaultTrimFormatW',  SpinTrimW.Value);
    INI.WriteInteger('System',  'DefaultTrimFormatH',  SpinTrimH.Value);
    // ###

    INI.Free;

    INI := TINIFile.Create(fMain.Config.ConvertFileName);
    INI.EraseSection('Zone');
    For i := 0 to Listview1.Items.Count - 1 do
    Begin
      INI.WriteString('Zone', Listview1.Items[i].Caption, Listview1.Items[i].SubItems[0]);
    end;
    INI.EraseSection('Edition');
    For i := 0 to Listview2.Items.Count - 1 do
    Begin
      INI.WriteString('Edition', Listview2.Items[i].Caption, Listview2.Items[i].SubItems[0]);
    end;
    INI.EraseSection('Section');
    For i := 0 to Listview3.Items.Count - 1 do
    Begin
      INI.WriteString('Section', Listview3.Items[i].Caption, Listview3.Items[i].SubItems[0]);
    end;
    INI.Free;

    INI := TINIFile.Create(fMain.Config.ConvertFileName);
    SL := TStringlist.Create;
    SL.Delimiter := ';';
    SL.StrictDelimiter := True;
    INI.EraseSection('Rule');
    For i := 0 to lwRuleName.Items.Count - 1 do
    Begin
      SL.Clear;
      SL.Add(BoolToStr(lwRuleName.Items[i].Checked));
      SL.Add(lwRuleName.Items[i].Caption);
      For x := 0 to lwRuleName.ColumnCount - 1 do
        SL.Add(lwRuleName.Items[i].SubItems[x]);
      INI.WriteString('Rule', IntToStr(i), SL.DelimitedText);
    end;
    SL.Free;


  except
  on E: Exception do
    begin
      fMain.StatusBar.Panels.Items[1].Text := E.Message;
      fMain.EventLog.Error(E.Message);
    end;
  end;
end;

procedure TfConvert.Panel68Click(Sender: TObject);
begin

end;

procedure TfConvert.Panel8Click(Sender: TObject);
begin

end;


procedure TfConvert.tbDeleteZone1Click(Sender: TObject);
var
  i: Integer;
begin
  For i := Listview2.Items.Count - 1 downto 0 do
     If Listview2.Items[i].Selected then
       Listview2.Items[i].Delete;
end;

procedure TfConvert.tbDeleteZone2Click(Sender: TObject);
var
  i: Integer;
begin
  For i := Listview3.Items.Count - 1 downto 0 do
     If Listview3.Items[i].Selected then
       Listview3.Items[i].Delete;
end;

procedure TfConvert.tbDeleteZoneClick(Sender: TObject);
var
  i: Integer;
begin
  For i := Listview1.Items.Count - 1 downto 0 do
    If Listview1.Items[i].Selected then
      Listview1.Items[i].Delete;
end;

procedure TfConvert.tbRulesAddClick(Sender: TObject);
var
  LI: TListItem;
  i: Integer;
begin
  Try
    LI := lwRuleName.Items.Add;
    Li.Caption:= 'New';
    For i := 0 to lwRuleName.ColumnCount - 1 do
       Li.SubItems.Add('');
  finally
  end;
end;

procedure TfConvert.tbRulesDeleteClick(Sender: TObject);
begin
  If lwRuleName.SelCount >= 1 then
    lwRuleName.Selected.Delete;
end;

procedure TfConvert.ToolButton1Click(Sender: TObject);
var
  LI: TListItem;
begin
  LI := Listview1.Items.Add;
  LI.Caption := 'New';
  LI.SubItems.Add('');
end;

procedure TfConvert.ToolButton2Click(Sender: TObject);
var
  LI: TListItem;
begin
  LI := Listview2.Items.Add;
  LI.Caption := 'New';
  LI.SubItems.Add('');
end;

procedure TfConvert.ToolButton3Click(Sender: TObject);
var
  LI: TListItem;
begin
  LI := Listview3.Items.Add;
  LI.Caption := 'New';
  LI.SubItems.Add('');
end;



end.

