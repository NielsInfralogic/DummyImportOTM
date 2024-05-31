unit uMain;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, sqldblib, sqldb, eventlog, mssqlconn, Forms, Controls,
  Graphics, Dialogs, ExtCtrls, ComCtrls, LazUTF8, LazFileUtils, FileUtil,
  Laz2_DOM, Laz2_XMLWrite, Laz2_XMLRead,
  INIfiles,
  BlowFish,
  LConvEncoding,
  RegExpr,
  uConvert, CreatePro,
  uSetProFinish;

type

   // ### NAN 2024-05-30

  TImportProductSection = Record
    SectionName : String;
    PageCount : Integer;
    SectionRemark : String;
  end;

  TImportProduct = Record
    PubName      : String;
    PubDate      : TDate;
    Editions     : Array of String;
    Sections     : Array of TImportProductSection;
    PageFormat   : String;
  end;

  // ### end

  TPro = Record
    PubName        : String;
    Zone           : String;
    Edition        : String;
    Section        : String;
    PubDate        : TDate;
    NoOfPages      : Integer;
    PrintingStart  : String;
    NumberOfCopies : Integer;
    PrintingEnd    : Integer;
    ProOK          : Boolean;
    FileName       : String;
    Trim           : Boolean;
  end;

  TConv = Record
    Key : String;
    Value : String;
  end;

  TRules = Record
    fName : String;
    fSec  : String;
    fEdi : String;
    fZone : String;
    tName : Array of String;
    tSec : Array of String;
    tEdi : Array of String;
    tZone : Array of String;
    SecRem : String;
    ForceUnique : Boolean;
    ForceCommon : Boolean;
    SetOnlyUnique : Boolean;
    SetOnlyPagecount : Boolean;
    Flying : Boolean;
    CreateMissing : Boolean;
    AppendCopyCount : Boolean;
    PartCopyCount : Boolean;
  end;


  TConvert = Record
   ConvZone : Array of TConv;
   ConvEdi : Array of TConv;
   ConvSec : Array of TConv;
   Rules : Array of TRules;
  end;

  TSystem = Record
    AppData : String;
    HomePath : String;
    Desktop : String;
    Comp_Name : String;
    DokPath   : String;
    DebugLevel : Integer;
    InputPath : String;
    DonePath : String;
    FileFilter : STring;
    OutputPath : String;
    NoMatchPath : String;
  end;

  TSQLdb = Record
    UserName     : String;
    UserPassword : String;
    HostName     : String;
    TableName    : String;
  end;

  TPPMNameId = Record
    Name : String;
    Id : Integer;
  end;

  TNames = Record
    Zone : Array of TPPMNameId;
    Editions : Array of TPPMNameId;
  end;

  { TConfig }

  TConfig = Class(TComponent)
  Public
    System : TSystem;
    INIfileName : String;
    CheckFinishfileName : String;
    ConvertFileName : String;
    LogfileName : String;
    SQLdb : TSQLdb;
    Convert : TConvert;
    Names : TNames;
    SendToCC : Boolean;
    Procedure LoadConfig;
    constructor Create(AOwner: TComponent);
  private
  End;


  { TfMain }

  TfMain = class(TForm)
    EventLog: TEventLog;
    il16x16: TImageList;
    il32x32: TImageList;
    lwLog: TListView;
    MSSQLConnection: TMSSQLConnection;
    SQLDBLibraryLoader: TSQLDBLibraryLoader;
    SQLQuery1: TSQLQuery;
    SQLTrans: TSQLTransaction;
    StatusBar: TStatusBar;
    Timer1: TTimer;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    tbSimulate: TToolButton;
    tbStatus: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    procedure FormCreate(Sender: TObject);
    procedure tbStatusClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure ToolButton3Click(Sender: TObject);
    procedure tbSimulateClick(Sender: TObject);
    procedure ToolButton7Click(Sender: TObject);
  private
    function Decrypt(chaine: string): string;
    procedure DoImport;
    function Encrypt(chaine: string): string;
    function IDToZone(IValue: Integer): String;
    procedure ReadConfig;
    function ReadyToSendtToCC(PPMID: Integer; CheckRem: Boolean): Boolean;
    procedure WriteLog(TextMsg: string);
    function ZoneToID(NameValue: String): Integer;

  public
    Config : TConfig;

  end;

var
  fMain: TfMain;
  RunningFile: THandle;

Const
  EncryptKey = 'Infra2Logic';

implementation

{$R *.lfm}

{ TConfig }

procedure TConfig.LoadConfig;
begin

end;

constructor TConfig.Create(AOwner: TComponent);
begin

end;

{ TfMain }

procedure TfMain.WriteLog(TextMsg: string);
var
  LI: TListItem;
  i: integer;
begin
  LI := lwLog.Items.Add;
  LI.Caption := FormatDateTime('DD.MM hh:nn', NOW);
  LI.SubItems.Add(TextMsg);
  if lwLog.Items.Count > 10000 then
  begin
    for i := 0 to 1000 do
      lwLog.Items.Delete(0);
  end;
  lwLog.items[lwLog.items.Count - 1].makevisible(False);
  lwLog.Repaint;
end;

function TfMain.IDToZone(IValue: Integer): String;
var
  i: Integer;
Begin
  Result := '';
  For i := 0 to Length(Config.Names.Zone) - 1 do
    If Config.Names.Zone[i].Id = iValue then
      Result := Config.Names.Zone[i].Name;
end;

function TfMain.ZoneToID(NameValue: String): Integer;
var
  i: Integer;
Begin
  Result := -1;
  For i := 0 to Length(Config.Names.Zone) - 1 do
    If COnfig.Names.Zone[i].Name = NameValue then
      Result := Config.Names.Zone[i].Id;
end;


function TfMain.ReadyToSendtToCC(PPMID: Integer; CheckRem: Boolean): Boolean;
var
  INI: TIniFile;
  SL, SLSec, SLZon, SLEdi, SLPag, SLRSec, SLRZon, SLREdi, SLRDay,
    SLRem: TStringList;
  ProType, MainPro: LongInt;
  ShortName , TmpS: String;
  x, i, y, di, RuleNr, d, s: Integer;
  Found: Boolean;
  PubDate: TDateTime;
  Que: TSQLQuery;
Begin
  Try
    Result := True;
    WriteLog('Check product for sending to CC: ' + IntToStr(PPMID));

    Que := TSQLquery.Create(nil);
    Que.DataBase := MSSQLConnection;
    Que.Transaction := SQLTrans;
    Que.PacketRecords:= -1;

    INI := TINIFile.Create('ProFinishToCC.ini');
    SL := TStringList.Create;
    SLPag := TStringList.Create;
    SLPag.Delimiter:= ',';
    SLPag.StrictDelimiter:= True;
    SLSec := TStringList.Create;
    SLSec.Delimiter:= ',';
    SLSec.StrictDelimiter:= True;
    SLRem := TStringList.Create;
    SLRem.Delimiter:= ',';
    SLRem.StrictDelimiter:= True;
    SLZon := TStringList.Create;
    SLZon.Delimiter:= ',';
    SLZon.StrictDelimiter:= True;
    SLEdi := TStringList.Create;
    SLEdi.Delimiter := ',';
    SLEdi.StrictDelimiter := True;
    SLRSec := TStringList.Create;
    SLRSec.Delimiter:= ',';
    SLRSec.StrictDelimiter:= True;
    SLRZon := TStringList.Create;
    SLRZon.Delimiter:= ',';
    SLRZon.StrictDelimiter:= True;
    SLREdi := TStringList.Create;
    SLREdi.Delimiter := ',';
    SLREdi.StrictDelimiter := True;
    SLRDay := TStringList.Create;
    SLRDay.Delimiter := ',';
    SLRDay.StrictDelimiter := True;

    INI.ReadSections(SL);
    Que.Close;
    Que.SQL.Text := 'Select Top 1 Short_Name, Pro_PageCount, Pro_SecName, Pro_SecRem, ZONES, EDITIONS, ProType, MAIN_PRO, Udg_Dato from data with (nolock)';
    Que.SQL.Add('Where ID=''' + IntToSTr(PPMID) + '''');
    Que.Open;

    If Que.RecordCount > 0 then
    Begin
      ShortName := Que.FieldByName('Short_Name').AsString;
      ProType   := Que.FieldByName('ProType').AsInteger;
      MainPro   := Que.FieldByName('MAIN_PRO').AsInteger;

      If ProType = 5 then //Er det en sub? Så hent hoved produktet
      Begin
        PPMID := MainPro;
        ProType := 4;
        Que.Close;
        Que.SQL.Text := 'Select Top 1 Short_Name, Pro_PageCount, Pro_SecName, Pro_SecRem, ZONES, EDITIONS, Udg_Dato from data with (nolock)';
        Que.SQL.Add('Where ID=''' + IntToSTr(PPMID) + '''');
        Que.Open;
      end;

      If Que.RecordCount > 0 then
      Begin
        RuleNr := -1;
        Found := False;
        SLRSec.DelimitedText := '';
        SLRZon.DelimitedText := '';
        SLREdi.DelimitedText := '';
        SLRDay.DelimitedText := '';
        Pubdate := Que.FieldByName('Udg_Dato').AsDateTime;

        For i := 0 to SL.Count - 1 do
        Begin
          If (ShortName = INI.ReadString(SL[i], 'Aliasname', '')) And (INI.ReadBool(SL[i], 'Use', False)) then
          Begin
            SLRDay.DelimitedText := INI.ReadString(SL[i], 'Days', '');

            For d := 0 to SLRDay.Count - 1 do
            Begin
              di := DayOfWeek(PubDate);
              TmpS := SLRDay[d];
                If ((di = 1) and (UTF8Uppercase(SLRDay[d]) = 'SØNDAG')) or
                   ((di = 2) and (UTF8Uppercase(SLRDay[d]) = 'MANDAG')) or
                   ((di = 3) and (UTF8Uppercase(SLRDay[d]) = 'TIRSDAG')) or
                   ((di = 4) and (UTF8Uppercase(SLRDay[d]) = 'ONSDAG')) or
                   ((di = 5) and (UTF8Uppercase(SLRDay[d]) = 'TORSDAG')) or
                   ((di = 6) and (UTF8Uppercase(SLRDay[d]) = 'FREDAG')) or
                   ((di = 7) and (UTF8Uppercase(SLRDay[d]) = 'LØRDAG')) then
                   Begin
                     Found := True;
                     RuleNr := i;
                     SLRSec.DelimitedText := INI.ReadString(SL[i], 'Section', '');
                     SLRZon.DelimitedText := INI.ReadString(SL[i], 'Zone', '');
                     SLREdi.DelimitedText := INI.ReadString(SL[i], 'Edition', '');
                     WriteLog('Found rule: ' + INI.ReadString(SL[i], 'Rulename', ''));
                     WriteLog('With Sec:' + SLRSec.CommaText + ' Zon:' + SLRZon.CommaText + ' Edi:' +  SLREdi.CommaText + ' Day:' + SLRDay.CommaText);
                     Break;
                   end;
            End;

          end;
        end;


        SLPag.DelimitedText := Que.FieldByName('Pro_PageCount').AsString;
        SLSec.DelimitedText := Que.FieldByName('Pro_SecName').AsString;
        SLRem.DelimitedText := Que.FieldByName('Pro_SecRem').AsString;
        SLZon.DelimitedText := Que.FieldByName('ZONES').AsString;
        SLEdi.DelimitedText := Que.FieldByName('EDITIONS').AsString;
        PubDate             := Que.FieldByName('Udg_Dato').AsDateTime;

        //Hvis det er et hovedpro skal data fra subber hentes
        If ProType = 4 then
        Begin
          Que.Close;
          Que.SQL.Text := 'Select Pro_PageCount, Pro_SecName, Pro_SecRem, ZONES, EDITIONS from data with (nolock)';
          Que.SQL.Add('Where ProType=5 And Main_Pro=''' + IntToSTr(PPMID) + '''');
          //Que.SQL.SaveToFile('7001.sql');
          Que.Open;

          While Not Que.EOF do
          Begin
            SLPag.DelimitedText := SLPag.DelimitedText + ',' + Que.FieldByName('Pro_PageCount').AsString;
            SLSec.DelimitedText := SLSec.DelimitedText + ',' + Que.FieldByName('Pro_SecName').AsString;
            SLRem.DelimitedText := SLRem.DelimitedText + ',' + Que.FieldByName('Pro_SecRem').AsString;
            SLZon.DelimitedText := SLZon.DelimitedText + ',' + Que.FieldByName('ZONES').AsString;
            SLEdi.DelimitedText := SLEdi.DelimitedText + ',' + Que.FieldByName('EDITIONS').AsString;
            Que.Next;
          End;
        end;

        If Found then //Er der en regl, ellers ud
        Begin
          //Check sektioner
          If SLRSec.Count > 0 Then
          Begin
            For i := 0 to SLRSec.Count - 1 do
            Begin
              Found := False;
              For y := 0 to SLSec.Count - 1 do
                If SLRSec[i] = SLSec[y] then
                  Found := True;
              If Not Found Then
              Begin
                WriteLog('Missing section ' + SLRSec[i]);
                Result := False;
              End;
            end;
          End;

          //Check zoner
          If SLRZon.Count > 0 Then
          Begin
            For i := 0 to SLRZon.Count - 1 do
            Begin
              Found := False;
              TmpS := SLZon.CommaText;
              For y := 0 to SLZon.Count - 1 do
                If SLRZon[i] = IdToZone(StrToIntDef(SLZon[y], 0)) then
                  Found := True;
              If Not Found Then
              Begin
                WriteLog('Missing zone ' + SLRZon[i]);
                Result := False;
              End;
            end;
          end;

          //Check edition
          If SLREdi.Count > 0 Then
          Begin
            For i := 0 to SLREdi.Count - 1 do
            Begin
              Found := False;
              For y := 0 to SLEdi.Count - 1 do
                //If SLREdi[i] = IdToEdition(StrToIntDef(SLEdi[y], 0)) then
                  Found := True;
              If Not Found Then
              Begin
                WriteLog('Missing editions .' + SLREdi[i]);
                Result := False;
              End;
            end;
          end;

          //Check section Rem for tom, men kun i dem der skal checkes
          If (CheckRem) and (SLRSec.Count > 0) Then
          Begin
            For s := 0 to SLRSec.Count - 1 do
            Begin
              For i := 0 to SLRem.Count - 1 do
              Begin
                If (SLRSec[s] = SLSec[i]) and (Trim(SLRem[i]) = '') then
                Begin
                  WriteLog('No remark in section: ' + SLRSec[s]);
                  Result := False;
                End;
              end;
            end;
          end;
        end;
      End Else
        Result := False;
    end else
      Result := False;

    //Check sidetal for 0 eller tom, Det skal gøres ved alle selv uden en regl
    For i := 0 to SLPag.Count - 1 do
      If StrToIntDef(SLPag[i], 0) <= 0 then
        Begin
          WriteLog('No pagecount.');
          Result := False;
        End;

  WriteLog('Check send to CC: ' + ShortName + ' ' + FormatDatetime('dd.mm.yy', Pubdate) + ' ' + BoolToStr(Result, 'Ok', 'Not ok'));
  finally
    INI.Free;
    SL.Free;
    SLPag.Free;
    SLSec.Free;
    SLZon.Free;
    SLEdi.Free;
    SLRSec.Free;
    SLRZon.Free;
    SLREdi.Free;
    SLRDay.Free;
    Que.Free;
  end;
end;


procedure TfMain.DoImport;
var
  MyFiles,
    Pro_PageCount, Pro_SecName, Pro_SecRem, Pro_SecUnique,
    SLPageName, SLSecName, SLUnqName, SLSecRem, SL, MyFile: TStringList;
  i, NoOfPages, PrintingStart, NumberOfCopies, PrintingEnd, x,
    SecCount, s, se, r, t, j, p, Total, y, z, n, rs, fl, rr, pp: Integer;
  Tmps1, Tmps2, Tmps3, Tmps4, TmpSec, NewF,
    NewE, TmpZone, Tmps, ss: String;
  ProOK, DoChanges, b: Boolean;
  PubDate: TDateTime;
  Que, Que1, Que2: TSQLQuery;
  PPMID: LongInt;
  RegexObj: TRegExpr;
  Pro : Array of TPro;
Begin
  Try
    Que := TSQLquery.Create(nil);
    Que.DataBase := MSSQLConnection;
    Que.Transaction := SQLTrans;
    Que.PacketRecords:= -1;
    Que1 := TSQLquery.Create(nil);
    Que1.DataBase := MSSQLConnection;
    Que1.Transaction := SQLTrans;
    Que1.PacketRecords:= -1;
    Que2 := TSQLquery.Create(nil);
    Que2.DataBase := MSSQLConnection;
    Que2.Transaction := SQLTrans;
    Que2.PacketRecords:= -1;
    MyFiles  := TStringList.Create;
    MyFile   := TStringList.Create;
    RegexObj := TRegExpr.Create;

    Try
      FindAllFiles(MyFiles, Config.System.InputPath, '*.txt', False);

      If MyFiles.Count > 0 then
      Begin
        WriteLog('Fundet ' + IntToSTr(MyFiles.Count) + ' filer.');
        EventLog.Debug('Antal filer: ' + IntToSTr(MyFiles.Count));

        for i := 0 to MyFiles.Count - 1 do
        begin
          WriteLog('Fil: ' + MyFiles[i]);
          EventLog.Debug('Fil navn: ' + MyFiles[i]);

          Setlength(Pro, 0);

          //Get pubdate from file name
          RegexObj.Expression := '\w*_(\d+)-(\d+)-(\d+)';
          if RegexObj.Exec(ExtractFilename(MyFiles[i])) then
            Pubdate := EncodeDate(StrToIntDef(RegexObj.Match[1], 1), StrToIntDef(RegexObj.Match[2], 1), StrToIntDef(RegexObj.Match[3], 1)) else
            PubDate := 0;

          //Filter
          RegexObj.Expression := Config.System.FileFilter;
          if RegexObj.Exec(ExtractFilename(MyFiles[i]) + '.' + ExtractFileExt(MyFiles[i])) then
          Begin
            MyFile.LoadFromFile(MyFiles[i]);
            For fl := 0 to MyFile.Count - 1 do
            Begin
              //Væk med alle alle ugyldige linjer
              If (Length(Trim(MyFile[fl])) < 108) or (Copy(MyFile[fl],1,1) = '-') then
                Continue;

              Setlength(Pro, Length(Pro) + 1);
              p := Length(Pro) - 1;

              Pro[p].PubName        := Trim(ISO_8859_15ToUTF8(Copy(MyFile[fl], 1, 51)));
              Pro[p].Section        := Trim(Copy(MyFile[fl], 52, 5));
              Pro[p].NoOfPages      := StrToIntDef(Trim(Copy(MyFile[fl], 90, 4)), 0);
              Pro[p].PubDate        := PubDate;
              Pro[p].Zone           := '1';
              Pro[p].Edition        := '1';
              Pro[p].NumberOfCopies := 0;
              Pro[p].ProOK          := False;
              Pro[p].FileName       := Trim(ISO_8859_15ToUTF8(Copy(MyFile[fl], 95, 14)));

              For x := 0 to length(Config.Convert.ConvZone) - 1 do
                If (Uppercase(Pro[p].Zone) = Uppercase(Config.Convert.ConvZone[x].Key)) then
                Begin
                  Pro[p].Zone := Config.Convert.ConvZone[x].Value;
                  WriteLog('Konverterer zone til: ' + Pro[p].Zone);
                end;

              For x := 0 to length(Config.Convert.ConvEdi) - 1 do
                If (Uppercase(Pro[p].Edition) = Uppercase(Config.Convert.ConvEdi[x].Key)) then
                Begin
                  Pro[p].Edition := Config.Convert.ConvEdi[x].Value;
                  WriteLog('Konverterer edition til: ' + Pro[p].Edition);
                end;

              If (Uppercase(Copy(Pro[p].FileName, 12,1)) = 'T') then
                Pro[p].Trim := True else
                Pro[p].Trim := False;


              //Check om Edition er valid
              Que.Close;
              Que.SQL.Text := 'Select top 1 ID from EDITIONSNAME';
              Que.SQL.Add('Where Name=''' + Pro[p].Edition + '''');
              Que.Open;
              If Que.RecordCount <= 0 then
              Begin
                WriteLog('Ukendt edition navn: ' + Pro[p].Edition + '. Sættes til 1');
                Pro[p].Edition := '1';
              end;
            End;
            //ALt i filen indlæst

            //Sortere hvis der er nogle der skal sættes sammen
            For p := 0 to Length(Pro) - 1 do
            Begin
              For r := 0 to Length(Config.Convert.Rules) - 1 do
              Begin
                If (UTF8Uppercase(Pro[p].PubName) = UTF8Uppercase(Config.Convert.Rules[r].fName)) and
                   (Config.Convert.Rules[r].PartCopyCount) Then
                Begin
                  For rr := 0 to Length(Config.Convert.Rules) - 1 do
                  Begin
                    If (Config.Convert.Rules[r].tName[0] = Config.Convert.Rules[rr].tName[0]) and (Config.Convert.Rules[rr].AppendCopyCount) Then
                    Begin
                      For pp := 0 to Length(Pro) - 1 do
                      Begin
                        If UTF8Uppercase(Pro[pp].PubName) = UTF8Uppercase(Config.Convert.Rules[rr].fName) then
                        Begin
                          Pro[p].NoOfPages := Pro[p].NoOfPages + Pro[pp].NoOfPages;
                          Pro[pp].PubName:= '';
                        End;
                      End;
                    End;
                  End;
                End;
              End;
            end;

            WriteLog('---- Seach for rules. ----');

            //Er der en regl
            For p := 0 to Length(Pro) - 1 do
            Begin
              If Pro[p].PubName <> '' then
              Begin
                WriteLog('');
                WriteLog('----------------------------------------------------------------------------------');
                WriteLog('Produkt fundet i fil: ' + Pro[p].PubName + ', ' + FormatDatetime('dd.mm YYYY', Pro[p].PubDate) + ', sektion: ' + Pro[p].Section);

                //Er der rules for denne?
                For r := 0 to Length(Config.Convert.Rules) - 1 do
                Begin
                  If (UTF8Uppercase(Pro[p].PubName) = UTF8Uppercase(Config.Convert.Rules[r].fName)) and
                     ((Config.Convert.Rules[r].fZone = '') or (Uppercase(Pro[x].Zone)   = Uppercase(Config.Convert.Rules[r].fZone))) And
                     ((Config.Convert.Rules[r].fEdi = '') or (Uppercase(Pro[x].Edition) = Uppercase(Config.Convert.Rules[r].fEdi))) then
                  Begin
                    If (Uppercase(Pro[p].Section) = Uppercase(Config.Convert.Rules[r].fSec)) or (Config.Convert.Rules[r].fSec = '') then
                    Begin //Den skal rules
                      WriteLog('Regl fundet: ' + Config.Convert.Rules[r].fName);
                      Pro[p].PubName := Config.Convert.Rules[r].tName[0];
                      If (Length(Config.Convert.Rules[r].tZone) > 0) and (Config.Convert.Rules[r].tZone[0] <> '') then
                        Pro[p].Zone := Config.Convert.Rules[r].tZone[0];
                      If (Length(Config.Convert.Rules[r].tSec) > 0) and (Config.Convert.Rules[r].tSec[0] <> '') then
                        Pro[p].Section := Config.Convert.Rules[r].tSec[0];

                      If not Config.Convert.Rules[r].Flying then
                      Begin
                        Que1.Close;
                        Que1.SQL.Text := 'Select ID, Long_Name, Long_Name_Add, Short_Name, Udg_Dato, Pro_PageCount, Pro_SecName, Pro_SecRem, Zones, Pro_Editions_Unique from DATA';
                        Que1.SQL.Add(   'Where CAST(UDG_DATO AS Date) = ''' + FormatDatetime('YYYY-MM-DD', Pro[p].PubDate) + '''');
                        Que1.SQL.Add(   'And Short_Name = ''' + Pro[p].PubName + '''');
                        Que1.SQL.Add(   'And Zones = ''' + IntToStr(ZoneToID(Pro[p].Zone)) + '''');
                        //Que1.SQL.SaveToFile('301.sql');
                        Que1.Open;

                        If (Que1.RecordCount > 0) then
                        Begin
                          WriteLog('Produktion: ' + Pro[p].PubName + ', ' + FormatDatetime('dd.mm YYYY', Pro[p].PubDate));
                        End else
                        Begin
                          If Config.Convert.Rules[r].CreateMissing then
                          Begin
                            PPMID := CreateNewPro(Pro[p].PubDate, 0, Pro[p].PubName);
                            Que1.Close;
                            Que1.SQL.Text := 'Select ID, Long_Name, Long_Name_Add, Short_Name, Udg_Dato, Pro_PageCount, Pro_SecName, Pro_SecRem, Zones, Pro_Editions_Unique from DATA';
                            Que1.SQL.Add(   'Where ID = ''' + IntToStr(PPMID) + '''');
                            //Que1.SQL.SaveToFile('302.sql');
                            Que1.Open;
                            WriteLog('Ny produktion oprettet: ' + Pro[p].PubName + ', ' + FormatDatetime('dd.mm YYYY', Pro[p].PubDate) + ' fra regl: ' + Config.Convert.Rules[r].fName);
                          End else
                          Begin
                            PPMID:= -1;
                            WriteLog('Produktion ikke fundet og oprettes ikke: ' + Pro[p].PubName + ', ' + FormatDatetime('dd.mm YYYY', Pro[p].PubDate) + ' fra regl: ' + Config.Convert.Rules[r].fName);
                          End;
                        End;
      ;
                        While not Que1.EOF do
                        Begin //Den findes, opdatere
                          DoChanges := False; //

                          PPMID := Que1.FieldByName('ID').AsInteger;
                          Pro_PageCount := TStringList.Create;
                          Pro_PageCount.Delimiter := ',';
                          Pro_SecName := TStringList.Create;
                          Pro_SecName.Delimiter := ',';
                          Pro_SecRem := TStringList.Create;
                          Pro_SecRem.Delimiter := ',';
                          Pro_SecUnique := TStringList.Create;
                          Pro_SecUnique.Delimiter := ';';          //OBS OBS ";"
                          Pro_SecUnique.StrictDelimiter := True;
                          Pro_PageCount.DelimitedText := Que1.FieldByName('Pro_PageCount').AsString;
                          Pro_SecName.DelimitedText   := Que1.FieldByName('Pro_SecName').AsString;
                          Pro_SecRem.DelimitedText    := Que1.FieldByName('Pro_SecRem').AsString;
                          Pro_SecUnique.DelimitedText := Que1.FieldByName('Pro_Editions_Unique').AsString;
                          TmpS := Que1.FieldByName('Pro_Editions_Unique').AsString;

                          // Først skal der evt rettes op i længde
                          EventLog.Debug('Found ' + Que1.FieldByName('Short_Name').AsString  + ' in PPMID:' + IntToStr(PPMID));
                          EventLog.Debug('Checking length before. Pro_PageCount=' + Pro_PageCount.CommaText +
                                                           ' Pro_SecName=' + Pro_SecName.CommaText +
                                                           ' Pro_SecRem=' + Pro_SecRem.CommaText +
                                                           ' Pro_SecUnique=' + Pro_SecUnique.CommaText
                                                           );

                          t := Pro_PageCount.Count;
                          While Pro_SecName.Count < t do
                            Pro_SecName.Add('');

                          While Pro_SecRem.Count < t do
                            Pro_SecRem.Add('');

                          While Pro_SecUnique.Count < t do
                            Pro_SecUnique.Add('');

                          While Pro_SecName.Count > t do
                            Pro_SecName.Delete(Pro_SecName.Count - 1);

                          While Pro_SecUnique.Count > t do
                            Pro_SecUnique.Delete(Pro_SecUnique.Count - 1);

                          While Pro_SecRem.Count > t do
                            Pro_SecRem.Delete(Pro_SecRem.Count - 1);

                          EventLog.Debug('Checking length after. Pro_PageCount=' + Pro_PageCount.CommaText +
                                                         ' Pro_SecName=' + Pro_SecName.CommaText +
                                                         ' Pro_SecRem=' + Pro_SecRem.CommaText +
                                                         ' Pro_SecUnique=' + Pro_SecUnique.CommaText
                                                         );
                          j := Pro_SecName.IndexOf(Pro[p].Section);
                          If j < 0 then
                          Begin         //Sec er der ikke. opret.
                            If (Pro_SecName.Count = 1) and (Pro_SecName[0] = '') then
                            Begin
                              Pro_SecName[0] := Pro[p].Section;
                            end else
                            Begin
                              Pro_PageCount.Add('0');
                              Pro_SecName.Add(Pro[p].Section);
                              Pro_SecRem.Add('');
                              Pro_SecUnique.Add('');

                            end;
                            j := Pro_SecName.IndexOf(Pro[p].Section);
                            DoChanges := True;
                            WriteLog('Sektion ' + Pro_SecName[j] + ' mangler i produkt. Oprettet.');
                          End;

                          If j >= 0 then //Og det skal den for pokker være
                          Begin
                            WriteLog('Sektion: [' + Pro_SecName[j] + '] fundet i ID:' + Que1.FieldByName('ID').AsString + ' ' + Que1.FieldByName('Long_Name').AsString + ' ' + Que1.FieldByName('Long_Name_Add').AsString + ' Zone ' + IdToZone(Que1.FieldByName('Zones').AsInteger) + ' Section ' + Pro_SecName[j]);

                            If (Not Config.Convert.Rules[r].SetOnlyUnique) and (Pro_PageCount[j] <> IntToStr(Pro[p].NoOfPages)) then
                            Begin
                              DoChanges := True;
                              WriteLog('Ændrer sidetal på sektion: ' + Pro_SecName[j] + ' ' + Pro_PageCount[j] + '->' + IntToStr(Pro[p].NoOfPages));
                              If (Config.Convert.Rules[r].AppendCopyCount) then
                              Begin
                                Pro_PageCount[j] := IntToStr(StrToIntDef(Pro_PageCount[j], 0) + Pro[p].NoOfPages);
                              End else
                                Pro_PageCount[j] := IntToStr(Pro[p].NoOfPages);
                            End;

                            If Pro_PageCount[j] = '' Then
                              Pro_PageCount[j] := '0';

                            If (Config.Convert.Rules[r].SecRem <> '') and (Pro_SecRem[j] <> Config.Convert.Rules[r].SecRem) then
                            Begin
                              DoChanges := True;
                              WriteLog('Indsætter sec remark: ' + Pro_SecRem[j] + '->' + Config.Convert.Rules[r].SecRem);
                              Pro_SecRem[j] := Config.Convert.Rules[r].SecRem;
                            End;

                            {If Not Config.Convert.Rules[r].SetOnlyPagecount then
                            Begin
                              If Config.Convert.Rules[r].ForceCommon then
                              Begin
                                If Pro_SecUnique[j] <> '' then DoChanges := True;
                                WriteLog('Set Unique empty.');
                                Pro_SecUnique[j] := '';
                              End Else
                              Begin
                                If Pro_SecUnique[j] <> DocUniquePage[s] then DoChanges := True;
                                WriteLog('Set Unique: ' + Pro_SecUnique[j] + '->' + DocUniquePage[s]);
                                Pro_SecUnique[j]   := DocUniquePage[s];
                              End;

                              If Config.Convert.Rules[r].ForceUnique then
                              Begin
                                WriteLog('Set Force unique on all pages.');
                                SL := TStringlist.Create;
                                For x := 1 to Pro[p].NoOfPages do
                                  SL.Add(Pro_SecUnique[j]);

                                If Pro_SecUnique[j] <> SL.CommaText then
                                  DoChanges := True;
                                Pro_SecUnique[j] := SL.CommaText;
                                SL.Free;
                              End;
                            End;}

                            {For x := 0 to Pro_PageCount.Count - 1 do
                            Begin
                              For y := x to Pro_PageCount.Count - 1 do
                              Begin
                                If Pro_SecName[x] > Pro_SecName[y] then
                                Begin
                                  TmpS := Pro_PageCount[x];
                                  Pro_PageCount[x] := Pro_PageCount[y];
                                  Pro_PageCount[y] := TmpS;

                                  TmpS := Pro_SecName[x];
                                  Pro_SecName[x] := Pro_SecName[y];
                                  Pro_SecName[y] := TmpS;

                                  TmpS := Pro_SecRem[x];
                                  Pro_SecRem[x] := Pro_SecRem[y];
                                  Pro_SecRem[y] := TmpS;

                                  TmpS := Pro_SecUnique[x];
                                  Pro_SecUnique[x] := Pro_SecUnique[y];
                                  Pro_SecUnique[y] := TmpS;

                                end;
                              End;
                            end;//Så er de sorteret}

                            //Tæller alle sider sammen
                            Total := 0;
                            For t := 0 to Pro_PageCount.Count - 1 do
                              Inc(Total, StrToIntDef(Pro_PageCount[t], 0));

                            If DoChanges then
                            Begin
                              Que2.Close;
                              Que2.SQL.Text := 'UPDATE DATA SET';
                              Que2.SQL.Add('  Pro_PageCount ='''  + Pro_PageCount.CommaText + '''');
                              Que2.SQL.Add(' ,SIDE_ANTAL    ='''  + IntToStr(Total)         + '''');
                              Que2.SQL.Add(' ,Pro_SecName   ='''  + Pro_SecName.CommaText   + '''');
                              Que2.SQL.Add(' ,DummyStatus   = 1');
                              Que2.SQL.Add(' ,Pro_Editions_Unique ='''  +  Pro_SecUnique.DelimitedText  + '''');
                              //If (NOT Config.RulesProConv[r].SetOnlyPageCount) or (Config.RulesProConv[r].UseForceCommon) then
                              Que2.SQL.Add(' ,Pro_SecRem    ='''  + Pro_SecRem.CommaText    + '''');
                              Que2.SQL.Add(' WHERE ID = ''' + IntToStr(PPMID) + '''');
                              //Que2.SQL.SaveToFile('300.sql');
                              If not tbSimulate.Down then
                                Que2.ExecSQL;

                              EventLog.Debug('End updating ' + Pro[p].PubName);

                              If ReadyToSendtToCC(PPMID, false) then
                              Begin
                                Que2.Close;
                                Que2.SQL.Text := 'UPDATE DATA SET';
                                Que2.SQL.Add(' SEND_TO_CC = 1');
                                Que2.SQL.Add(' WHERE ID = ''' + IntToStr(PPMID) + '''');
                                If not tbSimulate.Down then
                                  Que2.ExecSQL;
                              End;


                            end Else
                              WriteLog('Ingen ændringer. Alt passer.');
                          end;

                          Pro_PageCount.Free;
                          Pro_SecName.Free;
                          Pro_SecRem.Free;
                          Pro_SecUnique.Free;
                          Que1.Next;

                        If not tbSimulate.Down then
                          SQLTrans.commit else
                          WriteLog('Simulate');
                        end;

                        End else                        //Det er en flyvende rule
                        Begin
                          Que2.Close;
                          Que2.SQL.Text := 'Select ID, Long_Name, Long_Name_Add, Udg_Dato, ProType, Zones, Pro_PageCount, Pro_SecName, Pro_SecRem, Pro_Editions_Unique from DATA with (NOLOCK)';
                          Que2.SQL.Add(   'Where CAST(UDG_DATO AS Date) = ''' + FormatDatetime('YYYY-MM-DD', PubDate) + '''');
                          Que2.SQL.Add(   'And Upper(Pro_SecRem) like ''%#' + Uppercase(Pro[p].PubName) + '%''');
                          //Que2.SQL.SaveToFile('1100.sql');
                          Que2.Open;

                          If Que2.RecordCount > 0 then
                            WriteLog(Uppercase(Pro[p].PubName) + ' fundet i ' + Que2.FieldByName('Long_Name').AsString);

                          While not Que2.EOF do
                          Begin
                            DoChanges := False;
                            PPMID := Que2.FieldByName('ID').AsInteger;

                            SLPageName := TStringList.Create;
                            SLPageName.Delimiter     := ',';
                            SLPageName.DelimitedText := Que2.FieldByName('Pro_PageCount').AsString;
                            SLSecName := TStringList.Create;
                            SLSecName.Delimiter     := ',';
                            SLSecName.DelimitedText := Que2.FieldByName('Pro_SecName').AsString;
                            SLUnqName := TStringList.Create;
                            SLUnqName.Delimiter     := ';';
                            SLUnqName.StrictDelimiter := True;
                            SLUnqName.DelimitedText := Que2.FieldByName('Pro_Editions_Unique').AsString;
                            SLSecRem := TStringList.Create;
                            SLSecRem.Delimiter     := ',';
                            SLSecRem.DelimitedText := Que2.FieldByName('Pro_SecRem').AsString;

                            While SLUnqName.Count < SLPageName.Count do
                              SLUnqName.Add('');
                            While SLSecRem.Count < SLPageName.Count do
                              SLSecRem.Add('');

                            For x := 0 to SLSecName.Count - 1 do
                            Begin
                              Tmps := SLSecRem[x];
                              If Pos('#' + Uppercase(Pro[p].PubName), Uppercase(SLSecRem[x])) = 1 then
                              Begin
                                TmpZone := IdToZone(Que2.FieldByName('Zones').AsInteger);  //Som udgangspunkt tages produktets zone
                                If (Pos(':Z', Uppercase(SLSecRem[x])) > 0) Then
                                  TmpZone := Copy(SLSecRem[x], Pos(':Z', SLSecRem[x]) + 2, 1); //Okay den er ændret

                                If (TmpZone = Pro[p].Zone) then
                                Begin
                                  TmpSec := SLSecName[x];
                                  If (Pos(':S', Uppercase(SLSecRem[x])) > 0) then
                                    TmpSec := Copy(SLSecRem[x], Pos(':S', SLSecRem[x]) + 2, 1);

                                  //idxSec := DocSecName.IndexOf(TmpSec);

                                  If (Pro[p].Section = TmpSec) then                            //Vi har den
                                  Begin
                                    WriteLog('Found inserted section: ID=' + Que2.FieldByName('ID').AsString + ' ' + Que2.FieldByName('Long_Name').AsString + ' ' + Que2.FieldByName('Long_Name_Add').AsString + ' Zone=' + IdToZone(Que2.FieldByName('Zones').AsInteger) + ' Section=' + SLSecName[x] );
                                    If SLPageName[x] <> IntToSTr(Pro[p].NoOfPages) then
                                      DoChanges := True;
                                    SLPageName[x] := IntToSTr(Pro[p].NoOfPages);

                                    t := 0; //Total antal sider
                                    For y := 0 to SLPageName.Count - 1 do
                                      Inc(t, StrToIntDef(SLPageName[y], 0));

                                    {If SLUnqName[x] <> DocUniquePage[s] then
                                      DoChanges := True;
                                    SLUnqName[x] := DocUniquePage[s];}

                                    If DoChanges then
                                    Begin
                                      Que1.Close;
                                      Que1.SQL.Text := 'UPDATE DATA SET';
                                      Que1.SQL.Add('  SIDE_ANTAL = '''          + IntToSTr(t) + '''');
                                      Que1.SQL.Add(' ,DummyStatus = 1');
                                      Que1.SQL.Add(' ,Pro_PageCount = '''       + SLPageName.CommaText + '''');
                                      Que1.SQL.Add(' ,Pro_Editions_Unique = ''' + SLUnqName.DelimitedText + '''');

                                      Que1.SQL.Add(' Where ID = '''             + IntToStr(PPMID) + '''');
                                      If not tbSimulate.Down then
                                        Que1.ExecSQL else
                                        WriteLog('SIMULATE update');

                                      WriteLog('Pagecount=' + IntToSTr(t) + ' Pages=' + SLPageName.CommaText + ' Unique=' + SLUnqName.DelimitedText);

                                      If ReadyToSendtToCC(PPMID, false) then
                                      Begin
                                        Que1.Close;
                                        Que1.SQL.Text := 'UPDATE DATA SET';
                                        Que1.SQL.Add(' SEND_TO_CC = 1');
                                        Que1.SQL.Add(' Where ID = '''             + IntToStr(PPMID) + '''');
                                        If not tbSimulate.Down then
                                          Que1.ExecSQL else
                                      End;

                                    end else
                                      WriteLog('Nothing to changes. All match.');
                                  end;
                                end;
                              end;
                            end;

                            SLPageName.Free;
                            SLSecName.Free;
                            SLUnqName.Free;
                            SLSecRem.Free;
                            Que2.Next;
                          end;
                        end;
                      end;

                    If not tbSimulate.Down then
                      SQLTrans.commit;
                  end;
                end; //End of rules
              End; //Hvis den ikke er tom
            End; //All product finish

            NewF := ExtractFileName(MyFiles[i]);
            NewE := ExtractFileExt(MyFiles[i]);
            NewF := Copy(NewF, 1, Pos('.', NewF) - 1);
            NewF := NewF + ' ' + FormatDatetime('dd-mm-yy HH-NN-SS-ZZZ', Now) + NewE;

            If Not RenameFile(MyFiles[i], Config.System.DonePath + '\' + NewF) then
            Begin
              Writelog('Can''t move to done folder.' + MyFiles[i] + ' to ' + Config.System.DonePath + '\' + NewF);
              DeleteFile(MyFiles[i]);
            end;
          end else
          Begin
            {Writelog('No match' + MyFiles[i]);
            NewF := ExtractFileName(MyFiles[i]);
            NewE := ExtractFileExt(MyFiles[i]);
            NewF := NewF + NewE;
            RenameFile(MyFiles[i], Config.System.NoMatchPath + '\' + NewF);}
          End;
        End;
      End;  //Der er i ngen filer

      except
        on E: Exception do
        begin
          StatusBar.Panels.Items[1].Text := E.Message;
          EventLog.Error(E.Message);
        end;
      end;

  finally
    Que.Free;
    Que1.Free;
    Que2.Free;

    MyFiles.Free;
    MyFile.Free;
    RegexObj.Free;
  end;

end;

procedure TfMain.Timer1Timer(Sender: TObject);
begin
  Timer1.Enabled := False;
  tbStatus.ImageIndex := 9;
  DoImport;
  tbStatus.ImageIndex := 7;
  Timer1.Enabled := True;
end;

procedure TfMain.ToolButton1Click(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TfMain.ToolButton2Click(Sender: TObject);
begin

end;


procedure TfMain.ToolButton3Click(Sender: TObject);
begin
  Timer1.Enabled:= False;
  fConvert.ShowModal;
  Readconfig;
  Timer1.Enabled:= True;
end;

procedure TfMain.tbSimulateClick(Sender: TObject);
begin
  WriteLog('');
  WriteLog('');
  If tbSimulate.Down then
    WriteLog('                                -------- SIMULATE MODE ---------')
  else
    WriteLog('                                --------- NORMAL MODE ----------');
  WriteLog('');
  WriteLog('');

end;

procedure TfMain.ToolButton7Click(Sender: TObject);
begin
  FormSetProFinish.ShowModal;
end;


function TfMain.Encrypt(chaine: string): string;
var
  en : TBlowFishEncryptStream;
  s1 : TStringStream;
begin
  s1 := TStringStream.Create('');
  en := TBlowFishEncryptStream.Create('Infra2Logic', s1);
  en.Write(chaine[1], Length(chaine));
  en.Free;
  result := s1.datastring;
  s1.Free;
end;

function TfMain.Decrypt(chaine: string): string;
var
  de   : TBlowFishDeCryptStream;
  s2   : TStringStream;
  temp : String;
begin
  Try
    s2 := TStringStream.Create(chaine);
    de := TBlowFishDecryptStream.Create('Infra2Logic', s2);
    SetLength(temp, s2.Size);
    de.Read(temp[1], s2.Size);
    result := temp;

  finally
    s2.Free;
    de.Free;
  end;
end;



procedure TfMain.ReadConfig;
var
  INI: TIniFile;
  Que: TSQLQuery;
  slZone, slEdi, slRule, SL, SL1: TStringList;
  i, x, p: Integer;
begin
  Timer1.Enabled:= False;
  tbStatus.ImageIndex := 9;
  Config.System.AppData      := GetEnvironmentVariableUTF8('APPDATA') + '\' +  ExtractFileNameOnly(Application.ExeName);
  WriteLog('System path: ' + Config.System.AppData);
  Config.System.HomePath     := GetEnvironmentVariableUTF8('HOMEPATH');
  WriteLog('Home path: ' + Config.System.HomePath);
  Config.System.Desktop      := Config.System.HomePath + '\Desktop';
  WriteLog('Desktop path: ' + Config.System.Desktop);
  Config.System.Comp_Name    := GetEnvironmentVariableUTF8('COMPUTERNAME');
  WriteLog('Computer name: ' + Config.System.Comp_Name);

  Config.INIfileName := Config.System.AppData + '\' + ExtractFileNameOnly(Application.ExeName) + '.INI';
  Config.ConvertFileName := Config.System.AppData + '\Convert.INI';
  Config.CheckFinishfileName := Config.System.AppData + '\ProFinishToCC.ini';
  Config.LogfileName := Config.System.AppData + '\' + ExtractFileNameOnly(Application.ExeName) + '.LOG';

  //Har vi en ini fil i appdata dir?
  If NOT fileExistsUTF8(Config.INIfileName) then
  Begin
    //Det havde vi så ikke
    Showmessage('INI filename don''t exists: [' + Config.INIfileName + ']');
    ForceDirectoriesUTF8(Config.System.AppData);
    If fileExistsUTF8(Application.Location + '\' + ExtractFileNameOnly(Application.ExeName) + '.ini') then
    Begin
      If NOT CopyFile(Application.Location + '\' + ExtractFileNameOnly(Application.ExeName) + '.ini', Config.INIfileName) then
      Begin
        ShowMessage('The .ini file are missing');
        Application.Terminate;
      End else
        ShowMessage('INI file copy from app dir.');
    End else
    Begin //Der er ikke en skid, jeg laver en ny
      INI := TINIFile.Create(Config.INIfileName);
      INI.WriteString('DB', 'Database',      'PPM');
      INI.WriteString('DB', 'DB Host',       'localhost');
      INI.WriteBool(  'DB', 'Use win login', False);
      INI.WriteString('DB', 'DB Username',   '');
      INI.WriteString('DB', 'DB Password',   '');
      INI.WriteInteger('System', 'Debug level',  0);
      INI.WriteString('Path', 'Input',   '');
      INI.WriteString('Path', 'Output',   '');
      INI.WriteString('Path', 'Done',   '');
      INI.WriteString('Path', 'Filter',   '');

      INI.Free;
      ShowMessage('New INI config file created. Please edit: ' + Config.INIfileName);
    End;
  end;


  If NOT FileExistsUTF8(Config.INIfileName) then   //Hvis vi stadig ikke har den så ud. Så må de også selv om det.
  Begin
    Showmessage('INI file missing. [' + Config.INIfileName + ']');
    Application.Terminate;
  end;


  EventLog.Active          := False;
  EventLog.LogType         := ltFile;
  EventLog.FileName        := UTF8ToSys(Config.LogfileName);
  EventLog.TimeStampFormat := 'DD-MM-YY hh:mm:ss';
  EventLog.AppendContent   := True;
  EventLog.Active          := True;

  INI := TINIFile.Create(Config.INIfileName);

  Config.System.DebugLevel  := INI.ReadInteger('System', 'Debug level',  0);
  WriteLog('Debug level: ' + IntToStr(Config.System.DebugLevel));
  If Config.System.DebugLevel >= 0 then
  Begin
    EventLog.Debug('Get enviroment:');
    EventLog.Debug('Appdata= ' + Config.System.AppData);
    EventLog.Debug('Home path= ' + Config.System.HomePath);
    EventLog.Debug('Computername=' + Config.System.Comp_Name);
    EventLog.Debug('INI filename= ' + Config.INIfileName);
    EventLog.Debug('Log fileName= ' + Config.LogfileName);
  end;

  Config.System.DonePath := INI.ReadString('Path', 'Done', '');
  WriteLog('Done path: ' + Config.System.DonePath);
  Config.System.InputPath := INI.ReadString('Path', 'Input', '');
  WriteLog('Input path: ' + Config.System.InputPath);
  Config.System.OutputPath := INI.ReadString('Path', 'Output', '');
  WriteLog('Output path: ' + Config.System.OutputPath);
  Config.System.FileFilter := INI.ReadString('Path', 'Filter', '');
  WriteLog('Filefilter: ' + Config.System.FileFilter);
  Config.System.NoMatchPath := INI.ReadString('Path', 'No match', '');
  WriteLog('No match path: ' + Config.System.NoMatchPath);


  Config.SQLdb.TableName    := INI.ReadString('DB',    'Database', '');
  Config.SQLdb.HostName     := INI.ReadString('DB',    'DB Host',     '');

  If not INI.ReadBool('DB', 'Use win login', False) then
  Begin
    Config.SQLdb.UserName     := INI.ReadString('DB',    'DB Username',     '');
    If INI.ReadString('DB',  'DB Password', '') <> '' then
      Config.SQLdb.UserPassword := Decrypt(INI.ReadString('DB',  'DB Password', '')) else
      Config.SQLdb.UserPassword := '';
      //SQLdb.UserPassword := INI.ReadString('DB',  'DB Password', '');
  End Else
  Begin
    Config.SQLdb.UserName     := '';
    Config.SQLdb.UserPassword := '';
  End;


  Try
    SQLDBLibraryLoader.Enabled   := True;
    MSSQLConnection.UserName     := Config.SQLdb.UserName;
    MSSQLConnection.Password     := Config.SQLdb.UserPassword;
    MSSQLConnection.HostName     := Config.SQLdb.HostName;
    MSSQLConnection.DatabaseName := Config.SQLdb.TableName;
    MSSQLConnection.Connected    := True;
  except
    on E: Exception do
    Begin
      StatusBar.Panels.Items[1].Text := 'Fail to connect to SQL db. ';
      Config.SQLdb.UserPassword      := INI.ReadString('DB',  'DB Password', '');   //Vi prøver uden encrypt
      MSSQLConnection.Password       := Config.SQLdb.UserPassword;

      MSSQLConnection.Connected    := True;
      If MSSQLConnection.Connected then
      Begin
        StatusBar.Panels.Items[1].Text := 'Password not encrypted in ini file. Fixed.';
        EventLog.Error('Password not encrypted in ini file. Fixed.');
        INI.WriteString('DB',  'DB Password', Encrypt(Config.SQLdb.UserPassword));
      End;
    end;
  end;

  Que := TSQLquery.Create(nil);
  Que.DataBase := MSSQLConnection;
  Que.Transaction := SQLTrans;
  Que.PacketRecords:= -1;

  Que.Close;
  Que.SQL.Text := 'Select Top 1 Master_Dok from Config';
  Que.Open;

  Que.Close;
  Que.SQL.Text := 'Select Name, ID from ZonesNAME WITH (NOLOCK)';
  Que.Open;
  SetLength(Config.Names.Zone, Que.RecordCount);
  while not Que.EOF do
  begin
    Config.Names.Zone[Que.RecNo -  1].Name := Que.FieldByName('NAME').AsString;
    Config.Names.Zone[Que.RecNo -  1].ID   := Que.FieldByName('ID').AsInteger;
    Que.Next;
  end;

  INI.Free;

  INI := TINIFile.Create(fMain.Config.ConvertFileName);
  slZone := TStringlist.Create;
  INI.ReadSection('Zone', slZone);
  Setlength(Config.Convert.ConvZone, slZone.Count);
  For i := 0 to slZone.Count - 1 do
  Begin
    Config.Convert.ConvZone[i].Key := slZone[i];
    Config.Convert.ConvZone[i].Value := INI.ReadString('Zone', slZone[i], '');
  end;
  slZone.Free;

  slEdi := TStringlist.Create;
  INI.ReadSection('Edition', slEdi);
  Setlength(Config.Convert.ConvEdi, slEdi.Count);
  For i := 0 to slEdi.Count - 1 do
  Begin
    Config.Convert.ConvEdi[i].Key := slEdi[i];
    Config.Convert.ConvEdi[i].Value := INI.ReadString('Edition', slEdi[i], '');
  end;
  slEdi.Free;

  slRule := TStringlist.Create;
  SL := TStringlist.Create;
  SL.Delimiter:= ';';
  SL.StrictDelimiter:= True;
  SL1 := TStringlist.Create;
  SL1.Delimiter:= ',';
  SL1.StrictDelimiter:= True;
  INI.ReadSection('Rule', slRule);
  Setlength(Config.Convert.Rules, slRule.Count);
  For i := 0 to slRule.Count - 1 do
  Begin
    SL.Clear;
    SL1.Clear;

    SL.DelimitedText:= INI.ReadString('Rule', slRule[i], '');

    While SL.Count < 24 do
      SL.Add(''); //Dette blot hvis der ikke er nok col

    If StrToBoolDef(SL[0], False) then
    Begin
      Config.Convert.Rules[i].fName := SL[1];
      Config.Convert.Rules[i].fSec := SL[2];
      Config.Convert.Rules[i].fEdi := SL[3];
      Config.Convert.Rules[i].fZone := SL[4];
      SL1.DelimitedText := SL[6];
      Setlength(Config.Convert.Rules[i].tName, SL1.Count);
      For x := 0 to SL1.Count - 1 do
        Config.Convert.Rules[i].tName[x] := SL1[x];
      SL1.DelimitedText := SL[7];
      Setlength(Config.Convert.Rules[i].tSec, SL1.Count);
      For x := 0 to SL1.Count - 1 do
        Config.Convert.Rules[i].tSec[x] := SL1[x];
      SL1.DelimitedText := SL[8];
      Setlength(Config.Convert.Rules[i].tEdi, SL1.Count);
      For x := 0 to SL1.Count - 1 do
        Config.Convert.Rules[i].tEdi[x] := SL1[x];
      SL1.DelimitedText := SL[9];
      Setlength(Config.Convert.Rules[i].tZone, SL1.Count);
      For x := 0 to SL1.Count - 1 do
        Config.Convert.Rules[i].tZone[x] := SL1[x];
      Config.Convert.Rules[i].SecRem := SL[10];
      Config.Convert.Rules[i].ForceUnique := StrToBoolDef(SL[11], False);
      Config.Convert.Rules[i].ForceCommon := StrToBoolDef(SL[14], False);;
      Config.Convert.Rules[i].SetOnlyUnique    := StrToBoolDef(SL[17], False);;
      Config.Convert.Rules[i].SetOnlyPagecount := StrToBoolDef(SL[18], False);;
      Config.Convert.Rules[i].CreateMissing    := StrToBoolDef(SL[19], False);;
      Config.Convert.Rules[i].Flying           := StrToBoolDef(SL[20], False);
      Config.Convert.Rules[i].AppendCopyCount  := StrToBoolDef(SL[21], False);
      Config.Convert.Rules[i].PartCopyCount    := StrToBoolDef(SL[22], False);
    end;
  end;
  slRule.Free;
  SL.Free;
  SL1.Free;
  INI.Free;

  If (MSSQLConnection.Connected) and
     (Config.System.InputPath <> '') And (Config.System.DonePath <> '') and
     {(Config.System.OutputPath <> '') and} (Config.System.FileFilter <> '')
  then
  Begin
    Timer1.Enabled:= True;
    tbStatus.ImageIndex := 7;
  end else
    tbStatus.ImageIndex := 8;

end;

procedure TfMain.FormCreate(Sender: TObject);
begin
  Config := TConfig.Create(Nil);
  ReadConfig;

  RunningFile := FileCreate(Config.System.AppData + '\Running');
end;

procedure TfMain.tbStatusClick(Sender: TObject);
begin
//  Timer1.Enabled := not Timer1.Enabled;
end;


end.

