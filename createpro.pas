unit CreatePro;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, sqldb, DateUtils;

function CreateNewPro(Pubdate : TDate; ProID : Integer; ShortName : String) : Integer ;

implementation

Uses
  uMain;


function CreateNewPro(Pubdate : TDate; ProID : Integer; ShortName : String) : Integer ;
var
  DefLongName, DefShortName, DefOutputAlias,
  DefAcc                    , DefEditions, DefZones, DefFals, DefLogRem,
  DefRepRem, DefProRem, DefPakRem, DefDisRem, DefSecName, DefAfvId,
  DefPak_Pattern: String;
  DefID, Ann, DoAnnHol,
  Mat, DoMatHol, Ctp,
  DoCtpHol, Dis, DoDisHol,
  Pak, DoPakHol, Pro, DoProHol,
  Afl, DoAflHol, WeekKind,
  DefSats, DefPress, DefPaper,
  DefIL                     , DefPubKind, DefOplag, DefOmreg, DefCrimp,
  DefHeigth, DefWidth, Rel, DoRelHol, DefPakTemplate: LongInt;
  Ann_Time, Mat_Time,
  Ctp_Time, Dis_Time,
  Pak_Time, Pro_Time,
  Afl_Time                  : TTime;
  DefProTime, DefCtpTime, DefAflTime, DefDisTime, DefMatTime,
    DefPakTime, DefProTimeEnd, DefPakTimeEnd, Pro_Time_End, Rel_Time,
    DefRelTime, Pak_Time_End: TDateTime;
  DefUseTrimbox : Boolean;
  FS: TFormatSettings;
  dt: TDate;
  SL: TStringList;
  x: Integer;
  Que, Que1: TSQLQuery;
Begin
  Try
    FS := DefaultFormatSettings;
    FS.DateSeparator   := '-';
    FS.TimeSeparator := ':';
    FS.ShortDateFormat := 'YYYY-MM-DD';
    FS.ShortTimeFormat := 'HH:NN';

    Que := TSQLquery.Create(nil);
    Que.DataBase := fMain.MSSQLConnection;
    Que.Transaction := fMain.SQLTrans;
    Que.PacketRecords:= -1;

    Que1 := TSQLquery.Create(nil);
    Que1.DataBase := fMain.MSSQLConnection;
    Que1.Transaction := fMain.SQLTrans;
    Que1.PacketRecords:= -1;

    Que.Close;
    Result := -1;
    If ProID > 0 then
    Begin
      Que.Close;
      Que.SQL.Text := 'Select TOP 1 * from PRO_NAME WITH (NOLOCK)';
      Que.SQL.Add('    WHERE ID = ''' + IntToStr(ProID) + '''');
      Que.Open;
    End else
    If (ShortName <> '') And (ProID <= 0) then
    Begin
      Que.Close;
      Que.SQL.Text := 'Select TOP 1 * from PRO_NAME WITH (NOLOCK)';
      Que.SQL.Add('    WHERE SHORT_NAME = ''' + ShortName + '''');
      Que.Open;
    End;

    If Que.RecordCount > 0 then
    Begin
      DefLongName    := Que.FieldByName('Long_Name').AsString;
      DefShortName   := Que.FieldByName('Short_Name').AsString;

      // ### NAN 2024-05-31
      DefOutputAlias := Que.FieldByName('Output_Short_Name').AsString;
      // ###

      DefID          := Que.FieldByName('ID').AsInteger;
      DefPubKind     := Que.FieldByName('Pub_Kind').AsInteger;
      DefEditions    := Que.FieldByName('DEF_EDITIONS').AsString;
      DefZones       := Que.FieldByName('DEF_ZONES').AsString;
      DefPakTemplate := StrToIntDef(Que.FieldByName('DEF_Pak_TemplateID').AsString, 0);
      DefPak_Pattern := Que.FieldByName('Pak_Pattern').AsString;

      {If DefEditions = '' then
        DefEditions := IntToStr(EditiontoID('1'));
      If DefZones    = '' then
        DefZones    := IntToStr(ZonetoID('1'));}

      Ann          := Que.FieldByName('DEF_ANN_DEADLINE_DAGE').AsInteger;
      Ann_Time     := TimeOf(Que.FieldByName('DEF_ANN_DEADLINE').AsDateTime);
      DoAnnHol     := Que.FieldByName('DEF_ANN_DEADLINE_HELL').AsInteger;
      Rel          := Que.FieldByName('DEF_REL_DEADLINE_DAGE').AsInteger;
      Rel_Time     := TimeOf(Que.FieldByName('DEF_REL_DEADLINE').AsDateTime);
      DoRelHol     := Que.FieldByName('DEF_REL_DEADLINE_HELL').AsInteger;
      Mat          := Que.FieldByName('DEF_Mat_DEADLINE_DAGE').AsInteger;
      Mat_Time     := TimeOf(Que.FieldByName('DEF_Mat_DEADLINE').AsDateTime);
      DoMatHol     := Que.FieldByName('DEF_Mat_DEADLINE_HELL').AsInteger;
      Ctp          := Que.FieldByName('DEF_CTP_DEADLINE_DAGE').AsInteger;
      Ctp_Time     := TimeOf(Que.FieldByName('DEF_CTP_DEADLINE').AsDateTime);
      DoCtpHol     := Que.FieldByName('DEF_CTP_DEADLINE_HELL').AsInteger;
      Dis          := Que.FieldByName('DEF_DIS_DEADLINE_DAGE').AsInteger;
      Dis_Time     := TimeOf(Que.FieldByName('DEF_DIS_DEADLINE').AsDateTime);
      DoDisHol     := Que.FieldByName('DEF_DIS_DEADLINE_HELL').AsInteger;
      Pak          := Que.FieldByName('DEF_PAK_DEADLINE_DAGE').AsInteger;
      Pak_Time     := TimeOf(Que.FieldByName('DEF_PAK_DEADLINE').AsDateTime);
      Pak_Time_End := TimeOf(Que.FieldByName('DEF_PAK_END_DEADLINE').AsDateTime);
      DoPakHol     := Que.FieldByName('DEF_PAK_DEADLINE_HELL').AsInteger;
      Pro          := Que.FieldByName('DEF_PRO_DEADLINE_DAGE').AsInteger;
      Pro_Time     := TimeOf(Que.FieldByName('DEF_PRO_DEADLINE').AsDateTime);
      Pro_Time_End := TimeOf(Que.FieldByName('DEF_PRO_End_DEADLINE').AsDateTime);
      DoProHol     := Que.FieldByName('DEF_PRO_DEADLINE_HELL').AsInteger;
      Afl          := Que.FieldByName('DEF_AFL_DEADLINE_DAGE').AsInteger;
      Afl_Time     := TimeOf(Que.FieldByName('DEF_AFL_DEADLINE').AsDateTime);
      DoAflHol     := Que.FieldByName('DEF_AFL_DEADLINE_HELL').AsInteger;
      WeekKind     := Que.FieldByName('DEF_PUBDATE').AsInteger;
      DefSats      := Que.FieldByName('SATS_FORMAT').AsInteger;
      DefCrimp     := Que.FieldByName('Crimp').AsInteger;
      DefHeigth    := Que.FieldByName('DEF_SATS_H').AsInteger;
      DefWidth     := Que.FieldByName('DEF_SATS_W').AsInteger;
      DefPress     := Que.FieldByName('PRESS_ID').AsInteger;
      DefFals      := Que.FieldByName('FALS_ID').AsString;
      DefPaper     := Que.FieldByName('PAPIR_ID').AsInteger;
      DefIL        := Que.FieldByName('NR_IL').AsInteger;
      DefAcc       := Que.FieldByName('DEF_Acc').AsString;
      DefOplag     := Que.FieldByName('DEF_OPLAG').AsInteger;
      DefOmreg     := Que.FieldByName('DEF_OMREG').AsInteger;
      DefLogRem    := Que.FieldByName('DEF_Log_Rem').AsString;
      DefRepRem    := Que.FieldByName('DEF_Rep_Rem').AsString;
      DefProRem    := Que.FieldByName('DEF_Pro_Rem').AsString;
      DefPakRem    := Que.FieldByName('DEF_Pak_Rem').AsString;
      DefDisRem    := Que.FieldByName('DEF_Dis_Rem').AsString;
      DefSecName   := Que.FieldByName('Def_Section').AsString;
      DefAfvId     := Que.FieldByName('Afv_ID').AsString;

      // ### NAN 2024-05-31 - mangler at blive overført til DATA
      DefUseTrimbox := Que.FieldByName('Def_UseTrimBox').AsInteger <> 0;
      // ###

      dt := IncDay(PubDate, Pro);
      DefProTime   := EncodeDateTime(YearOf(dt), MonthOf(dt), DayOf(dt), HourOf(Pro_Time), MinuteOf(Pro_Time), 0, 0);
      DefProTimeEnd := EncodeDateTime(YearOf(dt), MonthOf(dt), DayOf(dt), HourOf(Pro_Time_End), MinuteOf(Pro_Time_End), 0, 0);

      If DefProTimeEnd < DefProTime then
        DefProTimeEnd := IncDay(DefProTimeEnd);

      dt := IncDay(PubDate, Ctp);
      DefCtpTime := EncodeDateTime(YearOf(dt), MonthOf(dt), DayOf(dt), HourOf(Ctp_Time), MinuteOf(Ctp_Time), 0, 0);

      dt := IncDay(PubDate, Rel);
      DefRelTime := EncodeDateTime(YearOf(dt), MonthOf(dt), DayOf(dt), HourOf(Rel_Time), MinuteOf(Rel_Time), 0, 0);

      dt := IncDay(PubDate, Afl);
      DefAflTime := EncodeDateTime(YearOf(dt), MonthOf(dt), DayOf(dt), HourOf(Afl_Time), MinuteOf(Afl_Time), 0, 0);

      dt := IncDay(PubDate, Dis);
      DefDisTime := EncodeDateTime(YearOf(dt), MonthOf(dt), DayOf(dt), HourOf(Dis_Time), MinuteOf(Dis_Time), 0, 0);

      dt := IncDay(PubDate, Mat);
      DefMatTime := EncodeDateTime(YearOf(dt), MonthOf(dt), DayOf(dt), HourOf(Mat_Time), MinuteOf(Mat_Time), 0, 0);

      dt := IncDay(PubDate, Pak);
      DefPakTime := EncodeDateTime(YearOf(dt), MonthOf(dt), DayOf(dt), HourOf(Pak_Time), MinuteOf(Pak_Time), 0, 0);
      DefPakTimeEnd := EncodeDateTime(YearOf(dt), MonthOf(dt), DayOf(dt), HourOf(Pak_Time_End), MinuteOf(Pak_Time_End), 0, 0);

      {If (DefOplag > 0) And (DefPress > 0) then
      Begin
        Que.Close;
        Que.SQL.Text := 'Select TOP 1 TRYK_TIME from PRESS WITH (NOLOCK)';
        Que.SQL.Add('    WHERE ID = ''' + IntToStr(DefPress) + '''');
        Que.Open;
        If Que.RecordCount > 0 Then
        Begin
          DefProTimeEnd := IncMinute(DefProTime, Round(((DefOplag / Que.FieldByName('TRYK_TIME').AsInteger) * 60)));
          DefPakTimeEnd := DefProTimeEnd;
        end;
      end else
      Begin
        DefProTimeEnd := IncMinute(DefProTime, 60);
        DefPakTimeEnd := DefProTimeEnd;
      end;


      If MinutesBetween(DefProTime, DefProTimeEnd) < 60 then
      Begin
        DefProTimeEnd := IncMinute(DefProTime, 60);
        DefPakTimeEnd := DefProTimeEnd;
      end; }


      Que1.Close;

      // ### NAN 2024-05-31
//    Que1.SQL.Text := 'Insert into DATA (Long_Name, Short_Name, ProType, ProID, Pub_Kind,  Pro_SecName, Pro_PageCount, Pro_SecRem, PAGEFORMAT, Crimp, SKAR_HOJDE, SKAR_BREDE, IL_ID, PRESS_ACC, PAPIR, Udg_dato, Pro_Dato, Pro_Dato_SLUT, Ctp_Dato, Mat_Dato, Afl_Dato, Pak_Dato, Pak_Dato_Slut, Dist_Dato, Rel_Dato, PRO_REALEASE, ZONES, Editions, Tryk_Fals, Press, OPLAG_FORVENTET, PRO_Omreg, LOG_BEM, REP_BEM, PRO_BEM, PAK_BEM, DIS_BEM, Afvikling, Pak_Pattern) OUTPUT INSERTED.ID Values (';
      Que1.SQL.Text := 'Insert into DATA (Long_Name, Short_Name, ProType, ProID, Pub_Kind,  Pro_SecName, Pro_PageCount, Pro_SecRem, PAGEFORMAT, Crimp, SKAR_HOJDE, SKAR_BREDE, IL_ID, PRESS_ACC, PAPIR, Udg_dato, Pro_Dato, Pro_Dato_SLUT, Ctp_Dato, Mat_Dato, Afl_Dato, Pak_Dato, Pak_Dato_Slut, Dist_Dato, Rel_Dato, PRO_REALEASE, ZONES, Editions, Tryk_Fals, Press, OPLAG_FORVENTET, PRO_Omreg, LOG_BEM, REP_BEM, PRO_BEM, PAK_BEM, DIS_BEM, Afvikling, Pak_Pattern, Output_Short_Name,UseTrimBox) OUTPUT INSERTED.ID Values (';
      // ###

      Que1.SQL.Add(    ' ''' + DefLongName + '''');
      Que1.SQL.Add(    ',''' + DefSHortName + '''');
      Que1.SQL.Add(    ',''' + '0' + '''');
      Que1.SQL.Add(    ',''' + IntToStr(DefID) + '''');
      Que1.SQL.Add(    ',''' + IntToStr(DefPubKind) + '''');
      If (DefSecName <> '')  then          //Oprettes med bestemt sektion navn
      Begin
        Que1.SQL.Add(    ',''' + DefSecName + '''');
        SL := TStringlist.Create;
        SL.Delimiter:= ',';
        SL.DelimitedText := DefSecName;
        For x := 0 to SL.Count - 1 do
          SL[x] := '0';
        Que1.SQL.Add(    ',''' + SL.CommaText + '''');

        For x := 0 to SL.Count - 1 do
          SL[x] := '';
        Que1.SQL.Add(    ',''' + SL.CommaText + '''');
        SL.Free;
      End else
      Begin
        Que1.SQL.Add(    ',''' + '' + '''');
        Que1.SQL.Add(    ',''' + '0' + '''');
        Que1.SQL.Add(    ',''' + '' + '''');
      End;

      Que1.SQL.Add(    ',''' + IntToStr(DefSats) + '''');
      Que1.SQL.Add(    ',''' + IntToStr(DefCrimp) + '''');
      Que1.SQL.Add(    ',''' + IntToStr(DefHeigth) + '''');
      Que1.SQL.Add(    ',''' + IntToStr(DefWidth) + '''');
      Que1.SQL.Add(    ',''' + IntToSTr(DefIL) + '''');
      Que1.SQL.Add(    ',''' + DefAcc + '''');
      Que1.SQL.Add(    ',''' + IntToStr(DefPaper) + '''');
      Que1.SQL.Add(    ',''' + Formatdatetime('yyyy-mm-dd',       PubDate,       FS) + '''');
      Que1.SQL.Add(    ',''' + Formatdatetime('yyyy-mm-dd hh:nn', DefProTime,    FS) + '''');
      Que1.SQL.Add(    ',''' + Formatdatetime('yyyy-mm-dd hh:nn', DefProTimeEnd, FS) + '''');
      Que1.SQL.Add(    ',''' + Formatdatetime('yyyy-mm-dd hh:nn', DefCtpTime,    FS) + '''');
      Que1.SQL.Add(    ',''' + Formatdatetime('yyyy-mm-dd hh:nn', DefMatTime,    FS) + '''');
      Que1.SQL.Add(    ',''' + Formatdatetime('yyyy-mm-dd hh:nn', DefAflTime,    FS) + '''');
      Que1.SQL.Add(    ',''' + Formatdatetime('yyyy-mm-dd hh:nn', DefPakTime,    FS) + '''');
      Que1.SQL.Add(    ',''' + Formatdatetime('yyyy-mm-dd hh:nn', DefPakTimeEnd, FS) + '''');
      Que1.SQL.Add(    ',''' + Formatdatetime('yyyy-mm-dd hh:nn', DefDisTime,    FS) + '''');
      Que1.SQL.Add(    ',''' + Formatdatetime('yyyy-mm-dd hh:nn', DefRelTime,    FS) + '''');
      Que1.SQL.Add(    ',''' + Formatdatetime('yyyy-mm-dd hh:nn', DefRelTime,    FS) + ''''); //Skal være samme: rel_dato og pro_realease
      Que1.SQL.Add(    ',''' + DefZones + '''');
      Que1.SQL.Add(    ',''' + DefEditions + '''');
      Que1.SQL.Add(    ',''' + DefFals + '''');
      Que1.SQL.Add(    ',''' + IntToStr(DefPress) + '''');
      Que1.SQL.Add(    ',''' + IntToStr(DefOplag) + '''');
      Que1.SQL.Add(    ',''' + IntToStr(DefOmreg) + '''');
      Que1.SQL.Add(    ',''' + DefLogRem + '''');
      Que1.SQL.Add(    ',''' + DefRepRem + '''');
      Que1.SQL.Add(    ',''' + DefProRem + '''');
      Que1.SQL.Add(    ',''' + DefPakRem + '''');
      Que1.SQL.Add(    ',''' + DefDisRem + '''');
      Que1.SQL.Add(    ',''' + DefAfvId + '''');
      Que1.SQL.Add(    ',''' + DefPak_Pattern + '''');

      // ### NAN 2024-05-31
      Que1.SQL.Add(    ',''' + DefOutputAlias + '''');
      Que1.SQL.Add(    ',''' + BoolToStr(DefUseTrimbox) + '''');
      // ###

      Que1.SQL.Add(    ')');
      //Que1.SQL.SaveToFile('300.sql');
      Que1.Open;
      Result := Que1.FieldByName('ID').AsInteger;
      fMain.SQLTrans.commit;

      Que1.Close;


      // 2024-07-05 - fix hovedprodukt hvis dette er et weekendtillæg der ved en fejl er lagt ind som B-sektion i hovedprodukt
      //              Store proc fjerner B-sektion fra hovedproduktet.
      if (Result > 0) then
      begin
        try
           Que1.SQL.Text := 'EXEC spPostCreate ' + IntToStr(Result);
          // Que1.SQL.Text := 'call spPostCreate(:param1)';
          // Que1.Params.BeginUpdate;
           //Que1.Params.ParamByName('param1').AsInteger := Result;
           // Que1.Params.EndUpdate;
           Que1.ExecSQL;
           fMain.EventLog.Debug('Called spPostCreate( ' + IntToStr(Result) + ')');
        except
          on E: Exception do
          begin
             fMain.EventLog.Debug('call spPostCreate error: ' + E.Message);
           end;
        end;

      end;


      {If DefPakTemplate > 0 then
        MakePakMat(Result, DefPakTemplate);}

      //skæremål
    end;
  except
    on E: Exception do
    begin
       {fMain.StatusBar.Items[3].Text := 'Fail to Create new plan. ' + E.Message; }
       fMain.EventLog.Error('Create New product fail.' + E.Message);
     end;
  end;
end;



end.

