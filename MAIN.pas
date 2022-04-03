unit MAIN;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, Grids, ExtCtrls, ToolWin, Menus, StdCtrls, DateUtils,
  OleAuto, Ole2, OleCtl, OleConst, ImgList, Printers, FileCtrl;

type
  TfrmMain = class(TForm)
    mmMain: TMainMenu;
    statMain: TStatusBar;
    tlbMain: TToolBar;
    pnlMain: TPanel;
    str1: TStringGrid;
    btnExport: TToolButton;
    btnOpen: TToolButton;
    btnSep1: TToolButton;
    btnAdd: TToolButton;
    grpDateTimeInfo: TGroupBox;
    grpDescription: TGroupBox;
    dtpStart: TDateTimePicker;
    dtpEnd: TDateTimePicker;
    lblStart: TLabel;
    lblEnd: TLabel;
    redDes: TRichEdit;
    File1: TMenuItem;
    mniNew: TMenuItem;
    mniOpen: TMenuItem;
    mniSave: TMenuItem;
    mniPrint: TMenuItem;
    mniClose: TMenuItem;
    Edit1: TMenuItem;
    Add1: TMenuItem;
    Help1: TMenuItem;
    Support1: TMenuItem;
    About1: TMenuItem;
    dtpDate: TDateTimePicker;
    btnSave1: TToolButton;
    btnSep2: TToolButton;
    btnSep3: TToolButton;
    btnAddRow: TToolButton;
    btnClearSheet: TToolButton;
    btnNow: TButton;
    chkStart: TCheckBox;
    chkEnd: TCheckBox;
    grpNow: TGroupBox;
    ilMain: TImageList;
    btnSep4: TToolButton;
    btnClose: TToolButton;
    mniExport: TMenuItem;
    btnRemoveRow: TToolButton;
    Row1: TMenuItem;
    mniEvent1: TMenuItem;
    dirlstMain: TDirectoryListBox;
    fllstMain: TFileListBox;
    mniSaveAs1: TMenuItem;
    procedure FormActivate(Sender: TObject);
    procedure mniCloseClick(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure btnExportClick(Sender: TObject);
    procedure btnSave1Click(Sender: TObject);
    procedure btnNowClick(Sender: TObject);
    procedure btnOpenClick(Sender: TObject);
    procedure btnAddRowClick(Sender: TObject);
    procedure btnRemoveRowClick(Sender: TObject);
    procedure btnClearSheetClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure dtpStartChange(Sender: TObject);
    procedure mniNewClick(Sender: TObject);
    procedure mniOpenClick(Sender: TObject);
    procedure mniExportClick(Sender: TObject);
    procedure mniSaveClick(Sender: TObject);
    procedure mniPrintClick(Sender: TObject);
    procedure PrintGrid(sGrid : TStringGrid; sTitle : string);
    procedure mniEvent1Click(Sender: TObject);
    procedure Row1Click(Sender: TObject);
    procedure About1Click(Sender: TObject);
    procedure Support1Click(Sender: TObject);
    procedure fllstMainClick(Sender: TObject);
    procedure mniSaveAs1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;
  TimeDiff, TimeStart, TimeEnd : TTime;
  Date : TDate;
  sDes : string;
  HoursBet, MinBet, SecBet, MSBet: Word;
  SHour, SMin, SSec, SMS, EHour, EMin, ESec, EMS : Word;
  sTimeDiff : String;
  iCount : Integer;

implementation

{$R *.dfm}

procedure TfrmMain.About1Click(Sender: TObject);
begin
  MessageDlg('Time Sheet ® Version 1' + #13 + '© 2022 Connor Bell' + #9 + 
  'All Rights Reserved' + #13 + 'Version Date: 2022/04/03', mtInformation, 
  [mbClose], 0); 
end;

procedure TfrmMain.btnAddClick(Sender: TObject);
begin
  // Add
  str1.ColWidths[4] := Length(redDes.Lines.Text) * 10;
  if Length(redDes.Lines.Text) = 0 then
    begin
      str1.ColWidths[4] := 400; 
    end;
  Date := dtpDate.Date;
  sDes := redDes.Lines.Text;
  TimeStart := dtpStart.Time;
  TimeEnd := dtpEnd.Time;
  DecodeTime(TimeStart, SHour, SMin, SSec, SMS);
  DecodeTime(TimeEnd, EHour, EMin, ESec, EMS);
  HoursBet := EHour - sHour;
  MinBet := EMin - SMin;
  SecBet := ESec - SSec;
  MSBet := EMS - SMS;
  TimeDiff := EncodeTime(HoursBet, MinBet, SecBet, MSBet);
  with str1 do
    begin
      Cells[0, iCount] := DateToStr(Date);
      Cells[1, iCount] := TimeToStr(TimeStart);
      Cells[2, iCount] := TimeToStr(TimeEnd);
      Cells[3, iCount] := TimeToStr(TimeDiff);
      Cells[4, iCount] := sDes;
    end;
  Inc(iCount);
  str1.RowCount := str1.RowCount + 1;  
  statMain.Panels[1].Text := 'Unsaved';
  str1.FixedRows := 1; 
  redDes.Clear; 
end;

procedure TfrmMain.btnAddRowClick(Sender: TObject);
begin
  str1.RowCount := str1.RowCount + 1; 
  str1.FixedRows := 1; 
end;

procedure TfrmMain.btnClearSheetClick(Sender: TObject); 
var
  sStat : string; 
  iMessage : Integer; 
  SD : TSaveDialog;
  i : Integer;
  CSV : TStrings;
  FileName : string;
begin
  sStat := statMain.Panels[1].Text; 
  if (sStat = 'Saved') or (sStat = '') then
    begin                                                   
      iCount := 1; 
      for I := 0 to str1.ColCount - 1 do
        begin
          str1.Cols[I].Clear;
        end;
      with str1 do
        begin
          Cells[0, 0] := 'Date';
          Cells[1, 0] := 'Start Time';
          Cells[2, 0] := 'End Time';
          Cells[3, 0] := 'Duration';
          Cells[4, 0] := 'Description';
        end;
      str1.RowCount := 1; 
      statMain.Panels[0].Text := '';
      statMain.Panels[1].Text := ''; 
    end
  else 
  if (sStat = 'Unsaved') then
    begin
      iMessage := MessageDlg('Unsaved Changes' + #13 +
      'Would you like to save changes?', mtWarning, [mbYes, mbNo, mbCancel], 0);
      if iMessage = mrYes then
        begin
          iCount := 1; 
          SD := TSaveDialog.Create(Self);
          SD.Filter := 'CSV separator separated(*.csv)|*.CSV';
          if SD.Execute = True then
            begin
              FileName := SD.FileName;
              if Copy(FileName, Pos('.', FileName), Length(FileName) -
              Pos('.', FileName) + 1) <> '.csv' then
                FileName := FileName + '.csv';
                Screen.Cursor := crHourGlass;
                CSV := TStringList.Create;
              try
                for i := 0 to str1.RowCount - 1 do
                  begin
                    CSV.Add(str1.Rows[i].CommaText);
                  end;
                CSV.SaveToFile(FileName);
              finally
                CSV.Free;
                Screen.Cursor := crDefault;
                MessageDlg('Sucessfully saved records' + #13 + FileName,
                 mtInformation, [mbOK], 0);
              end;
            end;
          statMain.Panels[0].Text := '';
          statMain.Panels[1].Text := '';
          for I := 0 to str1.ColCount - 1 do
            begin
              str1.Cols[I].Clear;
            end;
          str1.RowCount := 1;
           with str1 do
            begin
              Cells[0, 0] := 'Date';
              Cells[1, 0] := 'Start Time';
              Cells[2, 0] := 'End Time';
              Cells[3, 0] := 'Duration';
              Cells[4, 0] := 'Description';
            end;
        end
      else
      if iMessage = mrNo then
        begin
          iCount := 1; 
          statMain.Panels[0].Text := '';
          statMain.Panels[1].Text := '';
          for I := 0 to str1.ColCount - 1 do
            begin
              str1.Cols[I].Clear;
            end;
          str1.RowCount := 1;
           with str1 do
            begin
              Cells[0, 0] := 'Date';
              Cells[1, 0] := 'Start Time';
              Cells[2, 0] := 'End Time';
              Cells[3, 0] := 'Duration';
              Cells[4, 0] := 'Description';
            end;
        end;  
    end;
end;

procedure TfrmMain.btnCloseClick(Sender: TObject);
var
  sStat : string; 
  iMessage : Integer; 
  SD : TSaveDialog;
  i : Integer;
  CSV : TStrings;
  FileName : string;
begin
  sStat := statMain.Panels[1].Text; 
  if (sStat = 'Saved') or (sStat = '') then
    begin                                                   
      Application.Terminate; 
    end
  else 
  if (sStat = 'Unsaved') then
    begin
      iMessage := MessageDlg('Unsaved Changes' + #13 +
      'Would you like to save changes?', mtWarning, [mbYes, mbNo, mbCancel], 0);
      if iMessage = mrYes then
        begin
          SD := TSaveDialog.Create(Self);
          SD.Filter := 'CSV separator separated(*.csv)|*.CSV';
          if SD.Execute = True then
            begin
              FileName := SD.FileName;
              if Copy(FileName, Pos('.', FileName), Length(FileName) -
              Pos('.', FileName) + 1) <> '.csv' then
                FileName := FileName + '.csv';
                Screen.Cursor := crHourGlass;
                CSV := TStringList.Create;
              try
                for i := 0 to str1.RowCount - 1 do
                  begin
                    CSV.Add(str1.Rows[i].CommaText);
                  end;
                CSV.SaveToFile(FileName);
              finally
                CSV.Free;
                Screen.Cursor := crDefault;
                MessageDlg('Sucessfully saved records' + #13 + FileName,
                 mtInformation, [mbOK], 0);
              end;
            end;
            Application.Terminate; 
        end
      else
      if iMessage = mrNo then
        begin
          Application.Terminate; 
        end;  
    end;
end;

procedure TfrmMain.btnExportClick(Sender: TObject);
var
  xls, wb, Range : OleVariant; 
  arrData : Variant; 
  iRowCount, iColCount, I, J : Integer;
  SD : TSaveDialog;
begin
  // Export
      iRowCount := str1.RowCount;
      iColCount := str1.ColCount; 
      arrData := VarArrayCreate([1, iRowCount, 1, iColCount], varVariant); 
      for I := 1 to iRowCount do
        begin
          for J := 1 to iColCount do
            begin
              arrData[I, J] := str1.Cells[J - 1, I - 1]; 
            end;
        end;
      xls := CreateOleObject('Excel.Application'); 
      wb := xls.Workbooks.Add; 
      Range := wb.WorkSheets[1].Range[wb.WorkSheets[1].Cells[1, 1], 
      wb.Worksheets[1].Cells[iRowCount, iColCount]]; 
      Range.Value := arrData; 
      xls.Visible := True; 
end;

procedure TfrmMain.btnNowClick(Sender: TObject);
begin
  if chkStart.Checked = True then
    begin
      dtpStart.Time := Now; 
    end
  else 
  if chkEnd.Checked = True then
    begin
      dtpEnd.Time := Now; 
    end;
end;

procedure TfrmMain.btnOpenClick(Sender: TObject);
var 
  OD : TOpenDialog;
  i : Cardinal;
  CSV : TStrings;
  FileName : string;
begin
  str1.RowCount := 1; 
   
  OD := TOpenDialog.Create(Self);
  OD.Filter := 'CSV separator separated(*.csv)|*.CSV';
  CSV := TStringList.Create;
  if OD.Execute = True then
    begin
      FileName := OD.FileName;
      try
        CSV.LoadFromFile(FileName);
        str1.RowCount := CSV.Count;
        for i := 1 to CSV.Count do
          begin
            str1.Rows[i - 1].CommaText := CSV[i - 1];
          end;
      finally
        CSV.Free;
        MessageDlg('Successfully imported time sheet' + #13 + FileName,
        mtInformation, [mbOK], 0);
      end;
    end;
    iCount := str1.RowCount;
    statMain.Panels[0].Text := FileName; 
    statMain.Panels[1].Text := 'Saved'; 
    frmMain.Caption := 'Time Sheet - ' + FileName; 
end;

procedure TfrmMain.btnRemoveRowClick(Sender: TObject);
begin
  str1.RowCount := str1.RowCount - 1; 
end;

procedure TfrmMain.btnSave1Click(Sender: TObject);
var 
  SD : TSaveDialog;
  i : Integer;
  CSV : TStrings;
  FileName : string;
begin
 // Save
  FileName := fllstMain.FileName; 
  if Copy(FileName, Pos('.', FileName), Length(FileName) -
  Pos('.', FileName) + 1) <> '.csv' then
    FileName := FileName + '.csv';
    Screen.Cursor := crHourGlass;
    CSV := TStringList.Create;
  try
    for i := 0 to str1.RowCount - 1 do
      begin
        CSV.Add(str1.Rows[i].CommaText);
      end;
    CSV.SaveToFile(FileName);
  finally
    CSV.Free;
    Screen.Cursor := crDefault;
  end;
  statMain.Panels[0].Text := FileName; 
  statMain.Panels[1].Text := 'Saved';
  frmMain.Caption := 'Time Sheet - ' + FileName;  
end;

procedure TfrmMain.dtpStartChange(Sender: TObject);
var
  SHour, SMin, SSec, SMS, EHour, EMin, ESec, EMS : Word; 
begin
  DecodeTime(dtpStart.Time, sHour, sMin, SSec, SMS);
  DecodeTime(dtpEnd.Time, EHour, EMin, ESec, EMS);
  if SHour > EHour then
    begin
      MessageDlg('Start time cannot be greater than end time', mtError, 
      [mbOK], 0); 
      dtpStart.Time := dtpEnd.Time; 
    end;
end;

procedure TfrmMain.fllstMainClick(Sender: TObject);
var 
  OD : TOpenDialog;
  i : Cardinal;
  CSV : TStrings;
  FileName : string;
begin
  str1.RowCount := 1; 
  CSV := TStringList.Create;
  FileName := fllstMain.FileName; 
  CSV.LoadFromFile(FileName);
  str1.RowCount := CSV.Count;
  for i := 1 to CSV.Count do
    begin
      str1.Rows[i - 1].CommaText := CSV[i - 1];
    end;
  CSV.Free;
  MessageDlg('Successfully imported time sheet' + #13 + FileName,
  mtInformation, [mbOK], 0);
  iCount := str1.RowCount;
  statMain.Panels[0].Text := FileName; 
  statMain.Panels[1].Text := 'Saved'; 
  frmMain.Caption := 'Time Sheet - ' + FileName; 
end;

procedure TfrmMain.mniCloseClick(Sender: TObject);
var
  sStat : string; 
  iMessage : Integer; 
  SD : TSaveDialog;
  i : Integer;
  CSV : TStrings;
  FileName : string;
begin
  sStat := statMain.Panels[1].Text; 
  if (sStat = 'Saved') or (sStat = '') then
    begin                                                   
      Application.Terminate; 
    end
  else 
  if (sStat = 'Unsaved') then
    begin
      iMessage := MessageDlg('Unsaved Changes' + #13 +
      'Would you like to save changes?', mtWarning, [mbYes, mbNo, mbCancel], 0);
      if iMessage = mrYes then
        begin
          SD := TSaveDialog.Create(Self);
          SD.Filter := 'CSV separator separated(*.csv)|*.CSV';
          if SD.Execute = True then
            begin
              FileName := SD.FileName;
              if Copy(FileName, Pos('.', FileName), Length(FileName) -
              Pos('.', FileName) + 1) <> '.csv' then
                FileName := FileName + '.csv';
                Screen.Cursor := crHourGlass;
                CSV := TStringList.Create;
              try
                for i := 0 to str1.RowCount - 1 do
                  begin
                    CSV.Add(str1.Rows[i].CommaText);
                  end;
                CSV.SaveToFile(FileName);
              finally
                CSV.Free;
                Screen.Cursor := crDefault;
                MessageDlg('Sucessfully saved records' + #13 + FileName,
                 mtInformation, [mbOK], 0);
              end;
            end;
            Application.Terminate; 
        end
      else
      if iMessage = mrNo then
        begin
          Application.Terminate; 
        end;  
    end;
end;

procedure TfrmMain.mniEvent1Click(Sender: TObject);
begin
// Add
  str1.ColWidths[4] := Length(redDes.Lines.Text) * 10;
  if Length(redDes.Lines.Text) = 0 then
    begin
      str1.ColWidths[4] := 400; 
    end;
  Date := dtpDate.Date;
  sDes := redDes.Lines.Text;
  TimeStart := dtpStart.Time;
  TimeEnd := dtpEnd.Time;
  DecodeTime(TimeStart, SHour, SMin, SSec, SMS);
  DecodeTime(TimeEnd, EHour, EMin, ESec, EMS);
  HoursBet := EHour - sHour;
  MinBet := EMin - SMin;
  SecBet := ESec - SSec;
  MSBet := EMS - SMS;
  TimeDiff := EncodeTime(HoursBet, MinBet, SecBet, MSBet);
  with str1 do
    begin
      Cells[0, iCount] := DateToStr(Date);
      Cells[1, iCount] := TimeToStr(TimeStart);
      Cells[2, iCount] := TimeToStr(TimeEnd);
      Cells[3, iCount] := TimeToStr(TimeDiff);
      Cells[4, iCount] := sDes;
    end;
  Inc(iCount);
  str1.RowCount := str1.RowCount + 1;  
  statMain.Panels[1].Text := 'Unsaved';
  str1.FixedRows := 1; 
end;

procedure TfrmMain.mniExportClick(Sender: TObject);
var
  xls, wb, Range : OleVariant; 
  arrData : Variant; 
  iRowCount, iColCount, I, J : Integer;
  SD : TSaveDialog;
begin
  // Export
      iRowCount := str1.RowCount;
      iColCount := str1.ColCount; 
      arrData := VarArrayCreate([1, iRowCount, 1, iColCount], varVariant); 
      for I := 1 to iRowCount do
        begin
          for J := 1 to iColCount do
            begin
              arrData[I, J] := str1.Cells[J - 1, I - 1]; 
            end;
        end;
      xls := CreateOleObject('Excel.Application'); 
      wb := xls.Workbooks.Add; 
      Range := wb.WorkSheets[1].Range[wb.WorkSheets[1].Cells[1, 1], 
      wb.Worksheets[1].Cells[iRowCount, iColCount]]; 
      Range.Value := arrData; 
      xls.Visible := True; 
end;

procedure TfrmMain.mniNewClick(Sender: TObject);
var
  sStat : string; 
  iMessage : Integer; 
  SD : TSaveDialog;
  i : Integer;
  CSV : TStrings;
  FileName : string;
begin
  sStat := statMain.Panels[1].Text; 
  if (sStat = 'Saved') or (sStat = '') then
    begin                                                   
      iCount := 1; 
      for I := 0 to str1.ColCount - 1 do
        begin
          str1.Cols[I].Clear;
        end;
      with str1 do
        begin
          Cells[0, 0] := 'Date';
          Cells[1, 0] := 'Start Time';
          Cells[2, 0] := 'End Time';
          Cells[3, 0] := 'Duration';
          Cells[4, 0] := 'Description';
        end;
      str1.RowCount := 1; 
      statMain.Panels[0].Text := '';
      statMain.Panels[1].Text := ''; 
    end
  else 
  if (sStat = 'Unsaved') then
    begin
      iMessage := MessageDlg('Unsaved Changes' + #13 +
      'Would you like to save changes?', mtWarning, [mbYes, mbNo, mbCancel], 0);
      if iMessage = mrYes then
        begin
          iCount := 1; 
          SD := TSaveDialog.Create(Self);
          SD.Filter := 'CSV separator separated(*.csv)|*.CSV';
          if SD.Execute = True then
            begin
              FileName := SD.FileName;
              if Copy(FileName, Pos('.', FileName), Length(FileName) -
              Pos('.', FileName) + 1) <> '.csv' then
                FileName := FileName + '.csv';
                Screen.Cursor := crHourGlass;
                CSV := TStringList.Create;
              try
                for i := 0 to str1.RowCount - 1 do
                  begin
                    CSV.Add(str1.Rows[i].CommaText);
                  end;
                CSV.SaveToFile(FileName);
              finally
                CSV.Free;
                Screen.Cursor := crDefault;
                MessageDlg('Sucessfully saved records' + #13 + FileName,
                 mtInformation, [mbOK], 0);
              end;
            end;
          statMain.Panels[0].Text := '';
          statMain.Panels[1].Text := '';
          for I := 0 to str1.ColCount - 1 do
            begin
              str1.Cols[I].Clear;
            end;
          str1.RowCount := 1;
           with str1 do
            begin
              Cells[0, 0] := 'Date';
              Cells[1, 0] := 'Start Time';
              Cells[2, 0] := 'End Time';
              Cells[3, 0] := 'Duration';
              Cells[4, 0] := 'Description';
            end;
        end
      else
      if iMessage = mrNo then
        begin
          iCount := 1; 
          statMain.Panels[0].Text := '';
          statMain.Panels[1].Text := '';
          for I := 0 to str1.ColCount - 1 do
            begin
              str1.Cols[I].Clear;
            end;
          str1.RowCount := 1;
           with str1 do
            begin
              Cells[0, 0] := 'Date';
              Cells[1, 0] := 'Start Time';
              Cells[2, 0] := 'End Time';
              Cells[3, 0] := 'Duration';
              Cells[4, 0] := 'Description';
            end;
        end;  
    end;
end;

procedure TfrmMain.mniOpenClick(Sender: TObject);
var 
  OD : TOpenDialog;
  i : Cardinal;
  CSV : TStrings;
  FileName : string;
begin
  str1.RowCount := 1; 
   
  OD := TOpenDialog.Create(Self);
  OD.Filter := 'CSV separator separated(*.csv)|*.CSV';
  CSV := TStringList.Create;
  if OD.Execute = True then
    begin
      FileName := OD.FileName;
      try
        CSV.LoadFromFile(FileName);
        str1.RowCount := CSV.Count;
        for i := 1 to CSV.Count do
          begin
            str1.Rows[i - 1].CommaText := CSV[i - 1];
          end;
      finally
        CSV.Free;
        MessageDlg('Successfully imported time sheet' + #13 + FileName,
        mtInformation, [mbOK], 0);
      end;
    end;
    iCount := str1.RowCount;
    statMain.Panels[0].Text := FileName; 
    statMain.Panels[1].Text := 'Saved'; 
    frmMain.Caption := 'Time Sheet - ' + FileName; 
end;

procedure TfrmMain.mniPrintClick(Sender: TObject);
begin
with TPrintDialog.Create(nil) do
    begin
      try
        if Execute then
          begin
            PrintGrid(str1, 'Time Sheet' + statMain.Panels[0].Text);
          end;
      finally
        Free;
      end;
    end;
end;

procedure TfrmMain.mniSaveAs1Click(Sender: TObject);
var 
  SD : TSaveDialog;
  i : Integer;
  CSV : TStrings;
  FileName : string;
begin
 // Save
  SD := TSaveDialog.Create(Self);
  SD.Filter := 'CSV separator separated(*.csv)|*.CSV';
  if SD.Execute = True then
    begin
      FileName := SD.FileName;
      if Copy(FileName, Pos('.', FileName), Length(FileName) -
      Pos('.', FileName) + 1) <> '.csv' then
        FileName := FileName + '.csv';
        Screen.Cursor := crHourGlass;
        CSV := TStringList.Create;
      try
        for i := 0 to str1.RowCount - 1 do
          begin
            CSV.Add(str1.Rows[i].CommaText);
          end;
        CSV.SaveToFile(FileName);
      finally
        CSV.Free;
        Screen.Cursor := crDefault;
        MessageDlg('Sucessfully saved timesheet' + #13 + FileName,
        mtInformation, [mbOK], 0);
      end;
    end;
    statMain.Panels[0].Text := FileName; 
    statMain.Panels[1].Text := 'Saved';
    frmMain.Caption := 'Time Sheet - ' + FileName;
end;

procedure TfrmMain.mniSaveClick(Sender: TObject);
var 
  SD : TSaveDialog;
  i : Integer;
  CSV : TStrings;
  FileName : string;
begin
 FileName := fllstMain.FileName; 
  if Copy(FileName, Pos('.', FileName), Length(FileName) -
  Pos('.', FileName) + 1) <> '.csv' then
    FileName := FileName + '.csv';
    Screen.Cursor := crHourGlass;
    CSV := TStringList.Create;
  try
    for i := 0 to str1.RowCount - 1 do
      begin
        CSV.Add(str1.Rows[i].CommaText);
      end;
    CSV.SaveToFile(FileName);
  finally
    CSV.Free;
    Screen.Cursor := crDefault;
  end;
  statMain.Panels[0].Text := FileName; 
  statMain.Panels[1].Text := 'Saved';
  frmMain.Caption := 'Time Sheet - ' + FileName;
end;

procedure TfrmMain.PrintGrid(sGrid: TStringGrid; sTitle: string);
var
  iX1, iX2, iY1, iY2, iTmp, f: Integer;
  TR: TRect;
  sHeader : String;
begin
  sHeader := 'Time Sheet - ' + statMain.Panels[0].Text; 
  Printer.Title := sHeader;
  Printer.BeginDoc;
  Printer.Canvas.Pen.Color  := 0;
  Printer.Canvas.Font.Name  := 'Tahoma';
  Printer.Canvas.Font.Size  := 8;
  Printer.Canvas.TextOut(5, 100, Printer.Title);
  for f := 1 to str1.ColCount - 1 do
  begin
    iX1 := 0;
    for iTmp := 1 to (f - 1) do
      iX1 := iX1 + 5 * (str1.ColWidths[iTmp]);
    iY1 := 300;
    iX2 := 0;
    for iTmp := 1 to f do
      iX2 := iX2 + 5 * (str1.ColWidths[iTmp]);
    iY2 := 450;
    TR := Rect(iX1, iY1, iX2 - 30, iY2);
    Printer.Canvas.Font.Style := [fsBold];
    Printer.Canvas.Font.Size := 8;
    Printer.Canvas.TextRect(TR, iX1 + 50, 350, str1.Cells[f, 0]);
    Printer.Canvas.Font.Style := [];
    for iTmp := 1 to str1.RowCount - 1 do
    begin
      iY1 := 150 * iTmp + 300;
      iY2 := 150 * (iTmp + 1) + 300;
      TR := Rect(iX1, iY1, iX2 - 30, iY2);
      Printer.Canvas.TextRect(TR, iX1 + 50, iY1 + 50, str1.Cells[f, iTmp]);
    end;
  end;
  Printer.EndDoc;
end;

procedure TfrmMain.Row1Click(Sender: TObject);
begin
  str1.RowCount := str1.RowCount + 1; 
  str1.FixedRows := 1; 
end;

procedure TfrmMain.Support1Click(Sender: TObject);
begin
  MessageDlg('Email: cbell@jeppeboys.co.za' + #13 + 
  'Contact: +27 66 202 1724', mtInformation, [mbClose], 0); 
end;

procedure TfrmMain.FormActivate(Sender: TObject);
begin
  // Form Activate
  str1.ColWidths[4] := 400;
  iCount := 1;
  str1.RowCount := 1; 
  redDes.ReadOnly := False;
  with str1 do
    begin
      Cells[0, 0] := 'Date';
      Cells[1, 0] := 'Start Time';
      Cells[2, 0] := 'End Time';
      Cells[3, 0] := 'Duration';
      Cells[4, 0] := 'Description';
    end;
end;

end.
