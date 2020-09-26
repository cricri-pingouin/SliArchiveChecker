unit FmMain;

interface

uses
  Messages, SysUtils, Classes, Controls, Forms, Grids, XMLDoc, XMLIntf, sevenzip,
  StdCtrls, ExtCtrls, StrUtils;

type
  TfrmType = class(TForm)
    strgrdDropped: TStringGrid;
    lblDropped: TLabel;
    lblSelected: TLabel;
    strgrdSelected: TStringGrid;
    strgrdFiltered: TStringGrid;
    lblFiltered: TLabel;
    rgFileNum: TRadioGroup;
    rgFileSize: TRadioGroup;
    rgFileName: TRadioGroup;
    btnFilter: TButton;
    edtFileNum: TEdit;
    edtFileSize: TEdit;
    lblFileNum: TLabel;
    lblFileSize: TLabel;
    edtFileName: TEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure strgrdDroppedClick(Sender: TObject);
    procedure btnFilterClick(Sender: TObject);
    procedure strgrdFilteredClick(Sender: TObject);
  private
    procedure WMDropFiles(var Msg: TWMDropFiles); message WM_DROPFILES;
  public
    { Public declarations }

  end;

var
  frmType: TfrmType;
  HeadDB, DescDB: array of string; //XML file contents: header, description
  ClsidDB: array of TGUID; //XML file contents: TGUID
  DroppedArchivesCount: Integer;
  DroppedArchivesFullNames: array of string;
  DroppedArchivesCLSID: array of TGUID;
  DBsize: Integer; //Number of nodes in XML file
  FilesDropped: Boolean;       //Flag to trap error if clicking grid before any files dropped
  ReadFileHeaderSize: Integer; //Max length of header to read

implementation

uses
  ShellAPI, UFileCatcher;

{$R *.dfm}

function IntToStrDelimited(aNum: integer): string;
// Formats the integer aNum with the default Thousand Separator
var
  D: Double;
begin
  D := aNum;
  Result := Format('%.0n', [D]); // ".0" -> no decimals, n -> thousands separators
end;

function GetSizeOfFile(const FileName: string): Int64;
var
  Rec: TSearchRec;
begin
  Result := 0;
  if (FindFirst(FileName, faAnyFile, Rec) = 0) then
  begin
    Result := Rec.Size;
    FindClose(Rec);
  end;
end;

function ExtractFileNameWoExt(const FileName: string): string;
var
  i: integer;
begin
  i := LastDelimiter('.' + PathDelim + DriveDelim, FileName);
  if (i = 0) or (FileName[i] <> '.') then
    i := MaxInt;
  Result := ExtractFileName(Copy(FileName, 1, i - 1));
end;

function ReadFileHeader(const FileName: string): string;
var
  Stream: TFileStream;
  Buffer: array of AnsiChar;
  i: Integer;
begin
  SetLength(Buffer, ReadFileHeaderSize);
  //Populate buffer elements
  Stream := TFileStream.Create(FileName, fmOpenRead);
  try
    Stream.Read(Buffer[0], ReadFileHeaderSize);
  finally
    Stream.Free;
  end;
  Result := '';
  for i := 0 to ReadFileHeaderSize do
    Result := Result + IntToHex(Ord(Buffer[i]), 2) + ' ';
end;

function CompareVsHeader(const FileHeader: string; const BaseHeader: string): boolean;
var
  i: Integer;
begin
  Result := True;
  for i := 0 to Length(BaseHeader) do
    if (BaseHeader[i] <> ' ') and (BaseHeader[i] <> '?') and (BaseHeader[i] <> FileHeader[i]) then
    begin
      Result := False;
      Exit;
    end;
end;

procedure TfrmType.FormCreate(Sender: TObject);
var
  DOC: IXMLDocument;
  FileTypeNode: IXMLNode;
  i, j: Integer;
begin
  //Init StringGrid headers
  //Don't mind doing it here as I needed to copy/paste when clearing grids!
  strgrdDropped.Cells[0, 0] := 'File name';
  strgrdDropped.Cells[1, 0] := 'Extension';
  strgrdDropped.Cells[2, 0] := 'Size (KB)';
  strgrdDropped.Cells[3, 0] := 'Header';
  strgrdDropped.Cells[4, 0] := 'Type detected';
  strgrdSelected.Cells[0, 0] := 'File path';
  strgrdSelected.Cells[1, 0] := 'File name';
  strgrdSelected.Cells[2, 0] := 'Extension';
  strgrdSelected.Cells[3, 0] := 'Size (KB)';
  strgrdFiltered.Cells[0, 0] := 'File name';
  strgrdFiltered.Cells[1, 0] := 'Extension';
  strgrdFiltered.Cells[2, 0] := 'Size (KB)';
  strgrdFiltered.Cells[3, 0] := 'Type detected';
  //Load signatures
  DOC := LoadXMLDocument('archtype.xml');
  DBsize := DOC.ChildNodes.Nodes['filetypes'].ChildNodes.Count;
  SetLength(HeadDB, DBsize);
  SetLength(DescDB, DBsize);
  SetLength(ClsidDB, DBsize);
  ReadFileHeaderSize := 0;
  for i := 0 to DBsize - 1 do
  begin
    FileTypeNode := DOC.ChildNodes.Nodes['filetypes'].ChildNodes[i];
    HeadDB[i] := FileTypeNode.ChildNodes['header'].NodeValue;
    j := length(HeadDB[i]);
    if (j > ReadFileHeaderSize) then
      ReadFileHeaderSize := j;
    DescDB[i] := FileTypeNode.ChildNodes['description'].NodeValue;
    ClsidDB[i] := StringToGUID(FileTypeNode.ChildNodes['clsid'].NodeValue);
  end;
  ReadFileHeaderSize := (ReadFileHeaderSize + 1) div 3; //Header in the form XX YY ZZ, so 3 characters per char
  // Tell windows we accept file drops
  DragAcceptFiles(Self.Handle, True);
  FilesDropped := False; //Flag to check files qwere dropped before doing stuff
end;

procedure TfrmType.FormDestroy(Sender: TObject);
begin
  // Cancel acceptance of file drops
  DragAcceptFiles(Self.Handle, False);
end;

procedure TfrmType.WMDropFiles(var Msg: TWMDropFiles);
var
  i, j, FileCount: Word;
  Catcher: TFileCatcher; //File catcher class
  FullName, FileHeader, ThisDesc: string;
  ThisFileSize: Int64;
begin
  inherited;
  // Create file catcher object to hide all messy details
  Catcher := TFileCatcher.Create(Msg.Drop);
  FileCount := Pred(Catcher.FileCount) + 1; //Not sure why +1 needed
  //Clear TStringGrid
  strgrdDropped.Visible := False;
  for i := 0 to strgrdDropped.ColCount - 1 do
    strgrdDropped.Cols[i].Clear;
  strgrdDropped.RowCount := 2; //1 for header
  strgrdDropped.Cells[0, 0] := 'File name';
  strgrdDropped.Cells[1, 0] := 'Extension';
  strgrdDropped.Cells[2, 0] := 'Size (KB)';
  strgrdDropped.Cells[3, 0] := 'Header';
  strgrdDropped.Cells[4, 0] := 'Type detected';
  //Initialise
  DroppedArchivesCount := 0;
  SetLength(DroppedArchivesFullNames, 0);
  SetLength(DroppedArchivesCLSID, 0);
  try
    //Try to add each dropped file to display
    for i := 0 to FileCount - 1 do  //-1 due to base 0
    begin
      Application.ProcessMessages;
      //Get file properties
      FullName := Catcher.Files[i];
      //FullFilesNames.Add(FullName);
      ThisFileSize := GetSizeOfFile(FullName);
      FileHeader := ReadFileHeader(FullName);
      //Identify file header
      ThisDesc := '';
      for j := 0 to DBsize - 1 do
        if CompareVsHeader(FileHeader, HeadDB[j]) then
        begin
          ThisDesc := DescDB[j];
          Break;
        end;
      //If header recognised, add to list
      if (ThisDesc <> '') then
      begin
        //Store info outside grid
        DroppedArchivesCount := DroppedArchivesCount + 1;
        SetLength(DroppedArchivesFullNames, DroppedArchivesCount);
        SetLength(DroppedArchivesCLSID, DroppedArchivesCount);
        DroppedArchivesFullNames[DroppedArchivesCount - 1] := FullName;
        DroppedArchivesCLSID[DroppedArchivesCount - 1] := ClsidDB[j];
        //Populate StringGrid
        strgrdDropped.RowCount := DroppedArchivesCount + 1;
        strgrdDropped.Cells[0, DroppedArchivesCount] := ExtractFileNameWoExt(FullName);
        strgrdDropped.Cells[1, DroppedArchivesCount] := UpperCase(StringReplace(ExtractFileExt(FullName), '.', '', []));
        strgrdDropped.Cells[2, DroppedArchivesCount] := IntToStrDelimited(ThisFileSize div 1024);
        strgrdDropped.Cells[3, DroppedArchivesCount] := FileHeader;
        strgrdDropped.Cells[4, DroppedArchivesCount] := ThisDesc;

      end;
    end;
  finally
    Catcher.Free;
  end;
  //Notify Windows we handled message
  Msg.Result := 0;
  if DroppedArchivesCount > 0 then
  begin
    FilesDropped := True;
    lblDropped.Caption := 'Recognised compressed files/archives dropped: ' + IntToStr(DroppedArchivesCount);
  end
  else
    FilesDropped := False;
  strgrdDropped.Visible := True;

end;

procedure TfrmType.strgrdDroppedClick(Sender: TObject);
var
  i: Integer;
begin
  //Make sure files were dropped, otherwise program would crash
  if not FilesDropped then
    Exit;
  //Clear TStringGrid
  strgrdSelected.Visible := False;
  for i := 0 to strgrdSelected.ColCount - 1 do
    strgrdSelected.Cols[i].Clear;
  strgrdSelected.RowCount := 2; //1 for header
  strgrdSelected.Cells[0, 0] := 'File path';
  strgrdSelected.Cells[1, 0] := 'File name';
  strgrdSelected.Cells[2, 0] := 'Extension';
  strgrdSelected.Cells[3, 0] := 'Size (KB)';
  //Populate TStringGrid with selected archive contents
  with CreateInArchive(DroppedArchivesCLSID[strgrdDropped.Row - 1]) do
  begin
    OpenFile(DroppedArchivesFullNames[strgrdDropped.Row - 1]);
    for i := 0 to NumberOfItems - 1 do
      if not ItemIsFolder[i] then
      begin
        strgrdSelected.Cells[0, strgrdSelected.RowCount - 1] := ExtractFilePath(ItemPath[i]);
        strgrdSelected.Cells[1, strgrdSelected.RowCount - 1] := ExtractFileNameWoExt(ItemPath[i]);
        strgrdSelected.Cells[2, strgrdSelected.RowCount - 1] := UpperCase(StringReplace(ExtractFileExt(ItemPath[i]), '.', '', []));
        strgrdSelected.Cells[3, strgrdSelected.RowCount - 1] := IntToStr(ItemSize[i] div 1024);
        strgrdSelected.RowCount := strgrdSelected.RowCount + 1;
      end;
  end;
  //Only remove last row if matches were found!
  if strgrdSelected.RowCount > 2 then
    strgrdSelected.RowCount := strgrdSelected.RowCount - 1;
  strgrdSelected.Visible := true;
  lblSelected.Caption := 'Files in selected compressed file/archive: ' + IntToStr(strgrdSelected.RowCount - 1);
end;

procedure TfrmType.strgrdFilteredClick(Sender: TObject);
var
  i, SelectedRow: Integer;
begin
  SelectedRow := strgrdFiltered.Row;
  for i := 1 to strgrdDropped.RowCount - 1 do
    if ((strgrdFiltered.Cells[0, SelectedRow] = strgrdDropped.Cells[0, i]) and (strgrdFiltered.Cells[1, SelectedRow] = strgrdDropped.Cells[1, i]) and (strgrdFiltered.Cells[2, SelectedRow] = strgrdDropped.Cells[2, i]) and (strgrdFiltered.Cells[3, SelectedRow] = strgrdDropped.Cells[4, i])) then
    begin
      strgrdDropped.Row := i;
      Break;
    end;
end;

procedure TfrmType.btnFilterClick(Sender: TObject);
var
  i, j: Integer;
  FileNum, FileNumLimit: LongInt;
  Filter, StringFound: Boolean;
  MaxFileSize, MinFileSize, ThisSize, FileSizeLimit: Int64;
  MatchString: string;
begin
  //Make sure files were dropped, otherwise program would crash
  if not FilesDropped then
    Exit;
  //Clear TStringGrid
  strgrdFiltered.Visible := False;
  for i := 0 to strgrdFiltered.ColCount - 1 do
    strgrdFiltered.Cols[i].Clear;
  strgrdFiltered.RowCount := 2; //1 for header
  strgrdFiltered.Cells[0, 0] := 'File name';
  strgrdFiltered.Cells[1, 0] := 'Extension';
  strgrdFiltered.Cells[2, 0] := 'Size (KB)';
  strgrdFiltered.Cells[3, 0] := 'Type detected';
  for i := 0 to DroppedArchivesCount - 1 do
  begin
    Application.ProcessMessages;
    //Populate TStringGrid with selected archive contents
    with CreateInArchive(DroppedArchivesCLSID[i]) do
    begin
      OpenFile(DroppedArchivesFullNames[i]);
      Filter := True;
      //Filter file number
      //rgFileNum.ItemIndex: 0: at least, 1: at most, 2: exactly
      FileNumLimit := StrToInt64(edtFileNum.Text);
      if (FileNumLimit > 0) then
      begin
      //Count files only, i.e. exclude folders
        FileNum := 0;
        for j := 0 to NumberOfItems - 1 do
          if not ItemIsFolder[j] then
            FileNum := FileNum + 1;
        if ((rgFileNum.ItemIndex = 0) and (FileNum < FileNumLimit)) or ((rgFileNum.ItemIndex = 1) and (FileNum > FileNumLimit)) or ((rgFileNum.ItemIndex = 2) and (FileNum <> FileNumLimit)) then
          Filter := False;
      end;
      //File number matches filter: test file size
      //rgFileSize.ItemIndex: 0: at least 1 file smaller, 1: at least 1 file bigger, 2: no file smaller, 3: no file bigger
      FileSizeLimit := StrToInt64(edtFileSize.Text);
      if (Filter and (FileSizeLimit > 0)) then
      begin
        //Get min and max files sizes
        MaxFileSize := 0;
        MinFileSize := MaxLongInt;
        for j := 0 to NumberOfItems - 1 do
          if not ItemIsFolder[j] then
          begin
            ThisSize := ItemSize[j] div 1024;
            if (ThisSize > MaxFileSize) then
              MaxFileSize := ThisSize;
            if (ThisSize < MinFileSize) then
              MinFileSize := ThisSize;
          end;
        if ((rgFileSize.ItemIndex = 0) and (MinFileSize > FileSizeLimit)) or ((rgFileSize.ItemIndex = 1) and (MaxFileSize < FileSizeLimit)) or ((rgFileSize.ItemIndex = 2) and (MinFileSize < FileSizeLimit)) or ((rgFileSize.ItemIndex = 3) and (MaxFileSize > FileSizeLimit)) then
          Filter := False;
      end;
      //File size matches filter: test file names
      //rgFileName: 0: at least one file containing, 1: no file containing
      MatchString := edtFileName.Text;
      if (Filter and (MatchString <> '')) then //if empty string: skip
      begin
        for j := 0 to NumberOfItems - 1 do
          if not ItemIsFolder[j] then
          begin
            StringFound := ContainsText(ExtractFileName(ItemPath[j]), MatchString);
            //String found and we want it
            if ((rgFileName.ItemIndex = 0) and StringFound) then
              Break;
            //String found and we don't want it
            if ((rgFileName.ItemIndex = 1) and StringFound) then
            begin
              Filter := False;
              Break
            end;
          end;
        //String not found (in last file since we exited for loop) and we want it
        if ((rgFileName.ItemIndex = 0) and not StringFound) then
          Filter := False;
      end;
      //Filter still true: add to list
      if Filter then
      begin
        strgrdFiltered.Cells[0, strgrdFiltered.RowCount - 1] := strgrdDropped.Cells[0, i + 1];
        strgrdFiltered.Cells[1, strgrdFiltered.RowCount - 1] := strgrdDropped.Cells[1, i + 1];
        strgrdFiltered.Cells[2, strgrdFiltered.RowCount - 1] := strgrdDropped.Cells[2, i + 1];
        strgrdFiltered.Cells[3, strgrdFiltered.RowCount - 1] := strgrdDropped.Cells[4, i + 1];
        strgrdFiltered.RowCount := strgrdFiltered.RowCount + 1;
      end;
    end;
  end;
  //Only remove last row if matches were found!
  if strgrdFiltered.RowCount > 2 then
    strgrdFiltered.RowCount := strgrdFiltered.RowCount - 1;
  strgrdFiltered.Visible := true;
  if strgrdFiltered.Cells[0, 1] <> '' then
    lblFiltered.Caption := 'Matching archives: ' + IntToStr(strgrdFiltered.RowCount - 1)
  else
    lblFiltered.Caption := 'Matching archives: 0';
end;

end.

