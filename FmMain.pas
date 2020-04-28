Unit FmMain;

Interface

Uses
  Messages, SysUtils, Classes, Controls, Forms, Grids, XMLDoc, XMLIntf, sevenzip,
  StdCtrls, ExtCtrls, StrUtils; //, Dialogs;

Type
  TfrmType = Class(TForm)
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
    Procedure FormCreate(Sender: TObject);
    Procedure FormDestroy(Sender: TObject);
    Procedure strgrdDroppedClick(Sender: TObject);
    Procedure btnFilterClick(Sender: TObject);
    Procedure strgrdFilteredClick(Sender: TObject);
  Private
    Procedure WMDropFiles(Var Msg: TWMDropFiles); Message WM_DROPFILES;
  Public
    { Public declarations }

  End;

Var
  frmType: TfrmType;
  HeadDB, DescDB: Array Of String; //XML file contents: header, description
  ClsidDB: Array Of TGUID; //XML file contents: TGUID
  DroppedArchivesCount: Integer;
  DroppedArchivesFullNames: Array Of String;
  DroppedArchivesCLSID: Array Of TGUID;
  DBsize: Integer; //Number of nodes in XML file
  FilesDropped: Boolean;       //Flag to trap error if clicking grid before any files dropped
  ReadFileHeaderSize: Integer; //Max length of header to read

Implementation

Uses
  ShellAPI, UFileCatcher;

{$R *.dfm}

Function IntToStrDelimited(aNum: integer): String;
// Formats the integer aNum with the default Thousand Separator
Var
  D: Double;
Begin
  D := aNum;
  Result := Format('%.0n', [D]); // ".0" -> no decimals, n -> thousands separators
End;

Function GetSizeOfFile(Const FileName: String): Int64;
Var
  Rec: TSearchRec;
Begin
  Result := 0;
  If (FindFirst(FileName, faAnyFile, Rec) = 0) Then
  Begin
    Result := Rec.Size;
    FindClose(Rec);
  End;
End;

Function ExtractFileNameWoExt(Const FileName: String): String;
Var
  i: integer;
Begin
  i := LastDelimiter('.' + PathDelim + DriveDelim, FileName);
  If (i = 0) Or (FileName[i] <> '.') Then
    i := MaxInt;
  Result := ExtractFileName(Copy(FileName, 1, i - 1));
End;

Function ReadFileHeader(Const FileName: String): String;
Var
  Stream: TFileStream;
  Buffer: Array Of AnsiChar;
  i: Integer;
Begin
  SetLength(Buffer, ReadFileHeaderSize);
  //Populate buffer elements
  Stream := TFileStream.Create(FileName, fmOpenRead);
  Try
    Stream.Read(Buffer[0], ReadFileHeaderSize);
  Finally
    Stream.Free;
  End;
  Result := '';
  For i := 0 To ReadFileHeaderSize Do
    Result := Result + IntToHex(Ord(Buffer[i]), 2) + ' ';
End;

Function CompareVsHeader(Const FileHeader: String; Const BaseHeader: String): boolean;
Var
  i: Integer;
Begin
  Result := True;
  For i := 0 To Length(BaseHeader) Do
    If (BaseHeader[i] <> ' ') And (BaseHeader[i] <> '?') And (BaseHeader[i] <> FileHeader[i]) Then
    Begin
      Result := False;
      Exit;
    End;
End;

Procedure TfrmType.FormCreate(Sender: TObject);
Var
  DOC: IXMLDocument;
  FileTypeNode: IXMLNode;
  i, j: Integer;
Begin
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
  For i := 0 To DBsize - 1 Do
  Begin
    FileTypeNode := DOC.ChildNodes.Nodes['filetypes'].ChildNodes[i];
    HeadDB[i] := FileTypeNode.ChildNodes['header'].NodeValue;
    j := length(HeadDB[i]);
    If (j > ReadFileHeaderSize) Then
      ReadFileHeaderSize := j;
    DescDB[i] := FileTypeNode.ChildNodes['description'].NodeValue;
    ClsidDB[i] := StringToGUID(FileTypeNode.ChildNodes['clsid'].NodeValue);
  End;
  ReadFileHeaderSize := (ReadFileHeaderSize + 1) Div 3; //Header in the form XX YY ZZ, so 3 characters per char
  // Tell windows we accept file drops
  DragAcceptFiles(Self.Handle, True);
  FilesDropped := False; //Flag to check files qwere dropped before doing stuff
End;

Procedure TfrmType.FormDestroy(Sender: TObject);
Begin
  // Cancel acceptance of file drops
  DragAcceptFiles(Self.Handle, False);
End;

Procedure TfrmType.WMDropFiles(Var Msg: TWMDropFiles);
Var
  i, j, FileCount: Word;
  Catcher: TFileCatcher; //File catcher class
  FullName, FileHeader, ThisDesc: String;
  ThisFileSize: Int64;
Begin
  Inherited;
  // Create file catcher object to hide all messy details
  Catcher := TFileCatcher.Create(Msg.Drop);
  FileCount := Pred(Catcher.FileCount) + 1; //Not sure why +1 needed
  //Clear TStringGrid
  strgrdDropped.Visible := False;
  For i := 0 To strgrdDropped.ColCount - 1 Do
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
  Try
    //Try to add each dropped file to display
    For i := 0 To FileCount - 1 Do  //-1 due to base 0
    Begin
      //Get file properties
      FullName := Catcher.Files[i];
      //FullFilesNames.Add(FullName);
      ThisFileSize := GetSizeOfFile(FullName);
      FileHeader := ReadFileHeader(FullName);
      //Identify file header
      ThisDesc := '';
      For j := 0 To DBsize - 1 Do
        If CompareVsHeader(FileHeader, HeadDB[j]) Then
        Begin
          ThisDesc := DescDB[j];
          Break;
        End;
      //If header recognised, add to list
      If (ThisDesc <> '') Then
      Begin
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
        strgrdDropped.Cells[2, DroppedArchivesCount] := IntToStrDelimited(ThisFileSize Div 1024);
        strgrdDropped.Cells[3, DroppedArchivesCount] := FileHeader;
        strgrdDropped.Cells[4, DroppedArchivesCount] := ThisDesc;

      End;
    End;
  Finally
    Catcher.Free;
  End;
  //Notify Windows we handled message
  Msg.Result := 0;
  If DroppedArchivesCount > 0 Then
  Begin
    FilesDropped := True;
    lblDropped.Caption := 'Recognised compressed files/archives dropped: ' + IntToStr(DroppedArchivesCount);
  End
  Else
    FilesDropped := False;
  strgrdDropped.Visible := True;

End;

Procedure TfrmType.strgrdDroppedClick(Sender: TObject);
Var
  i: Integer;
Begin
  //Make sure files were dropped, otherwise program would crash
  If Not FilesDropped Then
    Exit;
  //Clear TStringGrid
  strgrdSelected.Visible := False;
  For i := 0 To strgrdSelected.ColCount - 1 Do
    strgrdSelected.Cols[i].Clear;
  strgrdSelected.RowCount := 2; //1 for header
  strgrdSelected.Cells[0, 0] := 'File path';
  strgrdSelected.Cells[1, 0] := 'File name';
  strgrdSelected.Cells[2, 0] := 'Extension';
  strgrdSelected.Cells[3, 0] := 'Size (KB)';
  //Populate TStringGrid with selected archive contents
  With CreateInArchive(DroppedArchivesCLSID[strgrdDropped.Row - 1]) Do
  Begin
    OpenFile(DroppedArchivesFullNames[strgrdDropped.Row - 1]);
    For i := 0 To NumberOfItems - 1 Do
      If Not ItemIsFolder[i] Then
      Begin
        strgrdSelected.Cells[0, strgrdSelected.RowCount - 1] := ExtractFilePath(ItemPath[i]);
        strgrdSelected.Cells[1, strgrdSelected.RowCount - 1] := ExtractFileNameWoExt(ItemPath[i]);
        strgrdSelected.Cells[2, strgrdSelected.RowCount - 1] := UpperCase(StringReplace(ExtractFileExt(ItemPath[i]), '.', '', []));
        strgrdSelected.Cells[3, strgrdSelected.RowCount - 1] := IntToStr(ItemSize[i] Div 1024);
        strgrdSelected.RowCount := strgrdSelected.RowCount + 1;
      End;
  End;
  //Only remove last row if matches were found!
  If strgrdSelected.RowCount > 2 Then
    strgrdSelected.RowCount := strgrdSelected.RowCount - 1;
  strgrdSelected.Visible := true;
  lblSelected.Caption := 'Files in selected compressed file/archive: ' + IntToStr(strgrdSelected.RowCount - 1);
End;

Procedure TfrmType.strgrdFilteredClick(Sender: TObject);
Var
  i, SelectedRow: Integer;
Begin
  SelectedRow := strgrdFiltered.Row;
  For i := 1 To strgrdDropped.RowCount - 1 Do
    If ((strgrdFiltered.Cells[0, SelectedRow] = strgrdDropped.Cells[0, i]) And (strgrdFiltered.Cells[1, SelectedRow] = strgrdDropped.Cells[1, i]) And (strgrdFiltered.Cells[2, SelectedRow] = strgrdDropped.Cells[2, i]) And (strgrdFiltered.Cells[3, SelectedRow] = strgrdDropped.Cells[4, i])) Then
    Begin
      strgrdDropped.Row := i;
      Break;
    End;
End;

Procedure TfrmType.btnFilterClick(Sender: TObject);
Var
  i, j: Integer;
  FileNum, FileNumLimit: LongInt;
  Filter, StringFound: Boolean;
  MaxFileSize, MinFileSize, ThisSize, FileSizeLimit: Int64;
  MatchString: String;
Begin
  //Make sure files were dropped, otherwise program would crash
  If Not FilesDropped Then
    Exit;
  //Clear TStringGrid
  strgrdFiltered.Visible := False;
  For i := 0 To strgrdFiltered.ColCount - 1 Do
    strgrdFiltered.Cols[i].Clear;
  strgrdFiltered.RowCount := 2; //1 for header
  strgrdFiltered.Cells[0, 0] := 'File name';
  strgrdFiltered.Cells[1, 0] := 'Extension';
  strgrdFiltered.Cells[2, 0] := 'Size (KB)';
  strgrdFiltered.Cells[3, 0] := 'Type detected';
  For i := 0 To DroppedArchivesCount - 1 Do
  Begin
  //Populate TStringGrid with selected archive contents
    With CreateInArchive(DroppedArchivesCLSID[i]) Do
    Begin
      OpenFile(DroppedArchivesFullNames[i]);
      Filter := True;
      //Filter file number
      //rgFileNum.ItemIndex: 0: at least, 1: at most, 2: exactly
      FileNumLimit := StrToInt64(edtFileNum.Text);
      If (FileNumLimit > 0) Then
      Begin
      //Count files only, i.e. exclude folders
        FileNum := 0;
        For j := 0 To NumberOfItems - 1 Do
          If Not ItemIsFolder[j] Then
            FileNum := FileNum + 1;
        If ((rgFileNum.ItemIndex = 0) And (FileNum < FileNumLimit)) Or ((rgFileNum.ItemIndex = 1) And (FileNum > FileNumLimit)) Or ((rgFileNum.ItemIndex = 2) And (FileNum <> FileNumLimit)) Then
          Filter := False;
      End;
      //File number matches filter: test file size
      //rgFileSize.ItemIndex: 0: at least 1 file smaller, 1: at least 1 file bigger, 2: no file smaller, 3: no file bigger
      FileSizeLimit := StrToInt64(edtFileSize.Text);
      If (Filter And (FileSizeLimit > 0)) Then
      Begin
        //Get min and max files sizes
        MaxFileSize := 0;
        MinFileSize := MaxLongInt;
        For j := 0 To NumberOfItems - 1 Do
          If Not ItemIsFolder[j] Then
          Begin
            ThisSize := ItemSize[j] Div 1024;
            If (ThisSize > MaxFileSize) Then
              MaxFileSize := ThisSize;
            If (ThisSize < MinFileSize) Then
              MinFileSize := ThisSize;
          End;
        If ((rgFileSize.ItemIndex = 0) And (MinFileSize > FileSizeLimit)) Or ((rgFileSize.ItemIndex = 1) And (MaxFileSize < FileSizeLimit)) Or ((rgFileSize.ItemIndex = 2) And (MinFileSize < FileSizeLimit)) Or ((rgFileSize.ItemIndex = 3) And (MaxFileSize > FileSizeLimit)) Then
          Filter := False;
      End;
      //File size matches filter: test file names
      //rgFileName: 0: at least one file containing, 1: no file containing
      MatchString := edtFileName.Text;
      If (Filter And (MatchString <> '')) Then //if empty string: skip
      Begin
        For j := 0 To NumberOfItems - 1 Do
          If Not ItemIsFolder[j] Then
          Begin
            StringFound := ContainsText(ExtractFileName(ItemPath[j]), MatchString);
            //String found and we want it
            If ((rgFileName.ItemIndex = 0) And StringFound) Then
              Break;
            //String found and we don't want it
            If ((rgFileName.ItemIndex = 1) And StringFound) Then
            Begin
              Filter := False;
              Break
            End;
          End;
        //String not found (in last file since we exited for loop) and we want it
        If ((rgFileName.ItemIndex = 0) And Not StringFound) Then
          Filter := False;
      End;
      //Filter still true: add to list
      If Filter Then
      Begin
        strgrdFiltered.Cells[0, strgrdFiltered.RowCount - 1] := strgrdDropped.Cells[0, i + 1];
        strgrdFiltered.Cells[1, strgrdFiltered.RowCount - 1] := strgrdDropped.Cells[1, i + 1];
        strgrdFiltered.Cells[2, strgrdFiltered.RowCount - 1] := strgrdDropped.Cells[2, i + 1];
        strgrdFiltered.Cells[3, strgrdFiltered.RowCount - 1] := strgrdDropped.Cells[4, i + 1];
        strgrdFiltered.RowCount := strgrdFiltered.RowCount + 1;
      End;
    End;
  End;
  //Only remove last row if matches were found!
  If strgrdFiltered.RowCount > 2 Then
    strgrdFiltered.RowCount := strgrdFiltered.RowCount - 1;
  strgrdFiltered.Visible := true;
  If strgrdFiltered.Cells[0, 1] <> '' Then
    lblFiltered.Caption := 'Matching archives: ' + IntToStr(strgrdFiltered.RowCount - 1)
  Else
    lblFiltered.Caption := 'Matching archives: 0';
End;

End.

