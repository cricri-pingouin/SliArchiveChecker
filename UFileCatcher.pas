Unit UFileCatcher;

Interface

Uses
  Windows, ShellAPI;

Type
  TFileCatcher = Class(TObject)
  Private
    fDropHandle: HDROP;
    Function GetFile(Idx: Integer): String;
    Function GetFileCount: Integer;
    Function GetPoint: TPoint;
  Public
    Constructor Create(DropHandle: HDROP);
    Destructor Destroy; Override;
    Property FileCount: Integer Read GetFileCount;
    Property Files[Idx: Integer]: String Read GetFile;
    Property DropPoint: TPoint Read GetPoint;
  End;

Implementation

{ TFileCatcher }

Constructor TFileCatcher.Create(DropHandle: HDROP);
Begin
  Inherited Create;
  fDropHandle := DropHandle;
End;

Destructor TFileCatcher.Destroy;
Begin
  DragFinish(fDropHandle);
  Inherited;
End;

Function TFileCatcher.GetFile(Idx: Integer): String;
Var
  FileNameLength: Integer;
Begin
  FileNameLength := DragQueryFile(fDropHandle, Idx, nil, 0);
  SetLength(Result, FileNameLength);
  DragQueryFile(fDropHandle, Idx, PChar(Result), FileNameLength + 1);
End;

Function TFileCatcher.GetFileCount: Integer;
Begin
  Result := DragQueryFile(fDropHandle, $FFFFFFFF, nil, 0);
End;

Function TFileCatcher.GetPoint: TPoint;
Begin
  DragQueryPoint(fDropHandle, Result);
End;

End.

