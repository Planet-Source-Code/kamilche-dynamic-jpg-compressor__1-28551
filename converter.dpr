library converter;

{ Important note about DLL memory management: ShareMem must be the
  first unit in your library's USES clause AND your project's (select
  Project-View Source) USES clause if your DLL exports any procedures or
  functions that pass strings as parameters or function results. This
  applies to all strings passed to and from your DLL--even those that
  are nested in records and classes. ShareMem is the interface unit to
  the BORLNDMM.DLL shared memory manager, which must be deployed along
  with your DLL. To avoid using BORLNDMM.DLL, pass string information
  using PChar or ShortString parameters. }

uses
  Windows,
  SysUtils,
  Classes,
  Dialogs,
  Graphics,
  Jpeg;

Function BMPToJPG(InputFilename : PChar;
                 OutputFilename : PChar;
                        Quality : Integer) : Integer; StdCall;
var
  Bitmap : TBitmap;
  Jpg : TJpegImage;
begin
  {Create Bitmap and Jpeg Image}
  Bitmap := TBitmap.Create;
  Jpg := TJpegImage.Create;

  try

    if (Quality < 1) or (Quality > 100) then Quality := 100;

    {Set Compression Quality}
    Jpg.CompressionQuality := Quality;
    Jpg.Compress;

    {Load Bitmap}
    Bitmap.LoadFromFile(InputFilename);

    {Assign Bitmap to Jpeg Image, save it}
    Jpg.Assign(Bitmap);
    Jpg.SaveToFile(ChangeFileExt(OutputFilename, '.jpg'));

    {Clean up}
    Jpg.Free;
    Bitmap.Free;

    {Return success}
    Result := 0;

  except
    {Return failure}
    Result := -1;

   end;

end;

//DLL Entry Point
procedure DLLEntryPoint(dwReason: DWORD); stdcall;
begin
{
   case dwReason of
      DLL_PROCESS_ATTACH: ShowMessage('Process Attach'); //1
      DLL_PROCESS_DETACH: ShowMessage('Process Detach'); //3
      DLL_THREAD_ATTACH : ShowMessage('Thread Attach');  //2
      DLL_THREAD_DETACH : ShowMessage('Thread Detach');  //0
   end;
}
end;

EXPORTS
  DLLEntryPoint,
  BMPToJPG;

begin
    {DLLEntryPoint is the procedure that is invoked every time
     a DLL's entry point is called.}
    DLLProc := @DLLEntryPoint;         {Assign the address of DLLEntryPoint to DLLProc}
    DLLEntryPoint(DLL_PROCESS_ATTACH); {Indicate that the DLL is attaching to the process}
end.
