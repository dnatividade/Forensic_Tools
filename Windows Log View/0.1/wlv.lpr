{
FONTE: https://stackoverflow.com/questions/13139865/how-to-read-contents-of-windows-event-log-using-delphi

TODO: formatar data e hora; parametrizar o sistema para outros logs de eventos

@dnat
}

{$APPTYPE CONSOLE}

//{$R *.res}

uses
  SysUtils,
  ActiveX,
  ComObj,
  Variants;


procedure  GetLogEvents;
const
  wbemFlagForwardOnly = $00000020;
var
  FSWbemLocator : OLEVariant;
  FWMIService   : OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject   : OLEVariant;
  oEnum         : IEnumvariant;
  iValue        : LongWord;
begin;
  FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWMIService   := FSWbemLocator.ConnectServer('localhost', 'root\CIMV2', '', '');
  //FWbemObjectSet:= FWMIService.ExecQuery('SELECT TimeGenerated,Category,ComputerName,EventCode,Message,RecordNumber FROM Win32_NTLogEvent  Where Logfile="System"','WQL',wbemFlagForwardOnly);
  FWbemObjectSet:= FWMIService.ExecQuery('SELECT TimeGenerated,Category,EventCode,Message FROM Win32_NTLogEvent  Where Logfile="System"','WQL',wbemFlagForwardOnly);
  oEnum         := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;
  while oEnum.Next(1, FWbemObject, iValue) = 0 do

  //6005 = Código de inicialização do Log de Eventos - usado como log de inicialização do Windows
  if (Integer(FWbemObject.EventCode) = 6005) then
  begin
    //Writeln(Format('Date/Time         %s',[String(FWbemObject.TimeGenerated)]));
    Writeln(Format('Date/Time         %s',[String(FWbemObject.TimeGenerated)]));
    Writeln(Format('Category          %s',[String(FWbemObject.Category)]));
    //Writeln(Format('Computer Name     %s',[String(FWbemObject.ComputerName)]));
    Writeln(Format('EventCode         %d',[Integer(FWbemObject.EventCode)]));
    Writeln(Format('Message           %s',[String(FWbemObject.Message)]));
    //Writeln(Format('RecordNumber      %d',[Integer(FWbemObject.RecordNumber)]));
    Writeln('--------------------');
    FWbemObject:= Unassigned;
  end;
end;

begin
 try
    CoInitialize(nil);
    try
      GetLogEvents;
    finally
      CoUninitialize;
    end;
 except
    on E:EOleException do
        Writeln(Format('EOleException %s %x', [E.Message,E.ErrorCode]));
    on E:Exception do
        Writeln(E.Classname, ':', E.Message);
 end;
 Writeln('Pressione Enter para sair');
 Readln;
end.

//@dnat
