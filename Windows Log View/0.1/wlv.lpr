{
FONTE: https://stackoverflow.com/questions/13139865/how-to-read-contents-of-windows-event-log-using-delphi

TODO: formatar data e hora; parametrizar o sistema para outros logs de eventos

@dnat
}

{$APPTYPE CONSOLE}

uses
  SysUtils,
  ActiveX,
  ComObj,
  Variants,
  Classes; // Adicionar a unidade Classes para manipulação de arquivos

// Função para converter o formato WMI para DD/MM/YYYY HH:mm:ss
function FormatWmiDateTime(const WmiDateTime: string): string;
var
  Year, Month, Day, Hour, Minute, Second: string;
begin
  // Extraindo partes da data e hora
  Year := Copy(WmiDateTime, 1, 4);
  Month := Copy(WmiDateTime, 5, 2);
  Day := Copy(WmiDateTime, 7, 2);
  Hour := Copy(WmiDateTime, 9, 2);
  Minute := Copy(WmiDateTime, 11, 2);
  Second := Copy(WmiDateTime, 13, 2);

  // Formatar para DD/MM/YYYY HH:mm:ss
  Result := Format('%s/%s/%s %s:%s:%s', [Day, Month, Year, Hour, Minute, Second]);
end;

procedure GetLogEvents;
const
  wbemFlagForwardOnly = $00000020;
var
  FSWbemLocator: OLEVariant;
  FWMIService: OLEVariant;
  FWbemObjectSet: OLEVariant;
  FWbemObject: OLEVariant;
  oEnum: IEnumvariant;
  iValue: LongWord;
  LogFile: TextFile;
begin
  // Criar o arquivo de log
  AssignFile(LogFile, 'LogEvents.txt');
  Rewrite(LogFile);
  try
    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService := FSWbemLocator.ConnectServer('localhost', 'root\CIMV2', '', '');
    FWbemObjectSet := FWMIService.ExecQuery(
    'SELECT TimeGenerated,Category,EventCode,Message FROM Win32_NTLogEvent Where Logfile="System"',
      'WQL',
      wbemFlagForwardOnly
    );
    oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;

    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    begin
      // Verificar se é o evento desejado
      //if Integer(FWbemObject.EventCode) = 6005 then
      if Integer(FWbemObject.EventCode) = 4624 then
      begin
        // Gravar as informações formatadas no arquivo
        Writeln(LogFile, Format('Date/Time:        %s', [FormatWmiDateTime(String(FWbemObject.TimeGenerated))]));
        Writeln(LogFile, Format('Category:         %s', [String(FWbemObject.Category)]));
        Writeln(LogFile, Format('EventCode:        %d', [Integer(FWbemObject.EventCode)]));
        Writeln(LogFile, Format('Message:          %s', [String(FWbemObject.Message)]));
        Writeln(LogFile, '--------------------');
      end;
      FWbemObject := Unassigned;
    end;
  finally
    // Fechar o arquivo
    CloseFile(LogFile);
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
    on E: EOleException do
      Writeln(Format('EOleException %s %x', [E.Message, E.ErrorCode]));
    on E: Exception do
      Writeln(E.Classname, ':', E.Message);
  end;
  Writeln('As informações foram salvas no arquivo "LogEvents.txt".');
  Writeln('Pressione Enter para sair.');
  Readln;
end.



//@dnat
