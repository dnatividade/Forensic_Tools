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
  Classes, // Para manipulação de arquivos
  DateUtils, // Para manipulação de datas
  StrUtils; // Para manipulação de strings

// Função para converter a data no formato WMI (yyyyMMddHHmmss.ssssss±UUU)
function GetWmiDateTime(const DateTime: TDateTime): string;
var
  Year, Month, Day, Hour, Min, Sec, MSec: Word;
begin
  DecodeDateTime(DateTime, Year, Month, Day, Hour, Min, Sec, MSec);
  Result := Format('%.4d%.2d%.2d%.2d%.2d%.2d.%.6d+000', [Year, Month, Day, Hour, Min, Sec, MSec]);
end;

// Função para converter a data WMI para TDateTime
function WmiToDateTime(const WmiDateTime: string): TDateTime;
var
  Year, Month, Day, Hour, Min, Sec, MSec: Word;
begin
  Year := StrToInt(Copy(WmiDateTime, 1, 4));
  Month := StrToInt(Copy(WmiDateTime, 5, 2));
  Day := StrToInt(Copy(WmiDateTime, 7, 2));
  Hour := StrToInt(Copy(WmiDateTime, 9, 2));
  Min := StrToInt(Copy(WmiDateTime, 11, 2));
  Sec := StrToInt(Copy(WmiDateTime, 13, 2));
  MSec := StrToInt(Copy(WmiDateTime, 16, 3));

  Result := EncodeDateTime(Year, Month, Day, Hour, Min, Sec, MSec);
end;

// Função para gerar a data limite para o filtro WMI (formato adequado)
function GetDateLimit(NumDays: Integer): string;
var
  DateLimit: TDateTime;
begin
  DateLimit := Now - NumDays; // Subtrai X dias da data atual
  Result := GetWmiDateTime(DateLimit); // Converte a data para o formato WMI
end;

// Função para extrair o nome do usuário da mensagem do evento
function ExtractUserNameFromMessage(const Message: string; EventCode: Integer): string;
var
  StartPos, EndPos: Integer;
begin
  if EventCode = 4647 then
  begin
    // No evento 4647, pode haver o nome do computador em vez do nome do usuário
    // Ajuste para tratar isso corretamente
    Result := 'Desconhecido'; // Retorna 'Desconhecido' para o caso de nome de computador
  end
  else
  begin
    // Procurar por "Nome da Conta:" e extrair o nome do usuário
    StartPos := Pos('Nome da Conta:', Message);
    if StartPos > 0 then
    begin
      // Avançar a posição para logo após "Nome da Conta:"
      StartPos := StartPos + Length('Nome da Conta:') + 1;
      EndPos := Pos(#13, Message, StartPos); // Procurar o final da linha

      if EndPos > StartPos then
        Result := Trim(Copy(Message, StartPos, EndPos - StartPos))
      else
        Result := Trim(Copy(Message, StartPos, Length(Message) - StartPos + 1)); // Até o final da mensagem
    end
    else
      Result := 'Desconhecido'; // Caso não encontre
  end;
end;

procedure GetLogEvents(DaysLimit: Integer; const ExcludedUsers: array of string; const OutputFile: string);
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
  DateLimit: string;
  Usuario: string;
  LogTime: TDateTime;
begin
  // Criar o arquivo de log
  AssignFile(LogFile, OutputFile);
  Rewrite(LogFile);
  try
    DateLimit := GetDateLimit(DaysLimit); // Calcula a data limite

    FSWbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    FWMIService := FSWbemLocator.ConnectServer('localhost', 'root\CIMV2', '', '');

    // Consulta WQL com filtro de tempo e IDs de eventos
    FWbemObjectSet := FWMIService.ExecQuery(
      Format('SELECT TimeGenerated,Category,EventCode,Message FROM Win32_NTLogEvent WHERE Logfile="Security" AND TimeGenerated >= "%s" AND (EventCode=4624 OR EventCode=4625 OR EventCode=4634 OR EventCode=4647)',
             [DateLimit]),
      'WQL',
      wbemFlagForwardOnly
    );

    oEnum := IUnknown(FWbemObjectSet._NewEnum) as IEnumVariant;

    while oEnum.Next(1, FWbemObject, iValue) = 0 do
    begin
      // Inicializar variável de usuário
      Usuario := 'Desconhecido';

      // Verificar qual evento foi encontrado e extrair o nome do usuário
      if Integer(FWbemObject.EventCode) = 4624 then // Logon
      begin
        Usuario := ExtractUserNameFromMessage(String(FWbemObject.Properties_('Message').Value), 4624);
      end
      else if Integer(FWbemObject.EventCode) = 4625 then // Falha de Logon
      begin
        Usuario := ExtractUserNameFromMessage(String(FWbemObject.Properties_('Message').Value), 4625);
      end
      else if Integer(FWbemObject.EventCode) = 4634 then // Logoff
      begin
        Usuario := ExtractUserNameFromMessage(String(FWbemObject.Properties_('Message').Value), 4634);
      end
      else if Integer(FWbemObject.EventCode) = 4647 then // Desconexão RDP
      begin
        Usuario := ExtractUserNameFromMessage(String(FWbemObject.Properties_('Message').Value), 4647);
      end;

      // Ignorar registros de usuários indesejados
      if not (Usuario in ExcludedUsers) then
      begin
        // Converte o tempo do evento para TDateTime
        LogTime := WmiToDateTime(String(FWbemObject.TimeGenerated));

        // Formatar a data e hora como D/MM/YYYY HH:mm:ss
        Writeln(LogFile, Format('%s,"%d","%s"',
          [FormatDateTime('d/mm/yyyy hh:nn:ss', LogTime),
           Integer(FWbemObject.EventCode),
           Usuario]));
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
      // Excluir usuários indesejados
      GetLogEvents(10, ['SISTEMA', 'SERVIÇO LOCAL', 'SERVIÇO DE REDE'], 'Logs_RDP.csv');
    finally
      CoUninitialize;
    end;
  except
    on E: EOleException do
      Writeln(Format('EOleException %s %x', [E.Message, E.ErrorCode]));
    on E: Exception do
      Writeln(E.Classname, ':', E.Message);
  end;
  Writeln('Relatório gerado em: Logs_RDP.csv');
  Writeln('Pressione Enter para sair.');
  Readln;
end.



//@dnat
