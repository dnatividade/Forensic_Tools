{
** SOFTWARE PARA LIVRE DISTRIBUIÇÃO E ALTERAÇÃO **

AUTHOR: DNAT
CONNECTIVA REDES DE COMPUTADORES LTDA
Redes, servidores e muito mais...
suporte@connectivaredes.com

Software para mostrar informações quanto a origem de arquivo, usando ADS (Alternate Data Stream)
Isso só funciona para:
- sistema de arquivos NTFS, pois estas informações estão contidas no ADS do arquivo;
- arquivos baixados da internet ou de algum outro computador da rede;
- navegadores ou programas que suportem e gerem informações em ADS (os navegadores como Microsoft Edge, Mozilla Firefox e Google Chrome suportam este recurso.
}

unit principal;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, Buttons, ExtCtrls,
  ComCtrls, StdCtrls, Windows, IniFiles;

type

  { TfrmPrincipal }

  TfrmPrincipal = class(TForm)
    btHelp: TBitBtn;
    btGetSrcFile: TBitBtn;
    log1: TMemo;
    OpenDialog1: TOpenDialog;
    Panel1: TPanel;
    StatusBar1: TStatusBar;
    procedure btHelpClick(Sender: TObject);
    procedure btGetSrcFileClick(Sender: TObject);
  private

  public

  end;

var
  frmPrincipal: TfrmPrincipal;

implementation

{$R *.lfm}

{ TfrmPrincipal }

procedure TfrmPrincipal.btGetSrcFileClick(Sender: TObject);
const
  cZone = ':Zone.Identifier';

var
   sFileName, sFilenameWithStream: string;
   fs: TFileStream;
   oif: TIniFile;

begin
  if OpenDialog1.Execute then
   begin
     sFileName:= OpenDialog1.FileName;
     sFilenameWithStream:= sFileName + cZone;
   end;
  try
    fs:= TFileStream.Create(sFilenameWithStream, fmOpenReadWrite or fmShareDenyNone);
    try
      fs.Seek(0, soFromBeginning);
      oif := TIniFile.Create(fs);
      try
        log1.Clear;
        log1.Lines.Add('Arquivo: '+sFileName);
        log1.Lines.Add('ZoneId: '      +oif.ReadString('ZoneTransfer', 'ZoneId', ''));
        log1.Lines.Add('ZoneTransfer: '+oif.ReadString('ZoneTransfer', 'ReferrerUrl', ''));
        log1.Lines.Add('HostUrl: '     +oif.ReadString('ZoneTransfer', 'HostUrl', ''));
      finally
        oif.Free;
      end;
    finally
      fs.Free;
    end;

  except
    on E: Exception do
    begin
      log1.Clear;
      log1.Lines.Add('Arquivo: '+sFileName);
      log1.Lines.Add('ERRO: Não foi possível acessar ADS');
    end;
  end;
end;

procedure TfrmPrincipal.btHelpClick(Sender: TObject);
begin
  log1.Clear;
  log1.Lines.Add('Modo de usar:');
  log1.Lines.Add('- clique no botão "Obter origem de um arquivo";');
  log1.Lines.Add('- será aberta uma caixa de diálogo para você escolher um arquvo qualquer;');
  log1.Lines.Add('- selecione o arquivo desejado.');
  log1.Lines.Add('');
  log1.Lines.Add('Serão mostradas informações quanto a origem do arquivo.');
  log1.Lines.Add('');
  log1.Lines.Add('Isso só funciona para:');
  log1.Lines.Add('- sistema de arquivos NTFS, pois estas informações estão contidas no ADS do arquivo;');
  log1.Lines.Add('- arquivos baixados da internet ou de algum outro computador da rede;');
  log1.Lines.Add('- navegadores ou programas que suportem e gerem informações em ADS (os navegadores como Microsoft Edge, Mozilla Firefox e Google Chrome suportam este recurso.');
  log1.Lines.Add('');
  log1.Lines.Add('@dnat');
end;

end.

