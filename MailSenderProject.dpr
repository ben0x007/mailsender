program MailSenderProject;

uses
  Forms,
  MainUnit in 'MainUnit.pas' {MailerForm};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := '±¨Ïú';
  Application.CreateForm(TMailerForm, MailerForm);
  Application.Run;
end.
