program MailSenderProject;

uses
  Forms,
  MainUnit in 'MainUnit.pas' {MailerForm};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := '����';
  Application.CreateForm(TMailerForm, MailerForm);
  Application.Run;
end.
