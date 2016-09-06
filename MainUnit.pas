unit MainUnit;

interface

uses
  inifiles, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, IdComponent, IdTCPConnection, IdTCPClient, IdMessageClient, ComObj,
  IdSMTP, ComCtrls, StdCtrls, Buttons, ExtCtrls, IdBaseComponent, IdMessage;

type
  TMailerForm = class(TForm)
    MailMessage: TIdMessage;
    pnlTop: TPanel;
    pnlBottom: TPanel;
    ledHost: TLabeledEdit;
    ledAttachment: TLabeledEdit;
    btnAttachment: TBitBtn;
    SMTP: TIdSMTP;
    Account: TLabeledEdit;
    Password: TLabeledEdit;
    btnSendMail: TBitBtn;
    StatusMemo: TMemo;
    AttachmentDialog: TOpenDialog;
    monthSel: TComboBox;
    Month: TLabel;
    cnName: TLabeledEdit;
    ledAttachment2: TLabeledEdit;
    btnAttachment2: TBitBtn;
    procedure btnSendMailClick(Sender: TObject);
    procedure SMTPStatus(ASender: TObject; const AStatus: TIdStatus;
      const AStatusText: string);
    procedure btnAttachmentClick(Sender: TObject);
    procedure btnAttachment2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure writeWorkLog(sqlstr: string);
    procedure RemoveDuplicates(const stringList : TStringList) ;
    function SplitString(const Source, ch: string): TStringList;
     function CheckInputInfo: Boolean;
  private
    procedure GetSettings;
    procedure SaveSettings;
    procedure SetStatusText(const Text: String);

  public
    { Public declarations }
  end;

var
  MailerForm: TMailerForm;

implementation

{$R *.dfm}

procedure TMailerForm.btnSendMailClick(Sender: TObject);

var I, J, K,M,N,X,Y,F,S,D,H: Integer;
Var E:real;
  MaxRow, MaxCol: Integer;
  conList,con2List, titList, tempList,titTmpList,tit2List,tit2TmpList, Strs: TStringList;
  ExcelApp, Sheet: Variant;
  IdBody, IdHtml: TIdText;
  emailList: TStringList;
var temp: string;

begin

  if not CheckInputInfo then
    Exit;

  writeWorkLog('----------------start----------------------');
  StatusMemo.Clear;
             
  conList := TStringList.Create;
  con2List := TStringList.Create;
  titList := TStringList.Create;
  tit2List := TStringList.Create;
  tempList := TStringList.Create;
  Strs := TStringList.Create;
  emailList:=TStringList.Create;
  writeWorkLog('----------------parse excel start----------------------');
   // 创建一个excel的ole对象
  ExcelApp := CreateOleObject('Excel.Application');
  try
     // 打开一个excel文件
    ExcelApp.WorkBooks.Open(ledAttachment.Text);
    conList.BeginUpdate;
    try
      ExcelApp.WorkSheets[1].Activate;
      Sheet := ExcelApp.WorkSheets[1];
     // 有数据的区域的行数和列数
      MaxRow := Sheet.UsedRange.Rows.count;
      MaxCol := Sheet.UsedRange.Columns.count;
     
      for I := 1 to MaxRow do
      begin
        Strs.Clear;
        for J := 1 to MaxCol do

        begin
          // 获得excel的数据第i行，第j列单元格内的数据
          Strs.Add(Sheet.Cells[I, J].Value);
        end;
      
        if I = 1 then
          titList.Add(Strs.CommaText)
        else
          //ShowMessage(Strs.CommaText);
          conList.Add(Strs.CommaText);
          writeWorkLog(Strs.CommaText);
      end;
    
    finally
         // 关闭工作区
      ExcelApp.WorkBooks.Close;

    end;
    if ledAttachment2.Text<>'' then
       begin
       try
        ExcelApp.WorkBooks.Open(ledAttachment2.Text);
        ExcelApp.WorkSheets[1].Activate;
        Sheet := ExcelApp.WorkSheets[1];
        MaxRow := Sheet.UsedRange.Rows.count;
        MaxCol := Sheet.UsedRange.Columns.count;

        for I := 1 to MaxRow do

          begin
            Strs.Clear;
            for J := 1 to MaxCol do
            begin
             Strs.Add(Sheet.Cells[I, J].Value);
            end;

            if I = 1 then
              tit2List.Add(Strs.CommaText)
            else
             con2List.Add(Strs.CommaText);

            writeWorkLog(Strs.CommaText);
         end;
    
      finally
       ExcelApp.WorkBooks.Close;
      end;
    end


  finally
      // 释放ole对象
    ExcelApp.Quit;
    conList.EndUpdate;

  end;

    writeWorkLog('----------------parse excel end----------------------');

   emailList := TStringList.Create;
   
    for k := 0 to conList.Count - 1 do
     begin
      tempList := SplitString(conList[k], ',');
     
      if(tempList[0]='') then  continue;
      if(tempList[tempList.Count-1]='') then  continue;

      emailList.Add(tempList[tempList.Count-1]);
      tempList.Clear;
    end;

    RemoveDuplicates(emailList);

 


   SMTP.Host := ledHost.Text;
   SMTP.Port := 25;
   SMTP.Username := Account.Text; // 帐户名
   SMTP.Password := Password.Text; // 密码
   SMTP.AuthenticationType := atLogin;

    titTmpList:= SplitString(titList[0], ',');
    if tit2List.Count<>0   then
       tit2TmpList:= SplitString(tit2List[0], ',');
       
    for M := 0 to emailList.Count - 1 do
     begin
          //setup mail message
      MailMessage:=TIdMessage.Create(nil);
      MailMessage.NoDecode := False;
      MailMessage.NoEncode := False;
      MailMessage.Encoding := meMIME;
      MailMessage.ContentType := 'multipart/mixed';
      MailMessage.From.Address := Account.Text;
      MailMessage.From.Name := cnName.Text;
      MailMessage.Recipients.EMailAddresses :=emailList[M]; //目的地址
      MailMessage.Subject := '报销单';

      IdBody := TIdText.Create(MailMessage.MessageParts);
      IdBody.ContentType := 'text/plain';
      IdBody.Body.Add('');
  


      IdHtml := TIdText.Create(MailMessage.MessageParts);
      IdHtml.ContentType := 'text/html;charset=gb2312';
      IdHtml.Body.Add('<html>');
      IdHtml.Body.Add('   <head>');
      IdHtml.Body.Add('       <title></title>');
      IdHtml.Body.Add('   </head>');
      IdHtml.Body.Add('   <body> ');
      IdHtml.Body.Add('<div style="font-size: 14px; line-height: 21px;">');
      IdHtml.Body.Add('<span style="font-family: 宋体; font-size: 27px; line-height: 48px;">Dear，'+''+'</span>');
      IdHtml.Body.Add('   </div>');

      IdHtml.Body.Add('<div style="font-size: 14px; line-height: 21px;">');
      IdHtml.Body.Add('<span style="font-family: 宋体; font-size: 27px; line-height: 48px;">下面是您'+monthSel.Text+'份实发的报销明细，请查收！如有疑问请与我联系！</span>');
      IdHtml.Body.Add('</div>');

      IdHtml.Body.Add('<table border="0" cellpadding="0" cellspacing="0" width="1300" style="border-collapse:collapse;">');
      IdHtml.Body.Add('<tbody>');
      IdHtml.Body.Add('<tr height="50" style="mso-height-source:userset;height:37.2pt">');
      for N := 0 to titTmpList.Count-2 do      //邮箱一列不显示
          begin
            IdHtml.Body.Add('<td height="50" class="xl94" width="'+'10%'+'" style="height: 17.2pt;  padding: 0px; color: windowtext; font-size: 9pt; ');
            IdHtml.Body.Add('font-family: 微软雅黑, sans-serif; vertical-align: middle; border: 0.5pt solid windowtext; text-align: center; background-color: rgb(53, 164, 67);">');
            IdHtml.Body.Add(titTmpList[N]);
            IdHtml.Body.Add('</td>');
         end;
      IdHtml.Body.Add('</tr>');

      for X:= 0 to conList.Count - 1 do
           begin
                 tempList := SplitString(conList[X], ',');
                 if(tempList[0]='') then  continue;
                 if(emailList[M]<>tempList[tempList.Count-1]) then continue;

                 IdHtml.Body.Add('<tr height="20" style="mso-height-source:userset;height:15.0pt">');
                 for Y := 0 to tempList.Count-2 do      //邮箱一列不显示
                   begin

                     IdHtml.Body.Add('<td height="20" class="xl101" style="height: 15pt; border: 0.5pt solid windowtext; padding: 0px; color: windowtext; font-size: 8pt;');
                     IdHtml.Body.Add(' font-family: 宋体; vertical-align: middle; white-space: nowrap;">');
                     IdHtml.Body.Add(tempList[Y]);
                     IdHtml.Body.Add('</td>');

                  end;
                 IdHtml.Body.Add('</tr>');
       end;
       IdHtml.Body.Add('</tbody>');
       IdHtml.Body.Add('</table>');

       if tit2List.Count <>0 then
       begin
       IdHtml.Body.Add('<br/>');

       IdHtml.Body.Add('<table border="0" cellpadding="0" cellspacing="0" width="1300" style="border-collapse:collapse;">');
       IdHtml.Body.Add('<tbody>');
       IdHtml.Body.Add('<tr height="50" style="mso-height-source:userset;height:37.2pt">');



      for N := 0 to tit2TmpList.Count-2 do      //邮箱一列不显示
          begin
            IdHtml.Body.Add('<td height="50" class="xl94" width="'+'10%'+'" style="height: 17.2pt;  padding: 0px; color: windowtext; font-size: 9pt; ');
            IdHtml.Body.Add('font-family: 微软雅黑, sans-serif; vertical-align: middle; border: 0.5pt solid windowtext; text-align: center; background-color: rgb(53, 164, 67);">');
            IdHtml.Body.Add(tit2TmpList[N]);
            IdHtml.Body.Add('</td>');
         end;
      IdHtml.Body.Add('</tr>');

      for X:= 0 to con2List.Count - 1 do
           begin
                 tempList := SplitString(con2List[X], ',');
                 if(tempList[0]='') then  continue;
                 if(emailList[M]<>tempList[tempList.Count-1]) then continue;

                 IdHtml.Body.Add('<tr height="20" style="mso-height-source:userset;height:15.0pt">');
                 for Y := 0 to tempList.Count-2 do      //邮箱一列不显示
                   begin

                     IdHtml.Body.Add('<td height="20" class="xl101" style="height: 15pt; border: 0.5pt solid windowtext; padding: 0px; color: windowtext; font-size: 8pt;');
                     IdHtml.Body.Add(' font-family: 宋体; vertical-align: middle; white-space: nowrap;">');
                     IdHtml.Body.Add(tempList[Y]);
                     IdHtml.Body.Add('</td>');

                  end;
                 IdHtml.Body.Add('</tr>');
       end;
       IdHtml.Body.Add('</tbody>');
       IdHtml.Body.Add('</table>');
       
       end;


       

       IdHtml.Body.Add('   </body>');
       IdHtml.Body.Add('</html>');


      try

      //SMTP.Authenticate;
      SMTP.Connect(1000);
      SMTP.Send(MailMessage);
      
      SetStatusText(emailList[M]+'发送成功');
      writeWorkLog(emailList[M]+'--->发送成功');
       inc(S);
      except on E: Exception do   begin
        StatusMemo.Lines.Insert(0, 'ERROR: ' + E.Message);
        writeWorkLog(emailList[M]+'-->ERROR: ' + E.Message);
         inc(F);
       end;
     end;

   SMTP.Disconnect;
   Sleep(200);


   end;


    titList.Clear;
    tit2List.Clear;
    titTmpList.Clear;
    tit2TmpList.Clear;
    conList.Clear;
    con2List.Clear;
    tempList.Clear;
    emailList.Clear;
    E:=0;
    writeWorkLog('成功发送'+IntToStr(S)+'封邮件,'+'失败'+IntToStr(F)+'封邮件');
    ShowMessage('成功发送'+IntToStr(S)+'封邮件,'+'失败'+IntToStr(F)+'封邮件');
    writeWorkLog('----------------end----------------------');

end; (* btnSendMail Click *)

procedure TMailerForm.SMTPStatus(ASender: TObject; const AStatus: TIdStatus;
  const AStatusText: string);
begin
  StatusMemo.Lines.Insert(0, 'Status: ' + AStatusText);
end; (* SMTP Status *)

procedure TMailerForm.SetStatusText(const Text: String);
begin

  StatusMemo.Lines.Insert(0,  Text);

end;


procedure TMailerForm.btnAttachmentClick(Sender: TObject);
begin
  if AttachmentDialog.Execute then
    ledAttachment.Text := AttachmentDialog.FileName;
end;

procedure TMailerForm.btnAttachment2Click(Sender: TObject);
begin
  if AttachmentDialog.Execute then
    ledAttachment2.Text := AttachmentDialog.FileName;
end;



procedure TMailerForm.FormCreate(Sender: TObject);
begin
  GetSettings;
end;

procedure TMailerForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  SaveSettings;
end;

procedure TMailerForm.GetSettings;
var
  ini: TIniFile;
begin
  Ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));
  try
    ledHost.Text := ini.ReadString('SMTP', 'Host', '');
    Account.Text := ini.ReadString('MAIL', 'Account', '');
    Password.Text := ini.ReadString('MAIL', 'Password', '');
    cnName.Text   := ini.ReadString('MAIL', 'CnName', '');
  finally
    ini.Free;
  end;
end; (* GetSettings *)

procedure TMailerForm.SaveSettings;
var
  ini: TIniFile;
begin
  Ini := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));
  try
    ini.WriteString('SMTP', 'Host', ledHost.Text);
    ini.WriteString('MAIL', 'Account', Account.Text);
    ini.WriteString('MAIL', 'Password', Password.Text);
    ini.WriteString('MAIL', 'CnName', cnName.Text);
  finally
    ini.Free;
  end;
end; (* SaveSettings *)




procedure TMailerForm.writeWorkLog(sqlstr: string);
var filev: TextFile;
  ss: string;
begin
  sqlstr := DateTimeToStr(Now) + ' Log: ' + sqlstr;
  //ss:='c:\ErpRunLog.txt';
  ss := ChangeFileExt(Application.ExeName, '.txt');
  if FileExists(ss) then
  begin
    AssignFile(filev, ss);
    append(filev);
    writeln(filev, sqlstr);
  end else begin
    AssignFile(filev, ss);
    ReWrite(filev);
    writeln(filev, sqlstr);
  end;
  CloseFile(filev);
end;



function TMailerForm.SplitString(const Source, ch: string): TStringList;
var
  Temp: string;
  I: Integer;
  chLength: Integer;
begin
  Result := TStringList.Create;
//如果是空自符串则返回空列表
  if Source = '' then Exit;
  Temp := Source;
  I := Pos(ch, Source);
  chLength := Length(ch);
  while I <> 0 do
  begin
    Result.Add(Copy(Temp, 0, I - chLength ));
    Delete(Temp, 1, I - 1 + chLength);
    I := pos(ch, Temp);
  end;
  Result.add(Temp);
end;



function TMailerForm.CheckInputInfo: Boolean;
begin
  Result := True;
  if Trim(ledHost.Text) = '' then
  begin
    Result := False;
    ShowMessage('请输入邮件服务器地址');
    Exit;
  end;
  if Trim(Account.Text) = '' then
  begin
    Result := False;
    ShowMessage('请输入发件人邮箱');
    Exit;
  end;
  if Trim(Password.Text) = '' then
  begin
    Result := False;
    ShowMessage('请输入邮箱密码');
    Exit;
  end;
  if Trim(cnName.Text) = '' then
  begin
    Result := False;
    ShowMessage('请输入发件人中文姓名');
    Exit;
  end;
  if Trim(monthSel.Text) = '月份' then
  begin
    Result := False;
    ShowMessage('请选择报销月份');
    Exit;
  end;
  if Trim(ledAttachment.Text) = '' then
  begin
    Result := False;
    ShowMessage('请选择要发送的文件');
    Exit;
  end;
end;

procedure TMailerForm.RemoveDuplicates(const stringList : TStringList) ;
 var
   buffer: TStringList;
   cnt: Integer;
 begin
   stringList.Sort;
   buffer := TStringList.Create;
   try
     buffer.Sorted := True;
     buffer.Duplicates := dupIgnore;
     buffer.BeginUpdate;
     for cnt := 0 to stringList.Count - 1 do
       buffer.Add(stringList[cnt]) ;
     buffer.EndUpdate;
     stringList.Assign(buffer) ;
   finally
     FreeandNil(buffer) ;
   end;
 end;

end.

