unit UfrmMain;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Forms,
  LYTray, Menus, StdCtrls, Buttons, ADODB,
  ActnList, AppEvnts, ComCtrls, ToolWin, ExtCtrls,
  registry,inifiles,Dialogs,
  StrUtils, DB,ComObj,Variants,Math;

type
  TfrmMain = class(TForm)
    LYTray1: TLYTray;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    ADOConnection1: TADOConnection;
    ApplicationEvents1: TApplicationEvents;
    CoolBar1: TCoolBar;
    ToolBar1: TToolBar;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ActionList1: TActionList;
    editpass: TAction;
    about: TAction;
    stop: TAction;
    ToolButton2: TToolButton;
    ToolButton5: TToolButton;
    ToolButton9: TToolButton;
    OpenDialog1: TOpenDialog;
    ADOConn_BS: TADOConnection;
    BitBtn3: TBitBtn;
    DateTimePicker1: TDateTimePicker;
    Label1: TLabel;
    Label2: TLabel;
    procedure N3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure ApplicationEvents1Activate(Sender: TObject);
    procedure ToolButton7Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure ToolButton5Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
  private
    { Private declarations }
    procedure WMSyscommand(var message:TWMMouse);message WM_SYSCOMMAND;
    procedure UpdateConfig;{配置文件生效}
    function LoadInputPassDll:boolean;
    function MakeDBConn:boolean;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses ucommfunction;

const
  CR=#$D+#$A;
  STX=#$2;ETX=#$3;ACK=#$6;NAK=#$15;
  sCryptSeed='lc';//加解密种子
  //SEPARATOR=#$1C;
  sCONNECTDEVELOP='错误!请与开发商联系!' ;
  IniSection='Setup';

var
  ConnectString:string;
  GroupName:string;//
  SpecType:string ;//
  CombinID:string;//
  LisFormCaption:string;//
  QuaContSpecNoG:string;
  QuaContSpecNo:string;
  QuaContSpecNoD:string;
  EquipChar:string;
  MrConnStr:string;
  ifConnSucc:boolean;
  ifRecLog:boolean;//是否记录调试日志

  hnd:integer;
  bRegister:boolean;

{$R *.dfm}

function ifRegister:boolean;
var
  HDSn,RegisterNum,EnHDSn:string;
  configini:tinifile;
  pEnHDSn:Pchar;
begin
  result:=false;
  
  HDSn:=GetHDSn('C:\')+'-'+GetHDSn('D:\')+'-'+ChangeFileExt(ExtractFileName(Application.ExeName),'');

  CONFIGINI:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));
  RegisterNum:=CONFIGINI.ReadString(IniSection,'RegisterNum','');
  CONFIGINI.Free;
  pEnHDSn:=EnCryptStr(Pchar(HDSn),sCryptSeed);
  EnHDSn:=StrPas(pEnHDSn);

  if Uppercase(EnHDSn)=Uppercase(RegisterNum) then result:=true;

  if not result then messagedlg('对不起,您没有注册或注册码错误,请注册!',mtinformation,[mbok],0);
end;

function GetConnectString:string;
var
  Ini:tinifile;
  userid, password, datasource, initialcatalog: string;
  ifIntegrated:boolean;//是否集成登录模式

  pInStr,pDeStr:Pchar;
  i:integer;
begin
  result:='';
  
  Ini := tinifile.Create(ChangeFileExt(Application.ExeName,'.INI'));
  datasource := Ini.ReadString('连接数据库', '服务器', '');
  initialcatalog := Ini.ReadString('连接数据库', '数据库', '');
  ifIntegrated:=ini.ReadBool('连接数据库','集成登录模式',false);
  userid := Ini.ReadString('连接数据库', '用户', '');
  password := Ini.ReadString('连接数据库', '口令', '107DFC967CDCFAAF');
  Ini.Free;
  //======解密password
  pInStr:=pchar(password);
  pDeStr:=DeCryptStr(pInStr,sCryptSeed);
  setlength(password,length(pDeStr));
  for i :=1  to length(pDeStr) do password[i]:=pDeStr[i-1];
  //==========

  result := result + 'user id=' + UserID + ';';
  result := result + 'password=' + Password + ';';
  result := result + 'data source=' + datasource + ';';
  result := result + 'Initial Catalog=' + initialcatalog + ';';
  result := result + 'provider=' + 'SQLOLEDB.1' + ';';
  //Persist Security Info,表示ADO在数据库连接成功后是否保存密码信息
  //ADO缺省为True,ADO.net缺省为False
  //程序中会传ADOConnection信息给TADOLYQuery,故设置为True
  result := result + 'Persist Security Info=True;';
  if ifIntegrated then
    result := result + 'Integrated Security=SSPI;';
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  ctext        :string;
  reg          :tregistry;
begin
  ConnectString:=GetConnectString;
  
  UpdateConfig;
  DateTimePicker1.DateTime:=now;
  if ifRegister then bRegister:=true else bRegister:=false;  

  lytray1.Hint:='数据接收服务'+ExtractFileName(Application.ExeName);

//=============================初始化密码=====================================//
    reg:=tregistry.Create;
    reg.RootKey:=HKEY_CURRENT_USER;
    reg.OpenKey('\sunyear',true);
    ctext:=reg.ReadString('pass');
    if ctext='' then
    begin
        reg:=tregistry.Create;
        reg.RootKey:=HKEY_CURRENT_USER;
        reg.OpenKey('\sunyear',true);
        reg.WriteString('pass','JIHONM{');
        //MessageBox(application.Handle,pchar('感谢您使用智能监控系统，'+chr(13)+'请记住初始化密码：'+'lc'),
        //            '系统提示',MB_OK+MB_ICONinformation);     //WARNING
    end;
    reg.CloseKey;
    reg.Free;
//============================================================================//
end;

procedure TfrmMain.N3Click(Sender: TObject);
begin
    if not LoadInputPassDll then exit;
    application.Terminate;
end;

procedure TfrmMain.N1Click(Sender: TObject);
begin
  show;
end;

procedure TfrmMain.ApplicationEvents1Activate(Sender: TObject);
begin
  hide;
end;

procedure TfrmMain.WMSyscommand(var message: TWMMouse);
begin
  inherited;
  if message.Keys=SC_MINIMIZE then hide;
  message.Result:=-1;
end;

procedure TfrmMain.ToolButton7Click(Sender: TObject);
begin
  if MakeDBConn then ConnectString:=GetConnectString;
end;

procedure TfrmMain.UpdateConfig;
var
  INI:tinifile;
  autorun:boolean;
begin
  ini:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));

  autorun:=ini.readBool(IniSection,'开机自动运行',false);
  ifRecLog:=ini.readBool(IniSection,'调试日志',false);

  GroupName:=trim(ini.ReadString(IniSection,'工作组',''));
  EquipChar:=trim(uppercase(ini.ReadString(IniSection,'仪器字母','')));//读出来是大写就万无一失了
  SpecType:=ini.ReadString(IniSection,'默认样本类型','');
  CombinID:=ini.ReadString(IniSection,'组合项目代码','');

  LisFormCaption:=ini.ReadString(IniSection,'检验系统窗体标题','');

  QuaContSpecNoG:=ini.ReadString(IniSection,'高值质控联机号','9999');
  QuaContSpecNo:=ini.ReadString(IniSection,'常值质控联机号','9998');
  QuaContSpecNoD:=ini.ReadString(IniSection,'低值质控联机号','9997');

  MrConnStr:=ini.ReadString(IniSection,'连接仪器数据库','');

  ini.Free;

  OperateLinkFile(application.ExeName,'\'+ChangeFileExt(ExtractFileName(Application.ExeName),'.lnk'),15,autorun);

  try
    ADOConn_BS.Connected := false;
    ADOConn_BS.ConnectionString := MrConnStr;
    ADOConn_BS.Connected := true;
    ifConnSucc:=true;
  except
    ifConnSucc:=false;
    showmessage('连接仪器数据库失败!');
  end;
end;

function TfrmMain.LoadInputPassDll: boolean;
TYPE
    TDLLFUNC=FUNCTION:boolean;
VAR
    HLIB:THANDLE;
    DLLFUNC:TDLLFUNC;
    PassFlag:boolean;
begin
    result:=false;
    HLIB:=LOADLIBRARY('OnOffLogin.dll');
    IF HLIB=0 THEN BEGIN SHOWMESSAGE(sCONNECTDEVELOP);EXIT; END;
    DLLFUNC:=TDLLFUNC(GETPROCADDRESS(HLIB,'showfrmonofflogin'));
    IF @DLLFUNC=NIL THEN BEGIN SHOWMESSAGE(sCONNECTDEVELOP);EXIT; END;
    PassFlag:=DLLFUNC;
    FREELIBRARY(HLIB);
    result:=passflag;
end;

function TfrmMain.MakeDBConn:boolean;
var
  newconnstr,ss: string;
  Label labReadIni;
begin
  result:=false;

  labReadIni:
  newconnstr := GetConnectString;
  
  try
    ADOConnection1.Connected := false;
    ADOConnection1.ConnectionString := newconnstr;
    ADOConnection1.Connected := true;
    result:=true;
  except
  end;
  if not result then
  begin
    ss:='服务器'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '数据库'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '集成登录模式'+#2+'CheckListBox'+#2+#2+'0'+#2+#2+#3+
        '用户'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '口令'+#2+'Edit'+#2+#2+'0'+#2+#2+'1';
    if ShowOptionForm('连接数据库','连接数据库',Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
      goto labReadIni else application.Terminate;
  end;
end;

procedure TfrmMain.ToolButton2Click(Sender: TObject);
var
  ss:string;
begin
  if LoadInputPassDll then
  begin
    ss:='连接仪器数据库'+#2+'DBConn'+#2+#2+'1'+#2+#2+#3+
      '工作组'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '仪器字母'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '检验系统窗体标题'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '默认样本类型'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '组合项目代码'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '开机自动运行'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      '调试日志'+#2+'CheckListBox'+#2+#2+'0'+#2+'注:强烈建议在正常运行时关闭'+#2+#3+
      '高值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '常值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '低值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2+#3;

  if ShowOptionForm('',Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
	  UpdateConfig;
  end;
end;

procedure TfrmMain.ToolButton5Click(Sender: TObject);
var
  ss:string;
begin
  ss:='RegisterNum'+#2+'Edit'+#2+#2+'0'+#2+'将该窗体标题栏上的字符串发给开发者,以获取注册码'+#2;
  if bRegister then exit;
  if ShowOptionForm(Pchar('注册:'+GetHDSn('C:\')+'-'+GetHDSn('D:\')+'-'+ChangeFileExt(ExtractFileName(Application.ExeName),'')),Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
    if ifRegister then bRegister:=true else bRegister:=false;
end;

procedure TfrmMain.BitBtn3Click(Sender: TObject);
VAR
  adotemp22:tadoquery;
  SamNo:string;
  ReceiveItemInfo:OleVariant;
  FInts:OleVariant;
  sName,sSex,sAge,sKB,sBLH,sBedNo,sLCZD,sSTATUS,s22,sPath,sFileName:String;
  ls:TStrings;
  i:integer;
begin
  if not ifConnSucc then
  begin
    showmessage('连接仪器数据库失败,不能发送!');
    exit;
  end;
  
  (sender as TBitBtn).Enabled:=false;  

  adotemp22:=tadoquery.Create(nil);
  adotemp22.Connection:=ADOConn_BS;
  adotemp22.Close;
  adotemp22.SQL.Text:='SET NAMES GB2312;';//GB2312//gbk//解决汉字乱码问题
  adotemp22.ExecSQL;
  adotemp22.SQL.Clear;
  adotemp22.SQL.Text:='select si.NAME,si.SEX,si.AGE,si.AGEUNIT,si.HOSPITALIZEDNUM,si.MEDICALRECORDNUM,si.DEPTNUM,si.BEDNUM,si.DATE,si.TIME,'+
                      ' usi.PID,usi.STATUS,usi.TURBIDITY,usi.COLOR,usi.DIAGNOSIS,'+
                      ' urvi.ITEM1,urvi.ITEM2,urvi.ITEM3,urvi.ITEM4,urvi.ITEM5,urvi.ITEM6,urvi.ITEM7,urvi.ITEM8,urvi.ITEM9,urvi.ITEM10,urvi.ITEM11,urvi.ITEM12,urvi.ITEM13,urvi.ITEM14,urvi.Path,urvi.FileName,'+
                      ' udri.ITEM21,udri.ITEM22,udri.ITEM23,udri.ITEM24,udri.ITEM25,udri.ITEM26,udri.ITEM27,udri.ITEM28,udri.ITEM29,udri.ITEM30,udri.ITEM31,udri.ITEM32,udri.ITEM33,udri.ITEM34 '+
                      ' from sick_info si '+
                      ' inner join urine_sample_info usi on si.SID=usi.UID '+
                      ' left join urine_ref_value_index urvi on usi.PID=urvi.SAMPLEID '+
                      ' left join urine_dry_results_index udri on usi.PID=udri.PID '+
                      ' where si.DATE='''+FormatDateTime('YYYYMMDD',DateTimePicker1.Date)+''' ';
  adotemp22.Open;
  while not adotemp22.Eof do
  begin
    SamNo:=adotemp22.fieldbyname('PID').AsString;
    sName:=adotemp22.fieldbyname('NAME').AsString;
    sSex:=adotemp22.fieldbyname('SEX').AsString;
    sAge:=adotemp22.fieldbyname('AGE').AsString+adotemp22.fieldbyname('AGEUNIT').AsString;
    sKB:=adotemp22.fieldbyname('DEPTNUM').AsString;
    s22:='';
    if(adotemp22.fieldbyname('HOSPITALIZEDNUM').AsString<>'')and(adotemp22.fieldbyname('MEDICALRECORDNUM').AsString<>'') then s22:='/';  
    sBLH:=adotemp22.fieldbyname('HOSPITALIZEDNUM').AsString+s22+adotemp22.fieldbyname('MEDICALRECORDNUM').AsString;
    sBedNo:=adotemp22.fieldbyname('BEDNUM').AsString;
    sLCZD:=adotemp22.fieldbyname('DIAGNOSIS').AsString;
    sSTATUS:=adotemp22.fieldbyname('STATUS').AsString;
    sPath:=adotemp22.fieldbyname('Path').AsString;
    if sPath<>'' then
      if sPath[length(sPath)]<>'\' then sPath:=sPath+'\';
      
    sFileName:=adotemp22.fieldbyname('FileName').AsString;
    ls:=TStringList.Create;
    ExtractStrings([#$20],[],pchar(sFileName),ls);

    ReceiveItemInfo:=VarArrayCreate([0,30+ls.Count-1],varVariant);

    ReceiveItemInfo[0]:=VarArrayof(['TURBIDITY',adotemp22.fieldbyname('TURBIDITY').AsString,'','']);
    ReceiveItemInfo[1]:=VarArrayof(['COLOR',adotemp22.fieldbyname('COLOR').AsString,'','']);
    ReceiveItemInfo[2]:=VarArrayof(['ITEM1',adotemp22.fieldbyname('ITEM1').AsString,'','']);
    ReceiveItemInfo[3]:=VarArrayof(['ITEM2',adotemp22.fieldbyname('ITEM2').AsString,'','']);
    ReceiveItemInfo[4]:=VarArrayof(['ITEM3',adotemp22.fieldbyname('ITEM3').AsString,'','']);
    ReceiveItemInfo[5]:=VarArrayof(['ITEM4',adotemp22.fieldbyname('ITEM4').AsString,'','']);
    ReceiveItemInfo[6]:=VarArrayof(['ITEM5',adotemp22.fieldbyname('ITEM5').AsString,'','']);
    ReceiveItemInfo[7]:=VarArrayof(['ITEM6',adotemp22.fieldbyname('ITEM6').AsString,'','']);
    ReceiveItemInfo[8]:=VarArrayof(['ITEM7',adotemp22.fieldbyname('ITEM7').AsString,'','']);
    ReceiveItemInfo[9]:=VarArrayof(['ITEM8',adotemp22.fieldbyname('ITEM8').AsString,'','']);
    ReceiveItemInfo[10]:=VarArrayof(['ITEM9',adotemp22.fieldbyname('ITEM9').AsString,'','']);
    ReceiveItemInfo[11]:=VarArrayof(['ITEM10',adotemp22.fieldbyname('ITEM10').AsString,'','']);
    ReceiveItemInfo[12]:=VarArrayof(['ITEM11',adotemp22.fieldbyname('ITEM11').AsString,'','']);
    ReceiveItemInfo[13]:=VarArrayof(['ITEM12',adotemp22.fieldbyname('ITEM12').AsString,'','']);
    ReceiveItemInfo[14]:=VarArrayof(['ITEM13',adotemp22.fieldbyname('ITEM13').AsString,'','']);
    ReceiveItemInfo[15]:=VarArrayof(['ITEM14',adotemp22.fieldbyname('ITEM14').AsString,'','']);
    ReceiveItemInfo[16]:=VarArrayof(['ITEM21',adotemp22.fieldbyname('ITEM21').AsString,'','']);
    ReceiveItemInfo[17]:=VarArrayof(['ITEM22',adotemp22.fieldbyname('ITEM22').AsString,'','']);
    ReceiveItemInfo[18]:=VarArrayof(['ITEM23',adotemp22.fieldbyname('ITEM23').AsString,'','']);
    ReceiveItemInfo[19]:=VarArrayof(['ITEM24',adotemp22.fieldbyname('ITEM24').AsString,'','']);
    ReceiveItemInfo[20]:=VarArrayof(['ITEM25',adotemp22.fieldbyname('ITEM25').AsString,'','']);
    ReceiveItemInfo[21]:=VarArrayof(['ITEM26',adotemp22.fieldbyname('ITEM26').AsString,'','']);
    ReceiveItemInfo[22]:=VarArrayof(['ITEM27',adotemp22.fieldbyname('ITEM27').AsString,'','']);
    ReceiveItemInfo[23]:=VarArrayof(['ITEM28',adotemp22.fieldbyname('ITEM28').AsString,'','']);
    ReceiveItemInfo[24]:=VarArrayof(['ITEM29',adotemp22.fieldbyname('ITEM29').AsString,'','']);
    ReceiveItemInfo[25]:=VarArrayof(['ITEM30',adotemp22.fieldbyname('ITEM30').AsString,'','']);
    ReceiveItemInfo[26]:=VarArrayof(['ITEM31',adotemp22.fieldbyname('ITEM31').AsString,'','']);
    ReceiveItemInfo[27]:=VarArrayof(['ITEM32',adotemp22.fieldbyname('ITEM32').AsString,'','']);
    ReceiveItemInfo[28]:=VarArrayof(['ITEM33',adotemp22.fieldbyname('ITEM33').AsString,'','']);
    ReceiveItemInfo[29]:=VarArrayof(['ITEM34',adotemp22.fieldbyname('ITEM34').AsString,'','']);

    for i :=0  to ls.Count-1 do
    begin
      ReceiveItemInfo[30+i]:=VarArrayof(['P'+inttostr(i+1),'','',sPath+ls[i]]);
    end;
    ls.Free;

    if bRegister then
    begin
      FInts :=CreateOleObject('Data2LisSvr.Data2Lis');
      FInts.fData2Lis(ReceiveItemInfo,rightstr('0000'+SamNo,4),
        FormatDateTime('YYYY-MM-DD',DateTimePicker1.Date)+' '+adotemp22.fieldbyname('TIME').AsString,
        (GroupName),(SpecType),sSTATUS,(EquipChar),
        (CombinID),
        sName+'{!@#}'+sSex+'{!@#}{!@#}'+sAge+'{!@#}'+sBLH+'{!@#}'+sKB+'{!@#}{!@#}'+sBedNo+'{!@#}'+sLCZD+'{!@#}{!@#}',
        (LisFormCaption),(ConnectString),
        (QuaContSpecNoG),(QuaContSpecNo),(QuaContSpecNoD),'',
        ifRecLog,true,'常规');
      if not VarIsEmpty(FInts) then FInts:= unAssigned;
    end;

    adotemp22.Next;
  end;
  adotemp22.Free;
  
  (sender as TBitBtn).Enabled:=true;
end;

initialization
    hnd := CreateMutex(nil, True, Pchar(ExtractFileName(Application.ExeName)));
    if GetLastError = ERROR_ALREADY_EXISTS then
    begin
        MessageBox(application.Handle,pchar('该程序已在运行中！'),
                    '系统提示',MB_OK+MB_ICONinformation);
        Halt;
    end;

finalization
    if hnd <> 0 then CloseHandle(hnd);

end.
