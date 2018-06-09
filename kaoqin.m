function kaoqin(identity)


tic;
global Port;
Port=serial('COM7');
set(Port,'BaudRate',9600);%设置波特率
set(Port,'DataBits',8);%8位数据位
set(Port,'StopBits',1);%1位停止位
%delete(instrfindall)
%instrfind
x=datestr(now,26);
i=x(9:10);
pathkaoqin='D:\matlab\bin\work\SRC\xu_lan_GVI\考勤日报表\';
pathkaoqin=strcat(pathkaoqin,i);
pathkaoqin=strcat(pathkaoqin,'.xlsx');

w={datestr(now,13)};
h={'迟到'};
f={'早退'};

x={strcat('MATLAB人脸考勤系统考勤表',x)};

xlswrite(pathkaoqin,x,'Sheet1',['A1']);

if (datestr(now,13)<='09:59:59')
    xlswrite(pathkaoqin,w,'Sheet1',['C',num2str(identity+3)]);
    fopen(Port);
        fwrite(Port,'1');
        fwrite(Port,'1');
        fclose(Port);
    if (datestr(now,13)>='09:00:00')        
        fopen(Port);
        fwrite(Port,'1');
        fwrite(Port,'1');
        fclose(Port);
        xlswrite(pathkaoqin,h,'Sheet1',['I',num2str(identity+3)]);                  
    end
elseif (datestr(now,13)<='12:59:59')
    xlswrite(pathkaoqin,w,'Sheet1',['D',num2str(identity+3)]);
    if (datestr(now,13)<='11:59:59')
        xlswrite(pathkaoqin,f,'Sheet1',['J',num2str(identity+3)]);
        fopen(Port);
         fwrite(Port,'1');
         fwrite(Port,'1');
         fclose(Port);
    end
elseif (datestr(now,13)<='14:59:59')
    xlswrite(pathkaoqin,w,'Sheet1',['E',num2str(identity+3)]);
    if (datestr(now,13)>='14:00:00')
         xlswrite(pathkaoqin,h,'Sheet1',['K',num2str(identity+3)]);    
          fopen(Port);
         fwrite(Port,'1');
         fwrite(Port,'1');
         fclose(Port);
    end
elseif (datestr(now,13)<='18:59:59')
    xlswrite(pathkaoqin,w,'Sheet1',['F',num2str(identity+3)]);           
    if (datestr(now,13)<='16:59:59')
         fopen(Port);
         fwrite(Port,'1');
         fwrite(Port,'1');
         fclose(Port);
        xlswrite(pathkaoqin,f,'Sheet1',['L',num2str(identity+3)]);         
    end
elseif (datestr(now,13)<='19:59:59') 
         
    xlswrite(pathkaoqin,w,'Sheet1',['G',num2str(identity+3)]);    
else (datestr(now,13)<='23:59:59')
    xlswrite(pathkaoqin,w,'Sheet1',['H',num2str(identity+3)]); 
    xx=datestr(now,13);
    aaa=datenum(xx);
    tianshu=datestr(now,26);
    i=tianshu(9:10);
    pathkaoqin111='D:\matlab\bin\work\SRC\xu_lan_GVI\考勤日报表\';
    pathkaoqin111=strcat(pathkaoqin111,i);
    pathkaoqin111=strcat(pathkaoqin111,'.xlsx');
    [num,txt,raw]=xlsread(pathkaoqin111);
    
    hhh=raw{identity+3,7}*24;
    mmm=(hhh-floor(hhh))*60;
    sss=floor((mmm-floor(mmm))*60);
    xxx=strcat(num2str(floor(hhh)),':',num2str(mmm),':',num2str(sss));
    
    xxx=cellstr(xxx);
    
    aa=datenum(xxx);
    a=aaa-aa;
    hh=floor(a*24);
    mm=((a*24)-hh)*60;
    ss=num2str((mm-floor(mm))*60);
    mm=floor(mm);
    Lastest={strcat(num2str(hh),'时',num2str(mm),'分',ss,'秒')};
    xlswrite(pathkaoqin,Lastest,'Sheet1',['M',num2str(identity+3)]);
end
toc;
