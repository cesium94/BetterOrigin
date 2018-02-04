    %% 
%     我之前都是用 matlab 比较多，处理完数据直接绘图输出，
%     易于编程便于修改。但实验室常用 origin 处理数据和绘图，
%     尝试了一下， origin 的内置函数比如数据拟合函数的种类和功能
%     确实要比 matlab 的 CurveFit 工具箱要多得多好用得多，
%     而且matlab默认画出来的图确实丑..
% 
%     而对于批量的数据处理和绘图，用origin想来是要疯掉的。
%     正好最近接到一个类似任务，小小调研了一下，寻得一神器，
%     或许对其他人也有用，分享之。
   %% 连接origin端口
%     Obtain Origin COM Server object.
% 
%     This will connect to an existing instance
%     of Origin, or create a new one if none exist
    originObj=actxserver('Origin.ApplicationSI');

%     Make the Origin session visible
    invoke(originObj, 'Execute', 'doc -mc 1;');
       
%     Clear "dirty" flag in Origin to suppress
%     prompt for saving current project
%     You may want to replace this with code 
%     to handling of saving current project
    invoke(originObj, 'IsModified', 'false');
    
%    Create a new origin project
    invoke(originObj, 'NewProject');
    %% matlab数据输入
%     Create some data to send over to Origin 
%     This can be replaced with real data 
%     such as result of computation in MATLAB
    mdata = [0.1:0.1:3;10 * sin(0.1:0.1:3);20 * cos(0.1:0.1:3)];
%     Transpose the data so that it can be 
%     placed in Origin worksheet columns
    mdata = mdata';
    %% 对origin WorkBook 进行操作
%     Create a new workbook named "mydata"
%     invoke(originObj, 'CreatePage', '2','mydata','origin') #
%     # do the same thing
    invoke(originObj, 'Execute'...
        , 'newbook name:="mydata" sheet:=1 option:=lsname;');
%     Send mdata over to the mydata worksheet
    invoke(originObj, 'PutWorksheet', 'mydata', mdata);
%     Set cols' properties
%     1>Y,4>X,6>Z,[L]>name,[U]>units
    invoke(originObj, 'Execute', ['wks.col1.type = 4;'...
                                  ,'wks.col2.type = 1;'...
                                  ,'wks.col3.type = 6;'...
                                  ,'col(A)[L]$ = X;'...
                                  ,'col(B)[L]$ = Y;'...
                                  ,'col(C)[L]$ = Z;'...
                                  ,'col(A)[U]$ = a.u.;'...
                                  ,'col(B)[U]$ = a.u.;'...
                                  ,'col(C)[U]$ = a.u.;'...
                                  ]);
    invoke(originObj, 'Execute', 'page.active = 1; layer - a;');
    %% 在Origin中绘图
%     "plotxy iy:=[mydata]Sheet1!(1,2:3) 
%      plot:=201 ogl:=[<new template:=origin name:=MyGraph>];" #
%      # is labtalk language
    invoke(originObj, 'Execute',...
         ['plotxy iy:=[mydata]Sheet1!(1,2:3)'...
                                 ,'plot:=201'...
        ,'ogl:=[<new template:=origin name:=MyGraph>];']);
%     Rescale the layer in the graph
     invoke(originObj, 'Execute', 'page.active = 1; layer - a;');
%     clear originObj;
   %%
%    这个程序通过 Origin 的 Automation Server 功能，
%    首先创建了一个 Origin COM Server object
%    “originObj” ，然后通过 invoke 语句操纵 “originObj” ，
%    “originObj” 有很多可操纵的参数
%    详见 Origin 帮助文档 
%    Programming > Automation Server > ApplicationSI 
% 
%    “originObj” 的 'Execute'
%    参数意味着可以传递origin自带的labtalk语言，
%    而labtalk语言理论上可以实现origin所有用点击能完成的功能，
%    嘿嘿，这样matlab就相当于完全控制origin啦。
%    详细信息同样参考origin帮助文档labtalk语言部分，
%    更多功能任你实现 
        
        
   
    
  
  
  
   
    
    

