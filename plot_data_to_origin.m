    
    %% Preperation
    % Obtain Origin COM Server object
    % This will connect to an existing instance of Origin, or create a new one if none exist
    originObj=actxserver('Origin.ApplicationSI');

    % Make the Origin session visible
    invoke(originObj, 'Execute', 'doc -mc 1;');
       
    % Clear "dirty" flag in Origin to suppress prompt for saving current project
    % You may want to replace this with code to handling of saving current project
    invoke(originObj, 'IsModified', 'false');
    
    % Load the custom project CreateOriginPlot.OPJ found in Samples area of Origin installation
%     invoke(originObj, 'NewProject');
      invoke(originObj, 'ActiveFolder');
   %%
   h = waitbar(0,'Processing...');
   for i=1:length(filenames)
    %% Data
     % Create some data to send over to Origin - create three vectors
    % This can be replaced with real data such as result of computation in MATLAB
    eval(['counts=c',num2str(i),';']);
    t=1:length(counts);
    t=t';
    max=zeros(length(counts),1)+maxm(i);
    min=zeros(length(counts),1)+minm(i);
    data = [t,counts,max,min];
    %% WorkBook

    filename=filenames{i};
    para=str2num(filename(1:end-8));
    if para>=0
       bookname=['plus',num2str(para)];
    else
       bookname=['minus',num2str(abs(para))]; 
    end
    % invoke(originObj, 'Execute', ['newbook name:="',bookname,'" sheet:=1 option:=lsname;']);
    invoke(originObj, 'CreatePage', '2',bookname,'origin');  
     % Send this data over to the Data1 worksheet
    invoke(originObj, 'PutWorksheet', bookname, data);
    %set col type(1:Y,4:X,6:Z),name,units
    invoke(originObj, 'Execute', ['wks.col1.type = 4;'...
                                  ,'wks.col2.type = 1;'...
                                  ,'wks.col3.type = 1;'...
                                  ,'wks.col3.type = 1;'...
                                  ,'col(A)[L]$ = Time;'...
                                  ,'col(B)[L]$ = Counts;'...
                                  ,'col(C)[L]$ = Max;'...
                                  ,'col(D)[L]$ = Min;'...
                                  ,'col(A)[U]$ = a.u.;'...
                                  ,'col(B)[U]$ = a.u.;'...
                                   ]);
   
    %% Plot
    invoke(originObj, 'Execute',...
        ['plotxy iy:=[',bookname,']Sheet1!(1,2:4) plot:=200 '...
        ,'ogl:=[<new template:=origin name:=',bookname,'>];'...
        ]);
    % Rescale the two layers in the graph and copy graph to clipboard
     invoke(originObj, 'Execute', 'page.active = 1; layer - a;');
    %% waitbar
     waitbar(i / length(filenames))
    
        end
   
    %%
    close(h)
%     clear originObj;
        
        
    %Need Help? See origin help£º Labtalk Graphing->Creating Graphs
    %%
    invoke(originObj, 'ActiveFolder');
    invoke(originObj, 'CreatePage', '2',bookname,'origin');  
%      Send this data over to the Data1 worksheet
    invoke(originObj, 'PutWorksheet', bookname, [delay*0.02/3,vis]);
%     set col type(1:Y,4:X,6:Z),name,units
    invoke(originObj, 'Execute', ['wks.col1.type = 4;'...
                                  ,'wks.col2.type = 1;'...
                                  ,'col(A)[L]$ =Delay;'...
                                  ,'col(B)[L]$ =',bookname,' Fringe Contrast;'...
                                  ,'col(A)[U]$ = ns;'...
                                  ,'col(B)[U]$ = a.u.;'...
                                   ]);
    
    invoke(originObj, 'Execute',...
        ['plotxy iy:=[',bookname,']Sheet1!(1,2) plot:=201 '...
        ,'ogl:=[<new template:=origin name:=',bookname,'>];'...
        ]);
%     Rescale the two layers in the graph and copy graph to clipboard
     invoke(originObj, 'Execute', 'page.active = 1; layer - a;');
%     

