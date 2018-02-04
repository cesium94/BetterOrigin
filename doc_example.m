    %% 
%     ��֮ǰ������ matlab �Ƚ϶࣬����������ֱ�ӻ�ͼ�����
%     ���ڱ�̱����޸ġ���ʵ���ҳ��� origin �������ݺͻ�ͼ��
%     ������һ�£� origin �����ú�������������Ϻ���������͹���
%     ȷʵҪ�� matlab �� CurveFit ������Ҫ��ö���õö࣬
%     ����matlabĬ�ϻ�������ͼȷʵ��..
% 
%     ���������������ݴ���ͻ�ͼ����origin������Ҫ����ġ�
%     ��������ӵ�һ����������СС������һ�£�Ѱ��һ������
%     �����������Ҳ���ã�����֮��
   %% ����origin�˿�
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
    %% matlab��������
%     Create some data to send over to Origin 
%     This can be replaced with real data 
%     such as result of computation in MATLAB
    mdata = [0.1:0.1:3;10 * sin(0.1:0.1:3);20 * cos(0.1:0.1:3)];
%     Transpose the data so that it can be 
%     placed in Origin worksheet columns
    mdata = mdata';
    %% ��origin WorkBook ���в���
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
    %% ��Origin�л�ͼ
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
%    �������ͨ�� Origin �� Automation Server ���ܣ�
%    ���ȴ�����һ�� Origin COM Server object
%    ��originObj�� ��Ȼ��ͨ�� invoke ������ ��originObj�� ��
%    ��originObj�� �кܶ�ɲ��ݵĲ���
%    ��� Origin �����ĵ� 
%    Programming > Automation Server > ApplicationSI 
% 
%    ��originObj�� �� 'Execute'
%    ������ζ�ſ��Դ���origin�Դ���labtalk���ԣ�
%    ��labtalk���������Ͽ���ʵ��origin�����õ������ɵĹ��ܣ�
%    �ٺ٣�����matlab���൱����ȫ����origin����
%    ��ϸ��Ϣͬ���ο�origin�����ĵ�labtalk���Բ��֣�
%    ���๦������ʵ�� 
        
        
   
    
  
  
  
   
    
    

