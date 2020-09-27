clc;
close all;
clear all;

fileList = dir('*.xlsx');
counter = 1;
FSize = size(fileList);
while counter <= FSize(1,1)

	% reading time sereis form xlsx
	DataRead=fileList(counter).name;
	TimeSeriesData=xlsread(DataRead);
	TimeSeriesData=TimeSeriesData(:,4);
    % plot the data
%     plot(TimeSeriesData);
%     title('Time Series Data Plot');
%     xlabel('Time');
%     ylabel('Time Series values');
    % pause(1)
%     k = waitforbuttonpress;
    MinMSE = inf;
%   
    TotalDataSize=length(TimeSeriesData);
    numTest=floor((TotalDataSize*16)/100);
    % numTest = 67;
    

%   
    NumItt=10;

    %--------------------------------------------------------------------------
    TransformedTechNos = zeros(11,1);
    RangeTS =0; MinTS=0;
    range2 = 0;
    x = 0;
    POS =0;
%     Lembda = 0;   %  used in the power transform
    ConstADD = 0; % Added value to the power transform
    
    Lembda = 0.5;
%     for tranformOuterloop = 1:11
    % TimeSeriesData  = TimeSeriesData;  
%     fprintf('T= %f\n',tranformOuterloop);
tranformOuterloop=1;

    TransformTechReq = 1;
        [FinalTransformedTS,POS,ConstADD,RangeTS,MinTS,range2,x] = TransformationTechnique(TimeSeriesData,tranformOuterloop,Lembda);
   
    %--------------------------------------------------------------------------

   
    % Calcutations for original time series 
    OriginalnumTotalData = length(TimeSeriesData);
    OriginalnumTrainData = OriginalnumTotalData - numTest;
    OriginalStartTestData = (OriginalnumTotalData - numTest)+1;
    OriginalTrainDATA = TimeSeriesData(1:end-numTest,:);
    OriginalTestDATA = TimeSeriesData(end-numTest+1:end,:);
    % variables for test and train data adjstment
   
    % Calcutations for FinalTransformedTS time series 
    numTotalData = length(FinalTransformedTS);
    numTrainData = numTotalData - numTest;
    StartTestData = (numTotalData - numTest)+1;
    TrainDATA = FinalTransformedTS(1:end-numTest,:);
%     TestDATA = FinalTransformedTS(end-numTest+1:end,:);
    ShiftReq = OriginalnumTotalData-numTotalData;
    
%     numTotalData=size(D_org,1);
%     n_val=1;
%     size_of_validation=numTest;
%     numTest=size_of_validation;
%     numTrainData=numTotalData-numTest;
%     warning off;

    
%     TrainDATA=D_org(1:numTrainData);
    % Test_D=D_org((s_train+1):s_org);
%     TestDATA=FinalTransformedTS(numTrainData-inpt_nodes+1:end); 
%     OriginalTestDATA=TimeSeriesData(numTrainData-inpt_nodes+1:numTotalData);
%     numTotalData=size(D_org,1); % SizeOfTransformedData=size(TranformedTS,1);
%     numTrainData=size(TrainDATA,1);  % SizeOfTrain = size(TrainData,1);
%     numTest=size(TestDATA,1);   % SizeOfTestData = size(TestData,1);   
    % plot the data
%     plot(FinalTransformedTS);
%     title('Time Series Data Plot');
%     xlabel('Time');
%     ylabel('Time Series values');


    %--------------------Model code----------------------------------------------------------

%     train_lim=numTrainData-inpt_nodes; 
%     test_lim=numTest-inpt_nodes;
%     p_train=[];t_train=[];
%     p_test=[];t_test=[];
    
    % for loop for making the Input matrix (input pattern)
%     for i=1:train_lim
%     p_train=[p_train TrainDATA(i:i+(inpt_nodes-1))];
%     t_train=[t_train TrainDATA(i+inpt_nodes)];
%     end
%     % for loop for making the output (test) matrix 
%     for i=1:test_lim
%         p_test=[p_test TestDATA(i:i+(inpt_nodes-1))];
%         t_test=[t_test OriginalTestDATA(i+inpt_nodes)];
%     end

    for lag=1:25;

    p_window = windowize(TrainDATA,1:lag+1);
    p_train = p_window(1:end-lag,1:lag); %training set
    t_train = p_window(1:end-lag,end); %training set
    predict_start=TrainDATA(end-lag+1:end,1); %starting point of iterative prediction

    % type = 'function estimation';
    L_fold=10; %l-fold crossvalidation
    % itt loop
    for u = 1:NumItt 
    fprintf('\nFileNo= %f\n',counter);    
    fprintf('T= %f\n',tranformOuterloop);
    fprintf('itt u= %f\n',u);
%     net=newff(p_train, t_train,hdn_nodes,{'logsig','purelin'},'trainlm','learngdm','mse');
   
    model = initlssvm(p_train,t_train,'f',[],[],'RBF_kernel','original');
    model = tunelssvm(model,'simplex','crossvalidatelssvm',{L_fold,'mse'});
    
    % model = initlssvm(p_train,t_train,'f',[],[],'RBF_kernel');
    % model= tunelssvm(model,'simplex','rcrossvalidatelssvm',{L_fold,'mae'},'wmyriad');
    % model=robustlssvm(model);

    D = predict(model,predict_start,numTest);

%     sigma = std(DifferencedTSForWN);
%     R = sigma.*randn(numTotalData,1);
%     for i=1:test_lim
%     t_for(i)=sim(net,T)+R(numTotalData-test_lim+i);
%     T(1:end-1)=T(2:end);
%     T(end)=t_for(i);
%     end
     
%      numTest = test_lim;
     OriginalTestDATA = TimeSeriesData(end-numTest+1:end,:);
     
    DifferencedTSForWN = diff(FinalTransformedTS);
    sigma = std(DifferencedTSForWN);
    R = sigma.*randn(numTotalData,1);
    t_for = D;
    for i=1:numTest
    t_for(i)=t_for(i)+R(numTotalData-numTest+i);
    end
     if TransformTechReq == 1   
     RetransformedForecastTS = zeros(numTest,1);
    [RetransformedForecastTS]= RetransformTech(t_for,tranformOuterloop,numTest,OriginalTestDATA,OriginalTrainDATA,POS,TimeSeriesData,Lembda,ConstADD,RangeTS,MinTS,range2,x);
     else 
         RetransformedForecastTS = t_for;
     end
    % TestDATA = RetransformedForecastTS(end-numTest+1:end,:);

    % fprintf('f s');


%     if ShiftReq == 0
    % figure
    % % if NormYes == 1
    % % h1 = plot(FinalTransformedTS,'Color',[.7,.7,.7]);
    % % else
    %     h1 = plot(TimeSeriesData,'Color',[.7,.7,.7]);
    % % end
    % grid on;
    % hold on;
    % h2 = plot(StartTestData:numTotalData,RetransformedForecastTS,'b','LineWidth',2);
    % % h3 = plot(StartTestData:numTotalData,t_for + 1.96*sqrt(YMSE),'r:','LineWidth',2);
    % % plot(StartTestData:numTotalData,t_for - 1.96*sqrt(YMSE),'r:','LineWidth',2);
    % legend([h1 h2 ],'Observed','Forecast');
    % title({'30-Period Forecasts and Approximate 95% ''Confidence Intervals'})
    % hold off
    % else
    %  SiftedTimeSeries = TimeSeriesData(ShiftReq+1:end);
    % figure
    % % if NormYes == 1
    % % h1 = plot(FinalTransformedTS,'Color',[.7,.7,.7]);
    % % else
    %     h1 = plot(SiftedTimeSeries,'Color',[.7,.7,.7]);
    % % end
    % grid on;
    % hold on;
    % h2 = plot(StartTestData:numTotalData,RetransformedForecastTS,'b','LineWidth',2);
    % % h3 = plot(StartTestData:numTotalData,t_for + 1.96*sqrt(YMSE),'r:','LineWidth',2);
    % % plot(StartTestData:numTotalData,t_for - 1.96*sqrt(YMSE),'r:','LineWidth',2);
    % legend([h1 h2 ],'Observed','Forecast');
    % title({'30-Period Forecasts and Approximate 95% ''Confidence Intervals'})
    % hold off    
%     end

    % -------------------------------------------------------------------------
    % Function calling for Accuracy measures
    [MFE,MAE,SSE,MSE,RMSE,MPE,MAPE]=AccuracyMeasures(OriginalTestDATA,RetransformedForecastTS,numTest);

    if MSE < MinMSE 
        MinMSE = MSE;
        Gamma=model.gam;
        Sig2=model.kernel_pars;
        MinMAE = MAE;
        MinMAPE = MAPE;
        MinLag = lag;
%         InputNode = inpt_nodes;
%         HiddenNode = hdn_nodes;
        MinForecase = RetransformedForecastTS;
        MinTransNo = tranformOuterloop;
        MinIttNo =u;
    end
    end % no of itteration
%     end % hidden node for loop
%     end % input node for loop
    end % lag loop
%     end % outer transform loop

%     fprintf('Minimum value of MSE = %f\n',MinMSE);
%     fprintf('Min error for Technic NO = %f\n',MinTransNo);
%     fprintf('Min error for ittration no = %f\n',MinIttNo);
%     fprintf('Min error for Forecaste=\n');
%     MinForecase
% %     figure
% %     % if NormYes == 1
% %     % h1 = plot(FinalTransformedTS,'Color',[.7,.7,.7]);
% %     % else
% %         h1 = plot(OriginalTestDATA,'Color',[.7,.7,.7]);
% %     % end
% %     grid on;
% %     hold on;
% %     h2 = plot(1:numTest,MinForecase,'b','LineWidth',2);
% %     % h3 = plot(StartTestData:numTotalData,t_for + 1.96*sqrt(YMSE),'r:','LineWidth',2);
% %     % plot(StartTestData:numTotalData,t_for - 1.96*sqrt(YMSE),'r:','LineWidth',2);
% %     legend([h1 h2 ],'Observed','Forecast');
% %     title({'30-Period Forecasts and Approximate 95% ''Confidence Intervals'})
% %     hold off

% 	finalResult = [MinSSE MinTransNo MinIttNo MinForecase];
% 	xlswrite(DataRead,transpose(finalResult),2);
    filename = 'LsSvmRes.xlsx';
%     Headings = {DataRead};
    sheet = counter;
    xlRangeH = 'D2';
    Heading1 = {DataRead};
    xlswrite(filename,Heading1,sheet,xlRangeH);
    
    
    xlRangeTestDataH = 'A7';
    TotalDataInfo = {'TotalData';OriginalnumTotalData;'TestData'};
    xlswrite(filename,TotalDataInfo,sheet,xlRangeTestDataH);
    
    xlRangeTestData = 'A10';
    xlswrite(filename,OriginalTestDATA,sheet,xlRangeTestData);
    
    xlRangeTestDatasizeH = 'C7';
    TotalDataInfo = {'TestData';numTest};
    xlswrite(filename,TotalDataInfo,sheet,xlRangeTestDatasizeH);
    %----------
%    (1) Ordinary Differencing',...
%     '(2) Seasonal Differencing',...
%     '(3) Ratio Transformation',...
%     '(4) Variance Stablization - Logarithm',...
%     '(5) Variance Stablization - Square Root',...
%     '(6) Normatlization Zero One',...
%     '(7) Normatlization at Desired Range',...
%     '(8) Trend Removal Ordinary( 1st order differencing)',...
%     '(9) Trend Removal (2nd orderdifferencing)',...
%     '(10) Power Transformation',...
%     '(11) No Action',...
%     '(12) Remove all Transformations if applied'};
    switch MinTransNo
        case 1
            TechName = '1st Differencing';
        case 2
            TechName = 'NAN';
        case 3
            TechName = 'Ratio Transformation';
        case 4 
            TechName = 'Log Transformation';
        case 5
            TechName = 'Square Root Transformation';
        case 6
            TechName = '0-1 Normatlization';
        case 7
            TechName = '0-1 Normatlization';
        case 8
            TechName = 'Trend Removal';
        case 9
            TechName = '2st Differencing';
        case 10
            TechName = 'Power Transformation';
        case 11
            TechName = 'NAN';
        
    end
    %-------------
    xlRangeT = 'D3';
    TTech = {'Trans. Tech.';TechName};
    xlswrite(filename,TTech,sheet,xlRangeT);

    A = {'MinMSE=';MinMSE;'TrainData';numTrainData;'LSSVM'};
    
    xlRange = 'D5';
    xlswrite(filename,A,sheet,xlRange)
    LagSvm = {'MinMAE';MinMAE;'MinMAPE';MinMAPE;'Lag=';MinLag;'L_fold=';L_fold;'Gamma';Gamma;'Sig2';Sig2;'NumItt';NumItt};
    
    xlRangelag = 'E3';
    xlswrite(filename,LagSvm,sheet,xlRangelag)
    xlRange1 = 'D10';
    xlswrite(filename,MinForecase,sheet,xlRange1);
%     xlswrite(FILE,ARRAY,SHEET,RANGE)

    counter = counter + 1;
end % end while loop

