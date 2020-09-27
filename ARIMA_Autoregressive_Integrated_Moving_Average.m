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
	TimeSeriesData=TimeSeriesData(:,1);
% plot the data
plot(TimeSeriesData);
title('Time Series Data Plot');
xlabel('Time');
ylabel('Time Series values');
% pause(1)
k = waitforbuttonpress;
MinMSE = inf;
prompt={'Enter Size of Test Data:'};
% Create all your text fields with the questions specified by the variable prompt.
titleofPromt='Prompt for Test Data Size'; 
% The main title of your input dialog interface.
answer=inputdlg(prompt,titleofPromt);
numTest = str2double(answer{1});
% numTest = 67;
% ForecastedMatrix = zeros(numTest,1);
acf = autocorr(TimeSeriesData);
% k = waitforbuttonpress;
autocorr(TimeSeriesData);
 k = waitforbuttonpress;

% pacf() to decide value of p in AR() 
parcorr(TimeSeriesData);
 k = waitforbuttonpress;
prompt={'Enter value of p (order of AR):','Enter value of d(order of differencing):','Enter value of q (order of MA):'};
        % Create all your text fields with the questions specified by the variable prompt.
        titleOfPromt='Prompt for Order of ARIMA'; 
        % The main title of your input dialog interface.
        answer=inputdlg(prompt,titleOfPromt);
        p = str2double(answer{1});
        d = str2double(answer{2});
        q = str2double(answer{3});

prompt={'Enter Number of Itterations:'};
% Create all your text fields with the questions specified by the variable prompt.
titleofPromt='Prompt for itterations'; 
% The main title of your input dialog interface.
answer=inputdlg(prompt,titleofPromt);
NumItt = str2double(answer{1});

%--------------------------------------------------------------------------
TransformedTechNos = zeros(11,1);
RangeTS =0; MinTS=0;
range2 = 0;
x = 0;
POS =0;
Lembda = 0;   %  used in the power transform
ConstADD = 0; % Added value to the power transform

Lembda = 0.5;
% for tranformOuterloop = 1:11
% TimeSeriesData  = TimeSeriesData;  
% fprintf('T= %f\n',tranformOuterloop);

tranformOuterloop=1;
TransformTechReq = 1;
    [FinalTransformedTS,POS,ConstADD,RangeTS,MinTS,range2,x] = TransformationTechnique(TimeSeriesData,tranformOuterloop,Lembda);
% % Input parameter POS = Period of seasonality
% else
%     FinalTransformedTS = TimeSeriesData;
% end    
MvalidTrail=12;
for jj=1:MvalidTrail
    MinMSE = inf;

%
% Calcutations for original time series 
OriginalnumTotalData = length(TimeSeriesData);
OriginalnumTrainData = OriginalnumTotalData - numTest;
OriginalStartTestData = (OriginalnumTotalData - numTest)+1;
OriginalTrainDATA = TimeSeriesData(1:end-numTest,:);
OriginalTestDATA = TimeSeriesData(end-numTest+1:end,:);


% Calcutations for FinalTransformedTS time series 
numTotalData = length(FinalTransformedTS);
numTrainData = numTotalData - numTest;
StartTestData = (numTotalData - numTest)+1;
TrainDATA = FinalTransformedTS(1:end-numTest,:);
TestDATA = FinalTransformedTS(end-numTest+1:end,:);
ShiftReq = OriginalnumTotalData-numTotalData;

ValidationSize=numTest;

    base=numTrainData-ValidationSize-MvalidTrail+1;
    NewTrainData=TrainDATA(1:base+jj-1,:);
    NewValidData=TrainDATA(base+jj:base+jj+ValidationSize-1,:);
    
    OriginalTestDATA=TimeSeriesData(base+jj:base+jj+ValidationSize-1,:);
    OriginalTrainDATA=TimeSeriesData(1:base+jj-1,:);
%     numTotalData=length(TrainDATA);
    numTrainData=length(NewTrainData);
    TrainDATA=NewTrainData;
    TestDATA=NewValidData;
    numTotalData=numTrainData+numTest;


%--------------------Model code----------------------------------------------------------
model=arima(p,d,q);

for u = 1:NumItt 
fprintf('itt u %f\n',u);
 t_for = zeros(numTest,1);
% for changing the estimation settings use the optimset as following:
options = optimset('fmincon');
options = optimset(options,'Algorithm','sqp','TolCon',1e-12,'MaxFunEvals',8000,'Display','iter','Diagnostics','on');
fprintf('est start');
fit = estimate(model,TrainDATA,'options',options);
ar_coeffs=(fit.AR)';
fi=cell2mat(ar_coeffs);
ma_coeffs=(fit.MA)';
theta=cell2mat(ma_coeffs);

 yt=zeros(numTotalData,1);
 yt(1:numTrainData)=TrainDATA;

%     model=arima(p,d,q);
%     fit = estimate(model,TrainDATA);
% 
%     [t_for,YMSE] = forecast(fit,numTest,'Y0',TrainDATA);
%     
%     %finding the order of AR using pacf values
%     
    sigma=sqrt(fit.Variance);
    wn=randn(numTotalData,1).*sigma;
    
    fprintf('Est end');
    for l=(numTrainData+1):numTotalData
        t1=fit.Constant;
        ar_sum=0;
        if p > 0
        for i=1:p
            ar_term(i)=(fi(i)*yt(l-i));
            ar_sum=ar_sum + ar_term(i);    
        end
        end
        
        ma_sum=0;
        if q > 0
        for j=1:q
            ma_term(j)=theta(j)*wn(l-j);
            ma_sum=ma_sum + ma_term(j);
        end
        end
        
        yt(l)=t1+ar_sum + ma_sum + wn(l);
        t_for(l-numTrainData,:)=yt(l);
        %ryt(l-train_size)=t_for(l-train_size)+yt(l-1);
    end
%    ForecastedMatrix = horzcat(ForecastedMatrix,t_for);
        t_for=t_for(1:numTest,:);
 if TransformTechReq == 1   
 RetransformedForecastTS = zeros(numTest,1);
[RetransformedForecastTS]= RetransformTech(t_for,tranformOuterloop,numTest,OriginalTestDATA,OriginalTrainDATA,POS,TimeSeriesData,Lembda,ConstADD,RangeTS,MinTS,range2,x);
 else 
     RetransformedForecastTS = t_for;
 end
% TestDATA = RetransformedForecastTS(end-numTest+1:end,:);
    
% end
 
% -------------------------------------------------------------------------
% Function calling for Accuracy measures
[MFE,MAE,SSE,MSE,RMSE,MPE,MAPE,SMAPE]=AccuracyMeasures(OriginalTestDATA,RetransformedForecastTS,numTest);

if MSE < MinMSE
        MinMSE = MSE;
        MinMAE = MAE;
        MinMAPE=MAPE;
        MinSSE=SSE;
        MinMFE=MFE;
        MinRMSE=RMSE;
        MinMPE=MPE;
        MinSMAPE=SMAPE;
        MinForecase = RetransformedForecastTS;
        MinTransNo = tranformOuterloop;
        MinIttNo =u;
end
end % no of itteration
% end % outer transform loop

% fprintf('Min value of SSE = %f\n',MinSSE);
% fprintf('Min error for Technic NO = %f\n',MinTransNo);
% fprintf('Min error for ittration no = %f\n',MinIttNo);
% fprintf('Min error for Forecaste  =\n');
% MinForecase
filename = 'ARIMAINRAUD.xlsx';
%     Headings = {DataRead};
    sheet = jj;
%     xlRangeH = 'D2';
%     Heading1 = {DataRead};
%     xlswrite(filename,Heading1,1,xlRangeH);
    
    tempjj = 2*jj;
    xlRangeTestDataH = [('A' + tempjj - 2)  '1'];
    TotalDataInfo = {'TotalData';numTotalData};
    xlswrite(filename,TotalDataInfo,1,xlRangeTestDataH);
    xlRangeTestDataH = [('A' + tempjj -1)  '1'];
    TotalDataInfo = {'TrianData';numTrainData};
    xlswrite(filename,TotalDataInfo,1,xlRangeTestDataH);
%     xlRangeTestDataH = 'A1';
%     TotalDataInfo = {'TestData';numTest};
%     xlswrite(filename,TotalDataInfo,1,xlRangeTestDataH);
    
    
    xlRangeTestData = [('A' + tempjj - 2)  '20'];   
%    xlRangeTestData = 'A20';
    TotalDataInfo = {'Val_Data'};
    xlswrite(filename,TotalDataInfo,1,xlRangeTestData);
    
    xlRangeTestData = [('A' + tempjj - 2 )  '21'];
%    xlRangeTestData = 'A21';
    xlswrite(filename,OriginalTestDATA,1,xlRangeTestData);
    
   
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
            TechName = 'S_Diff';
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
    xlRangeT = 'A3';
    TTech = {'Trans. Tech.';'NumItt';'MSE';'MAE';'MAPE';'SSE';'MFE';'RMSE';'MPE';'SMAPE'};
    xlswrite(filename,TTech,1,xlRangeT);
    
    
    xlRangeT = [('A' + tempjj -1 )  '3'];
   % xlRangeT = 'C3';
    TTech2 = {TechName;NumItt;MinMSE;MinMAE;MinMAPE;MinSSE;MinMFE;MinRMSE;MinMPE;MinSMAPE};
    xlswrite(filename,TTech2,1,xlRangeT);
    
    xlRangeTestData = [('A' + tempjj -1)  '20'];   
  % xlRangeTestData = 'C20';
    TotalDataInfo = {'ARIMA'};
    xlswrite(filename,TotalDataInfo,1,xlRangeTestData);
    xlRangeTestData = [('A' + tempjj -1)  '21'];   
   % xlRangeTestData = 'C21';
    xlswrite(filename,MinForecase,1,xlRangeTestData);

end

    counter = counter + 1;
end
