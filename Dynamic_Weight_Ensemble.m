%%
clc;
close all;
clear all;
% Input excel file must contain 1st column as test data 
% and rest of the colums as Model forecasts
%%
fileList = dir('1*.xlsx');   % Test data and it's forecast
fileList2= dir('2*.xlsx');  % Validation data and it's forecast
counter = 1;
[FSize,ColSize] = size(fileList);
while counter <= FSize

	% reading time sereis form xlsx
	DataRead=fileList(counter).name;
    DataRead2=fileList2(counter).name;
    [type,sheetname] = xlsfinfo(DataRead);
    [noofRow,NumOfSheets]=size(sheetname);
    
    for readSheet=1:NumOfSheets
    XlRangeColumn=xlsread(DataRead,readSheet,'C2');
    numTest=XlRangeColumn;
	TimeSeriesData=xlsread(DataRead,readSheet,'A20:J200');
    %AllModelMSEError=xlsread(DataRead,readSheet,'B5:J5'); % For reading MSE of all Models
    ActualTestData=TimeSeriesData(:,1);         % Test data 1 vector
    AllModelForecasts=TimeSeriesData(:,2:end);  % Test data forecaste 9 models
    
    % Dynamic Ensemble variables and matrices
    TimeSeriesData2_val=xlsread(DataRead2,readSheet,'A20:J200');
    Y_vd=TimeSeriesData2_val(:,1); % Y_vd= Validation data 
    
%     Y_Star=Y_Star_val_test;  % Y_Star= Validation data 
    V_validation=TimeSeriesData2_val(:,2:end); % V_validation= all models validation forecast
%     Y_Star_i=Y_Star_i_val_For;   % Y_Star_i= all models validation forecast
    N_ts=XlRangeColumn;             % no of test data or k data points
    F=AllModelForecasts;            % F= Test data forecaste for all n=9 models
    [No_row,n]=size(AllModelForecasts);
    w=zeros(N_ts,n);    %  w = weight matrix inicialized to 0 and obtained form validation data 
    ScaleFactor=1;
    
    %%
    % Dynamic ensemble algo
    w(1,:)=1/n;
    Y_cap_final=[];
    for k=1:N_ts
%         w_tran=w';
%         Multi=zeros(1,9);
%         Multi1=F(k,:);
%         Multi2=w(k,:);
%         Multi=(Multi1).*(Multi2);
%         Y_cap_final(k,:)=sum(Multi);
%         Multi=F(k,:).*w(k,:);
        Y_cap_final(k,:)=sum((F(k,:).*w(k,:)))
        Y_Star=vertcat(Y_vd,Y_cap_final)  % Y*
        Y_Star_i=vertcat(V_validation,F(1:k,:))
        % for loop of i to calculate errors (mse)
        for i=1:n
            t_for=Y_Star_i(:,i)
            numTest2=N_ts+k;
            [MFE,MAE,SSE,MSE,RMSE,MPE,MAPE,SMAPE]=AccuracyMeasures(Y_Star,t_for,numTest2,ScaleFactor);
            w(k+1,i)=1/MSE;
                       
        end
        w(k+1,:)=w(k+1,:)./sum(w(k+1,:))
      
    end % end of for loop k
    
    
%     %%
%     % Code for Error based Weighted Ensemble Tech
%     [rowofErrors,NumOfModels]=size(AllModelMSEError);
%     e=AllModelMSEError;
%     SumofInvertedError=0;
%     for i=1:NumOfModels
%         SumofInvertedError=SumofInvertedError+(1/e(:,i));
%     end
%     Modelweights=zeros(1,NumOfModels);
%     ForecasteMatrix=AllModelForecasts;
%     for i=1:NumOfModels
%         Modelweights(:,i)=(1/e(:,i))/SumofInvertedError;
%         ForecasteMatrix(:,i)=ForecasteMatrix(:,i).*Modelweights(:,i);
%     end
%     ErrBasedFore=sum(ForecasteMatrix,2);
%     
%     
%     
%     
%     
%     %%
%     % Simple Average of all models
%     SimpleMean=mean(AllModelForecasts,2);
%     
%     % Trimed mean code
%     [Modelrow,NumModel]=size(AllModelForecasts);
%     TrimValue=ceil((NumModel*20)/100);
%     SortedData=sort(AllModelForecasts,2);
%     TrimedMean=mean(SortedData(:,(TrimValue+1:end-TrimValue)),2);
%     
%     % Winsorised Mean code
%     for i=1:TrimValue
%         SortedData(:,i)=SortedData(:,TrimValue+1);
%         SortedData(:,end-i+1)=SortedData(:,end-TrimValue);
%     end
%     WinsorisedMean=mean(SortedData,2);
%     
%     % Median code
%     
%     if  rem(NumModel,2)
%         % odd 
%         MedianEnsemble=SortedData(:,(NumModel+1)/2);
%     else
%         % even
%         MedianEnsemble1=SortedData(:,NumModel/2);
%         MedianEnsemble2=SortedData(:,(NumModel/2)+1);
%         MedianEnsemble=((MedianEnsemble1+MedianEnsemble2)./2);
%     end
%     
    
 






%%
 % code for Simple average excel write
 filename = '1Test Data Results of 10 TS for Dynamic Ensemble.xlsx';
 sheet = readSheet;
 % code for ErrorBased excel write
 EnsembleTech={'Dyn_Weigt'};
 xlRangeH = 'T20';
 xlswrite(filename,EnsembleTech,sheet,xlRangeH);
 xlRangeH2 = 'T21';
 xlswrite(filename,Y_cap_final,sheet,xlRangeH2);
 
%  
%  
%  TestDataInfo={'TestData'};
%  xlTestRange='L20';
%  xlswrite(filename,TestDataInfo,sheet,xlTestRange);
%  xlTestRange2='L21';
%  xlswrite(filename,ActualTestData,sheet,xlTestRange2);
%  
%  SimpleAverage={'SimpleAverage'};
%  xlRangeH = 'M20';
%  xlswrite(filename,SimpleAverage,sheet,xlRangeH);
%  xlRangeH2 = 'M21';
%  xlswrite(filename,SimpleMean,sheet,xlRangeH2);
%  
%  % code for Trimed Mean excel write
%  TrimedMeanXL={'TrimedMean'};
%  xlRangeH = 'N20';
%  xlswrite(filename,TrimedMeanXL,sheet,xlRangeH);
%  xlRangeH2 = 'N21';
%  xlswrite(filename,TrimedMean,sheet,xlRangeH2);
%  
%  % code for Winsorised Mean excel write
%  EnsembleTech={'WinsorisedMean'};
%  xlRangeH = 'O20';
%  xlswrite(filename,EnsembleTech,sheet,xlRangeH);
%  xlRangeH2 = 'O21';
%  xlswrite(filename,WinsorisedMean,sheet,xlRangeH2);
%  
%  % code for Median excel write
%  EnsembleTech={'Median'};
%  xlRangeH = 'P20';
%  xlswrite(filename,EnsembleTech,sheet,xlRangeH);
%  xlRangeH2 = 'P21';
%  xlswrite(filename,MedianEnsemble,sheet,xlRangeH2);
%  % code for ErrorBased excel write
%  EnsembleTech={'E_Based'};
%  xlRangeH = 'Q20';
%  xlswrite(filename,EnsembleTech,sheet,xlRangeH);
%  xlRangeH2 = 'Q21';
%  xlswrite(filename,ErrBasedFore,sheet,xlRangeH2);

    end % end for reading sheet of excel file
    counter = counter + 1;
end % end while loop    
