function[MFE,MAE,SSE,MSE,RMSE,MPE,MAPE,SMAPE] = AccuracyMeasures(OriginalTestDATA,RetransformedForecastTS,numTest2,ScaleFactor)
% Forecasting Error = actual data - forecasted data
OriginalTestDATA=OriginalTestDATA./ScaleFactor;
RetransformedForecastTS=RetransformedForecastTS./ScaleFactor;
E = OriginalTestDATA - RetransformedForecastTS;
N= numTest2;

% Mean Forecast Error (MFE) = 1/N (Sigma E)
SumofError = sum(E);
MFE = (SumofError)/N;
% fprintf('Mean_Forecast_Error= %f\n',MFE);
% Mean Absolute Error (MAE) = 1/N (Sigma |E|)
% First calculate the "error" part.
% E= Actual - Predicted;
% Then take the "absolute" value of the "error".
% absoluteErr = abs(E);
% Finally take the "mean" of the "absoluteErr".
% meanAbsoluteErr = mean(absoluteErr);
% or MAE=mean(abs(E));
MAE = mae(E);
% fprintf('Mean_Absolute_Error= %f\n',MAE);
% Sum_of_Squared_Error (SSE)
SSE = sse(E);
% fprintf('Sum_of_Squared_Error= %f\n',SSE);
% Mean_Squared_Error (MSE)
MSE = mse(E);
% fprintf('Mean_Squared_Error= %f\n',MSE);
% Root_Mean_Squared_Error (RMSE)
RMSE=sqrt(mean((E).^2));
% fprintf('Root_Mean_Squared_Error= %f\n',RMSE);
% Mean_Percentage_Error (MPE)
MPE= ((sum(E./OriginalTestDATA))*100)/N;
% fprintf('Mean_Percentage_Error= %f\n',MPE);
% Mean_Absolute_Percentage_Error (MAPE)
MAPE= ((sum(abs(E./OriginalTestDATA)))*100)/N;
% fprintf('Mean_Absolute_Percentage_Error= %f\n',MAPE); 
% Symmetric Mean Absolute Percentage Error (SMAPE)
SMAPE=((sum(abs(E)./((OriginalTestDATA+RetransformedForecastTS)/2)))*100)/N;
end

