function [RetransformedForecastTS] = RetransformTech(t_for,tranformOuterloop,numTest,OriginalTestDATA,OriginalTrainDATA,POS,TimeSeriesData,Lembda,ConstADD,RangeTS,MinTS,range2,x)

%for i = 1:length(TransformedTechNos)
%     TechNo = TransformedTechNos(i,:);
%     '(1) Ordinary Differencing',...
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
    switch tranformOuterloop
        case 0
            fprintf('case 0');
        
        %--------Starting of  case 1 ------------------------------------
        case 1
            % ----Start Retransformation without using Test data-----------------
            RetransformedForecastTS(1,:) = OriginalTrainDATA(end,:) + t_for(1,:);
            for j=2:numTest
                RetransformedForecastTS(j,:) = t_for(j,:) +  RetransformedForecastTS(j-1,:);
            end
            %----- End without test data
            % -----Start Retransformation with Test data-----------------
%             for j= 2:numTest
%                 RetransformedForecastTS(j,:) = t_for(j,:) +  OriginalTestDATA(j-1,:);
%                 
%             end
        %--------Ending of case 1----------------------------------------
        %--------Starting of  case 2 ------------------------------------
        case 2
            % ----Start Retransformation without using Test data-----------------
%             RetransformedForecastTS=t_for;
            for j=1:POS
            RetransformedForecastTS(j,:) = t_for(j,:) + OriginalTrainDATA(end-POS+j,:);
            end
            
            for k= POS+1:numTest
                RetransformedForecastTS(k,:) = t_for(k,:) +  RetransformedForecastTS(k-POS,:);
            end
            % ----Start Retransformation with test data-----------------
%             for j= 1:numTest
%                 RetransformedForecastTS(j,:) = t_for(j,:) +  TimeSeriesData(end-numTest-POS+j,:);
%                 
%             end                                                          
%             
        case 3
% %             RetransformedForecastTS=t_for;
% %             fprintf('No Retransformation required');
            % ----Start Ratio Retransformation without using Test data-----------------
            RetransformedForecastTS(1,:) = t_for(1,:); % will remain uncommented in both the retransformation codes
            for j=2:numTest
                RetransformedForecastTS(j,:) = (t_for(j,:)+1).*RetransformedForecastTS(j-1,:);
            end
            %----- End without test data
            % -----Start Ratio Retransformation with Test data-----------------
%             for j= 2:numTest
%                 RetransformedForecastTS(j,:) = (t_for(j,:)+1).*OriginalTestDATA(j-1,:);
%                 
%             end
        %--------Ending of case 3----------------------------------------
        
        
         case 4
            % ----Start Log Retransformation without using Test data-----------------
        minTS = min(TimeSeriesData);
        if minTS <=0
             RetransformedForecastTS = t_for; 
        else
            RetransformedForecastTS = exp(t_for);
        end 
            
         case 5
            % ----Start Log Retransformation without using Test data-----------------
            minTS = min(TimeSeriesData);
        if minTS <=0
             RetransformedForecastTS = t_for; 
        else
           RetransformedForecastTS = (t_for).^2;   
        end 
            
         
            
         case 6
            % ----Start Zero-One Normatliztion Retransformation without using Test data-----------------
            RetransformedForecastTS = ((t_for).*RangeTS)+MinTS;  
            
         case 7
            % ----Start X-Y Normatliztion Retransformation without using Test data-----------------
            RetransformedForecastTS = ((t_for)-x)./range2;
            RetransformedForecastTS = ((RetransformedForecastTS).*RangeTS)+MinTS;     
            
         case 8
            % ----Start Retransformation without using Test data-----------------
            RetransformedForecastTS(1,:) = OriginalTrainDATA(end,:) + t_for(1,:);
            for j=2:numTest
                RetransformedForecastTS(j,:) = t_for(j,:) +  RetransformedForecastTS(j-1,:);
            end
            %----- End without test data
            % -----Start Retransformation with Test data-----------------
%             for j= 2:numTest
%                 RetransformedForecastTS(j,:) = t_for(j,:) +  OriginalTestDATA(j-1,:);
%                 
%             end
        %--------Ending of case 1----------------------------------------   
          case 9
            % ----Start Retransformation without using Test data-----------------
            RetransformedForecastTS(1,:) = OriginalTrainDATA(end-1,:) + t_for(1,:);
            RetransformedForecastTS(2,:) = OriginalTrainDATA(end,:) + t_for(2,:);
            for j=3:numTest
                RetransformedForecastTS(j,:) = t_for(j,:) +  RetransformedForecastTS(j-1,:);
            end
            %----- End without test data
            % -----Start Retransformation with Test data-----------------
%             for j= 3:numTest
%                 RetransformedForecastTS(j,:) = t_for(j,:) +  OriginalTestDATA(j-1,:);
%                 
%             end
        %--------Ending of case 1----------------------------------------    
         case 10
            if Lembda == 0
                
                RetransformedForecastTS = exp(t_for);
                for j= 1:numTest
                RetransformedForecastTS(j,:) = RetransformedForecastTS(j,:) - ConstADD;
                end
            else
               
                RetransformedForecastTS = ((t_for .* Lembda) + 1).^(1/Lembda);
                for j= 1:numTest
                RetransformedForecastTS(j,:) = RetransformedForecastTS(j,:) - ConstADD;
                end
            
            end
            
            case 11
                RetransformedForecastTS=t_for;
%             fprintf('No Retransformation required');
            
            case 12
                RetransformedForecastTS=t_for;
%             fprintf('All transfomation removed');
            
      
        %--------Ending of case 10----------------------------------------
        
    end    
    
  
    










% end % end of for loop



end % end of function
