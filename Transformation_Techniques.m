function[FinalTransformedTS,POS,ConstADD,RangeTS,MinTS,range2,x]=TransformationTechnique(TimeSeriesData,tranformOuterloop,Lembda)
ConstADD = 0; % Added value to the power transform
N=length(TimeSeriesData);
FinalTransformedTS = zeros(N,1);
TransformedTS = zeros(N,1);
range2 =0;
x =0;
POS = 0;
MinTS = 0; % mininmum of time series  used in 0-1 normalization
RangeTS=0; % Max - min of time series used in 0-1 normalization
OriginalTS=TimeSeriesData;
% FinalTransformedTS = TimeSeriesData;
% NormYes = 0;
% TransREQ = 1;
% TransformedTechNos = zeros(10,1);
% while TransREQ == 1
    
    TStoBeTranformed = TimeSeriesData;
%     % ------- Prompt box code-----------------------
%     quest='Please Enter The Number of Transformation Technique you desire:';
% alt={'(1) Ordinary Differencing',...
%     '(2) Seasonal Differencing',...
%     '(3) Ratio Transformation',...
%     '(4) Variance Stablization - Logarithm',...
%     '(5) Variance Stablization - Square Root',...
%     '(6) Normatlization Zero One',...
%     '(7) Normatlization at Desired Range',...
%     '(8) Trend Removal Ordinary( 1st order differencing)',...
%     '(9) Trend Removal (2nd orderdifferencing)',...
%     '(10) Power Transformation',...
%     '(11) Remove all Transformations if applied',...
%     '(12) No Action'};
% string=[quest,'\n'];
% for j=1:length(alt)
%     string=[string,'\n',alt{j}];
% end
% string=[string,'\n'];
% string=sprintf(string);%this makes \n work
% prompt={string};
% def={'1'};
% dlgTitle='Transformation Technique';
% lineNo=1;
% SelectedTech = str2double(inputdlg(prompt,dlgTitle,lineNo,def));
%-----------------end promt box-------------------------------------------
%------------------switch case for Technique selection--------------------
% Choice = 1;
switch tranformOuterloop
    case 1
       
        TransformedTS = diff(TStoBeTranformed);
        FinalTransformedTS = TransformedTS;
%         k = waitforbuttonpress;
%         %---------------
% %          prompt={'Please Enter 1 if you satisfied with Transformation, 0 Otherwise:'};
% %         % Create all your text fields with the questions specified by the variable prompt.
% %         titleOfPromt='Transformation / Preprocessing Techniques Verification'; 
% %         % The main title of your input dialog interface.
% %         answer=inputdlg(prompt,titleOfPromt);
% %         Choice = str2double(answer{1}); 
%             
%         if Choice == 1
%             FinalTransformedTS = TransformedTS;
%             %------------Code for Transformation Technique no -------------
%             for checkTrans = 1:10
%                 if( TransformedTechNos(checkTrans,:) == 0)
%                     TransformedTechNos(checkTrans,:)= SelectedTech;
%                      break
%                 end 
%             end
%             %-------------------End Transformation Technique no------------
%         else 
%             FinalTransformedTS = TStoBeTranformed;
%         end
        
    case 2
%         %------------- Seasonal Differencing Start-------------------------
%        FinalTransformedTS = TStoBeTranformed;
        N = length(TStoBeTranformed);
%         prompt={'Please Enter Period of seasonality:'};
%         % Create all your text fields with the questions specified by the variable prompt.
%         titleOfPromt='Period of seasonality'; 
%         % The main title of your input dialog interface.
%         answer=inputdlg(prompt,titleOfPromt);
%         POS = str2double(answer{1});
        POS = 12;
%         figure
%         plot(TStoBeTranformed)
%         xlim([0,N])
%         title('Time Series Data')
%         waitforbuttonpress;
        D12 = LagOp({1 -1},'Lags',[0,POS]);

        TransformedTS = filter(D12,TStoBeTranformed);
        diffS= length(TStoBeTranformed) - length(TransformedTS);
%         figure
%         plot(diffS+1:N,TransformedTS)
%         xlim([0,N])
%         title('Seasonal Differenced Time sereis ');
        %---------------
%          prompt={'Please Enter 1 if you satisfied with Transformation, 0 Otherwise:'};
%         % Create all your text fields with the questions specified by the variable prompt.
%         titleOfPromt='Transformation / Preprocessing Techniques Verification'; 
%         % The main title of your input dialog interface.
%         answer=inputdlg(prompt,titleOfPromt);
%         Choice = str2double(answer{1}); 
%         if Choice == 1
            FinalTransformedTS = TransformedTS;
%             %------------Code for Transformation Technique no -------------
%             for checkTrans = 1:10
%                 if( TransformedTechNos(checkTrans,:) == 0)
%                     TransformedTechNos(checkTrans,:)= SelectedTech;
%                      break
%                 end 
%             end
            %-------------------End Transformation Technique no------------
%         else 
%             FinalTransformedTS = TStoBeTranformed;
%         end
% %    %-------------------------Start of Case 3-------------------------------------
%    
    case 3
%         MinTStoBeTranformed = min(TStoBeTranformed);
%        fprintf('MinTStoBeTranformed  3Case=%f\n',MinTStoBeTranformed);
%        fprintf('Case 3\n');
%         FinalTransformedTS = OriginalTS;
        N = length(TStoBeTranformed);
        TransformedTS = zeros(N-1,1);
        for i=1:N-1
            if TStoBeTranformed(i,:)==0
                 TransformedTS(i,:) = mean(TransformedTS);
            else
                TransformedTS(i,:) = (TStoBeTranformed(i+1,:)./TStoBeTranformed(i,:))-1;
            end
        end
        FinalTransformedTS = TransformedTS;
        %---------------Ploting the Transformed time series-------------
        
%         plot(TransformedTS);
%         title('Ratio Transformed values Time Series');
%         xlabel('Time')
%         ylabel('Ratio Transformed values of Time Series')
%         k = waitforbuttonpress;
        %-------------------verifing for updation--------------------------
%          prompt={'Please Enter 1 if you satisfied with Transformation, 0 Otherwise:'};
%         % Create all your text fields with the questions specified by the variable prompt.
%         titleOfPromt='Transformation / Preprocessing Techniques Verification'; 
%         % The main title of your input dialog interface.
%         answer=inputdlg(prompt,titleOfPromt);
%         Choice = str2double(answer{1}); 
%         if Choice == 1
            % Total no of rows in Final time series becomes N-1
            % Only here Size of FinalTransformedTS = N-1:1
%             FinalTransformedTS = FinalTransformedTS(1:end-1,:);
%             for j=1:N-1
%             FinalTransformedTS(j,:) = TransformedTS(j,:);
%             end
%             %------------Code for Transformation Technique no -------------
%             for checkTrans = 1:10
%                 if( TransformedTechNos(checkTrans,:) == 0)
%                     TransformedTechNos(checkTrans,:)= SelectedTech;
%                      break
%                 end 
%             end
%             %-------------------End Transformation Technique no------------
%         else 
%             % Here again Size of FinalTransformedTS = N:1
%             FinalTransformedTS = TStoBeTranformed;
%         end
%         
       
    case 4
%        MinTStoBeTranformed = min(TStoBeTranformed);
%        fprintf('MinTStoBeTranformed Case 4=%f',MinTStoBeTranformed);
%        fprintf('Case 4');
%        TStoBeTranformed
        minTS = min(TimeSeriesData);
        if minTS <=0
             FinalTransformedTS = TimeSeriesData; 
        else
        TransformedTS = log(TStoBeTranformed);
        FinalTransformedTS = TransformedTS;
   
        end
       %---------------Ploting the Transformed time series-------------
        
%         plot(TransformedTS);
%         title('Log Transformed values Time Series');
%         xlabel('Time')
%         ylabel('Log Transformed values of Time Series')
%         waitforbuttonpress;
        %-------------------verifing for updation--------------------------
%         prompt={'Please Enter 1 if you satisfied with Transformation, 0 Otherwise:'};
%         % Create all your text fields with the questions specified by the variable prompt.
%         titleOfPromt='Transformation / Preprocessing Techniques Verification'; 
%         % The main title of your input dialog interface.
%         answer=inputdlg(prompt,titleOfPromt);
%         Choice = str2double(answer{1}); 
%         if Choice == 1
            
%             FinalTransformedTS
            %------------Code for Transformation Technique no -------------
%             for checkTrans = 1:10
%                 if( TransformedTechNos(checkTrans,:) == 0)
%                     TransformedTechNos(checkTrans,:)= SelectedTech;
%                      break
%                 end 
%             end
%             %-------------------End Transformation Technique no------------
%         else 
%             FinalTransformedTS = TStoBeTranformed;
%         end
%         % ----------------end of case4 ----------------------------------
    
    case 5
        minTS = min(TimeSeriesData);
        if minTS <0
             FinalTransformedTS = TimeSeriesData; 
        else
        TransformedTS = sqrt(TStoBeTranformed);
        FinalTransformedTS = TransformedTS;
        end
       %---------------Ploting the Transformed time series-------------
%         
%         plot(TransformedTS);
%         title('Square Root Transformed values Time Series');
%         xlabel('Time')
%         ylabel('Square Root values of Time Series')
%         waitforbuttonpress;
%         %-------------------verifing for updation--------------------------
%         prompt={'Please Enter 1 if you satisfied with Transformation, 0 Otherwise:'};
%         % Create all your text fields with the questions specified by the variable prompt.
%         titleOfPromt='Transformation / Preprocessing Techniques Verification'; 
%         % The main title of your input dialog interface.
%         answer=inputdlg(prompt,titleOfPromt);
%         Choice = str2double(answer{1}); 
%         if Choice == 1
            
            %------------Code for Transformation Technique no -------------
%             for checkTrans = 1:10
%                 if( TransformedTechNos(checkTrans,:) == 0)
%                     TransformedTechNos(checkTrans,:)= SelectedTech;
%                      break
%                 end 
%             end
%             %-------------------End Transformation Technique no------------
%         else 
%             FinalTransformedTS = TStoBeTranformed;
%         end
        % ----------------end of case5----------------------------------
    case 6
        % ---Normatlization Zero One starts------------------------------
        N = length(TStoBeTranformed);
        MaxTS = max(TStoBeTranformed);
        MinTS = min(TStoBeTranformed);
        RangeTS = MaxTS - MinTS;
        TransformedTS = zeros(N:1);
        for i=1:N
            TransformedTS(i,:) = ((TStoBeTranformed(i,:)-MinTS)./RangeTS);
        end
        %---------------Ploting the Transformed time series-------------
        
%         plot(TransformedTS);
%         title('Normalized Time Series');
%         xlabel('Time')
%         ylabel('Normalized values of Time Series')
%         k = waitforbuttonpress;
        %-------------------verifing for updation--------------------------
%         prompt={'Please Enter 1 if you satisfied with Transformation, 0 Otherwise:'};
%         % Create all your text fields with the questions specified by the variable prompt.
%         titleOfPromt='Transformation / Preprocessing Techniques Verification'; 
%         % The main title of your input dialog interface.
%         answer=inputdlg(prompt,titleOfPromt);
%         Choice = str2double(answer{1}); 
%         if Choice == 1
%             NormYes = 1;
            FinalTransformedTS = TransformedTS;
            %------------Code for Transformation Technique no -------------
%             for checkTrans = 1:10
%                 if( TransformedTechNos(checkTrans,:) == 0)
%                     TransformedTechNos(checkTrans,:)= SelectedTech;
%                      break
%                 end 
%             end
%             %-------------------End Transformation Technique no------------
%         else 
%             FinalTransformedTS = TStoBeTranformed;
%         end
        % ----------------end of case6 ----------------------------------
        
      
    case 7
        % ---Normatlization Zero One starts------------------------------
        N = length(TStoBeTranformed);
        MaxTS = max(TStoBeTranformed);
        MinTS = min(TStoBeTranformed);
        RangeTS = MaxTS - MinTS;
        TransformedTS = zeros(N:1);
        for i=1:N
            TransformedTS(i,:) = ((TStoBeTranformed(i,:)-MinTS)./RangeTS);
        end
%         % Then scale to [x,y]:
%         prompt={'Enter value of x (starting range):','Enter value of y (ending range)'};
%         % Create all your text fields with the questions specified by the variable prompt.
%         titleOfPromt='Transformation / Preprocessing Techniques Verification'; 
%         % The main title of your input dialog interface.
%         answer=inputdlg(prompt,titleOfPromt);
%         x = str2double(answer{1});
%         y = str2double(answer{2});
        x = 0 ;
        y = 1;
        range2 = y - x;
        XYnormalizedTS = zeros(N:1);
        for i=1:N
            XYnormalizedTS(i,:) = ((TransformedTS(i,:)*range2)+x);
        end
        %---------------Ploting the Transformed time series-------------
        TransformedTS = XYnormalizedTS;
%         plot(TransformedTS);
%         title('Normalized Time Series');
%         xlabel('Time')
%         ylabel('Normalized values of Time Series')
%         k = waitforbuttonpress;
        %-------------------verifing for updation--------------------------
%         prompt={'Please Enter 1 if you satisfied with Transformation, 0 Otherwise:'};
%         % Create all your text fields with the questions specified by the variable prompt.
%         titleOfPromt='Transformation / Preprocessing Techniques Verification'; 
%         % The main title of your input dialog interface.
%         answer=inputdlg(prompt,titleOfPromt);
%         Choice = str2double(answer{1}); 
%         if Choice == 1
%             NormYes = 1;
            FinalTransformedTS = TransformedTS;
%             %------------Code for Transformation Technique no -------------
%             for checkTrans = 1:10
%                 if( TransformedTechNos(checkTrans,:) == 0)
%                     TransformedTechNos(checkTrans,:)= SelectedTech;
%                      break
%                 end 
%             end
%             %-------------------End Transformation Technique no------------
%         else 
%             FinalTransformedTS = TStoBeTranformed;
%         end
%         % ----------------end of case7 ----------------------------------
    case 8
        % -------------Start case 8---------------------------------------
        N = length(TStoBeTranformed);

%         figure
%         plot(TStoBeTranformed)
%         xlim([0,N])
%         title('Time Series data')
%         waitforbuttonpress;
        D1 = LagOp({1,-1},'Lags',[0,1]);
        FirstOrderTransformedTS = filter(D1,TStoBeTranformed);

%         figure
%         plot(2:N,FirstOrderTransformedTS)
%         xlim([0,N])
%         title('First Differenced Time Series data')
%         waitforbuttonpress;
%         prompt={'Please Enter 1 if you satisfied with 1st order differencing, 0 Otherwise:'};
%         % Create all your text fields with the questions specified by the variable prompt.
%         titleOfPromt='Transformation / Preprocessing Techniques Verification'; 
%         % The main title of your input dialog interface.
%         answer=inputdlg(prompt,titleOfPromt);
%         Choice = str2double(answer{1}); 
%         if Choice == 1
           
                 FinalTransformedTS = FirstOrderTransformedTS;
%                       %------------Code for Transformation Technique no -------------
%                         for checkTrans = 1:10
%                             if( TransformedTechNos(checkTrans,:) == 0)
%                                 TransformedTechNos(checkTrans,:)= SelectedTech;
%                                  break
%                             end 
%                         end
%                         %-------------------End Transformation Technique no------------
%             else 
%             FinalTransformedTS = OriginalTS;
%         end
%     
%         
%         % -------------end case 8---------------------------------------
                
                 
   
    case 9
        % -------------Start case 8---------------------------------------
        N = length(TStoBeTranformed);

%         figure
%         plot(TStoBeTranformed)
%         xlim([0,N])
%         title('Time Series data')
%         waitforbuttonpress;
        D1 = LagOp({1,-1},'Lags',[0,1]);
        FirstOrderTransformedTS = filter(D1,TStoBeTranformed);

%         figure
%         plot(2:N,FirstOrderTransformedTS)
%         xlim([0,N])
%         title('First Differenced Time Series data')
%         waitforbuttonpress;
%         prompt={'Please Enter 1 if you satisfied with 1st order differencing, 0 Otherwise:'};
%         % Create all your text fields with the questions specified by the variable prompt.
%         titleOfPromt='Transformation / Preprocessing Techniques Verification'; 
%         % The main title of your input dialog interface.
%         answer=inputdlg(prompt,titleOfPromt);
%         Choice = str2double(answer{1}); 
%         if Choice == 1
%             prompt={'Please Enter 1 if want to apply 2nd order differencing, 0 Otherwise:'};
%             % Create all your text fields with the questions specified by the variable prompt.
%             titleOfPromt='Transformation / Preprocessing Techniques Verification'; 
%             % The main title of your input dialog interface.
%             answer=inputdlg(prompt,titleOfPromt);
%             Choice = str2double(answer{1}); 
%             if Choice == 1
                D2 = D1*D1;
                SecondOrderTransformedTS = filter(D2,FirstOrderTransformedTS);

%                 figure
%                 plot(3:N-1,SecondOrderTransformedTS)
%                 xlim([0,N])
%                 title('Second Differenced Time Series data')
%                         prompt={'Please Enter 1 if you satisfied with 2nd order differencing, 0 Otherwise:'};
%                         % Create all your text fields with the questions specified by the variable prompt.
%                         titleOfPromt='Transformation / Preprocessing Techniques Verification'; 
%                         % The main title of your input dialog interface.
%                         answer=inputdlg(prompt,titleOfPromt);
%                         Choice = str2double(answer{1}); 
%                 if Choice == 1
                            FinalTransformedTS = SecondOrderTransformedTS;
%             %------------Code for Transformation Technique no -------------
%             for checkTrans = 1:10
%                 if( TransformedTechNos(checkTrans,:) == 0)
%                     TransformedTechNos(checkTrans,:)= SelectedTech;
%                      break
%                 end 
%             end
%             %-------------------End Transformation Technique no------------
% %                         else 
% %                             FinalTransformedTS = FirstOrderTransformedTS;
% %                         end
% %                 
% % 
% %             else 
% %                 FinalTransformedTS = FirstOrderTransformedTS;
% %             end
%                 else 
%                     FinalTransformedTS = OriginalTS;
%                 end
%         %----- start 2nd Differenced code-----------
        
        % -------------end case 9---------------------------------------
    case 10
        minTS = min(TStoBeTranformed);
        if minTS <= 0

             ConstADD = abs(minTS);
             N = length(TStoBeTranformed);
             PositiveTransformedTS = zeros(N:1);
                for i=1:N
                    PositiveTransformedTS(i,:) = (TStoBeTranformed(i,:)+ConstADD);
                end
           
             
             if Lembda == 0
                  TransformedTS = log(PositiveTransformedTS);
                 
                        FinalTransformedTS = TransformedTS;
                       
             else
                    N = length(TStoBeTranformed);
                    TransformedTS = zeros(N:1);
                    for i=1:N
                        TransformedTS(i,:) = (((PositiveTransformedTS(i,:))^Lembda)-1)/Lembda;
                    end
                  FinalTransformedTS = TransformedTS; 
             end
                
                
             
        else
            
             
             if Lembda == 0
                  TransformedTS = log(TStoBeTranformed);
                  FinalTransformedTS = TransformedTS;
                       
                      
             else
%                    
                    TransformedTS= ((TStoBeTranformed.^Lembda)-1)./Lembda;
                        FinalTransformedTS = TransformedTS;
                       
                 
             end
             
           
        end
        
        
    case 11
        FinalTransformedTS = OriginalTS;
%         TransformedTechNos = zeros(10,1);
    case 12
         FinalTransformedTS = OriginalTS;
        fprintf('no action');
    otherwise
       warning('Unexpected Selection.');
end
% prompt={'Enter 1 to apply any other Transformation Technique, 0 Otherwise:'};
% % Create all your text fields with the questions specified by the variable prompt.
% titleofPrompt='Transformation Techniques Request'; 
% % The main title of your input dialog interface.
% answer=inputdlg(prompt,titleofPrompt);
% TransREQ = str2double(answer{1});



% end % end while loop
    


end
