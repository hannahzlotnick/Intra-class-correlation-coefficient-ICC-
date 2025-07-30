%%Name: Zlotnick ICC scoring
%%Date: 5/5/2025

function icc_values = calculate_icc_from_selected_range(filename)
    % Get sheet names from the Excel file
    [~, sheetNames] = xlsfinfo(filename);
    
    % Initialize an array to hold ICC values
    icc_values = zeros(length(sheetNames), 1);
    
    for i = 1:length(sheetNames)
        % Read data from the specified range of the current sheet
        data = readtable(filename, 'Sheet', sheetNames{i}, 'Range', 'B2:E53');
        
        % Convert the table to an array (ensure no non-numeric columns)
        data_matrix = table2array(data);
        
        % Calculate ICC for the current sheet
        icc_values(i) = calculate_icc(data_matrix);
        
        % Display the ICC result for the current sheet
        disp(['ICC for sheet "', sheetNames{i}, '": ', num2str(icc_values(i))]);
    end
end

function icc = calculate_icc(data)
    % Calculate ICC (2,1)
    [N, k] = size(data);
    
    % Grand mean
    grand_mean = mean(data(:));
    
    % Mean for each subject
    subject_means = mean(data, 2);
    
    % Mean for each rater
    rater_means = mean(data, 1);
    
    % Total sum of squares
    SStotal = sum((data(:) - grand_mean).^2);
    
    % Between-groups sum of squares
    SSbetween = N * sum((subject_means - grand_mean).^2);
    
    % Within-groups sum of squares
    SSwithin = sum((data - subject_means).^2, 2);
    SSwithin = sum(SSwithin);
    
    % Calculate ICC(2,1)
    icc = (SSbetween - (k - 1) * SSwithin) / (SStotal + (k - 1) * SSwithin);
end

% Example usage:
filename = 'Overall Scores.xlsx'; % Replace with your Excel file name
icc_values = calculate_icc_from_selected_range(filename);