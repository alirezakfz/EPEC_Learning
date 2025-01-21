% Clear workspace, command window, and figures% Clear workspace, command window, and figures
clear; clc; close all;

% File and data setup
filename = 'All Results.xlsx';
sheet = 'Sheet1';
ranges = containers.Map;
ranges('diag') = 'N36:N59';
ranges('FP') = 'T36:T59';
ranges('SFP') = 'Z36:Z59';
ranges('comp') = 'AF36:AF59';

diag_prices = readmatrix(filename, 'Sheet', sheet, 'Range', ranges('diag'));
FP_prices = readmatrix(filename, 'Sheet', sheet, 'Range', ranges('FP'));
SFP_prices = readmatrix(filename, 'Sheet', sheet, 'Range', ranges('SFP'));
comp_prices = readmatrix(filename, 'Sheet', sheet, 'Range', ranges('comp'));

diag_bids = 30*[readmatrix(filename,"Sheet",sheet,'Range',"C7:E30"),readmatrix(filename,"Sheet",sheet,'Range','B7:B30')];
FP_bids = 30*[readmatrix(filename,"Sheet",sheet,'Range',"G7:I30"),readmatrix(filename,"Sheet",sheet,'Range',"F7:F30")];
SFP_bids = 30*[readmatrix(filename,"Sheet",sheet,'Range',"K7:M30"),readmatrix(filename,"Sheet",sheet,'Range',"J7:J30")];
comp_bids = [readmatrix(filename,"Sheet",sheet,'Range',"B77:E100")];

%% Plots
% Variables for plots
t = 1:24;  % Time intervals (MTU)
euro = char(8364);  % Euro symbol
data = {diag_prices, FP_prices, SFP_prices};  % Grouped data
titles = {'Competitive vs. Diagonalization', 'Competitive vs. Fictitious Play', 'Competitive vs. Smooth Fictitious Play'};
legend_labels = {'Competitive', 'Diagonalization'; 'Competitive', 'Fictitious Play'; 'Competitive', 'Smooth Fictitious Play'};
colors = [0 0.4470 0.7410; 0.8500 0.3250 0.0980; 0.9290 0.6940 0.1250; 0.4660 0.6740 0.1880]; % Competitive + 3 methods

% Plot setup
figure(1);
set(gcf, 'Position', [100, 100, 1200, 400]); % Increase figure size for better visibility
for i = 1:3
    subplot(1, 3, i);
    % Set bar colors using the 'FaceColor' property
    b = bar(t, [comp_prices, data{i}], 'grouped');
    b(1).FaceColor = colors(1, :); % Competitive color
    b(2).FaceColor = colors(i+1, :); % Unique color for each method (Diagonalization, FP, SFP)
    
    xlabel('MTU (h)', 'FontSize', 14,'FontWeight','bold');
    if i == 1  % Add ylabel only to the first subplot
        ylabel({'SMP', ['(', euro, '/MWh)']}, 'FontSize', 14,'FontWeight','bold');
    end
    xticks(1:24);
    lgd = legend(legend_labels{i, :});
    lgd.FontSize = 12;
    lgd.Box = 'off';
    ylim([0, 120]);
    grid on;
    title(titles{i});  % Add a title for each subplot
end

%% 
figure(2)
for DA=1:4
    subplot(4,1,DA)

    b = bar(t, [comp_bids(:,DA),diag_bids(:,DA)], 'grouped');
    xticks(1:24);
    ylabel({'Demand Bid','(MW)'},'FontSize', 10,'FontWeight','bold');
    if DA==4
        xlabel('MTU (h)', 'FontSize', 14,'FontWeight','bold');
        lgd = legend('Competitive','Diagonalization');
        lgd.FontSize = 10;
        lgd.Box = 'off';
    end    
    ylim([0, 120]);
    grid on;
end

%% 
figure(3)
for DA=1:4
    subplot(4,1,DA)

    b = bar(t, [comp_bids(:,DA),FP_bids(:,DA)], 'grouped');
    xticks(1:24);
    ylabel({'Demand Bid','(MW)'},'FontSize', 10,'FontWeight','bold');
    if DA==4
        xlabel('MTU (h)', 'FontSize', 14,'FontWeight','bold');
        lgd = legend('Competitive','Fictitious Play');
        lgd.FontSize = 10;
        lgd.Box = 'off';
    end    
    ylim([0, 120]);
    grid on;
end

%% 
figure(4)
for DA=1:4
    subplot(4,1,DA)

    b = bar(t, [comp_bids(:,DA),SFP_bids(:,DA)], 'grouped');
    xticks(1:24);
    ylabel({'Demand Bid','(MW)'},'FontSize', 10,'FontWeight','bold');
    if DA==4
        xlabel('MTU (h)', 'FontSize', 14,'FontWeight','bold');
        lgd = legend('Competitive','Smooth Fictitious Play');
        lgd.FontSize = 10;
        lgd.Box = 'off';
    end    
    ylim([0, 120]);
    grid on;
end