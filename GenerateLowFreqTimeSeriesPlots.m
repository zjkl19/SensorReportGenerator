%% GenerateLowFreqTimeSeriesPlots.m
% 此脚本读取低频 Excel 数据（StaticData/LowFreq.xlsx），将 "-" 视为缺失（NaN），
% 并按列区间分组绘制时程曲线（温湿度、桥面风荷载、主梁挠度、裂缝监测、结构温度）。
% 静态数据放在 StaticData 文件夹，输出结果放在 LowFreqResults 文件夹，
% 生成 PNG 和 FIG 两种格式的图像。

clc; clear;

%% 1. 定义文件夹
staticDataFolder = 'StaticData';      % 静态（低频）数据文件夹
outputFolder     = 'LowFreqResults';  % 绘图结果输出文件夹

% 如果不存在则创建
if ~exist(staticDataFolder, 'dir')
    mkdir(staticDataFolder);
end
if ~exist(outputFolder, 'dir')
    mkdir(outputFolder);
end

%% 2. 导入数据：将 "-" 当作缺失，将所有列先读成字符，再把时间列读为 datetime
filename = fullfile(staticDataFolder, '静态数据（3.1-3.10） - 改.xlsx');
opts = detectImportOptions(filename);
opts = setvaropts(opts, opts.VariableNames, 'TreatAsMissing', {'-'});  
opts = setvaropts(opts, opts.VariableNames, 'Type', 'char');

% 第1列为时间，格式 "2025/3/1 0:00" 等
timeVar = opts.VariableNames{1};
opts = setvaropts(opts, timeVar, 'Type','datetime', 'InputFormat','yyyy/M/d H:mm');

T = readtable(filename, opts);

%% 3. 提取时间向量
time = T.(timeVar);

%% 4. 定义分组：列索引及组名
groups = {
    2:3,   '温湿度';
    4:5,   '桥面风荷载';
    6:23,  '主梁挠度';
    24:40, '裂缝监测';
    41:54, '结构温度'
};

%% 5. 循环分组，转换数据并绘图
for i = 1:size(groups,1)
    cols  = groups{i,1};
    gname = groups{i,2};
    vars  = T.Properties.VariableNames(cols);
    
    n = height(T);
    m = numel(cols);
    Y = nan(n, m);
    
    % 将每列字符或缺失转换为 double
    for j = 1:m
        colStr = string(T.(vars{j}));       % 转为 string 数组
        Y(:,j) = str2double(colStr);        % "0.158"->0.158, ""->NaN
    end
    
    % 绘制时程曲线
    hFig = figure('Visible','off');
    plot(time, Y, 'LineWidth',1.5);
    xlabel('时间');
    ylabel(gname);
    title([gname ' 时程曲线'], 'Interpreter','none');
    datetick('x','yyyy-mm-dd HH:MM','keepticks');
    legend(vars, 'Interpreter','none', 'Location','best');
    grid on;
    
    % 保存为 PNG
    pngFile = fullfile(outputFolder, sprintf('%s_TimeSeries.png', gname));
    saveas(hFig, pngFile);
    % 保存为 FIG
    figFile = fullfile(outputFolder, sprintf('%s_TimeSeries.fig', gname));
    saveas(hFig, figFile);
    
    close(hFig);
    fprintf('已生成：%s and %s\n', pngFile, figFile);
end
