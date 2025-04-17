%% GenerateLowFreqTimeSeriesPlots.m
% 此脚本：
% 1）读取静态低频 Excel 数据（StaticData/LowFreq.xlsx），将 "-" 视为缺失（NaN）；
% 2）按分组（温湿度、桥面风荷载、主梁挠度、裂缝监测、结构温度）绘制时程曲线，
%    并保存为 PNG 和 FIG；
% 3）将所有 PNG 图插入到一个 Word 报告（LowFreqResults/LowFreqTimeSeriesReport.docx）中。

clc; clear;

%% 1. 定义文件夹
staticDataFolder = 'StaticData';      % 静态（低频）数据文件夹
outputFolder     = 'LowFreqResults';  % 绘图结果输出文件夹

if ~exist(staticDataFolder, 'dir')
    mkdir(staticDataFolder);
end
if ~exist(outputFolder, 'dir')
    mkdir(outputFolder);
end

%% 2. 读取 Excel 数据，Treat "-" 为缺失，并保留原始列名
filename = fullfile(staticDataFolder, '静态数据（3.1-3.10） - 改.xlsx');
opts = detectImportOptions(filename);
opts = setvaropts(opts, opts.VariableNames, 'TreatAsMissing', {'-'});
opts = setvaropts(opts, opts.VariableNames, 'Type', 'char');
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

%% 5. 启动 Word COM，准备报告文档
wordApp = actxserver('Word.Application');
wordApp.Visible = true;         % 如需隐藏界面可设 false
doc = wordApp.Documents.Add();

%% 6. 循环分组，生成图像并插入 Word
for i = 1:size(groups,1)
    cols  = groups{i,1};
    gname = groups{i,2};
    vars  = T.Properties.VariableNames(cols);
    
    % 构造数据矩阵（字符→double，"-"/""→NaN）
    n = height(T); m = numel(cols);
    Y = nan(n, m);
    for j = 1:m
        colStr = string(T.(vars{j}));
        Y(:,j)  = str2double(colStr);
    end
    
    % 绘制时程曲线
    hFig = figure('Visible','off');
    plot(time, Y, 'LineWidth',1.5);
    xlabel('时间'); ylabel(gname);
    title([gname ' 时程曲线'], 'Interpreter','none');
    datetick('x','yyyy-mm-dd HH:MM','keepticks');
    legend(vars, 'Interpreter','none', 'Location','best');
    grid on;
    
    % 保存 PNG 和 FIG
    pngFile = fullfile(outputFolder, sprintf('%s_TimeSeries.png', gname));
    figFile = fullfile(outputFolder, sprintf('%s_TimeSeries.fig', gname));
    saveas(hFig, pngFile);
    saveas(hFig, figFile);
    close(hFig);
    fprintf('已生成：%s 和 %s\n', pngFile, figFile);
    
    % 插入到 Word 文档（绝对路径、LinkToFile=0、SaveWithDocument=1）
    sel = wordApp.Selection;
    sel.TypeText([gname ' 时程曲线']);
    sel.TypeParagraph;
    imgFullPath = fullfile(pwd, pngFile);
    sel.InlineShapes.AddPicture(imgFullPath, 0, 1);
    sel.TypeParagraph;
end

%% 7. 保存并关闭 Word 文档（仅用 SaveAs2，并指定文档格式16）
reportFile = fullfile(pwd, outputFolder, 'LowFreqTimeSeriesReport.docx');
% wdFormatDocumentDefault = 16
doc.SaveAs2(reportFile);
doc.Close;
wordApp.Quit;
delete(wordApp);

fprintf('Word 报告已保存至：%s\n', reportFile);
