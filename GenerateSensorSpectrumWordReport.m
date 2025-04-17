%% GenerateSensorSpectrumWordReport.m
% 此脚本对指定传感器通道进行频谱分析（采样频率50Hz）
% 读取 Data 文件夹下的 .mat 文件，变量名为 Datas
% 频谱分析结果（文件名、传感器编号、备注及频谱图）写入 Word 报告
% 传感器–备注映射存放在同级目录下的 sensor_mapping.csv

clc; clear;

%% 定义数据存放文件夹
dataFolder = 'Data';

%% 从外部 CSV 加载传感器映射表
% CSV 文件必须含列名：sensorID,remark
mapTable = readtable('sensor_mapping.csv', 'TextType','string');

%% 定义采样频率
Fs = 50;  % Hz

%% 获取 Data 文件夹下所有 .mat 文件
matFiles = dir(fullfile(dataFolder, '*.mat'));

%% 构建文件列表：提取传感器编号及排序用数字
fileInfo = struct('name', {}, 'sensorID', {}, 'X', {}, 'Y', {}, 'valid', {});
for k = 1:length(matFiles)
    fname = matFiles(k).name;
    % 正则提取传感器编号（如 "AI3-01"）
    tok = regexp(fname, '^(AI\d{1,2}-\d{2})_', 'tokens');
    if ~isempty(tok) && any(mapTable.sensorID == tok{1}{1})
        sid = tok{1}{1};
        % 拆分数字部分用于排序
        numtok = regexp(sid, 'AI(\d+)-(\d+)', 'tokens');
        fileInfo(k).name     = fname;
        fileInfo(k).sensorID = sid;
        fileInfo(k).X        = str2double(numtok{1}{1});
        fileInfo(k).Y        = str2double(numtok{1}{2});
        fileInfo(k).valid    = true;
    else
        % 不在映射表中则标记为无效
        fileInfo(k).name     = fname;
        fileInfo(k).sensorID = '';
        fileInfo(k).X        = Inf;
        fileInfo(k).Y        = Inf;
        fileInfo(k).valid    = false;
    end
end

%% 转换为表并过滤、排序
T = struct2table(fileInfo);
T = T(T.valid, :);  % 仅保留有效项
T = sortrows(T, {'X','Y'}, {'ascend','ascend'});

%% 创建 Word 文档
wordApp = actxserver('Word.Application');
wordApp.Visible = true;         % 如需隐藏界面可设 false
doc     = wordApp.Documents.Add();

%% 循环处理每个传感器文件
for k = 1:height(T)
    fname    = T.name{k};
    sensorID = T.sensorID{k};
    fullpath = fullfile(dataFolder, fname);
    
    % 加载 Datas
    S = load(fullpath, 'Datas');
    if ~isfield(S,'Datas')
        warning('%s 中无 Datas，跳过。', fname);
        continue;
    end
    data = S.Datas;
    if ~isnumeric(data) || size(data,2)~=1
        warning('%s Datas 不是列向量，跳过。', fname);
        continue;
    end
    
    % 查备注
    idx    = find(mapTable.sensorID == sensorID);
    remark = mapTable.remark(idx);
    
    %% FFT 频谱分析
    L  = length(data);
    Y  = fft(data);
    P2 = abs(Y/L);
    P1 = P2(1:floor(L/2)+1);
    if numel(P1)>1
        P1(2:end-1) = 2*P1(2:end-1);
    end
    f = Fs*(0:floor(L/2))/L;
    
    %% 绘制频谱图
    hFig = figure('Visible','off');
    plot(f, P1, 'LineWidth', 1.5);
    xlabel('频率 (Hz)');
    ylabel('幅值');
    title(sprintf('传感器: %s   备注: %s', sensorID, remark), 'Interpreter','none');
    
    %% 保存临时图片
    imgFile = fullfile(tempdir, sprintf('%s_%s.png', sensorID, datestr(now,'yyyymmdd_HHMMSS')));
    saveas(hFig, imgFile);
    close(hFig);
    
    %% 插入 Word
    sel = wordApp.Selection;
    sel.TypeText(sprintf('文件: %s   传感器: %s   备注: %s', fname, sensorID, remark));
    sel.TypeParagraph;
    sel.InlineShapes.AddPicture(imgFile);
    sel.TypeParagraph;
end

%% 保存并关闭
outDoc = fullfile(pwd, 'SensorSpectrumReport.docx');
doc.SaveAs2(outDoc);
doc.Close;
wordApp.Quit;
delete(wordApp);

fprintf('频谱报告已保存至: %s\n', outDoc);
