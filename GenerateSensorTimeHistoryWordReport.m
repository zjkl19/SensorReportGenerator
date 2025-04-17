%% GenerateSensorTimeSeriesWordReport.m
% 将传感器映射放到外部 CSV 文件：sensor_mapping.csv
% （其它说明同前）

clc; clear;

%% 定义存放.mat文件的文件夹路径
dataFolder = 'Data';

%% 从外部 CSV 加载传感器映射表
% CSV 文件第一列为 sensorID，第二列为 remark
mapTable = readtable('sensor_mapping.csv', 'TextType','string');

%% 获取 Data 文件夹下所有 .mat 文件
matFiles = dir(fullfile(dataFolder, '*.mat'));

%% 构建文件信息结构体，提取传感器编号及用于排序的数字信息
fileInfo = struct('name', {}, 'sensorID', {}, 'X', {}, 'Y', {}, 'valid', {});
for k = 1:length(matFiles)
    fname = matFiles(k).name;
    fileInfo(k).name = fname;
    % 提取传感器编号（如 "AI3-01"）
    tok = regexp(fname, '^(AI\d{1,2}-\d{2})_', 'tokens');
    if ~isempty(tok) && any(mapTable.sensorID==tok{1}{1})
        sid = tok{1}{1};
        fileInfo(k).sensorID = sid;
        % 拆数字用于排序
        tok2 = regexp(sid, 'AI(\d+)-(\d+)', 'tokens');
        fileInfo(k).X = str2double(tok2{1}{1});
        fileInfo(k).Y = str2double(tok2{1}{2});
        fileInfo(k).valid = true;
    else
        fileInfo(k).sensorID = 'Unknown';
        fileInfo(k).X = Inf;
        fileInfo(k).Y = Inf;
        fileInfo(k).valid = false;
    end
end

%% 转为表格并过滤与排序
T = struct2table(fileInfo);
T = T(T.valid, :);
T = sortrows(T, {'X','Y'}, {'ascend','ascend'});

%% 创建 Word 文档
wordApp = actxserver('Word.Application');
wordApp.Visible = true;
doc = wordApp.Documents.Add();

%% 循环处理每个通道，绘制时程并写入报告
for k = 1:height(T)
    fname     = T.name{k};
    sensorID  = T.sensorID{k};
    fullpath  = fullfile(dataFolder, fname);
    S         = load(fullpath, 'Datas');
    if ~isfield(S,'Datas'), warning([fname ' 没有 Datas，跳过']); continue; end
    data      = S.Datas;
    if ~isnumeric(data) || size(data,2)>1
        warning([fname ' Datas 不是列向量，跳过']); continue;
    end

    % 从映射表中取备注
    idx       = find(mapTable.sensorID==sensorID);
    remark    = mapTable.remark(idx);

    % 绘制时程图
    h = figure('Visible','off');
    plot(data,'LineWidth',1.5);
    xlabel('样本点'); ylabel('数值');
    title(['传感器: ' sensorID ' 备注: ' remark],'Interpreter','none');

    % 临时保存图片
    imgFile = fullfile(tempdir, sensorID + "_" + datestr(now,'yyyymmdd_HHMMSS') + ".png");
    saveas(h, imgFile);
    close(h);

    % 插入到 Word
    sel = wordApp.Selection;
    sel.TypeText(sprintf('文件: %s   传感器: %s   备注: %s', fname, sensorID, remark));
    sel.TypeParagraph;
    sel.InlineShapes.AddPicture(imgFile);
    sel.TypeParagraph;
end

%% 保存并关闭
outDoc = fullfile(pwd, 'SensorDataReport.docx');
doc.SaveAs2(outDoc);
doc.Close;
wordApp.Quit;
delete(wordApp);

disp(['报告已保存: ' outDoc]);
