%% GenerateSensorTimeSeriesWordReport.m
% 此脚本对指定传感器通道的时程数据进行分析
% 采样数据存放于 Data 文件夹下的 .mat 文件中，每个文件应包含变量 Datas
% 脚本将直接绘制时程图（数据随样本点变化图），并利用COM接口将文件信息及图像写入Word报告中
% 仅处理映射表中定义的传感器
% 请确保系统为Windows且安装有Microsoft Word

clc; clear;

%% 定义存放.mat文件的文件夹路径
dataFolder = 'Data';

%% 定义传感器映射表（传感器编号及对应备注），参照频谱版映射
mapping = {
    'AI3-01', 'ZDCQG-01-K15-X1/G18';
    'AI3-02', 'ZDCQG-03-K16-X1/G19';
    'AI3-03', 'ZDCQG-05-K16-X1/G20';
    'AI3-04', 'ZDCQG-07-K16-X1/G21';
    'AI3-05', 'ZDCQG-09-K17-X1/G22';
    'AI3-06', 'DZY-01-D15-P15-01';
    'AI3-07', 'DZY-01-D15-P15-02';
    'AI3-08', 'DZY-01-D15-P15-03';
    'AI4-01', 'DZY-02-D16-P16-01';
    'AI4-02', 'DZY-02-D16-P16-02';
    'AI4-03', 'DZY-02-D16-P16-03';
    'AI7-01', 'ZDCQG-02-K15-X4/G18';
    'AI7-02', 'ZDCQG-04-K16-X4/G19';
    'AI7-03', 'ZDCQG-06-K16-X4/G20';
    'AI7-04', 'ZDCQG-08-K16-X4/G21';
    'AI7-05', 'ZDCQG-10-K17-X4/G22';
    'AI7-06', 'SLCGQ-06-K16-ZFFDG2';
    'AI7-07', 'SLCGQ-07-K16-ZFFDG3';
    'AI7-08', 'SLCGQ-08-K16-ZFFDG4';
    'AI8-01', 'SLCGQ-09-K16-ZFFDG5';
    'AI8-02', 'SLCGQ-10-K16-ZFFDG6';
    'AI8-05', 'SLCGQ-11-K16-YFFDG2';
    'AI8-06', 'SLCGQ-12-K16-YFFDG3';
    'AI8-07', 'SLCGQ-13-K16-YFFDG4';
    'AI8-08', 'SLCGQ-14-K16-YFFDG5';
    'AI9-01', 'SLCGQ-01-K16-ZDG1';
    'AI9-02', 'SLCGQ-02-K16-ZDG5';
    'AI9-03', 'SLCGQ-03-K16-ZDG8';
    'AI9-04', 'SLCGQ-04-K16-ZDG11';
    'AI9-05', 'ZDCQG-13-K16-ZGD/A20';
    'AI9-06', 'ZDCQG-14-K16-ZGD/A20';
    'AI9-07', 'ZDCQG-15-K16-YFGD/A20';
    'AI9-08', 'ZDCQG-16-K16-YFGD/A20';
    'AI16-01', 'SLCGQ-05-K16-ZDG15';
    'AI16-02', 'SLCGQ-15-K16-YFFDG6';
    'AI16-03', 'ZDCQG-11-K16-ZFGD/A20';
    'AI16-04', 'ZDCQG-12-K16-ZFGD/A20'
};

%% 获取Data文件夹下所有.mat文件
matFiles = dir(fullfile(dataFolder, '*.mat'));

%% 构建文件信息结构体，提取传感器编号及用于排序的数字信息
fileInfo = struct('name', {}, 'sensorID', {}, 'X', {}, 'Y', {}, 'valid', {});
for k = 1:length(matFiles)
    fname = matFiles(k).name;
    fileInfo(k).name = fname;
    % 利用正则表达式从文件名中提取传感器编号（例如 "AI3-01_"）
    token = regexp(fname, '^(AI\d{1,2}-\d{2})_', 'tokens');
    if ~isempty(token)
        sensorID = token{1}{1};
        fileInfo(k).sensorID = sensorID;
        % 仅处理映射表中存在的传感器
        if any(strcmp(mapping(:,1), sensorID))
            % 提取传感器编号中的数字，用于排序（例如 "AI3-01" 提取出3和01）
            token2 = regexp(sensorID, 'AI(\d+)-(\d+)', 'tokens');
            if ~isempty(token2)
                fileInfo(k).X = str2double(token2{1}{1});
                fileInfo(k).Y = str2double(token2{1}{2});
                fileInfo(k).valid = true;
            else
                fileInfo(k).sensorID = 'Unknown';
                fileInfo(k).X = Inf;
                fileInfo(k).Y = Inf;
                fileInfo(k).valid = false;
            end
        else
            fileInfo(k).valid = false;
        end
    else
        fileInfo(k).sensorID = 'Unknown';
        fileInfo(k).X = Inf;
        fileInfo(k).Y = Inf;
        fileInfo(k).valid = false;
    end
end

%% 将结构体转换为表格，并过滤掉无效条目（即映射表中不存在的传感器）
T = struct2table(fileInfo);
T = T(T.valid == true, :);
% 按传感器编号中的数字部分进行升序排序
T = sortrows(T, {'X','Y'}, {'ascend','ascend'});

%% 使用COM接口创建Word文档
wordApp = actxserver('Word.Application');
wordApp.Visible = true;  % 如需隐藏Word窗口，可设置为false
doc = wordApp.Documents.Add();

%% 循环处理每个文件（按排序后顺序），绘制时程图并写入Word报告
for k = 1:height(T)
    matFileName = T.name{k};
    sensorID = T.sensorID{k};
    
    % 构造完整的.mat文件路径
    fullFilePath = fullfile(dataFolder, matFileName);
    
    % 加载.mat文件，检查是否包含变量 Datas
    S = load(fullFilePath, 'Datas');
    if ~isfield(S, 'Datas')
        warning(['文件 ' fullFilePath ' 不包含变量 ''Datas''，跳过。']);
        continue;
    end
    data = S.Datas;
    if ~isnumeric(data) || ~(isvector(data) && size(data,2)==1)
        warning(['文件 ' fullFilePath ' 中变量 ''Datas'' 不是数值列向量，跳过。']);
        continue;
    end
    
    % 获取传感器备注信息
    idx = strcmp(mapping(:,1), sensorID);
    if any(idx)
        sensorRemark = mapping{idx,2};
    else
        sensorRemark = '无测点编号';
    end
    
    %% 绘制时程图（数据随样本点变化）
    hFig = figure('Visible', 'off');
    plot(data, 'LineWidth', 1.5);
    xlabel('样本点');
    ylabel('数值');
    title(['传感器: ' sensorID ', 备注: ' sensorRemark], 'Interpreter', 'none');
    
    %% 保存图像为临时PNG文件
    tempImage = fullfile(tempdir, [sensorID '_' datestr(now, 'yyyymmdd_HHMMSS') '.png']);
    saveas(hFig, tempImage);
    close(hFig);
    
    %% 将文件信息和图像插入到Word文档中
    selection = wordApp.Selection;
    selection.TypeText(['文件: ' matFileName '   传感器: ' sensorID '   备注: ' sensorRemark]);
    selection.TypeParagraph;
    selection.InlineShapes.AddPicture(tempImage);
    selection.TypeParagraph;
end

%% 将Word文档保存到当前文件夹
outputDoc = fullfile(pwd, 'SensorDataReport.docx');
doc.SaveAs2(outputDoc);
doc.Close;
wordApp.Quit;
delete(wordApp);

disp(['Word文档已保存: ' outputDoc]);
