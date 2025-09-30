function fbx_generate_matrix_and_plots(srcXlsx, srcSheet, outDbXlsx, outMatrixXlsx, outPlotDir, step, mode, varargin)
% fbx_generate_matrix_and_plots
% -------------------------------------------------------------------------
% 一键式管线：从 All_tidy 读取二元 U0..U3，构建规范化参数库，生成三元
% 混合焓单 Sheet（矩阵格式），并导出所有 Fe–B–X PNG 图片。
%
% 步骤：
%   1) 调用 fbx_all_in_one_matrix(...) 生成/覆盖：
%        - outDbXlsx（Pairs_Used + README）
%        - outMatrixXlsx（Sheet = FBX_MATRIX_<MODE>，包含 c_Fe/c_B/c_X 及所有 Hmix_X）
%   2) 调用 fbx_export_all_single(...)，从步骤 1 生成的矩阵 Sheet 中识别所有
%        Hmix_* 列，为每个 X 输出三元图 PNG（命名：FBX_<X>.png）。
%
% 用法：
%   fbx_generate_matrix_and_plots                             % 全默认
%   fbx_generate_matrix_and_plots(infile,'All_tidy',dbOut,matrixOut,plotDir,0.01,'pair','CLim',[-30 5]);
%
% 说明：
%   - Name-Value 形式的附加参数（varargin）会原样传递给 fbx_export_all_single，
%     可用于调整配色、等值线、刻度等。
%   - mode 支持 'pair'（默认）与 'global'，与 fbx_all_in_one_matrix 一致。
%
% 作者：ChatGPT

    if nargin < 1 || isempty(srcXlsx)
        srcXlsx = 'C:\\Fe_BMAT\\Fe_BM\\Fe-B-X.xlsx';
    end
    if nargin < 2 || isempty(srcSheet)
        srcSheet = 'All_tidy';
    end
    if nargin < 3 || isempty(outDbXlsx)
        outDbXlsx = 'C:\\Fe_BMAT\\Fe_BM\\Hmix_FB_X_ternary.xlsx';
    end
    if nargin < 4 || isempty(outMatrixXlsx)
        outMatrixXlsx = 'C:\\Fe_BMAT\\Fe_BM\\Hmix_FB_X_matrix.xlsx';
    end
    if nargin < 5 || isempty(outPlotDir)
        outPlotDir = 'C:\\Fe_BMAT\\Fe_BM\\plots\\FBX_all';
    end
    if nargin < 6 || isempty(step)
        step = 0.01;
    end
    if nargin < 7 || isempty(mode)
        mode = 'pair';
    end

    mode = lower(string(mode));
    if mode ~= "pair" && mode ~= "global"
        error('mode 仅支持 ''pair'' 或 ''global''。');
    end

    fprintf('=== STEP 1/2 ===\n');
    fbx_all_in_one_matrix(srcXlsx, srcSheet, outDbXlsx, outMatrixXlsx, step, mode);

    fprintf('=== STEP 2/2 ===\n');
    sheetName = char("FBX_MATRIX_" + upper(mode));
    if strlength(string(outPlotDir)) == 0
        warning('OutDir 为空，跳过 PNG 导出。');
    else
        fbx_export_all_single('DataXlsx', outMatrixXlsx, ...
                              'Sheet',    sheetName, ...
                              'OutDir',   outPlotDir, ...
                              varargin{:});
    end

    fprintf('全部流程完成。\n');
end
