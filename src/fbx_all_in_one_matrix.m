
function fbx_all_in_one_matrix(srcXlsx, srcSheet, outDbXlsx, outXlsx, step, mode)
% fbx_all_in_one_matrix
% -------------------------------------------------------------------------
% 单文件一键：构建参数库 + 以“单Sheet矩阵格式”输出三元混合焓：
%   列：c_Fe | c_B | c_X | Hmix_<X1> | Hmix_<X2> | ... （不含中文注释）
%   行：网格点（cFe+cB+cX=1），步长默认 0.01（整数格），去除浮点尾数噪声。
%
% 口径：
%   mode='pair'（默认）  : y = c_B/(c_A+c_B)；ΔH_ij = (c_A+c_B)*ΔH^bin(y)，三对相加；
%   mode='global'        : ΔH_ij = 4 c_A c_B Σ U_k (c_A - c_B)^k，三对相加。
%
% 用法：
%   fbx_all_in_one_matrix                                 % 全默认（step=0.01，pair）
%   fbx_all_in_one_matrix(srcXlsx,srcSheet,outDbXlsx,outXlsx,0.01,'pair')
%
if nargin < 1 || isempty(srcXlsx),    srcXlsx    = 'C:\Fe_BMAT\Fe_BM\Fe-B-X.xlsx'; end
if nargin < 2 || isempty(srcSheet),   srcSheet   = 'All_tidy'; end
if nargin < 3 || isempty(outDbXlsx),  outDbXlsx  = 'C:\Fe_BMAT\Fe_BM\Hmix_FB_X_ternary.xlsx'; end
if nargin < 4 || isempty(outXlsx),    outXlsx    = 'C:\Fe_BMAT\Fe_BM\Hmix_FB_X_matrix.xlsx'; end
if nargin < 5 || isempty(step),       step       = 0.01; end
if nargin < 6 || isempty(mode),       mode       = 'pair'; end
mode = lower(string(mode));

fprintf('STEP 1/2: 构建参数库（Pairs_Used）...\n');
[Umap, Zmap] = build_U_from_AllTidy(srcXlsx, srcSheet);
write_PairsUsed_and_README(outDbXlsx, Umap, srcXlsx, srcSheet);

fprintf('STEP 2/2: 生成单Sheet矩阵（step=%.4g, mode=%s）...\n', step, mode);

% —— 自动识别可用 X（同时存在 Fe–X 与 B–X 的元素）——
keys = string(Umap.keys)';  toks = split(keys,"-");  % Nx2: [A B]
Fe_set = strings(0,1); B_set = strings(0,1);
for i=1:size(toks,1)
    a=toks(i,1); b=toks(i,2);
    if a=="Fe", Fe_set(end+1,1)=b; elseif b=="Fe", Fe_set(end+1,1)=a; end %#ok<AGROW>
    if a=="B",  B_set(end+1,1) =b; elseif b=="B",  B_set(end+1,1) =a; end %#ok<AGROW>
end
Fe_set = unique(Fe_set);  B_set = unique(B_set);
Xs = intersect(Fe_set, B_set);
Xs = Xs( Xs~="Fe" & Xs~="B" & strlength(Xs)>0 );
X_list = cellstr(Xs)';
if isempty(X_list), error('未识别到同时具有 Fe–X 与 B–X 的元素。'); end

% —— 构造列名：c_Fe | c_B | c_X | Hmix_<X...> ——
numX = numel(X_list);
varNames = cell(1, 3+numX);
varNames(1:3) = {'c_Fe','c_B','c_X'};
for j=1:numX
    varNames{3+j} = ['Hmix_' char(X_list{j})];
end

% —— 扫描（整数网格 + 去噪）——
N = round(1/step);
rows = (N+1)*(N+2)/2;
tol = 1e-12;  % 去噪阈值

% 预分配矩阵
M = zeros(rows, 3+numX);

% 预取三对 U 的句柄：Fe–B 对所有行一致；Fe–X、B–X 因 X 不同而不同
[U_FeB, okFB, A_FB, B_FB] = getU('Fe','B',Umap,Zmap);
if ~okFB, error('缺少 Fe–B 的 U 参数'); end
U_FeX = cell(numX,1);  U_BX = cell(numX,1);
A_FeX = strings(numX,1); B_FeX = strings(numX,1);
A_BX  = strings(numX,1); B_BX  = strings(numX,1);
for j=1:numX
    Xsym = string(X_list{j});
    [U_FeX{j}, okF, A_FeX(j), B_FeX(j)] = getU('Fe', Xsym, Umap, Zmap);
    [U_BX{j} , okB, A_BX(j) , B_BX(j) ] = getU('B' , Xsym, Umap, Zmap);
    if ~(okF && okB)
        error('缺少 Fe–%s 或 B–%s 的 U 参数。', char(Xsym), char(Xsym));
    end
end

% 填充
r = 0;
for l = 0:N
    for m = 0:(N-l)
        r = r + 1;
        iFe = l; iB = m; iX = N - l - m;   % 整数和恒等 N
        cFe = iFe / N; cB = iB / N; cX = iX / N;

        % 先算 Fe–B，对所有 X 通用
        [cA_FB, cB_FB] = mapAB(A_FB, B_FB, cFe, cB, cX, "X"); % X占位，不用到
        switch mode
            case "pair",   H_FeB = Hpair_pairMode(U_FeB, cA_FB, cB_FB);
            case "global", H_FeB = Hpair_global(   U_FeB, cA_FB, cB_FB);
            otherwise, error('未知模式：%s', char(mode));
        end

        % 写基础三列
        M(r,1) = cFe; M(r,2) = cB; M(r,3) = cX;

        % 逐 X 写 Hmix_X
        for j=1:numX
            [cA_FX, cB_FX] = mapAB(A_FeX(j), B_FeX(j), cFe, cB, cX, X_list{j});
            [cA_BX, cB_BX] = mapAB(A_BX(j) , B_BX(j) , cFe, cB, cX, X_list{j});
            switch mode
                case "pair"
                    H_FeX = Hpair_pairMode(U_FeX{j}, cA_FX, cB_FX);
                    H_BX  = Hpair_pairMode(U_BX{j} , cA_BX, cB_BX);
                case "global"
                    H_FeX = Hpair_global(   U_FeX{j}, cA_FX, cB_FX);
                    H_BX  = Hpair_global(   U_BX{j} , cA_BX, cB_BX);
            end
            M(r, 3+j) = H_FeB + H_FeX + H_BX;
        end
    end
end

% 去噪
M(abs(M) < tol) = 0;

% 写出到一个 Sheet：'FBX_MATRIX_<MODE>'
sheetName = char("FBX_MATRIX_" + upper(mode));
T = array2table(M, 'VariableNames', varNames);
try
    writetable(T, outXlsx, 'Sheet', sheetName, 'WriteMode','overwritesheet');
catch
    writetable(T, outXlsx, 'Sheet', sheetName);
end

fprintf('完成：已写出矩阵到 %s（Sheet=%s），行数=%d，列数=%d\n', outXlsx, sheetName, size(M,1), size(M,2));

% ========================= 子函数：构建 U 映射 =========================
function [Umap, Zmap] = build_U_from_AllTidy(xlsxPath, sheetName)
    T = readtable(xlsxPath, 'Sheet', sheetName, 'PreserveVariableNames', true);
    T = repairHeadersIfNeeded(T, sheetName);
    names = string(T.Properties.VariableNames);
    Acol  = findCol(names, ["a(row)","a_row","a (row)","arow"]);
    Pcol  = findCol(names, ["param","parameter"]);
    assert(~isempty(Acol) && ~isempty(Pcol), '未找到 "A (row)" 或 "Param" 列。');

    Zmap = symbolToZMap();
    Bcols = strings(0,1);
    for i = 1:numel(names)
        nm = string(names(i));
        nmClean = regexprep(nm, '\s+', '');
        if isKey(Zmap, char(nmClean))
            Bcols(end+1,1) = nm; %#ok<AGROW>
        end
    end
    assert(~isempty(Bcols), '未识别到元素列。');

    Umap = containers.Map('KeyType','char','ValueType','any');  % key='A-B'
    for r = 1:height(T)
        A = normalizeSymbol(T.(Acol)(r));
        p0 = normalizeParamToken(T.(Pcol)(r));   % U0->0..U3->3
        if A=="" || isnan(p0), continue; end
        pidx = p0 + 1;  % 1..4
        if ~isKey(Zmap, char(A)), continue; end
        ZA = Zmap(char(A));
        for jc = 1:numel(Bcols)
            Bsym = normalizeSymbol(Bcols(jc));
            if ~isKey(Zmap, char(Bsym)), continue; end
            val = toNumSafe(T.(char(Bcols(jc)))(r));
            if isnan(val), continue; end
            ZB = Zmap(char(Bsym));
            if (ZA < ZB) || (ZA==ZB && strlength(A) <= strlength(Bsym))
                Acanon = A; Bcanon = Bsym; sgn = +1;
            else
                Acanon = Bsym; Bcanon = A; sgn = -1; % 源为 B–A：奇次项取负后存入规范 A–B
            end
            key = char(Acanon + "-" + Bcanon);
            if ~isKey(Umap, key), Umap(key) = [NaN NaN NaN NaN]; end
            Ucur = Umap(key);
            if pidx==2 || pidx==4, Ucur(pidx) = sgn*val; else, Ucur(pidx) = val; end
            Umap(key) = Ucur;
        end
    end
    % 缺失补零
    keys2 = string(Umap.keys)';
    for k=1:numel(keys2)
        U = Umap(char(keys2(k))); U(isnan(U)) = 0; Umap(char(keys2(k))) = U;
    end
end

% ========================= 子函数：写 Pairs_Used 与 README =========================
function write_PairsUsed_and_README(outXlsx, Umap, srcXlsx, srcSheet)
    keys = string(Umap.keys)';
    PL = cell(numel(keys)+1, 5);
    PL(1,:) = {'Pair (A-lowZ – B-highZ)','U0','U1','U2','U3'};
    for k=1:numel(keys)
        U = Umap(char(keys(k)));
        PL{k+1,1} = char(keys(k));
        PL{k+1,2} = U(1); PL{k+1,3} = U(2); PL{k+1,4} = U(3); PL{k+1,5} = U(4);
    end
    try, writecell(PL, outXlsx, 'Sheet', 'Pairs_Used'); catch, xlswrite(outXlsx, PL, 'Pairs_Used'); end

    readme = {
    '字段','说明';
    '输入', srcXlsx;
    '工作表', srcSheet;
    '布局', 'All_tidy（网格：A(row)+Param+元素列）';
    '参数规范', '统一为 A–B（A=低Z、B=高Z），若源为 B–A，读入时奇次项 U1/U3 取负一次';
    '三元（pair）', 'y=c_j/(c_i+c_j)，ΔH_{ij}= (c_i+c_j)*ΔH^{bin}(y)，总和三对';
    '三元（global）', 'ΔH_{ij}=4 c_i c_j Σ U_k (c_i-c_j)^k，三对相加';
    '输出文件', outXlsx;
    };
    try, writecell(readme, outXlsx, 'Sheet', 'README'); catch, xlswrite(outXlsx, readme, 'README'); end
end

% ========================= 子函数：计算 & 映射 =========================
function H = Hpair_pairMode(U,cA,cB)
    w = cA + cB; if w <= 0, H=0; return; end
    yB = cB / w; t = 1 - 2*yB;
    Hbin = 4 .* yB .* (1 - yB) .* ( U(1) + U(2).*t + U(3).*t.^2 + U(4).*t.^3 );
    H = w .* Hbin;
end

function H = Hpair_global(U,cA,cB)
    t = cA - cB;
    H = 4 .* cA .* cB .* ( U(1) + U(2).*t + U(3).*t.^2 + U(4).*t.^3 );
end

function [Ucanon, ok, A, B] = getU(E1, E2, Umap, Zmap)
    ok=false; Ucanon=[]; A=""; B="";
    if ~isKey(Zmap, char(E1)) || ~isKey(Zmap, char(E2)), return; end
    z1 = Zmap(char(E1)); z2 = Zmap(char(E2)); if z1==z2, return; end
    if z1 < z2, key = char(E1 + "-" + E2); A=E1; B=E2;
    else,        key = char(E2 + "-" + E1); A=E2; B=E1; end
    if ~isKey(Umap, key), return; end
    Ucanon = Umap(key); ok=true;
end

function [cA,cB] = mapAB(A,B,cFe,cB_,cX_,Xsym)
    switch char(A)
        case 'Fe', cA = cFe;
        case 'B',  cA = cB_;
        otherwise, if char(A)==char(Xsym), cA = cX_; else, error('未知 A：%s', char(A)); end
    end
    switch char(B)
        case 'Fe', cB = cFe;
        case 'B',  cB = cB_;
        otherwise, if char(B)==char(Xsym), cB = cX_; else, error('未知 B：%s', char(B)); end
    end
end

% ========================= 子函数：工具 =========================
function T = repairHeadersIfNeeded(T, sht)
    vnames = string(T.Properties.VariableNames);
    bad = sum(contains(vnames, ["unnamed","var"], 'IgnoreCase', true));
    if bad >= max(3, round(0.2*numel(vnames)))
        C = table2cell(T);
        hdr = string(C(1,:));
        if any(strlength(hdr)>0)
            hn = matlab.lang.makeUniqueStrings(matlab.lang.makeValidName(cellstr(hdr)));
            T = cell2table(C(2:end,:), 'VariableNames', hn);
            fprintf('Sheet %s: 已修复表头（提升首行为变量名）。\n', string(sht));
        end
    end
end

function col = findCol(names, candidates)
    col = '';
    for i = 1:numel(names)
        nm = normalizeHeader(names(i));
        for j = 1:numel(candidates)
            if nm == normalizeHeader(candidates(j))
                col = char(names(i));
                return;
            end
        end
    end
end

function tok = normalizeParamToken(p)
    if ismissing(p), tok = NaN; return; end
    t = lower(string(p)); t = regexprep(t,'\s+','');
    t = regexprep(t,'^omega','u'); t = regexprep(t,'^ω','u'); t = regexprep(t,'^Ω','u'); t = regexprep(t,'^w','u');
    if startsWith(t,"u0"), tok = 0; return; end
    if startsWith(t,"u1"), tok = 1; return; end
    if startsWith(t,"u2"), tok = 2; return; end
    if startsWith(t,"u3"), tok = 3; return; end
    tok = NaN;
end

function s = normalizeHeader(sin)
    s = lower(string(sin));
    s = strrep(s,'（','('); s = strrep(s,'）',')');
    s = regexprep(s,'\s+','');
    s = regexprep(s,'[_/\|\\\-\–—]+','');
    s = regexprep(s,'[\(\)\[\]\{\}]+','');
end

function s = normalizeSymbol(sin)
    if ismissing(sin), s = ""; return; end
    s = string(sin);
    if strlength(s)==0, s = ""; return; end
    s = strtrim(erase(s, ["'",'"']));
    s = lower(s);
    s = upper(extractBetween(s,1,1)) + extractAfter(s,1);
end

function x = toNumSafe(v)
    if ismissing(v) || (isstring(v) && v==""), x = NaN; return; end
    if ischar(v) || isstring(v)
        vv = strrep(char(string(v)),',','.');
        x = str2double(vv);
    else
        x = double(v);
    end
end

function Zmap = symbolToZMap()
    syms = {'H','He','Li','Be','B','C','N','O','F','Ne', ...
            'Na','Mg','Al','Si','P','S','Cl','Ar','K','Ca', ...
            'Sc','Ti','V','Cr','Mn','Fe','Co','Ni','Cu','Zn', ...
            'Ga','Ge','As','Se','Br','Kr','Rb','Sr','Y','Zr', ...
            'Nb','Mo','Tc','Ru','Rh','Pd','Ag','Cd','In','Sn', ...
            'Sb','Te','I','Xe','Cs','Ba','La','Ce','Pr','Nd', ...
            'Pm','Sm','Eu','Gd','Tb','Dy','Ho','Er','Tm','Yb', ...
            'Lu','Hf','Ta','W','Re','Os','Ir','Pt','Au','Hg', ...
            'Tl','Pb','Bi','Po','At','Rn','Fr','Ra','Ac','Th', ...
            'Pa','U','Np','Pu','Am','Cm','Bk','Cf','Es','Fm', ...
            'Md','No','Lr','Rf','Db','Sg','Bh','Hs','Mt','Ds', ...
            'Rg','Cn','Nh','Fl','Mc','Lv','Ts','Og'};
    vals = num2cell(1:numel(syms));
    Zmap = containers.Map(syms, vals);
end

end % function
