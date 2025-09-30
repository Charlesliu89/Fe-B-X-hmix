
function build_Hmix_FB_X_ternary(xlsxPath, outXlsx, sheetName)
% build_Hmix_FB_X_ternary
% -------------------------------------------------------------------------
% 从 Excel（默认：C:\Fe_BMAT\Fe_BM\Fe-B-X.xlsx 的 All_tidy 表）读取二元参数（U0..U3，网格型布局：
%  A (row) + Param + 多列元素），将所有元素对统一为**规范 A–B（A=低 Z，B=高 Z）**，若源为 B–A，
% 读入时对**奇次项 U1/U3 变号**一次。随后：
%  (1) 计算 **B–X** 与 **Fe–X** 二元混合焓曲线（x=右侧 X 的摩尔分数；若 Z_X < Z_Left，
%      为保持“左–右”取向，对 U1/U3 再次变号），网格 x=0:0.001:1；
%  (2) 提供**三元混合焓计算器**：对任意 Fe–B–X 组成 (c_Fe,c_B,c_X)，按**两两元素对相加**：
%      \Delta H^tern = \sum_{pairs i<j} 4 c_i c_j [U0^{ij} + U1^{ij}(c_i-c_j) + U2^{ij}(c_i-c_j)^2 + U3^{ij}(c_i-c_j)^3]
%      其中 U^{ij} 为**规范 A–B**的参数（已在读入阶段完成 B–A → A–B 的奇次项变号）。
%
% 用法：
%   >> build_Hmix_FB_X_ternary                            % 使用默认路径/表/输出
%   >> build_Hmix_FB_X_ternary(infile, outfile, 'All_tidy')
%
% 计算器（运行后可直接调用）：
%   >> [Htot,parts] = Hmix3_calc('Si', 0.70, 0.20, 0.10);  % Fe70 B20 Si10
%   >> [Htot,parts] = Hmix3_calc('Fe', 0.83, 0.17, 0.00);  % 退化为二元（X=Fe 或 c_X=0 亦可）
%
% 输出：
%   outfile（默认 C:\Fe_BMAT\Fe_BM\Hmix_FB_X_ternary.xlsx）包含：
%   - 'B-X'、'Fe-X'：按**取向规则**输出的二元曲线（x=右侧 X 的分数，步长 0.001）
%   - 'Pairs_Used'：规范 A–B 的 U0..U3 参数表
%   - 'README'：规则与公式说明
%
% 作者：ChatGPT

if nargin < 1 || isempty(xlsxPath), xlsxPath = 'C:\Fe_BMAT\Fe_BM\Fe-B-X.xlsx'; end
if nargin < 2 || isempty(outXlsx),  outXlsx  = 'C:\Fe_BMAT\Fe_BM\Hmix_FB_X_ternary.xlsx'; end
if nargin < 3 || isempty(sheetName), sheetName = 'All_tidy'; end

dx = 0.001;
x  = (0:dx:1)';   % 二元统一网格（x = 右侧 X 的分数）

% ===== 1) 读取 All_tidy（网格型）并汇总到规范 A–B 的 U-map =====
T = readtable(xlsxPath, 'Sheet', sheetName, 'PreserveVariableNames', true);
T = repairHeadersIfNeeded(T, sheetName);
names = string(T.Properties.VariableNames);
Acol  = findCol(names, ["a(row)","a_row","a (row)","arow"]);
Pcol  = findCol(names, ["param","parameter"]);
assert(~isempty(Acol) && ~isempty(Pcol), '未找到 "A (row)" 或 "Param" 列，请检查 All_tidy。');

Zmap = symbolToZMap();
Bcols = strings(0,1);
for i = 1:numel(names)
    nm = string(names(i));
    nmClean = regexprep(nm, '\s+', '');
    if isKey(Zmap, char(nmClean))
        Bcols(end+1,1) = nm; %#ok<AGROW>
    end
end
assert(~isempty(Bcols), '未识别到元素列，请检查 All_tidy 的列名。');

% 规范 U(A-lowZ – B-highZ) 的映射
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
keys = string(Umap.keys)';
for k=1:numel(keys)
    U = Umap(char(keys(k))); U(isnan(U)) = 0; Umap(char(keys(k))) = U;
end
assert(~isempty(keys), 'All_tidy 中未解析到任何 U 参数。');

% ===== 2) 按“取向规则”输出二元 B–X / Fe–X 曲线（x=右侧 X 的分数） =====
elems = unique(reshape(split(keys, "-"), [], 1), 'stable');
elems = elems(elems~="");
for L = ["B","Fe"]
    if ~any(elems==L), continue; end
    Xs = setdiff(elems, L, 'stable');
    C = cell(numel(x)+1, 1 + numel(Xs));
    C{1,1} = sprintf('x (fraction of RIGHT element in %s–X: X)', char(L));
    C(2:end,1) = num2cell(x);
    col = 0;
    for j=1:numel(Xs)
        X = Xs(j);
        [Ucanon, ok] = getU_canon(L, X, Umap, Zmap); if ~ok, continue; end
        % 取向规则：若 Z_X < Z_Left，则对奇次项 U1/U3 变号（以保持 “Left–X”）
        if Zmap(char(X)) < Zmap(char(L))
            Ueff = [Ucanon(1), -Ucanon(2), Ucanon(3), -Ucanon(4)];
        else
            Ueff = Ucanon;
        end
        y = Hmix_binary(Ueff, x);
        col = col + 1;
        C{1,1+col} = char(L + "-" + X);
        C(2:end,1+col) = num2cell(y);
    end
    if col > 0
        C = C(:,1:(1+col));
        writecell_auto(C, outXlsx, char(L + "-X"));
    end
end

% ===== 3) 输出 Pairs_Used / README =====
PL = cell(numel(keys)+1, 5);
PL(1,:) = {'Pair (A-lowZ – B-highZ)','U0','U1','U2','U3'};
for k=1:numel(keys)
    U = Umap(char(keys(k)));
    PL{k+1,1} = char(keys(k));
    PL{k+1,2} = U(1); PL{k+1,3} = U(2); PL{k+1,4} = U(3); PL{k+1,5} = U(4);
end
writecell_auto(PL, outXlsx, 'Pairs_Used');

readme = {
'字段','说明';
'输入', xlsxPath;
'工作表', sheetName;
'布局', 'All_tidy（网格：A(row)+Param+元素列）';
'参数规范', '统一为 A–B（A=低Z、B=高Z），若源为 B–A，读入时奇次项 U1/U3 取负一次';
'二元取向', 'B–X/Fe–X：x=右侧 X 分数；当 Z_X<Z_Left 时再对 U1/U3 取负';
'三元计算', 'Fe–B–X：ΔH = ∑_{pairs} 4 c_i c_j [U0 + U1(c_i-c_j) + U2(c_i-c_j)^2 + U3(c_i-c_j)^3]';
'网格', '二元曲线使用 x=0:0.001:1；三元计算为任意给定 (c_Fe,c_B,c_X)';
'函数', 'Hmix3_calc(Xsym,cFe,cB,cX) 返回总值与三对分项，并可按需写入结果';
'输出文件', outXlsx;
};
writecell_auto(readme, outXlsx, 'README');

fprintf('已输出二元曲线与参数表至：%s\n', outXlsx);
fprintf('三元计算器已就绪：使用 [Htot,parts] = Hmix3_calc(Xsym,cFe,cB,cX) 调用。\n');

% ====== 4) —— 三元混合焓计算器（供用户即时调用） ======
% 用法： [Htot,parts] = Hmix3_calc('Si', 0.70, 0.20, 0.10);
% parts 结构体含：H_FeB, H_FeX, H_BX 及 U 参数等
    function [Htot, parts] = Hmix3_calc(Xsym, cFe, cB, cX)
        Xsym = normalizeSymbol(Xsym);
        if ~isKey(Zmap, char(Xsym)), error('未知元素符号：%s', char(Xsym)); end
        % 自动归一化 + 0.001 网格化（可保持你的网格口径）
        v = [cFe, cB, cX]; v(v<0)=0; s = sum(v);
        if s<=0, error('浓度全为 0。'); end
        v = v / s;
        v = round(v/0.001)*0.001;  % 对齐 0.001 网格
        % 再归一化保证总和为 1（避免累计误差）
        v = v / sum(v);
        cFe = v(1); cB = v(2); cX = v(3);

        % 三个对的 U（规范 A–B）
        [U_FeB, ok1, A1, B1] = getU_canon('Fe', 'B', Umap, Zmap); if ~ok1, error('缺少 Fe–B 的 U 参数'); end
        [U_FeX, ok2, A2, B2] = getU_canon('Fe', Xsym, Umap, Zmap); if ~ok2, error('缺少 Fe–%s 的 U 参数', char(Xsym)); end
        [U_BX , ok3, A3, B3] = getU_canon('B' , Xsym, Umap, Zmap); if ~ok3, error('缺少 B–%s 的 U 参数', char(Xsym)); end

        % 取规范 A–B 的 c_A, c_B
        [cA1,cB1] = pickABconcs(A1,B1,cFe,cB);   H_FeB = Hmix_pair(U_FeB, cA1, cB1);
        [cA2,cB2] = pickABconcs(A2,B2,cFe,cX);   H_FeX = Hmix_pair(U_FeX, cA2, cB2);
        [cA3,cB3] = pickABconcs(A3,B3,cB ,cX);   H_BX  = Hmix_pair(U_BX , cA3, cB3);

        Htot = H_FeB + H_FeX + H_BX;

        parts = struct();
        parts.c = struct('Fe',cFe,'B',cB,'X',cX,'Xsym',Xsym);
        parts.pairs = struct();
        parts.pairs.FeB = struct('A',A1,'B',B1,'U',U_FeB,'H',H_FeB);
        parts.pairs.FeX = struct('A',A2,'B',B2,'U',U_FeX,'H',H_FeX);
        parts.pairs.BX  = struct('A',A3,'B',B3,'U',U_BX ,'H',H_BX );
        fprintf('Fe–B–%s at [Fe=%.3f, B=%.3f, %s=%.3f]  =>  ΔH = %.6g  (FeB=%.6g, FeX=%.6g, BX=%.6g)\n', ...
            char(Xsym), cFe, cB, char(Xsym), cX, Htot, H_FeB, H_FeX, H_BX);
    end

% ====== 工具函数区域 ======
function y = Hmix_binary(U, xB)
    t = 1 - 2.*xB;
    y = 4 .* xB .* (1 - xB) .* ( U(1) + U(2).*t + U(3).*t.^2 + U(4).*t.^3 );
end

function H = Hmix_pair(U, cA, cB)
    % 三元对的贡献（规范 A–B，用**全系分数**）： 4 c_A c_B Σ U_n (c_A - c_B)^n
    t = (cA - cB);
    H = 4 .* cA .* cB .* ( U(1) + U(2).*t + U(3).*t.^2 + U(4).*t.^3 );
end

function [Ucanon, ok, A, B] = getU_canon(E1, E2, Umap, Zmap)
    % 返回规范 A–B 的 U 以及 A/B 的符号（字符串）
    ok=false; Ucanon=[]; A=""; B="";
    if ~isKey(Zmap, char(E1)) || ~isKey(Zmap, char(E2)), return; end
    z1 = Zmap(char(E1)); z2 = Zmap(char(E2));
    if z1 == z2, return; end
    if z1 < z2, key = char(E1 + "-" + E2); A=E1; B=E2;
    else,        key = char(E2 + "-" + E1); A=E2; B=E1;
    end
    if ~isKey(Umap, key), return; end
    Ucanon = Umap(key); ok=true;
end

function [cA,cB] = pickABconcs(A, B, c1, c2)
    % 已知 (A,B) 的顺序（A=低Z，B=高Z）以及 (c1,c2) 分别是这两个元素的分数
    % 返回 (cA,cB) 按规范顺序的浓度
    cA = c1; cB = c2;
end

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
    if ismissing(s) || s=="", s=""; return; end
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
    Zmap = containers.Map(syms, num2cell(1:numel(syms)));
end

function writecell_auto(C, xlsx, sheet)
    if exist('writecell','file')==2
        writecell(C, xlsx, 'Sheet', sheet);
    else
        xlswrite(xlsx, C, sheet);
    end
end

end % function
