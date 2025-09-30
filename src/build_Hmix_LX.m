
function build_Hmix_LX(xlsxPath, outXlsx, sheetName, xMode)
% build_Hmix_LX  Compute ΔH_mix(x) for B–X and Fe–X with a clear x-axis choice.
%
% xMode:
%   'left'  (default) -> x = fraction of LEFT element in "L–X" (i.e., x_B in B–X, x_Fe in Fe–X).
%   'right'           -> x = fraction of RIGHT element X in "L–X"（你之前使用的口径）。
%
% Orientation rule（仅在 xMode='right' 时需要）：当 Z_X < Z_left，为保证“左–右”取向，奇次项 U1/U3 变号。
% 在 xMode='left' 下，内部使用**规范 A–B（A低Z、B高Z）**的参数与自变量 x_B：
%   若 Z_left < Z_X：x_B = 1 - x_left；若 Z_left > Z_X：x_B = x_left；无需再对 U 变号。
%
% 公式（系数4）：ΔH = 4 x (1-x) [ U0 + U1(1-2x) + U2(1-2x)^2 + U3(1-2x)^3 ]
%
% 用法：
%   build_Hmix_LX                                  % 默认：从 C:\Fe_BMAT\Fe_BM\Fe-B-X.xlsx 读，左轴口径
%   build_Hmix_LX(infile, outfile, 'All_tidy');   % 指定表，左轴口径
%   build_Hmix_LX(infile, outfile, 'All_tidy','right');  % 使用右轴口径（X 分数）
%
% 作者：ChatGPT
if nargin < 1 || isempty(xlsxPath), xlsxPath = 'C:\Fe_BMAT\Fe_BM\Fe-B-X.xlsx'; end
if nargin < 2 || isempty(outXlsx),  outXlsx  = 'C:\Fe_BMAT\Fe_BM\Hmix_Fe_B_oriented.xlsx'; end
if nargin < 3 || isempty(sheetName), sheetName = 'All_tidy'; end
if nargin < 4 || isempty(xMode),     xMode = 'left'; end
xMode = lower(string(xMode));

dx = 0.001;
x  = (0:dx:1)';  % shared grid

% ---- Read & repair headers (grid expected) ----
T = readtable(xlsxPath, 'Sheet', sheetName, 'PreserveVariableNames', true);
T = repairHeadersIfNeeded(T);
names = string(T.Properties.VariableNames);

Acol  = findCol(names, ["a(row)","a_row","a (row)","arow"]);
Pcol  = findCol(names, ["param","parameter"]);
assert(~isempty(Acol) && ~isempty(Pcol), '未找到 "A (row)" 或 "Param" 列，请确认表结构。');

Zmap = symbolToZMap();
Bcols = strings(0,1);
for i = 1:numel(names)
    nm = string(names(i));
    nmClean = regexprep(nm, '\s+', '');
    if isKey(Zmap, char(nmClean))
        Bcols(end+1,1) = nm; %#ok<AGROW>
    end
end
assert(numel(Bcols)>=2, '未识别到元素列，请检查表头。');

% ---- Gather canonical U(A-lowZ – B-highZ) ----
seen = containers.Map('KeyType','char','ValueType','any');
elemsSeen = containers.Map('KeyType','char','ValueType','logical');
for r = 1:height(T)
    A = normalizeSymbol(T.(Acol)(r));
    P = normalizeParamToken(T.(Pcol)(r));  % U0->0 .. U3->3
    if A=="" || isnan(P), continue; end
    pidx = P + 1;
    if ~isKey(Zmap, char(A)), continue; end
    ZA = Zmap(char(A));
    elemsSeen(char(A)) = true;
    for jc = 1:numel(Bcols)
        Bsym = normalizeSymbol(Bcols(jc));
        if ~isKey(Zmap, char(Bsym)), continue; end
        elemsSeen(char(Bsym)) = true;
        val = toNumSafe(T.(char(Bcols(jc)))(r));
        if isnan(val), continue; end
        ZB = Zmap(char(Bsym));
        if (ZA < ZB) || (ZA==ZB && strlength(A) <= strlength(Bsym))
            Acanon = A; Bcanon = Bsym; sgn = +1;
        else
            Acanon = Bsym; Bcanon = A; sgn = -1;
        end
        key = char(Acanon + "-" + Bcanon);
        if ~isKey(seen, key), seen(key) = [NaN NaN NaN NaN]; end
        Ucur = seen(key);
        if pidx==2 || pidx==4, Ucur(pidx) = sgn * val; else, Ucur(pidx) = val; end
        seen(key) = Ucur;
    end
end
% fill missing
keysCanon = string(seen.keys)';
for k = 1:numel(keysCanon)
    U = seen(char(keysCanon(k))); U(isnan(U)) = 0; seen(char(keysCanon(k))) = U;
end

allElems = string(elemsSeen.keys)';
leftList = ["B","Fe"];

for L = leftList
    if ~any(allElems==L), continue; end
    Xcands = setdiff(allElems, L, 'stable');
    C = cell(numel(x)+1, 1 + numel(Xcands));
    if xMode=="left"
        C{1,1} = sprintf('x (fraction of LEFT element in %s–X: %s)', char(L), char(L));
    else
        C{1,1} = sprintf('x (fraction of RIGHT element in %s–X: X)', char(L));
    end
    for i=1:numel(x), C{1+i,1} = x(i); end

    colIdx = 0;
    for xi = 1:numel(Xcands)
        Xsym = Xcands(xi);
        [Ucanon, ok] = getCanonicalU(L, Xsym, seen, Zmap);
        if ~ok, continue; end

        if xMode=="left"
            % x_left -> x_B (canonical variable)
            if Zmap(char(L)) < Zmap(char(Xsym))
                % canonical A–B = L–X; x_B = x_X = 1 - x_left
                xB = 1 - x;
            else
                % canonical A–B = X–L; B = L; x_B = x_left
                xB = x;
            end
            y = Hmix_canonical(Ucanon, xB);
        else % xMode == "right"
            % x_right = x
            Uo = Ucanon;
            if Zmap(char(Xsym)) < Zmap(char(L))
                % orientation flip for odd terms when showing L–X but X is lower-Z
                Uo = [Ucanon(1), -Ucanon(2), Ucanon(3), -Ucanon(4)];
            end
            y = Hmix_canonical(Uo, x);
        end

        colIdx = colIdx + 1;
        C{1, 1+colIdx} = char(L + "-" + Xsym);
        for i=1:numel(x), C{1+i, 1+colIdx} = y(i); end
    end

    if colIdx < numel(Xcands), C = C(:, 1:(1+colIdx)); end
    writecell_auto(C, outXlsx, char(L + "-X"));
end

% Pairs_Used
PU = cell(numel(keysCanon)+1, 5);
PU(1,:) = {'Pair (A-lowZ – B-highZ)','U0','U1','U2','U3'};
for k=1:numel(keysCanon)
    U = seen(char(keysCanon(k)));
    PU{k+1,1} = char(keysCanon(k));
    PU{k+1,2} = U(1); PU{k+1,3} = U(2); PU{k+1,4} = U(3); PU{k+1,5} = U(4);
end
writecell_auto(PU, outXlsx, 'Pairs_Used');

% README
readme = {
'字段','说明';
'x 定义','xMode=left: x=左元素分数；xMode=right: x=右元素 X 分数';
'规范参数','先统一 A–B（A低Z、B高Z）；若原数据是 B–A，读入时对 U1/U3 变号以转为 A–B';
'左轴计算','xMode=left: 若 Z_left<Z_X，用 x_B=1-x_left；若 Z_left>Z_X，用 x_B=x_left；无需再变号';
'右轴计算','xMode=right: 若 Z_X<Z_left，为保持 L–X 取向，对 U1/U3 变号';
'混合焓公式','ΔH = 4 x(1-x)[U0 + U1(1-2x) + U2(1-2x)^2 + U3(1-2x)^3]（系数4）';
'共享网格','0..1 步长 0.001（所有列共用）';
'文件', outXlsx;
'来源工作表', sheetName;
};
writecell_auto(readme, outXlsx, 'README');

fprintf('已输出：%s（xMode=%s）\n  工作表：B-X, Fe-X, Pairs_Used, README\n', outXlsx, xMode);

% --------------- helpers ----------------
function y = Hmix_canonical(U, xB)
    t = 1 - 2.*xB;
    y = 4 .* xB .* (1 - xB) .* (U(1) + U(2).*t + U(3).*t.^2 + U(4).*t.^3);
end

function T = repairHeadersIfNeeded(T)
    vnames = string(T.Properties.VariableNames);
    bad = sum(contains(vnames, ["unnamed","var"], 'IgnoreCase', true));
    if bad >= max(3, round(0.2*numel(vnames)))
        C = table2cell(T);
        hdr = string(C(1,:));
        if any(strlength(hdr)>0)
            hn = matlab.lang.makeUniqueStrings(matlab.lang.makeValidName(cellstr(hdr)));
            T = cell2table(C(2:end,:), 'VariableNames', hn);
            fprintf('已修复表头：使用首行作为变量名。\n');
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

function s = normalizeHeader(sin)
    s = lower(string(sin));
    s = strrep(s,'（','('); s = strrep(s,'）',')');
    s = regexprep(s,'\s+','');
    s = regexprep(s,'[_/\|\\\-\–—]+','');
    s = regexprep(s,'[\(\)\[\]\{\}]+','');
end

function tok = normalizeParamToken(p)
    if ismissing(p), tok = NaN; return; end
    t = lower(string(p)); t = regexprep(t,'\s+','');
    if startsWith(t,"u0"), tok = 0; return; end
    if startsWith(t,"u1"), tok = 1; return; end
    if startsWith(t,"u2"), tok = 2; return; end
    if startsWith(t,"u3"), tok = 3; return; end
    t2 = regexprep(t,'^omega','u'); t2 = regexprep(t2,'^ω','u'); t2 = regexprep(t2,'^Ω','u'); t2 = regexprep(t2,'^w','u');
    if startsWith(t2,"u0"), tok = 0; return; end
    if startsWith(t2,"u1"), tok = 1; return; end
    if startsWith(t2,"u2"), tok = 2; return; end
    if startsWith(t2,"u3"), tok = 3; return; end
    tok = NaN;
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

function [Ucanon, ok] = getCanonicalU(E1, E2, seen, Zmap)
    if ~isKey(Zmap, char(E1)) || ~isKey(Zmap, char(E2)), ok=false; Ucanon=[]; return; end
    z1 = Zmap(char(E1)); z2 = Zmap(char(E2));
    if z1 == z2, ok=false; Ucanon=[]; return; end
    if z1 < z2, key = char(E1 + "-" + E2); else, key = char(E2 + "-" + E1); end
    if ~isKey(seen, key), ok=false; Ucanon=[]; return; end
    Ucanon = seen(key); ok=true;
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

function x = toNumSafe(v)
    if ismissing(v) || (isstring(v) && v==""), x = NaN; return; end
    if ischar(v) || isstring(v)
        vv = strrep(char(string(v)),',','.');
        x = str2double(vv);
    else
        x = double(v);
    end
end

function writecell_auto(C, xlsx, sheet)
    if exist('writecell','file')==2
        writecell(C, xlsx, 'Sheet', sheet);
    else
        xlswrite(xlsx, C, sheet);
    end
end

end % function
