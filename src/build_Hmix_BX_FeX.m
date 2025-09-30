
function build_Hmix_BX_FeX(xlsxPath, outXlsx, sheetList)
% build_Hmix_BX_FeX
% -------------------------------------------------------------
% 从一个 Excel 工作簿的多个工作表中读取元素对参数 (U0..U3)，支持三种布局：
%   1) grid：行含 A (row) 与 Param=U0..U3，列为各元素 B（H, He, Li, ..., Fe, ...）
%   2) wide：列包含 A/B 或 Pair，再加 U0..U3 列
%   3) long：列包含 A/B 或 Pair + Param(U0..U3) + Value
% 统一规范为 A–B（A=低Z，B=高Z），若源为 B–A，则对奇次项 U1/U3 变号一次。
% 在统一网格 x=0:0.001:1 上计算二元体系的混合焓（系数4）：
%   ΔH(x) = 4 x(1-x)[U0 + U1(1-2x) + U2(1-2x)^2 + U3(1-2x)^3]
%
% 输出两张主表（符合你的取向要求）：
%   B-X：x 定义为 **右侧 X 的分数**；若 Z_X < Z_B，则对 U1/U3 取负（取向变号）
%   Fe-X：x 定义为 **右侧 X 的分数**；若 Z_X < Z_Fe，则对 U1/U3 取负（取向变号）
% （可选）还可导出 LEFT 口径（x=左元素分数）的两张表，见“可调项”。
%
% 用法：
%   >> build_Hmix_BX_FeX                                 % 默认读取/输出路径见下
%   >> build_Hmix_BX_FeX(infile, outfile)               % 指定 I/O
%   >> build_Hmix_BX_FeX(infile, outfile, {'SheetA','SheetB'}) % 指定工作表
%
% 默认路径（按你的目录）：
if nargin < 1 || isempty(xlsxPath), xlsxPath = 'C:\Fe_BMAT\Fe_BM\Fe-B-X.xlsx'; end
if nargin < 2 || isempty(outXlsx),  outXlsx  = 'C:\Fe_BMAT\Fe_BM\Hmix_BX_FeX.xlsx'; end
if nargin < 3, sheetList = {}; end

% ---------------- 可调项 ----------------
exportLeftSheets = false;   % 需要同时导出 LEFT 口径（x=左元素分数）时置为 true

% 统一 x 网格
dx = 0.001;
x  = (0:dx:1)';                 % 列向量

% ---- 工作表列表 ----
if isempty(sheetList)
    try
        sheetList = sheetnames(xlsxPath);
    catch
        [~,sheetList] = xlsfinfo(xlsxPath);
    end
end
if ischar(sheetList), sheetList = {sheetList}; end

% ---- 元素 Z 表 ----
Zmap = symbolToZMap();

% ---- 聚合：canonical U(A-lowZ – B-highZ) across sheets ----
seenU  = containers.Map('KeyType','char','ValueType','any');   % key='A-B' value=[U0..U3]
srcMap = containers.Map('KeyType','char','ValueType','char');  % 记录来源工作表（最后一次覆盖）
missCt = containers.Map('KeyType','char','ValueType','double');% U 缺失数（原始）
skipLog = {}; % {sheet, row, reason}

for s = 1:numel(sheetList)
    sh = sheetList{s};
    try
        T = readtable(xlsxPath, 'Sheet', sh, 'PreserveVariableNames', true);
    catch ME
        skipLog(end+1,:) = {sh, 0, ['读表失败: ' ME.message]}; %#ok<AGROW>
        continue;
    end
    T = repairHeadersIfNeeded(T, sh);
    names = string(T.Properties.VariableNames);

    % 识别布局并累加
    [layout, info] = detectLayout(T, names, Zmap);
    switch layout
        case "grid"
            [seenU,srcMap,missCt,skipLog] = accumulate_grid(T, info, Zmap, sh, seenU, srcMap, missCt, skipLog);
        case "wide"
            [seenU,srcMap,missCt,skipLog] = accumulate_wide(T, info, Zmap, sh, seenU, srcMap, missCt, skipLog);
        case "long"
            [seenU,srcMap,missCt,skipLog] = accumulate_long(T, info, Zmap, sh, seenU, srcMap, missCt, skipLog);
        otherwise
            skipLog(end+1,:) = {sh, 0, '未识别布局（非 grid / wide / long）'}; %#ok<AGROW>
    end
end

keysCanon = string(seenU.keys)';
assert(~isempty(keysCanon), '未从任何工作表解析到可用的 U0..U3 参数。');

% ---- 元素全集 ----
elems = unique(reshape(split(keysCanon, "-"), [], 1), 'stable');
elems = elems(elems~="");  % 去空
% 如果没找到 B 或 Fe，仍继续，只会产生相应的警告
if ~any(elems=="B"),  warning('数据中未出现元素 "B"。'); end
if ~any(elems=="Fe"), warning('数据中未出现元素 "Fe"。'); end

% ---- 输出主表：RIGHT 口径（x=右侧 X 分数；若 X<Left 则对 U1/U3 变号） ----
leftList = ["B","Fe"];
for L = leftList
    if ~any(elems==L), continue; end
    Xcands = setdiff(elems, L, 'stable');
    C = cell(numel(x)+1, 1 + numel(Xcands));
    C{1,1} = sprintf('x (fraction of RIGHT element in %s–X: X)', char(L));
    for i=1:numel(x), C{1+i,1} = x(i); end

    colIdx = 0;
    for xi = 1:numel(Xcands)
        Xsym = Xcands(xi);
        [Ucanon, ok] = getU_for_pair(L, Xsym, seenU, Zmap);
        if ~ok, continue; end

        % RIGHT 口径核心：x = X 分数；若 Z_X < Z_L，则对 U1/U3 变号；变量仍用 x
        if Zmap(char(Xsym)) < Zmap(char(L))
            Ueff = [Ucanon(1), -Ucanon(2), Ucanon(3), -Ucanon(4)];
        else
            Ueff = Ucanon;
        end
        y = Hmix(Ueff, x);   % 注意：此处不做 x -> 1-x 的映射，避免二次处理

        colIdx = colIdx + 1;
        C{1, 1+colIdx} = char(L + "-" + Xsym);
        for i=1:numel(x), C{1+i, 1+colIdx} = y(i); end
    end

    if colIdx < numel(Xcands), C = C(:, 1:(1+colIdx)); end
    writecell_auto(C, outXlsx, char(L + "-X"));
end

% ---- （可选）输出 LEFT 口径（x=左元素分数；仅做变量映射，不再取向变号） ----
if exportLeftSheets
    for L = leftList
        if ~any(elems==L), continue; end
        Xcands = setdiff(elems, L, 'stable');
        C = cell(numel(x)+1, 1 + numel(Xcands));
        C{1,1} = sprintf('x (fraction of LEFT element in %s–X: %s)', char(L), char(L));
        for i=1:numel(x), C{1+i,1} = x(i); end

        colIdx = 0;
        for xi = 1:numel(Xcands)
            Xsym = Xcands(xi);
            [Ucanon, ok] = getU_for_pair(L, Xsym, seenU, Zmap);
            if ~ok, continue; end

            % LEFT 口径：若 Z_L < Z_X，canonical 为 L–X，x_B = 1 - x_left；否则 x_B = x_left
            if Zmap(char(L)) < Zmap(char(Xsym))
                xB = 1 - x;
            else
                xB = x;
            end
            y = Hmix(Ucanon, xB);  % 不再取向变号

            colIdx = colIdx + 1;
            C{1, 1+colIdx} = char(L + "-" + Xsym + "_LEFT");
            for i=1:numel(x), C{1+i, 1+colIdx} = y(i); end
        end

        if colIdx < numel(Xcands), C = C(:, 1:(1+colIdx)); end
        writecell_auto(C, outXlsx, char(L + "-X_LEFT"));
    end
end

% ---- Pairs_Used / README / Skip_Log ----
PL = cell(numel(keysCanon)+1, 7);
PL(1,:) = {'Pair (A-lowZ – B-highZ)','U0','U1','U2','U3','SourceSheet','MissingCount'};
for k=1:numel(keysCanon)
    key = char(keysCanon(k));
    U = seenU(key);
    PL{k+1,1} = key;
    PL{k+1,2} = U(1); PL{k+1,3} = U(2); PL{k+1,4} = U(3); PL{k+1,5} = U(4);
    if isKey(srcMap, key), PL{k+1,6} = srcMap(key); else, PL{k+1,6} = ''; end
    if isKey(missCt, key), PL{k+1,7} = missCt(key); else, PL{k+1,7} = 0; end
end
writecell_auto(PL, outXlsx, 'Pairs_Used');

readme = {
'字段','说明';
'输入工作簿', xlsxPath;
'工作表', strjoin(string(sheetList), ', ');
'布局支持','grid（A(row)+Param+元素列）/ wide（A,B+U0..U3）/ long（A,B+Param+Value）';
'参数规范','统一为 A–B（A=低Z、B=高Z）；若源行为 B–A，读取时对 U1/U3 取负一次';
'x 定义（主表）','B-X/Fe-X：x=右侧 X 的分数；当 Z_X<Z_Left 时对 U1/U3 取负（取向变号）';
'LEFT 口径（可选）','若启用 *_LEFT 表：x=左元素分数，仅做变量映射，不再二次取向变号';
'混合焓公式','ΔH=4 x(1-x)[U0 + U1(1-2x) + U2(1-2x)^2 + U3(1-2x)^3]（系数4）';
'统一网格','x=0..1, 步长 0.001，所有列共用';
'输出文件', outXlsx;
};
writecell_auto(readme, outXlsx, 'README');

if isempty(skipLog)
    SL = {'Sheet','Row','Reason'};
else
    SL = [{'Sheet','Row','Reason'}; skipLog];
end
writecell_auto(SL, outXlsx, 'Skip_Log');

suffix = '';
if exportLeftSheets, suffix = ', B-X_LEFT, Fe-X_LEFT'; end
fprintf('已输出：%s\n  工作表：B-X, Fe-X%s, Pairs_Used, README, Skip_Log\n', outXlsx, suffix);

% =====================================================
% =============== 计算与核心函数 ======================
% =====================================================
function y = Hmix(U, xB)
% xB：规范 A–B 中 B（高Z）的摩尔分数
    t = 1 - 2.*xB;
    y = 4 .* xB .* (1 - xB) .* (U(1) + U(2).*t + U(3).*t.^2 + U(4).*t.^3);
end

function [Ucanon, ok] = getU_for_pair(E1, E2, seenU, Zmap)
% 返回 (E1,E2) 对应的规范 U(A–B)，ok=false 表示缺失
    if ~isKey(Zmap, char(E1)) || ~isKey(Zmap, char(E2)), ok=false; Ucanon=[]; return; end
    z1 = Zmap(char(E1)); z2 = Zmap(char(E2));
    if z1 == z2, ok=false; Ucanon=[]; return; end
    if z1 < z2, key = char(E1 + "-" + E2); else, key = char(E2 + "-" + E1); end
    if ~isKey(seenU, key), ok=false; Ucanon=[]; return; end
    Ucanon = seenU(key); ok=true;
end

% =====================================================
% ===============    聚合：grid / wide / long =========
% =====================================================
function [seenU,srcMap,missCt,skipLog] = accumulate_grid(T, info, Zmap, sh, seenU, srcMap, missCt, skipLog)
    Acol = char(info.ACol); Pcol = char(info.ParamCol); Bcols = string(info.Bcols);
    for r = 1:height(T)
        A = normalizeSymbol(T.(Acol)(r));
        pidx = mapParamToken(T.(Pcol)(r));  % U0->1..U3->4
        if A=="" || isnan(pidx), continue; end
        if ~isKey(Zmap, char(A)), continue; end
        ZA = Zmap(char(A));
        for jc = 1:numel(Bcols)
            Bsym = normalizeSymbol(Bcols(jc));
            if ~isKey(Zmap, char(Bsym)), continue; end
            ZB = Zmap(char(Bsym));
            val = toNumSafe(T.(char(Bcols(jc)))(r));
            if isnan(val), continue; end
            if (ZA < ZB) || (ZA==ZB && strlength(A) <= strlength(Bsym))
                Acanon = A; Bcanon = Bsym; sgn = +1;
            else
                Acanon = Bsym; Bcanon = A; sgn = -1;
            end
            key = char(Acanon + "-" + Bcanon);
            if ~isKey(seenU, key), seenU(key) = [NaN NaN NaN NaN]; missCt(key) = 4; end
            Ucur = seenU(key);
            % 依据奇次项取向变号：
            if pidx==2 || pidx==4, Ucur(pidx) = sgn * val; else, Ucur(pidx) = val; end
            seenU(key) = Ucur; srcMap(key) = sh;
            missCt(key) = 4 - sum(~isnan(Ucur));
        end
    end
    % 缺失补零
    keys = seenU.keys;
    for k = 1:numel(keys)
        U = seenU(keys{k}); U(isnan(U)) = 0; seenU(keys{k}) = U;
    end
end

function [seenU,srcMap,missCt,skipLog] = accumulate_wide(T, info, Zmap, sh, seenU, srcMap, missCt, skipLog)
    vnames = string(T.Properties.VariableNames);
    vA = info.colA; vB = info.colB; vPair = info.colPair; vZA = info.colZA; vZB = info.colZB;
    Ucols = info.Ucols; % string of 4 names
    for r = 1:height(T)
        [sa,sb,ZA,ZB,ok] = parsePairFromRow(T, r, vnames, vA, vB, vPair, vZA, vZB, Zmap);
        if ~ok, skipLog(end+1,:) = {sh, r, '元素对缺失/无效'}; %#ok<AGROW>
            continue;
        end
        U = [toNumSafe(T.(char(Ucols(1)))(r)), ...
             toNumSafe(T.(char(Ucols(2)))(r)), ...
             toNumSafe(T.(char(Ucols(3)))(r)), ...
             toNumSafe(T.(char(Ucols(4)))(r))];
        if all(isnan(U)), skipLog(end+1,:) = {sh, r, 'U0..U3 全缺失'}; %#ok<AGROW>
            continue;
        end
        if (ZA < ZB) || (ZA==ZB && strlength(sa) <= strlength(sb))
            Acanon = sa; Bcanon = sb; Ucanon = U;
        else
            Acanon = sb; Bcanon = sa; Ucanon = [ U(1), -U(2), U(3), -U(4) ];
        end
        key = char(Acanon + "-" + Bcanon);
        seenU(key) = fillMissingToZero(Ucanon);
        srcMap(key) = sh; missCt(key) = sum(isnan(U));
    end
end

function [seenU,srcMap,missCt,skipLog] = accumulate_long(T, info, Zmap, sh, seenU, srcMap, missCt, skipLog)
    vnames = string(T.Properties.VariableNames);
    pcol = info.paramCol; vcol = info.valueCol;
    vA = info.colA; vB = info.colB; vPair = info.colPair; vZA = info.colZA; vZB = info.colZB;
    tmp = containers.Map('KeyType','char','ValueType','any');  % 先按原始顺序收集再规范
    for r = 1:height(T)
        [sa,sb,ZA,ZB,ok] = parsePairFromRow(T, r, vnames, vA, vB, vPair, vZA, vZB, Zmap);
        if ~ok, skipLog(end+1,:) = {sh, r, '元素对缺失/无效'}; %#ok<AGROW>
            continue;
        end
        pidx = mapParamToken(T.(char(pcol))(r));
        val  = toNumSafe(T.(char(vcol))(r));
        if isnan(pidx) || isnan(val), skipLog(end+1,:) = {sh, r, '参数名或数值缺失'}; %#ok<AGROW>
            continue;
        end
        key0 = char(sa + "|" + sb);  % 暂时不规范，收集完再统一
        if ~isKey(tmp, key0), tmp(key0) = [NaN NaN NaN NaN]; end
        Ucur = tmp(key0);
        Ucur(pidx) = val;
        tmp(key0) = Ucur;
    end
    % 统一规范 + 写入 seenU
    keys0 = tmp.keys;
    for k = 1:numel(keys0)
        key0 = string(keys0{k});
        parts = split(key0,"|"); sa = parts(1); sb = parts(2);
        ZA = Zmap(char(sa)); ZB = Zmap(char(sb));
        U = tmp(keys0{k});
        if (ZA < ZB) || (ZA==ZB && strlength(sa) <= strlength(sb))
            Acanon = sa; Bcanon = sb; Ucanon = U;
        else
            Acanon = sb; Bcanon = sa; Ucanon = [ U(1), -U(2), U(3), -U(4) ];
        end
        key = char(Acanon + "-" + Bcanon);
        seenU(key) = fillMissingToZero(Ucanon);
        srcMap(key) = sh; missCt(key) = sum(isnan(U));
    end
end

% =====================================================
% ===============   布局识别 & 解析辅助  ==============
% =====================================================
function [layout, info] = detectLayout(T, names, Zmap)
    layout = ""; info = struct();
    % grid：必须有 Param + A(row)，并且至少 3 列是元素符号
    pcol = findFirstAmongFuzzy(names, ["param","parameter"]);
    acol = findFirstAmongFuzzy(names, ["a(row)","a_row","a (row)","arow"]);
    if ~isnan(pcol) && ~isnan(acol)
        elemCols = strings(0,1);
        for i = 1:numel(names)
            nm = string(names(i)); nmClean = regexprep(nm, '\s+', '');
            if isKey(Zmap, char(nmClean)), elemCols(end+1,1) = nm; end %#ok<AGROW>
        end
        if numel(elemCols) >= 3
            layout = "grid"; info.ACol = names(acol); info.ParamCol = names(pcol); info.Bcols = elemCols; return;
        end
    end
    % wide：U0..U3 四列齐全 + A/B 或 Pair
    colU0 = findParamCol(names,0); colU1=findParamCol(names,1); colU2=findParamCol(names,2); colU3=findParamCol(names,3);
    vA = findFirstAmongFuzzy(names, ["a","el_a","elem_a","element_a","e1","el1","element1","left","left_elem"]);
    vB = findFirstAmongFuzzy(names, ["b","el_b","elem_b","element_b","e2","el2","element2","right","right_elem"]);
    vPair = findFirstFuzzy(names, "pair");
    vZA = findFirstAmongFuzzy(names, ["z_a","za","z1","zleft"]);
    vZB = findFirstAmongFuzzy(names, ["z_b","zb","z2","zright"]);
    if all(~isnan([colU0 colU1 colU2 colU3])) && (~isnan(vPair) || (~isnan(vA) && ~isnan(vB)))
        layout = "wide";
        info.Ucols = names([colU0 colU1 colU2 colU3]);
        info.colA=vA; info.colB=vB; info.colPair=vPair; info.colZA=vZA; info.colZB=vZB;
        return;
    end
    % long：存在 param 名列 + 可数值化的 value 列 + (A/B or Pair)
    candParam = [];
    for i=1:numel(names)
        v = T.(char(names(i)));
        if iscellstr(v) || isstring(v) || iscategorical(v)
            tokens = unique(string(v));
            m = 0;
            for t = 1:numel(tokens), if ~isnan(mapParamToken(tokens(t))), m = m+1; end, end
            if m >= 3, candParam(end+1) = i; end %#ok<AGROW>
        end
    end
    candVal = [];
    for i=1:numel(names)
        v = T.(char(names(i)));
        if isnumeric(v)
            if any(~isnan(double(v))), candVal(end+1)=i; end %#ok<AGROW>
        elseif iscellstr(v) || isstring(v) || iscategorical(v)
            nums = arrayfun(@(s) str2double(strrep(char(string(s)),',','.')), v);
            if any(~isnan(nums)), candVal(end+1)=i; end %#ok<AGROW>
        end
    end
    if ~isempty(candParam) && ~isempty(candVal) && (~isnan(vPair) || (~isnan(vA) && ~isnan(vB)))
        layout = "long";
        info.paramCol = names(candParam(1)); info.valueCol = names(candVal(1));
        info.colA=vA; info.colB=vB; info.colPair=vPair; info.colZA=vZA; info.colZB=vZB;
        return;
    end
    layout = "";
end

function [sa,sb,ZA,ZB,ok] = parsePairFromRow(T, r, names, vA, vB, vPair, vZA, vZB, Zmap)
% 兼容性解析：A/B 可能为空、也可能把"Al-Fe"一类的字符串放在 A 或 B 单元格里；
% 若检测到分隔符或无法识别为单一元素，则尝试从该单元格或 Pair 列中拆分出两种元素。

    sa=""; sb=""; ZA=NaN; ZB=NaN; ok=false;

    valA = "";
    valB = "";
    if ~isnan(vA), valA = T.(char(names(vA)))(r); end
    if ~isnan(vB), valB = T.(char(names(vB)))(r); end

    % 先尝试把 A/B 作为单一元素读取
    sa0 = normalizeSymbol(valA);
    sb0 = normalizeSymbol(valB);
    isSymA = isElementSymbol(sa0, Zmap);
    isSymB = isElementSymbol(sb0, Zmap);

    % 如果 A/B 不是有效单元素，或任一包含明显分隔符，则尝试拆分
    needSplit = false;
    if ~isSymA || ~isSymB
        needSplit = true;
    end
    if looksLikePairString(valA) || looksLikePairString(valB)
        needSplit = true;
    end

    if ~needSplit && isSymA && isSymB
        sa = sa0; sb = sb0;
    else
        % 优先从 A 单元格拆分
        [s1a, s1b] = splitPairString(valA);
        if isElementSymbol(s1a,Zmap) && isElementSymbol(s1b,Zmap)
            sa = s1a; sb = s1b;
        else
            % 再从 B 单元格拆分
            [s2a, s2b] = splitPairString(valB);
            if isElementSymbol(s2a,Zmap) && isElementSymbol(s2b,Zmap)
                sa = s2a; sb = s2b;
            else
                % 再从 Pair 列拆分
                if ~isnan(vPair)
                    pairStr = T.(char(names(vPair)))(r);
                    [s3a, s3b] = splitPairString(pairStr);
                    if isElementSymbol(s3a,Zmap) && isElementSymbol(s3b,Zmap)
                        sa = s3a; sb = s3b;
                    end
                end
            end
        end
    end

    % 若仍无效，直接返回 ok=false
    if ~(isElementSymbol(sa,Zmap) && isElementSymbol(sb,Zmap))
        ok=false; return;
    end

    % 原子序（若无 Z 列则由符号映射）
    if ~isnan(vZA) && ~isnan(vZB)
        ZA = toNumSafe(T.(char(names(vZA)))(r));
        ZB = toNumSafe(T.(char(names(vZB)))(r));
    else
        ZA = symbol2Z(sa, Zmap);
        ZB = symbol2Z(sb, Zmap);
    end
    ok = ~(isnan(ZA) || isnan(ZB));
end

function tf = isElementSymbol(sym, Zmap)
    if ismissing(sym) || strlength(string(sym))==0, tf=false; return; end
    s = char(normalizeSymbol(sym));
    tf = isKey(Zmap, s);
end

function tf = looksLikePairString(x)
    if ismissing(x), tf=false; return; end
    str = string(x);
    if strlength(str)==0, tf=false; return; end
    tf = ~isempty(regexp(char(str), '[-–—/\,;:\s\._]+', 'once'));
end
function T = repairHeadersIfNeeded(T, sheetName)
    vnames = string(T.Properties.VariableNames);
    bad = sum(contains(vnames, ["unnamed","var"], 'IgnoreCase', true));
    if bad >= max(3, round(0.2*numel(vnames)))
        C = table2cell(T);
        hdr = string(C(1,:));
        if any(strlength(hdr)>0)
            hn = matlab.lang.makeUniqueStrings(matlab.lang.makeValidName(cellstr(hdr)));
            T = cell2table(C(2:end,:), 'VariableNames', hn);
            fprintf('Sheet %s: 已修复表头（提升首行为变量名）。\n', string(sheetName));
        end
    end
end

function idx = findFirstFuzzy(vnames, token)
    token = lower(token);
    for i=1:numel(vnames)
        name = lower(string(vnames(i))); name = normalizeHeader(name);
        if contains(name, regexprep(token,'\s+','')), idx=i; return; end
    end
    idx = nan;
end

function idx = findFirstAmongFuzzy(vnames, candidates)
    idx = nan;
    for c=1:numel(candidates)
        idx = findFirstFuzzy(vnames, candidates(c));
        if ~isnan(idx), return; end
    end
end

function idx = findParamCol(vnames, num0to3)
    idx = nan;
    for i=1:numel(vnames)
        raw = string(vnames(i));
        name = lower(raw); norm = normalizeHeader(name);
        n = num2str(num0to3);
        patterns = ["u"+n,"omega"+n,"ω"+n,"Ω"+n,"w"+n, ...
                    "u_"+n,"omega_"+n,"w_"+n, "u"+ "_" + n, "omega"+ "_" + n, "w"+ "_" + n];
        ok = false;
        for p=1:numel(patterns), if contains(norm, patterns(p)), ok=true; break; end, end
        if ~ok
            rgx = "(\b|[^a-z0-9])(u|omega|ω|Ω|w)\s*[_\.\-\(\)\[\]\s]*" + n + "(\b|[^a-z0-9])";
            ok = ~isempty(regexp(name, rgx, 'once'));
        end
        if ok, idx=i; return; end
    end
end

function pidx = mapParamToken(tok)
    if ismissing(tok), pidx = NaN; return; end
    t = lower(normalizeHeader(string(tok)));
    if t=="" || ismissing(t), pidx = NaN; return; end
    if contains(t,"u0") || contains(t,"omega0") || contains(t,"ω0") || contains(t,"Ω0") || contains(t,"w0"), pidx = 1; return; end
    if contains(t,"u1") || contains(t,"omega1") || contains(t,"ω1") || contains(t,"Ω1") || contains(t,"w1"), pidx = 2; return; end
    if contains(t,"u2") || contains(t,"omega2") || contains(t,"ω2") || contains(t,"Ω2") || contains(t,"w2"), pidx = 3; return; end
    if contains(t,"u3") || contains(t,"omega3") || contains(t,"ω3") || contains(t,"w3"), pidx = 4; return; end
    pidx = NaN;
end

function s = normalizeHeader(sin)
    s = lower(string(sin));
    s = strrep(s,'（','('); s = strrep(s,'）',')');
    s = strrep(s,'【','['); s = strrep(s,'】',']');
    s = regexprep(s,'\s+','');
    s = regexprep(s,'[_/\|\\\-\–—]+','');
    s = regexprep(s,'[\(\)\[\]\{\}]+','');
    s = regexprep(s,'[\+\*\^\%\,\;:]+','');
end

function [sa, sb] = splitPairString(pairStr)
    if ismissing(pairStr) || (isstring(pairStr) && pairStr=="") , sa=""; sb=""; return; end
    txt = char(string(pairStr));
    parts = regexp(txt, '[-–—/\\,;:\s\._]+', 'split');
    parts = parts(~cellfun('isempty',parts));
    if numel(parts) < 2, sa=""; sb=""; return; end
    sa = normalizeSymbol(parts{1});
    sb = normalizeSymbol(parts{2});
end

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

function Z = symbol2Z(sym, Zmap)
    if isstring(sym) && ismissing(sym), error("元素符号缺失（<missing>）。"); end
    s = normalizeSymbol(sym);
    if s=="", error("元素符号为空。"); end
    schar = char(s);
    if ~isKey(Zmap, schar)
        % 尝试从复合字符串中提取唯一有效元素（例如 "[Fe]" 类似）
        tokens = regexp(schar, '[A-Za-z]+', 'match');
        hits = {};
        for i=1:numel(tokens)
            cand = char(normalizeSymbol(tokens{i}));
            if isKey(Zmap, cand)
                hits{end+1} = cand; %#ok<AGROW>
            end
        end
        if numel(hits)==1
            Z = Zmap(hits{1}); return;
        end
        error("未知元素符号：%s", s);
    end
    Z = Zmap(schar);
end

end

function U = fillMissingToZero(U)
    U = double(U);
    U(isnan(U)) = 0;
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
