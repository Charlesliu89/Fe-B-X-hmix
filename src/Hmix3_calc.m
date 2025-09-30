
function [Htot, parts] = Hmix3_calc(Xsym, cFe, cB, cX, dbXlsx, mode)
% Hmix3_calc  Ternary mixing enthalpy for Fe–B–X.
% ------------------------------------------------------------------
% REQUIREMENT:
%   Run build_Hmix_FB_X_ternary(...) first to generate an Excel with
%   'Pairs_Used' sheet (canonical A–B U0..U3).
%
% USAGE:
%   [Htot, parts] = Hmix3_calc('Si', 0.70, 0.20, 0.10);                 % default 'pair' mode
%   [Htot, parts] = Hmix3_calc('Cu', 0.60, 0.20, 0.20, 'C:\...\out.xlsx','pair');
%   [Htot, parts] = Hmix3_calc('Si', 0.70, 0.20, 0.10, [], 'global');   % global-c_i mode
%
% MODES:
%   'pair'  (default) ->  For each pair i–j, use pair-normalized variable:
%                         y_j = c_j / (c_i + c_j), ΔH_ij^bin = 4 y_j(1-y_j) Σ U_k (1-2y_j)^k,
%                         contribution to total = (c_i + c_j) * ΔH_ij^bin.
%   'global'          ->  Use full-composition form:
%                         ΔH_ij = 4 c_i c_j Σ U_k (c_i - c_j)^k, and sum over pairs.
%
% OUTPUTS:
%   Htot  - total mixing enthalpy
%   parts - details (final concentrations, each pair's U and ΔH, per-pair y and weights)
%
    % ---- guards & defaults ----
    if nargin < 4
        fprintf('USAGE:\n  [Htot, parts] = Hmix3_calc(''%s'', %.2f, %.2f, %.2f)\n', 'Si', 0.70, 0.20, 0.10);
        fprintf('Modes: ''pair'' (default) or ''global''. Optionally pass dbXlsx as 5th arg.\n');
        Htot=[]; parts=struct(); return;
    end
    if nargin < 5 || isempty(dbXlsx)
        dbXlsx = 'C:\Fe_BMAT\Fe_BM\Hmix_FB_X_ternary.xlsx';
    end
    if nargin < 6 || isempty(mode)
        mode = 'pair';
    end
    mode = lower(string(mode));

    % ---- read canonical U ----
    [Umap, Zmap] = read_pairs_used(dbXlsx);

    % ---- normalize inputs (snap 0.001) ----
    Xsym = normalizeSymbol(Xsym);
    if strcmpi(Xsym,'Fe') || strcmpi(Xsym,'B')
        if cX > 0, error('Xsym cannot be ''Fe'' or ''B'' when cX>0. For a binary limit set cX=0.'); end
    end
    v = [cFe, cB, cX];
    if all(v==0), error('All concentrations are zero.'); end
    v(v<0)=0; v=v/sum(v); v=round(v/0.001)*0.001; v=v/sum(v);
    cFe=v(1); cB=v(2); cX=v(3);

    % ---- fetch U(FeB), U(FeX), U(BX) in canonical order ----
    [U_FeB, ok1, A1, B1] = getU('Fe','B',Umap,Zmap);
    if ~ok1, error('Fe–B U parameters missing in Pairs_Used.'); end
    if cX > 0
        [U_FeX, ok2, A2, B2] = getU('Fe',Xsym,Umap,Zmap);
        [U_BX , ok3, A3, B3] = getU('B' ,Xsym,Umap,Zmap);
        if ~ok2, error('Fe–%s U parameters missing.', char(Xsym)); end
        if ~ok3, error('B–%s U parameters missing.' , char(Xsym)); end
    else
        U_FeX=[]; U_BX=[]; A2="";B2="";A3="";B3="";
    end

    % ---- compute per mode ----
    switch mode
        case "pair"
            % pair-normalized y_j = c_j / (c_i + c_j), weighted by (c_i + c_j)
            [cA1,cB1] = mapAB(A1,B1,cFe,cB,cX,Xsym);
            [H_FeB,y_FeB,w_FeB] = Hpair_pairMode(U_FeB,cA1,cB1);
            if cX>0
                [cA2,cB2] = mapAB(A2,B2,cFe,cB,cX,Xsym);   % (Fe,X)
                [cA3,cB3] = mapAB(A3,B3,cFe,cB,cX,Xsym);   % (B,X)
                [H_FeX,y_FeX,w_FeX] = Hpair_pairMode(U_FeX,cA2,cB2);
                [H_BX ,y_BX ,w_BX ] = Hpair_pairMode(U_BX ,cA3,cB3);
            else
                H_FeX=0; H_BX=0; y_FeX=NaN; y_BX=NaN; w_FeX=0; w_BX=0;
                cA2=NaN; cB2=NaN; cA3=NaN; cB3=NaN;
            end
            Htot = H_FeB + H_FeX + H_BX;

        case "global"
            % full-composition (global-c) form
            [cA1,cB1] = mapAB(A1,B1,cFe,cB,cX,Xsym);
            H_FeB = Hpair_global(U_FeB,cA1,cB1);
            if cX>0
                [cA2,cB2] = mapAB(A2,B2,cFe,cB,cX,Xsym);
                [cA3,cB3] = mapAB(A3,B3,cFe,cB,cX,Xsym);
                H_FeX = Hpair_global(U_FeX,cA2,cB2);
                H_BX  = Hpair_global(U_BX ,cA3,cB3);
            else
                H_FeX=0; H_BX=0;
                y_FeX=NaN; y_BX=NaN; w_FeX=0; w_BX=0;
                cA2=NaN; cB2=NaN; cA3=NaN; cB3=NaN;
            end
            Htot = H_FeB + H_FeX + H_BX;
            % define y as NaN in global mode
            [y_FeB,y_FeX,y_BX] = deal(NaN);

        otherwise
            error('Unknown mode: %s (use ''pair'' or ''global'')', char(mode));
    end

    % ---- pack details ----
    parts = struct();
    parts.mode = char(mode);
    parts.c = struct('Fe',cFe,'B',cB,'X',cX,'Xsym',Xsym);
    parts.pairs = struct();
    parts.pairs.FeB = struct('A',A1,'B',B1,'U',U_FeB,'cA',cA1,'cB',cB1,'yB',y_FeB,'weight',exist('w_FeB','var')*w_FeB + ~exist('w_FeB','var')*NaN,'H',H_FeB);
    if cX>0
        parts.pairs.FeX = struct('A',A2,'B',B2,'U',U_FeX,'cA',cA2,'cB',cB2,'yB',y_FeX,'weight',exist('w_FeX','var')*w_FeX + ~exist('w_FeX','var')*NaN,'H',H_FeX);
        parts.pairs.BX  = struct('A',A3,'B',B3,'U',U_BX ,'cA',cA3,'cB',cB3,'yB',y_BX ,'weight',exist('w_BX' ,'var')*w_BX  + ~exist('w_BX' ,'var')*NaN,'H',H_BX );
    else
        parts.pairs.FeX = struct('A','','B','','U',[],'cA',NaN,'cB',NaN,'yB',NaN,'weight',NaN,'H',0);
        parts.pairs.BX  = struct('A','','B','','U',[],'cA',NaN,'cB',NaN,'yB',NaN,'weight',NaN,'H',0);
    end
end % -- end main

% ======================= helpers =======================
function [Umap, Zmap] = read_pairs_used(dbXlsx)
    T = readtable(dbXlsx, 'Sheet', 'Pairs_Used', 'PreserveVariableNames', true);
    vnames = string(T.Properties.VariableNames);
    pairColIdx = find(contains(lower(vnames), 'pair'), 1, 'first');
    assert(~isempty(pairColIdx), 'Cannot find a "Pair" column in Pairs_Used.');
    pairs = T.(vnames(pairColIdx));
    if iscell(pairs), pairs = string(pairs); elseif ~isstring(pairs), pairs = string(pairs); end
    colU0 = find(contains(lower(vnames), 'u0') | contains(lower(vnames),'omega0'), 1, 'first');
    colU1 = find(contains(lower(vnames), 'u1') | contains(lower(vnames),'omega1'), 1, 'first');
    colU2 = find(contains(lower(vnames), 'u2') | contains(lower(vnames),'omega2'), 1, 'first');
    colU3 = find(contains(lower(vnames), 'u3') | contains(lower(vnames),'omega3'), 1, 'first');
    assert(~isempty(colU0) && ~isempty(colU1) && ~isempty(colU2) && ~isempty(colU3), 'Cannot find U0..U3 columns.');
    U0 = toNumCol(T.(vnames(colU0))); U1 = toNumCol(T.(vnames(colU1)));
    U2 = toNumCol(T.(vnames(colU2))); U3 = toNumCol(T.(vnames(colU3)));
    Umap = containers.Map('KeyType','char','ValueType','any');
    for i=1:numel(pairs)
        p = strtrim(pairs(i)); if strlength(p)==0, continue; end
        Ui = [U0(i) U1(i) U2(i) U3(i)];
        if any(isnan(Ui)), continue; end
        Umap(char(p)) = Ui;
    end
    Zmap = symbolToZMap();
end

function arr = toNumCol(col)
    if isnumeric(col), arr = double(col);
    else, arr = arrayfun(@(z) toNumSafe(z), col);
    end
end

function [Ucanon, ok, A, B] = getU(E1,E2,Umap,Zmap)
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
        otherwise  % X
            if char(A)==char(Xsym), cA = cX_; else, error('Unknown A symbol: %s', char(A)); end
    end
    switch char(B)
        case 'Fe', cB = cFe;
        case 'B',  cB = cB_;
        otherwise
            if char(B)==char(Xsym), cB = cX_; else, error('Unknown B symbol: %s', char(B)); end
    end
end

function [H,yB,w] = Hpair_pairMode(U,cA,cB)
    w = cA + cB;
    if w <= 0, H=0; yB=NaN; return; end
    yB = cB / w; t = 1 - 2*yB;
    Hbin = 4 .* yB .* (1 - yB) .* ( U(1) + U(2).*t + U(3).*t.^2 + U(4).*t.^3 );
    H = w .* Hbin;   % scale to per mole of total mixture
end

function H = Hpair_global(U,cA,cB)
    t = cA - cB;
    H = 4 .* cA .* cB .* ( U(1) + U(2).*t + U(3).*t.^2 + U(4).*t.^3 );
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

function s = normalizeSymbol(sin)
    if ismissing(sin), s=""; return; end
    s = string(sin);
    if strlength(s)==0, s=""; return; end
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
