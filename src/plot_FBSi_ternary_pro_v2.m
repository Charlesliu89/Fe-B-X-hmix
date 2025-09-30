
function plot_FBSi_ternary_pro_v2(varargin)
% plot_FBSi_ternary_pro_v2 — Fe–B–Si 三元混合焓视图（优化版）
% 美化可调 | 垂直外刻度（0.10/0.01）| 三边轴名中点外置（C_X）| 悬浮窗口下标 | 完整等值线掩膜
%
% 默认数据：
%   DataXlsx = 'C:\Fe_BMAT\Fe_BM\Fe-B-X.xlsx'
%   Sheet    = 'FBX_MATRIX_PAIR'   % 必含列：c_Fe | c_B | c_X | Hmix_Si
%
% 示例：
%   plot_FBSi_ternary_pro_v2
%   plot_FBSi_ternary_pro_v2('TitleFontSize',14,'TitleYOffset',0.08,'Levels',12,'CLim',[-30 5])
%   plot_FBSi_ternary_pro_v2('TickLabelGap',0.06,'EdgeNameOffset',[0.14 0.16 0.16])

%% ========= 可调参数（集中管理） =========
cfg.DataXlsx       = 'C:\Fe_BMAT\Fe_BM\Fe-B-X.xlsx';
cfg.Sheet          = 'FBX_MATRIX_PAIR';

% 散点/配色/色条
cfg.MarkerSize     = 12;
cfg.Colormap       = 'turbo';
cfg.CLim           = [];                % 固定色标范围，如 [-30 5]；空=[] 自动
cfg.CBLabel        = '\Delta H_{mix}';  % 色条标签（TeX）

% 等值线
cfg.ShowContour    = true;
cfg.Levels         = 10;
cfg.ContourColor   = 'k';

% 标题/字体
cfg.Title          = 'Fe–B–Si  \Delta H_{mix}'; % TeX
cfg.TitleFontSize  = 12;
cfg.TitleYOffset   = 0.08;
cfg.LabelFontSize  = 11;
cfg.FigPos         = [80 60 980 860];

% 三元网格（可关）
cfg.DrawGrid       = true;
cfg.GridMinorStep  = 0.01;
cfg.GridMajorStep  = 0.10;
cfg.GridColorMinor = [0 0 0];
cfg.GridColorMajor = [0 0 0];
cfg.GridLWMinor    = 0.7;
cfg.GridLWMajor    = 1.2;
cfg.GridAlphaMinor = 0.06;
cfg.GridAlphaMajor = 0.18;

% 刻度（垂直于边向外；0.10 主刻度带数值；0.01 次刻度不标数）
cfg.DrawTicks      = true;
cfg.TickMajorStep  = 0.10;
cfg.TickMinorStep  = 0.01;
cfg.TickLenMajor   = 0.035;  % 坐标单位（与三角高同量纲）
cfg.TickLenMinor   = 0.018;
cfg.TickLabelGap   = 0.05;   % 数值相对边的外移距离
cfg.TickFontSize   = 9;
cfg.TickLabelFormat= '%.2f';
cfg.TickAsPercent  = false;  % 若 true，以百分比显示（自动 *100）

% 三边轴名（位于边中点，外法线偏移；TeX 下标）
cfg.EdgeNames      = {'C_{B}','C_{Fe}','C_{Si}'};  % 底、左、右
cfg.EdgeNameOffset = [0.12, 0.14, 0.14];
cfg.EdgeNameFontSz = 12;
cfg.EdgeNameWeight = 'normal'; % 或 'bold'

% 顶点标签（纯文本）
cfg.VertexLabels   = {'Fe','B','Si'};
cfg.VertexOffsets  = [-0.03 -0.06; 1.02 -0.06; 0.49 sqrt(3)/2+0.03];

% 导出
cfg.SavePng        = '';      % 如 'C:\...\FBSi.png'；空=不导出
cfg.DPI            = 220;

%% ========= Name-Value 覆盖 =========
ip = inputParser;
fn = fieldnames(cfg);
for i=1:numel(fn), addParameter(ip,fn{i},cfg.(fn{i})); end
parse(ip,varargin{:});
cfg = ip.Results;

%% ========= 读数据 =========
T  = readtable(cfg.DataXlsx,'Sheet',cfg.Sheet,'PreserveVariableNames',true);
assert(all(ismember({'c_Fe','c_B','c_X'}, T.Properties.VariableNames)), ...
       '该表需包含列：c_Fe / c_B / c_X。');
vn = string(T.Properties.VariableNames);
colSi = vn(strcmpi(vn,'Hmix_Si'));
if isempty(colSi)
    cand = vn(startsWith(lower(vn),'hmix_si'));
    assert(~isempty(cand),'未找到 Hmix_Si 列（或同前缀列）。');
    colSi = cand(1);
end

cFe = T.c_Fe;  cB = T.c_B;  cSi = T.c_X;   % X 即 Si
H   = T.(colSi);

% 三元 -> 二维坐标（Fe, B, Si 顶点）
x2 = cB + 0.5*cSi;
y2 = (sqrt(3)/2)*cSi;

%% ========= 绘图 =========
fig = figure('Name','Fe–B–Si','Color','w','Position',cfg.FigPos);
ax  = axes(fig); hold(ax,'on'); axis(ax,'equal'); axis(ax,'off'); set(ax,'FontSize',cfg.LabelFontSize);

% 扩展坐标范围以显示外刻度/轴名
g = sqrt(3)/2;  xlim(ax, [-0.14, 1.14]);  ylim(ax, [-0.16, g + 0.12]);

% 三角边框
plot(ax,[0 1 0.5 0], [0 0 g 0], 'k-','LineWidth',1.25);

% 三元网格
if cfg.DrawGrid
    drawTernaryGrid(ax, cfg.GridMinorStep, cfg.GridLWMinor, cfg.GridColorMinor, cfg.GridAlphaMinor);
    drawTernaryGrid(ax, cfg.GridMajorStep, cfg.GridLWMajor, cfg.GridColorMajor, cfg.GridAlphaMajor);
end

% 散点热力（可下采样）
if cfg.MarkerSize>0
    ds  = 1;
    if isfield(cfg,'Downsample'), ds = max(1, round(cfg.Downsample)); end
    idx = 1:ds:numel(H);
    scatter(ax, x2(idx), y2(idx), cfg.MarkerSize, H(idx), 'filled');
end
colormap(ax, cfg.Colormap);
if ~isempty(cfg.CLim)
    if ~isempty(which('clim'))  % 兼容新版/旧版
        clim(ax, cfg.CLim);
    else
        set(ax,'CLim',cfg.CLim);
    end
end
cb = colorbar(ax); cb.Label.String = cfg.CBLabel; cb.Label.Interpreter='tex';

% 等值线
if cfg.ShowContour
    if exist('tricontour','file')==2
        tri = delaunay(x2,y2);
        tricontour(tri,x2,y2,H,cfg.Levels,cfg.ContourColor);
    else
        F = scatteredInterpolant(x2,y2,H,'natural','none');
        [gxv,gyv] = meshgrid(linspace(0,1,320), linspace(0, g, 280));
        ymax = sqrt(3) * min(gxv, 1 - gxv);        % 正确三角掩膜
        mask = (gxv >= 0) & (gxv <= 1) & (gyv >= 0) & (gyv <= ymax);
        Z = nan(size(gxv)); Z(mask) = F(gxv(mask), gyv(mask));
        contour(ax,gxv,gyv,Z,cfg.Levels,cfg.ContourColor);
    end
end

% 刻度（垂直于边、向外）
if cfg.DrawTicks
    drawEdgeTicksPerpOutside(ax, cfg);
end

% 三边轴名（中点外置，TeX 下标）
drawEdgeAxisNames(ax, cfg);

% 顶点标签（纯文本）
text(ax, cfg.VertexOffsets(1,1), cfg.VertexOffsets(1,2), cfg.VertexLabels{1}, 'FontSize',cfg.LabelFontSize, 'Interpreter','none');
text(ax, cfg.VertexOffsets(2,1), cfg.VertexOffsets(2,2), cfg.VertexLabels{2}, 'FontSize',cfg.LabelFontSize, 'Interpreter','none');
text(ax, cfg.VertexOffsets(3,1), cfg.VertexOffsets(3,2), cfg.VertexLabels{3}, 'FontSize',cfg.LabelFontSize, 'Interpreter','none');

% 标题（TeX）
ttl = title(ax, cfg.Title, 'Interpreter','tex','FontSize',cfg.TitleFontSize);
try ttl.Position(2) = ttl.Position(2) + cfg.TitleYOffset; end %#ok<TRYNC>

% 悬浮读数（仅悬浮窗口使用 TeX 下标）
setupDataTips_TeX(ax, cFe, cB, cSi, H);

% 导出（带回退；warning 使用标识符）
if strlength(string(cfg.SavePng))>0
    outdir = char(fileparts(string(cfg.SavePng)));
    if ~isempty(outdir) && ~exist(outdir,'dir'), mkdir(outdir); end
    try
        exportgraphics(fig, cfg.SavePng, 'Resolution', cfg.DPI);
    catch ME
        id = ME.identifier; if isempty(id), id = 'plot_FBSi_ternary:exportgraphics'; end
        warning(id, 'exportgraphics 失败：%s；将尝试使用 print 回退。', ME.message);
        try
            print(fig, cfg.SavePng, '-dpng', sprintf('-r%d', cfg.DPI));
        catch ME2
            id2 = ME2.identifier; if isempty(id2), id2 = 'plot_FBSi_ternary:print'; end
            warning(id2, 'print 也失败：%s；未保存。', ME2.message);
        end
    end
end
end % ===== 主函数结束 =====


%% ======= 工具：三元网格 =======
function drawTernaryGrid(ax, step, lw, rgb, alphaMix)
if step<=0, return; end
if nargin<5, alphaMix = 0.1; end
g = sqrt(3)/2;
col = (1-alphaMix)*rgb + alphaMix*[1 1 1];
for t=0:step:1
    % c_Si = t （水平）
    y = g*t; x1 = 0.5*t; x2 = 1-0.5*t; line(ax,[x1 x2],[y y],'Color',col,'LineWidth',lw);
    % c_B = t  （60°）
    x1=t; y1=0; x2=0.5+0.5*t; y2=g*(1-t); line(ax,[x1 x2],[y1 y2],'Color',col,'LineWidth',lw);
    % c_Fe = t （-60°）
    x1=1-t; y1=0; x2=0.5*(1-t); y2=g*(1-t); line(ax,[x1 x2],[y1 y2],'Color',col,'LineWidth',lw);
end
end

%% ======= 工具：刻度（外法线、主/次） =======
function drawEdgeTicksPerpOutside(ax, cfg)
g = sqrt(3)/2;
n_base  = [0, -1];      % base outward
n_left  = [-g, 0.5];    % left outward (unit)
n_right = [ g, 0.5];    % right outward (unit)

% 百分比显示开关
asPct = isfield(cfg,'TickAsPercent') && cfg.TickAsPercent;

% Base (Fe–B) : P=(t,0)
if cfg.TickMinorStep>0
    t = 0:cfg.TickMinorStep:1;
    for k=2:numel(t)-1
        P=[t(k),0]; P2=P+cfg.TickLenMinor*n_base;
        line(ax,[P(1) P2(1)],[P(2) P2(2)],'Color','k','LineWidth',0.6);
    end
end
t = 0:cfg.TickMajorStep:1;
for k=2:numel(t)-1
    P=[t(k),0]; P2=P+cfg.TickLenMajor*n_base;
    line(ax,[P(1) P2(1)],[P(2) P2(2)],'Color','k','LineWidth',1.0);
    lbl=P+(cfg.TickLenMajor+cfg.TickLabelGap)*n_base;
    val = t(k);
    if asPct, val = val*100; end
    text(ax,lbl(1),lbl(2),sprintf(cfg.TickLabelFormat,val), ...
        'HorizontalAlignment','center','VerticalAlignment','top', ...
        'FontSize',cfg.TickFontSize,'Interpreter','none');
end

% Left (Fe–Si) : P=(0.5 s, g s)
if cfg.TickMinorStep>0
    s=0:cfg.TickMinorStep:1;
    for k=2:numel(s)-1
        P=[0.5*s(k), g*s(k)]; P2=P+cfg.TickLenMinor*n_left;
        line(ax,[P(1) P2(1)],[P(2) P2(2)],'Color','k','LineWidth',0.6);
    end
end
s=0:cfg.TickMajorStep:1;
for k=2:numel(s)-1
    P=[0.5*s(k), g*s(k)]; P2=P+cfg.TickLenMajor*n_left;
    line(ax,[P(1) P2(1)],[P(2) P2(2)],'Color','k','LineWidth',1.0);
    val = 1 - s(k); % c_Fe
    if asPct, val = val*100; end
    lbl=P+(cfg.TickLenMajor+cfg.TickLabelGap)*n_left;
    text(ax,lbl(1),lbl(2),sprintf(cfg.TickLabelFormat,val), ...
        'HorizontalAlignment','right','VerticalAlignment','middle', ...
        'FontSize',cfg.TickFontSize,'Interpreter','none');
end

% Right (B–Si) : Q=(1-0.5 s, g s)
if cfg.TickMinorStep>0
    s=0:cfg.TickMinorStep:1;
    for k=2:numel(s)-1
        Q=[1-0.5*s(k), g*s(k)]; Q2=Q+cfg.TickLenMinor*n_right;
        line(ax,[Q(1) Q2(1)],[Q(2) Q2(2)],'Color','k','LineWidth',0.6);
    end
end
s=0:cfg.TickMajorStep:1;
for k=2:numel(s)-1
    Q=[1-0.5*s(k), g*s(k)]; Q2=Q+cfg.TickLenMajor*n_right;
    line(ax,[Q(1) Q2(1)],[Q(2) Q2(2)],'Color','k','LineWidth',1.0);
    val = s(k); % c_Si
    if asPct, val = val*100; end
    lbl=Q+(cfg.TickLenMajor+cfg.TickLabelGap)*n_right;
    text(ax,lbl(1),lbl(2),sprintf(cfg.TickLabelFormat,val), ...
        'HorizontalAlignment','left','VerticalAlignment','middle', ...
        'FontSize',cfg.TickFontSize,'Interpreter','none');
end
end

%% ======= 工具：三边轴名（中点外置，TeX 下标） =======
function drawEdgeAxisNames(ax, cfg)
g = sqrt(3)/2;
P_base=[0.5,0];      n_base=[0,-1];
P_left=[0.25,g*0.5]; n_left=[-g,0.5];
P_right=[0.75,g*0.5];n_right=[ g,0.5];

% 至少在大刻度值之外
autoOff = max([cfg.EdgeNameOffset(:)'; repmat(cfg.TickLenMajor+cfg.TickLabelGap+0.02,1,3)],[],1);
Pb = P_base + autoOff(1)*n_base;
Pl = P_left + autoOff(2)*n_left;
Pr = P_right+ autoOff(3)*n_right;

text(ax,Pb(1),Pb(2),cfg.EdgeNames{1},'HorizontalAlignment','center','VerticalAlignment','top',...
    'FontSize',cfg.EdgeNameFontSz,'Interpreter','tex','FontWeight',cfg.EdgeNameWeight);
text(ax,Pl(1),Pl(2),cfg.EdgeNames{2},'HorizontalAlignment','right','VerticalAlignment','middle',...
    'FontSize',cfg.EdgeNameFontSz,'Interpreter','tex','FontWeight',cfg.EdgeNameWeight);
text(ax,Pr(1),Pr(2),cfg.EdgeNames{3},'HorizontalAlignment','left','VerticalAlignment','middle',...
    'FontSize',cfg.EdgeNameFontSz,'Interpreter','tex','FontWeight',cfg.EdgeNameWeight);
end

%% ======= 工具：悬浮读数（TeX 下标） =======
function setupDataTips_TeX(ax, cFe, cB, cSi, H)
sc = findobj(ax,'Type','Scatter');
if isempty(sc), return; end
sc = sc(1);
if isprop(sc,'DataTipTemplate')    % 新版
    r1 = matlab.graphics.datatip.DataTipTextRow('c_{Fe}',        cFe); r1.Format='%.3f';
    r2 = matlab.graphics.datatip.DataTipTextRow('c_{B}',         cB ); r2.Format='%.3f';
    r3 = matlab.graphics.datatip.DataTipTextRow('c_{Si}',        cSi); r3.Format='%.3f';
    r4 = matlab.graphics.datatip.DataTipTextRow('\DeltaH_{mix}',  H ); r4.Format='%.4f';
    sc.DataTipTemplate.DataTipRows = [r1; r2; r3; r4];
    sc.DataTipTemplate.Interpreter = 'tex';
else                                % 旧版：回调
    dcm = datacursormode(ancestor(sc,'figure'));
    set(dcm,'UpdateFcn',@(obj,evt)localTip(evt,cFe,cB,cSi,H));
end
end

function out = localTip(evt,cFe,cB,cSi,H) 
i = evt.DataIndex;
out = {sprintf('c_Fe: %.3f',cFe(i)), ...
       sprintf('c_B : %.3f',cB(i)),  ...
       sprintf('c_Si: %.3f',cSi(i)), ...
       sprintf('ΔH_mix: %.4f',H(i))};
end
