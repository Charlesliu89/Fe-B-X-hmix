function H = calc_hmix(Fe, B, X)
%CALC_HMIX 示例函数：计算混合焓（虚拟公式）
%   H = calc_hmix(Fe, B, X) 返回一个虚拟的混合焓
%
%   Fe, B, X = 原子分数，总和=1

if abs(Fe + B + X - 1) > 1e-6
    error('原子分数总和必须为1');
end

% 这里只是举例：真实公式以后再替换
H = Fe*B*(-20) + Fe*X*(-5) + B*X*(-10);
end
