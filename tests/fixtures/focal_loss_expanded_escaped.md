因此，CaDDN 的深度损失完整展开后可以写成：

\\[ \boxed{ L\_{\mathrm{depth}}= \frac{1}{W_FH_F} \sum\_{u=1}^{W_F} \sum\_{v=1}^{H_F} -\alpha\_{u,v} (1-p_t(u,v))^2 \log p_t(u,v) } \\]

其中 \\(\alpha\_{u,v}\\) 根据当前像素属于前景还是背景，分别取 3.25 或 0.25。
