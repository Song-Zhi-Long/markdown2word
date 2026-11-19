具体而言，我们采用 Focal Loss（焦点损失）[30]：

\\[ L\_{\mathrm{depth}}= \frac{1}{W_F\cdot H_F} \sum\_{u=1}^{W_F}\sum\_{v=1}^{H_F} FL\big(D(u,v),\hat{D}(u,v)\big) \tag{3} \\]

其中，\\(D\\) 表示预测得到的深度分布
