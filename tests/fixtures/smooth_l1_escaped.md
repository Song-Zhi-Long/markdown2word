随后，使用 Smooth L1 Loss（平滑 L1 损失）计算预测回归量与真实回归量之间的差异：

\\[ \boxed{ L\_{\mathrm{reg}}= \sum\_{b\in\\{x,y,z,w,l,h,\theta\\}} \operatorname{SmoothL1}(\Delta b) } \\]

Smooth L1 可以理解为结合了 L1 和 L2 特性的损失函数：
