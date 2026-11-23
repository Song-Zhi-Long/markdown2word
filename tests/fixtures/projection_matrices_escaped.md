我们先不考虑 N 个点，只看其中一个三维点：

\\[ p_i= \begin{bmatrix} x_i\\\y_i\\\z_i \end{bmatrix} \\]

这里假设三维点已经处于相机坐标系下。

某个三维采样点为：

\\[ p_i= \begin{bmatrix} 2\\\1\\\10 \end{bmatrix} \\]

代入公式：

\\[ \begin{aligned} Kp_i &= \begin{bmatrix} 100&0&320\\\ 0&100&240\\\ 0&0&1 \end{bmatrix} \begin{bmatrix} 2\\\1\\\10 \end{bmatrix}\\\ &= \begin{bmatrix} 3400\\\2500\\\10 \end{bmatrix} \end{aligned} \\]

所以论文得到的投影结果是：
