# **RARE：面向单目三维目标检测的学习排序与检索**

# **RARE: Learn to RAnk and REtrieve for Monocular 3D Object Detection**

\<div align="center">\
\
朴惠贞（Hyeonjeong Park，韩国）¹、熊佩熙（Peixi Xiong，国家/地区待核实）²、阮晓倩（Xiaoqian Ruan，中国）¹、贾典（Dian Jia，国家/地区待核实）¹、于佩（Pei Yu，中国）³、唐伟（Wei Tang，中国）¹\
\
¹ 芝加哥伊利诺伊大学（University of Illinois Chicago）　² 英特尔（Intel）　³ 微软（Microsoft）\
\
{hpark233,xruan9,djia7,tangw}@uic.edu, peixi.xiong\@intel.com, pei.yu\@microsoft.com\
\
\</div>

\<div align="left">\
\
\*\*研究方向：\*\* 单目三维目标检测（Monocular 3D Object Detection，Mono3D）、基于排序的置信度学习（Learning to Rank）、多假设三维定位（Multi-hypothesis 3D Localization）。\
\
\</div>

## **摘要**

仅从单张 RGB 图像进行单目三维目标检测仍然面临挑战，这源于两个根本性难题：三维定位本身具有不适定性，即同一个二维观测可能对应多个合理的配置；以及置信度估计不可靠，无法反映真实的定位精度。现有方法预测确定性的三维边界框，这些边界框往往会坍缩为不合理的均值估计，并依赖对定位误差高度敏感的绝对置信度分数。本文提出 RARE，一种通过学习排序与检索来同时解决这两个难题的统一框架。RARE 将置信度估计建模为一个排序问题，学习根据检测结果的相对质量对其进行排序，而不是回归绝对数值。它能够提供更加鲁棒、稳定且对定位不确定性不那么敏感的置信度估计。在此基础上，RARE 学习为每个目标构建一个查询集合，以预测多个具有多样性且合理的三维配置，并检索其中排名最高的预测结果。它显式地对单目三维感知的多模态特性进行建模，从而产生更加合理的定位结果。大量实验验证了 RARE 的有效性。代码已公开于：[https://github.com/HyeonjeongPark37/RARE。](https://github.com/HyeonjeongPark37/RARE%E3%80%82?utm_source=chatgpt.com)
