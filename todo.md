代办事项

1. outlier label的内容优化成"n=*"这样的格式，以免选择yAxisDensity模式的时候产生歧义
2. OUTLIER_LABEL_RATIO改为plot option的可选参数，默认值0.001即千分之一
3. 完成1和2后，提交到当前分支。然后新增一个分支"qoder\implement_multi-histogram"并checkout
4. 在新分支构思并实现分组的正态分布图，有两种形式，一种是接受一个二维数组或2个等长的一维数组，其中一列是数值，另一列是分组信息。按分组信息分割原始数据，然后对每个子组分别生成独立的直方图，或在同一个画布中叠加各子组的直方图（shade透明度按50%预设）