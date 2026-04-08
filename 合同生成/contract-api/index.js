const express = require('express');
const cors = require('cors');
const path = require('path');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// 中间件
app.use(cors());                // 允许跨域（如果前端用不同端口访问）
app.use(express.json());        // 解析 JSON 请求体
app.use(express.urlencoded({ extended: true }));

// 托管前端静态文件（假设前端文件放在项目根目录的 public 文件夹）
app.use(express.static(path.join(__dirname, '../public')));

// 可选：托管模板文件目录，以便直接访问（如果不通过下载接口）
app.use('/templates', express.static(path.join(__dirname, '../templates')));

// 挂载 API 路由
const templatesRouter = require('./routes/templates');
app.use('/api/templates', templatesRouter);

const purchaseOrdersRouter = require('./routes/purchase-orders');
app.use('/api/purchase-orders', purchaseOrdersRouter);

const contractsRouter = require('./routes/contracts');
app.use('/api/contracts', contractsRouter);

// 对于所有其他未匹配的请求，返回 index.html（支持前端路由）
app.get(/.*/, (req, res) => {
    res.sendFile(path.join(__dirname, '../public/index.html'));
});

// 启动服务器
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
