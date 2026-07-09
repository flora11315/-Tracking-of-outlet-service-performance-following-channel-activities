# 渠道活动服务追踪 Agent Demo 工作台

## 本地演示流程

1. 启动本地工作台：

   ```bash
   python3 outputs/channel_activity_demo/demo_workbench.py --host 127.0.0.1 --port 8001
   ```

2. 打开浏览器：

   ```text
   http://127.0.0.1:8001/
   ```

3. 汇报时按页面顺序讲：

   - 输入数据：活动表、经营追踪表、可选问卷反馈表。
   - 点击“运行 Skill 分析”：展示从读取数据到生成提醒、邮件和附件的全流程。
   - 看板结论：覆盖活动、节点提醒、高优先级、优秀案例、高频材料。
   - 企业微信提醒模拟：说明 T+7/T+30/T+90 节点如何即时提醒渠道销售。
   - 邮件正文预览：说明全部结果定期进入邮件。
   - Excel 附件：下载完整明细，用于真实工作中的复盘和留痕。

## Render 公网部署

当前仓库已补充 `render.yaml` 和 `requirements.txt`，可以通过 Render Web Service 部署。

推荐步骤：

1. 将整个项目推到 GitHub 仓库。
2. 在 Render 新建 Web Service，选择该 GitHub 仓库。
3. Render 会自动读取根目录的 `render.yaml`。
4. 部署后打开 Render 分配的公网 URL。

Render 配置逻辑：

- Build Command：`pip install -r outputs/channel_activity_demo/requirements.txt`
- Start Command：`python outputs/channel_activity_demo/demo_workbench.py --host 0.0.0.0 --port $PORT`

注意：当前 demo 使用模拟数据，适合公网展示。后续如果接入真实渠道数据，应改为内网部署或加登录鉴权。
